#!/usr/bin/env python3
"""
fetch_sources.py - שלבים 2 ו-3.

לוקח קובץ refs.json (רשימת ההפניות), מביא את הטקסט האמיתי מספריא,
ובונה קובץ דוקס אחד עם כל המקורות לפי הסדר של הדף.
מקור שכבר מופיע בדף המקורות המקורי נלקח מהדף עצמו (צריך לתת --docx).
מקור שלא נמצא מסומן בדוקס, ולא מומצא.

שימוש:
    python fetch_sources.py refs.json
    python fetch_sources.py refs.json --docx "הדף המקורי.docx"

דרישות: pip install python-docx requests
"""
import argparse
import html
import json
import re
import sys
import time
from pathlib import Path
from urllib.parse import quote

import requests
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.table import Table
from docx.text.paragraph import Paragraph

API = "https://www.sefaria.org/api"
HEADERS = {"User-Agent": "mekorot-helper/1.0"}
FONT = "David"


# ---------------------------------------------------------------- ספריא ----

def flatten(x):
    """הטקסט בספריא מגיע כרשימות מקוננות. הופך הכול לרשימה שטוחה של מקטעים."""
    if x is None:
        return []
    if isinstance(x, str):
        return [x]
    out = []
    for item in x:
        out.extend(flatten(item))
    return out


FOOTNOTE = re.compile(
    r'<sup[^>]*class="footnote-marker"[^>]*>.*?</sup>\s*<i[^>]*class="footnote"[^>]*>.*?</i>', re.S)


def clean(s):
    s = FOOTNOTE.sub("", s)
    s = re.sub(r"<br\s*/?>", " ", s)
    s = re.sub(r"<[^>]+>", "", s)
    s = html.unescape(s)
    return re.sub(r"\s+", " ", s).strip()


def norm(s):
    """צורה מנורמלת להשוואה: בלי ניקוד, בלי פיסוק, בלי גרשיים."""
    s = re.sub(r"[\u0591-\u05C7]", "", s)
    s = re.sub(r"[^\u05D0-\u05EA ]", " ", s)
    s = re.sub(r"(?<![\u05D0-\u05EA])רבי(?![\u05D0-\u05EA])", "ר", s)  # רבי = ר'
    return re.sub(r"\s+", " ", s).strip()


def skel(s):
    """צורה מנורמלת עוד יותר: בלי ו' וי', כדי שכתיב מלא/חסר (תשמור/תשמר) לא ימנע התאמה."""
    return re.sub(r"\s+", " ", re.sub(r"[וי]", "", norm(s))).strip()


def variants(words):
    """הביטוי המלא, ואחר כך גרסאות מקוצרות שלו (בלי המילים הראשונות, ואז בלי האחרונות).
    כך 'רבה בר חנן' מתאים גם לטקסט 'רבה בר בר חנן'."""
    w = skel(words).split()
    out = [" ".join(w)]
    for i in range(1, len(w) - 1):
        out.append(" ".join(w[i:]))
    for j in range(len(w) - 1, 1, -1):
        out.append(" ".join(w[:j]))
    return [v for v in out if v]


def fetch_ref(ref):
    url = f"{API}/texts/{quote(ref, safe=':,-.')}"
    r = requests.get(url, params={"context": 0, "commentary": 0}, headers=HEADERS, timeout=30)
    r.raise_for_status()
    data = r.json()
    if "error" in data:
        raise ValueError(data["error"])
    segs = [clean(t) for t in flatten(data.get("he"))]
    segs = [t for t in segs if t]
    if not segs:
        raise ValueError("no Hebrew text returned")
    return data.get("ref", ref), segs


def resolve_name(query):
    """שואל את ספריא מה ההפניה הנכונה לטקסט חופשי."""
    r = requests.get(f"{API}/name/{quote(query)}", headers=HEADERS, timeout=30)
    r.raise_for_status()
    d = r.json()
    if d.get("is_ref") and d.get("ref"):
        return d["ref"]
    return None


def get_segments(src):
    """מנסה: ההפניה כמו שהיא, אחר כך חיפוש בספריא לפי ההפניה, ואחר כך לפי השם בעברית."""
    attempts = []
    if src.get("sefaria_ref"):
        attempts.append(("direct", src["sefaria_ref"]))
        attempts.append(("name", src["sefaria_ref"]))
    attempts.append(("name", f"{src['work_he']} {src['location_he']}".strip()))
    last_err = "not found"
    for kind, q in attempts:
        try:
            ref = q if kind == "direct" else resolve_name(q)
            if not ref:
                last_err = f"Sefaria did not recognize: {q}"
                continue
            return fetch_ref(ref)
        except Exception as e:  # noqa: BLE001
            last_err = str(e)
        time.sleep(0.3)
    raise RuntimeError(last_err)


# -------------------------------------------------------- בחירת קטע מדויק ----

def pick_dibbur(segs, dh):
    d = skel(dh)
    # התאמה לפי מילים שלמות: ד"ה "את" לא יתאים ל"אתא"
    start = re.compile(re.escape(d) + r"(\s|$)")
    near = re.compile(r"(^|\s)" + re.escape(d) + r"(\s|$)")
    hits = [s for s in segs if start.match(skel(s))]
    if not hits:
        hits = [s for s in segs if near.search(skel(s)[: len(d) + 40])]
    return hits


def pick_range(segs, start, end, is_commentary):
    """חיתוך לפי 'מהמילים ... עד ...'. החיתוך הוא ברמת מקטע, לא ברמת מילה."""
    def find(words, begin):
        for probe in variants(words):
            for i in range(begin, len(segs)):
                if probe in skel(segs[i]):
                    return i
        return None

    i0 = 0
    if start:
        i0 = find(start, 0)
        if i0 is None:
            if is_commentary:  # בפירוש לא מביאים את כל הדף סתם
                return [], "המילים 'מהמילים' לא נמצאו בפירוש"
            return segs, "המילים 'מהמילים' לא נמצאו, הובא הטקסט כולו"
    if end:
        i1 = find(end, i0)
        if i1 is None:
            if is_commentary:
                return segs[i0:i0 + 1], "המילים 'עד' לא נמצאו, הובאה ההערה הראשונה בלבד"
            return segs[i0:], "המילים 'עד' לא נמצאו, הובא עד סוף הטקסט"
        return segs[i0:i1 + 1], None
    if is_commentary:
        return segs[i0:i0 + 1], None  # בפירוש, מקטע אחד הוא הערה אחת
    return segs[i0:], None


# רש"י ותוספות בנויים בספריא לפי שורות בגמרא. הפניה לדף שלם מחזירה רק את השורה הראשונה,
# לכן עוברים שורה אחרי שורה עד שנגמרות ההערות.
LINE_LEVEL = re.compile(r"^(Rashi|Tosafot) on .+ \d+[ab]$")
_daf_cache = {}


def next_amud(ref):
    m = re.match(r"^(.+ )(\d+)([ab])$", ref)
    if not m:
        return None
    base, num, side = m.groups()
    return f"{base}{num}b" if side == "a" else f"{base}{int(num) + 1}a"


def book(ref):
    """שם הספר בלבד (בלי הספרות ואותיות הדף), להשוואה בין הפניות."""
    return re.sub(r"[^A-Za-z]", "", re.split(r"\d", ref, maxsplit=1)[0]).lower()


def first_words(segs, k=15):
    return " | ".join(" ".join(s.split()[:3]) for s in segs[:k])


def fetch_daf_commentary(ref):
    if ref in _daf_cache:
        return ref, _daf_cache[ref]
    out, misses, last_err = [], 0, "no commentary found"
    for n in range(1, 121):
        try:
            _, segs = fetch_ref(f"{ref}:{n}")
            out.extend(segs)
            misses = 0
        except Exception as e:  # noqa: BLE001
            last_err = str(e)
            misses += 1
            if (out and misses >= 8) or misses >= 12:
                break
        time.sleep(0.15)
    if not out:
        raise RuntimeError(last_err)
    _daf_cache[ref] = out
    return ref, out


# ------------------------------------------------------ מקור מתוך הדף עצמו ----

def read_lines(path):
    doc = Document(path)
    lines = []
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            p = Paragraph(child, doc)
            text = p.text.strip()
            if not text:
                continue
            runs = [r for r in p.runs if r.text.strip()]
            lines.append((text, bool(runs) and all(r.bold for r in runs)))
        elif child.tag.endswith("}tbl"):
            for row in Table(child, doc).rows:
                cells = [c.text.strip() for c in row.cells if c.text.strip()]
                if cells:
                    lines.append((" | ".join(cells), False))
    return lines


def text_from_doc(doc_lines, heading):
    """מוצא כותרת מודגשת בדף, ומחזיר את הפסקאות שמתחתיה עד הכותרת המודגשת הבאה."""
    h = norm(heading)
    start = None
    for i, (t, bold) in enumerate(doc_lines):
        n = norm(t)
        if bold and (n == h or (len(h) >= 6 and (h in n or n in h) and len(n) >= 6)):
            start = i
            break
    if start is None:
        return None
    body = []
    for t, bold in doc_lines[start + 1:]:
        if bold:
            break
        body.append(t)
    return body or None


# ------------------------------------------------------------ קישורים לספריא ----

def ref_url(ref):
    """הופך הפניה כמו 'Bava Batra 4b' לכתובת של ספריא."""
    ref = ref.strip()
    m = re.match(r"^(.+?) (\d[^ ]*)$", ref)
    if m:
        title, sec = m.groups()
        path = title.replace(" ", "_") + "." + sec.replace(":", ".")
    else:
        path = ref.replace(" ", "_")
    return "https://www.sefaria.org/" + quote(path, safe="._,-'")


def search_url(query):
    return "https://www.sefaria.org/search?q=" + quote(query)


# ------------------------------------------------------------- בניית הדוקס ----

def style_run(r, size, bold):
    r.font.name = FONT
    r.font.size = Pt(size)
    r.bold = bold
    rPr = r._r.get_or_add_rPr()
    rPr.rFonts.set(qn("w:cs"), FONT)
    szcs = OxmlElement("w:szCs")
    szcs.set(qn("w:val"), str(size * 2))
    rPr.find(qn("w:sz")).addnext(szcs)
    if bold:
        rPr.find(qn("w:b")).addnext(OxmlElement("w:bCs"))
    rPr.append(OxmlElement("w:rtl"))


def add_par(doc, text, size=12, bold=False):
    p = doc.add_paragraph()
    p._p.get_or_add_pPr().append(OxmlElement("w:bidi"))  # כיוון מימין לשמאל
    p.paragraph_format.space_after = Pt(6)
    style_run(p.add_run(text), size, bold)
    return p


def add_link_par(doc, text, url, size=10):
    p = doc.add_paragraph()
    p._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    r_id = doc.part.relate_to(url, RT.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), FONT)
    fonts.set(qn("w:hAnsi"), FONT)
    fonts.set(qn("w:cs"), FONT)
    rPr.append(fonts)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    rPr.append(color)
    for tag in ("w:sz", "w:szCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), str(size * 2))
        rPr.append(el)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rPr.append(underline)
    rPr.append(OxmlElement("w:rtl"))
    run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)
    link.append(run)
    p._p.append(link)


def build_docx(results, out_path):
    doc = Document()
    failed = [r for r in results if not r["ok"]]
    add_par(doc, "דף מקורות מלא", 20, True)
    summary = f"הובאו {len(results) - len(failed)} מתוך {len(results)} מקורות."
    if failed:
        summary += " לא נמצאו: " + ", ".join(str(r["n"]) for r in failed) + "."
    add_par(doc, summary, 11)

    last_section = object()
    for r in results:
        s = r["src"]
        if s.get("section") != last_section:
            last_section = s.get("section")
            if last_section:
                add_par(doc, last_section, 16, True)
        add_par(doc, f"{r['n']}. {s['work_he']} - {s['location_he']}", 14, True)
        if r.get("link"):
            add_link_par(doc, r["link_text"], r["link"])
        if r["ok"]:
            for chunk in r["chunks"]:
                add_par(doc, chunk, 12)
            add_par(doc, "מקור הטקסט: " + r["origin"], 9)
        else:
            add_par(doc, "[לא נמצא. יש להוסיף ידנית] " + r["error"], 11)
        for note in r["notes"]:
            add_par(doc, "הערה: " + note, 9)
    doc.save(out_path)


# ------------------------------------------------------------------- ראשי ----

def _process(n, src, doc_lines):
    res = {"n": n, "src": src, "ok": False, "chunks": [], "origin": "", "notes": [], "error": ""}
    if src.get("note"):
        res["notes"].append(src["note"])

    # מקור שכבר מופיע בדף עצמו
    if src.get("already_in_doc") and doc_lines and src.get("doc_heading"):
        body = text_from_doc(doc_lines, src["doc_heading"])
        if body:
            res.update(ok=True, chunks=body, origin="הדף המקורי")
            return res
        res["notes"].append("לא נמצא בדף המקורי תחת הכותרת, נבדק בספריא")

    req = (src.get("sefaria_ref") or "").strip()
    try:
        if LINE_LEVEL.match(req):
            ref, segs = fetch_daf_commentary(req)
        else:
            ref, segs = get_segments(src)
            if req and book(ref) != book(req):
                res["notes"].append(f"ספריא תיקנה את ההפניה ל-{ref}. כדאי לוודא שזה המקור הנכון")
    except Exception as e:  # noqa: BLE001
        res["error"] = str(e)
        return res

    res["ref"] = ref
    is_commentary = " on " in ref
    if src.get("dibbur_hamatchil"):
        hits = pick_dibbur(segs, src["dibbur_hamatchil"])
        if not hits and LINE_LEVEL.match(req) and next_amud(req):
            nxt = next_amud(req)
            try:
                _, segs2 = fetch_daf_commentary(nxt)
                hits = pick_dibbur(segs2, src["dibbur_hamatchil"])
                if hits:
                    res["notes"].append(f"הד\"ה לא נמצא ב-{req}, ונמצא ב-{nxt}")
            except Exception:  # noqa: BLE001
                pass
        if not hits:
            res["error"] = (f"ד\"ה '{src['dibbur_hamatchil']}' לא נמצא בטקסט של {ref}. "
                            f"ההערות שכן קיימות שם מתחילות ב: {first_words(segs)}")
            return res
        if len(hits) > 1:
            res["notes"].append(f"נמצאו {len(hits)} הערות שמתחילות בד\"ה הזה, הובאו כולן")
        chunks = hits
    elif src.get("range_start_words") or src.get("range_end_words") or is_commentary:
        chunks, warn = pick_range(segs, src.get("range_start_words"), src.get("range_end_words"), is_commentary)
        if warn:
            res["notes"].append(warn)
        if not (src.get("range_start_words") or src.get("range_end_words")):
            chunks = segs  # פירוש בלי תחום מוגדר: מביאים את כל מה שיש במקום
        if not chunks:
            res["error"] = warn or "לא נמצא קטע מתאים"
            return res
    else:
        chunks = segs

    # מקטעים קצרים (גמרא) מחוברים לפסקה אחת, פירושים נשארים כל אחד בפסקה
    if not is_commentary and not src.get("dibbur_hamatchil"):
        chunks = [" ".join(chunks)]
    res.update(ok=True, chunks=chunks, origin=f"ספריא ({ref})")
    return res


def process(n, src, doc_lines):
    res = _process(n, src, doc_lines)
    ref = res.get("ref")
    if ref:
        res["link"] = ref_url(ref)
        res["link_text"] = "פתח בספריא" if res["ok"] else "פתח בספריא לחיפוש ידני"
    elif res["ok"] and src.get("sefaria_ref"):
        res["link"] = ref_url(src["sefaria_ref"])
        res["link_text"] = "פתח בספריא (ההפניה לא אומתה)"
    elif not res["ok"]:
        res["link"] = search_url(f"{src['work_he']} {src['location_he']}".strip())
        res["link_text"] = "חיפוש בספריא"
    return res


def main():
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("refs", help="refs.json")
    ap.add_argument("--docx", help="the original sheet (for sources that are already inside it)")
    ap.add_argument("--out", help="output file name (default: next to refs.json)")
    args = ap.parse_args()

    refs_path = Path(args.refs)
    sources = json.loads(refs_path.read_text(encoding="utf-8"))

    in_sheet = sum(1 for s in sources if s.get("already_in_doc"))
    if in_sheet and not args.docx:
        sys.exit(f"STOP: {in_sheet} sources in this list are already inside the original sheet, "
                 "and they must be taken from it (not from Sefaria).\n"
                 "Run again and add:  --docx \"<path to the original sheet>\"")
    if args.docx and not Path(args.docx).exists():
        sys.exit(f"STOP: file not found: {args.docx}")

    doc_lines = read_lines(args.docx) if args.docx else None
    out_path = Path(args.out) if args.out else refs_path.with_suffix(".docx")

    results = []
    for n, src in enumerate(sources, start=1):
        print(f"[{n}/{len(sources)}] working...", flush=True)
        res = process(n, src, doc_lines)
        results.append(res)
        print(f"[{n}/{len(sources)}] {'OK (' + ('sheet' if res['origin'] == 'הדף המקורי' else 'Sefaria') + ')' if res['ok'] else 'NOT FOUND'}")
        time.sleep(0.3)

    build_docx(results, out_path)
    failed = [str(r["n"]) for r in results if not r["ok"]]
    print(f"Done. {len(results) - len(failed)}/{len(results)} sources found.")
    if failed:
        print("Not found: " + ", ".join(failed))
    print(f"Open this file: {out_path}")


if __name__ == "__main__":
    main()
