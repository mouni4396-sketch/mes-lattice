"""
readers/pdf_reader.py  —  Agent 6 (Intake) stage 1, PDF format reader.

DETERMINISTIC. No AI. Emits the SAME concept-record shape as html_reader.py and
docx_reader.py, so PDFs flow through the identical mapper -> draft_writer chain.

Uses pdfplumber (pure Python, Render-safe) for text + tables. No system tools.

IMPORTANT CAVEAT: PDFs have NO heading styles (unlike HTML <h2> or Word
'Heading 2'). So "sections" here are INFERRED with a heuristic: a short line
(few words, not ending in a period) that is followed by longer body text is
treated as a heading. This is best-effort — PDF structure is genuinely lossy.
The AI mapper tolerates imperfect sectioning because it also gets the full prose.

For SCANNED PDFs (no text layer) this returns little/no text — those need OCR,
which is out of scope here. The endpoint should warn if a PDF yields no prose.

Record shape (identical to the other readers):
    {concept_name, source, section, prose, sections, lists, tables}
"""

import re
from pathlib import Path

import pdfplumber


def _clean(text: str) -> str:
    return re.sub(r"[ \t]+", " ", (text or "")).strip()


def _looks_like_heading(line: str) -> bool:
    """Heuristic: short, title-ish, no terminal period, not a bullet."""
    s = line.strip()
    if not s or len(s) > 80:
        return False
    words = s.split()
    if len(words) > 9:
        return False
    if s.endswith((".", ",", ";", ":")):
        return False
    if s[0] in "-*•·":
        return False
    # mostly title / capitalised words, or ends like a section label
    caps = sum(1 for w in words if w[:1].isupper())
    return caps >= max(1, len(words) // 2)


def _looks_like_bullet(line: str) -> bool:
    s = line.strip()
    return bool(s) and s[0] in "-*•·" or bool(re.match(r"^\(?\d+[.)]\s", s))


def _extract_structure(full_text: str):
    """Turn a flat text blob into inferred sections + lists."""
    lines = [ln for ln in (full_text or "").splitlines()]
    sections = []
    lists = []
    current = {"heading": "(top)", "level": 1, "parts": []}
    pending_list = []

    def flush_list():
        nonlocal pending_list
        if pending_list:
            lists.append(pending_list)
            pending_list = []

    def flush_section():
        if current["parts"]:
            sections.append({
                "heading": current["heading"],
                "level": current["level"],
                "text": _clean(" ".join(current["parts"])),
            })

    for raw in lines:
        line = _clean(raw)
        if not line:
            continue
        if _looks_like_bullet(line):
            pending_list.append(re.sub(r"^\(?\d+[.)]\s*|^[-*•·]\s*", "", line))
            continue
        if _looks_like_heading(line):
            flush_list()
            flush_section()
            current = {"heading": line, "level": 2, "parts": []}
        else:
            flush_list()
            current["parts"].append(line)

    flush_list()
    flush_section()
    return sections, lists


def read_pdf_file(path) -> dict:
    fp = Path(path)
    all_text_parts = []
    tables = []

    with pdfplumber.open(str(fp)) as pdf:
        # title: PDF metadata, else first non-empty line, else filename
        meta_title = (pdf.metadata or {}).get("Title")
        for page in pdf.pages:
            txt = page.extract_text() or ""
            if txt:
                all_text_parts.append(txt)
            for t in (page.extract_tables() or []):
                if not t:
                    continue
                headers = [_clean(x or "") for x in t[0]]
                body = [[_clean(x or "") for x in row] for row in t[1:]
                        if any(row)]
                tables.append({"headers": headers, "rows": body})

    full_text = "\n".join(all_text_parts)

    title = _clean(meta_title) if meta_title else None
    if not title:
        for ln in full_text.splitlines():
            if _clean(ln):
                title = _clean(ln)
                break
    if not title:
        title = fp.stem

    sections, lists = _extract_structure(full_text)
    prose = _clean(full_text.replace("\n", " "))

    return {
        "concept_name": title,
        "source": str(fp),
        "section": fp.stem,
        "prose": prose,
        "sections": sections,
        "lists": lists,
        "tables": tables,
    }


def read_pdf_dir(directory, pattern="*.pdf") -> list:
    root = Path(directory)
    files = sorted(root.rglob(pattern))
    files = [f for f in files if "__MACOSX" not in str(f)]
    return [read_pdf_file(f) for f in files]


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        print("usage: python -m readers.pdf_reader <file-or-dir> [out.json]")
        raise SystemExit(1)
    target = Path(sys.argv[1])
    recs = read_pdf_dir(target) if target.is_dir() else [read_pdf_file(target)]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    if out:
        Path(out).write_text(json.dumps(recs, indent=2), encoding="utf-8")
        print(f"wrote {len(recs)} records to {out}")
    else:
        print(json.dumps(recs[0], indent=2)[:2000])
        print(f"\n... {len(recs)} record(s) total")
