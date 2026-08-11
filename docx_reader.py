"""
readers/docx_reader.py  —  Agent 6 (Intake) stage 1, Word (.docx) format reader.

DETERMINISTIC. No AI. Emits the SAME concept-record shape as html_reader.py, so
Word docs flow through the identical mapper -> draft_writer chain. The mapper
never knows or cares whether a concept came from HTML or Word.

Record shape (identical to html_reader):
    {
      "concept_name": str,     # doc Title property, or first Heading 1, or filename
      "source": str,           # file path (citation)
      "section": str,          # filename stem (docs rarely have a URL path)
      "prose": str,            # full readable text
      "sections": [ {"heading": str, "level": int, "text": str} ],
      "lists": [ [str, ...], ... ],
      "tables": [ {"headers": [...], "rows": [[...], ...]} ],
    }

Uses python-docx (pure Python, Render-safe). No pandoc / system deps.

Usage:
    from readers.docx_reader import read_docx_file, read_docx_dir
    records = read_docx_dir("docs/vendorX")
"""

import re
from pathlib import Path

import docx                       # pip install python-docx
from docx.document import Document as _Doc
from docx.table import Table
from docx.text.paragraph import Paragraph


def _clean(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def _heading_level(style_name: str):
    """Return heading level int if the paragraph style is a heading, else None."""
    if not style_name:
        return None
    m = re.match(r"Heading\s+(\d+)", style_name, re.I)
    if m:
        return int(m.group(1))
    if style_name.lower() in ("title",):
        return 1
    return None


def _is_list(style_name: str) -> bool:
    s = (style_name or "").lower()
    return "list" in s or "bullet" in s


def _iter_block_items(parent):
    """Yield paragraphs and tables in document order (python-docx doesn't do
    this natively). Works at document-body level."""
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    body = parent.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def _table_to_record(tbl: Table):
    rows = tbl.rows
    if not rows:
        return None
    headers = [_clean(c.text) for c in rows[0].cells]
    body = []
    for r in rows[1:]:
        cells = [_clean(c.text) for c in r.cells]
        if any(cells):
            body.append(cells)
    return {"headers": headers, "rows": body}


def read_docx_file(path) -> dict:
    fp = Path(path)
    doc = docx.Document(str(fp))

    # title: core property, else first heading/title paragraph, else filename
    title = None
    try:
        if doc.core_properties.title:
            title = doc.core_properties.title.strip()
    except Exception:
        pass

    sections = []
    lists = []
    tables = []
    current = {"heading": "(top)", "level": 0, "parts": []}
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

    for block in _iter_block_items(doc):
        if isinstance(block, Table):
            flush_list()
            t = _table_to_record(block)
            if t:
                tables.append(t)
            continue

        # paragraph
        txt = _clean(block.text)
        if not txt:
            continue
        style = block.style.name if block.style else ""
        lvl = _heading_level(style)

        if lvl is not None:
            # new heading -> close previous section + any open list
            flush_list()
            flush_section()
            if title is None:
                title = txt          # first heading becomes the title fallback
            current = {"heading": txt, "level": lvl, "parts": []}
        elif _is_list(style):
            pending_list.append(txt)
        else:
            flush_list()
            current["parts"].append(txt)

    flush_list()
    flush_section()

    if not title:
        title = fp.stem

    prose = _clean(" ".join(
        s["text"] for s in sections
    )) or _clean(" ".join(p.text for p in doc.paragraphs))

    return {
        "concept_name": title,
        "source": str(fp),
        "section": fp.stem,
        "prose": prose,
        "sections": sections,
        "lists": lists,
        "tables": tables,
    }


def read_docx_dir(directory, pattern="*.docx") -> list:
    root = Path(directory)
    files = sorted(root.rglob(pattern))
    files = [f for f in files if "__MACOSX" not in str(f)
             and not f.name.startswith("~$")]        # skip Word lock files
    return [read_docx_file(f) for f in files]


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        print("usage: python -m readers.docx_reader <file-or-dir> [out.json]")
        raise SystemExit(1)
    target = Path(sys.argv[1])
    recs = read_docx_dir(target) if target.is_dir() else [read_docx_file(target)]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    if out:
        Path(out).write_text(json.dumps(recs, indent=2), encoding="utf-8")
        print(f"wrote {len(recs)} records to {out}")
    else:
        print(json.dumps(recs[0], indent=2)[:2000])
        print(f"\n... {len(recs)} record(s) total")
