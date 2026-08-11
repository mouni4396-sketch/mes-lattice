"""
readers/html_reader.py  —  Agent 6 (Intake) stage 1, HTML format reader.

DETERMINISTIC. No AI. Turns one or many HTML doc pages into the neutral
'concept record' shape that the AI mapper (stage 2) consumes.

A concept record is intentionally vendor-agnostic and format-agnostic — the
PDF and docx readers will emit the SAME shape, so the mapper never knows or
cares which format a concept came from.

    {
      "concept_name": str,      # page H1 / <title>, cleaned
      "source": str,            # citation (path or URL)
      "section": str,           # e.g. "business-data/flow/create_flow"
      "prose": str,             # full readable text (deboilerplated)
      "sections": [             # heading -> text under it (structure kept)
         {"heading": str, "level": int, "text": str}
      ],
      "lists": [ [str, ...], ... ],   # bullet lists (often constraints/fields)
      "tables": [ {"headers": [...], "rows": [[...], ...]} ],  # if any
    }

Usage:
    from readers.html_reader import read_html_file, read_html_dir
    records = read_html_dir("docs/cm/business-data")
    # -> list[dict] ready to hand to the mapper, or dump to JSON to eyeball
"""

import re
import json
from pathlib import Path

from bs4 import BeautifulSoup

TITLE_SUFFIX = " - Critical Manufacturing Documentation Portal"
_BOILER = ["script", "style", "nav", "header", "footer", "noscript", "aside"]
# mermaid/diagram noise seen in CM pages
_NOISE_RE = re.compile(r"(graph TD;|classDef |stroke-width:|fill:#|-->)")


def _clean(text: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _strip_title_suffix(title: str) -> str:
    if title and title.endswith(TITLE_SUFFIX):
        return title[: -len(TITLE_SUFFIX)].strip()
    return title


def _section_from_path(fp: Path, root: Path) -> str:
    """business-data/flow/create_flow style id from the file location."""
    try:
        rel = fp.relative_to(root.parent if root.name else root)
    except ValueError:
        rel = fp
    parts = list(rel.parts)
    if parts and parts[-1] == "index.html":
        parts = parts[:-1]
    elif parts:
        parts[-1] = parts[-1].replace(".html", "")
    return "/".join(parts)


def _extract_sections(soup: BeautifulSoup):
    """
    Walk the document, grouping text under its nearest preceding heading.
    Preserves the doc's own structure (Overview / Setup / Preconditions / ...),
    which is the main signal the mapper uses to locate fields and behaviour.
    """
    sections = []
    current = {"heading": "(top)", "level": 0, "parts": []}

    for el in soup.find_all(["h1", "h2", "h3", "h4", "p", "li"]):
        name = el.name
        txt = _clean(el.get_text(" ", strip=True))
        if not txt:
            continue
        if name in ("h1", "h2", "h3", "h4"):
            # flush previous
            if current["parts"]:
                sections.append({
                    "heading": current["heading"],
                    "level": current["level"],
                    "text": _clean(" ".join(current["parts"])),
                })
            heading = txt.rstrip("#").strip()      # CM headings end with '#'
            current = {"heading": heading, "level": int(name[1]), "parts": []}
        else:
            if not _NOISE_RE.search(txt):
                current["parts"].append(txt)

    if current["parts"]:
        sections.append({
            "heading": current["heading"],
            "level": current["level"],
            "text": _clean(" ".join(current["parts"])),
        })
    return sections


def _extract_lists(soup: BeautifulSoup):
    lists = []
    for ul in soup.find_all(["ul", "ol"]):
        items = []
        for li in ul.find_all("li", recursive=False):
            t = _clean(li.get_text(" ", strip=True))
            if t and not _NOISE_RE.search(t):
                items.append(t)
        if items:
            lists.append(items)
    return lists


def _extract_tables(soup: BeautifulSoup):
    tables = []
    for tb in soup.find_all("table"):
        rows = tb.find_all("tr")
        if not rows:
            continue
        headers = [_clean(c.get_text(" ", strip=True))
                   for c in rows[0].find_all(["th", "td"])]
        body = []
        for r in rows[1:]:
            cells = [_clean(c.get_text(" ", strip=True))
                     for c in r.find_all(["th", "td"])]
            if any(cells):
                body.append(cells)
        tables.append({"headers": headers, "rows": body})
    return tables


def read_html_file(path, root=None) -> dict:
    """Parse one HTML file into a concept record."""
    fp = Path(path)
    root = Path(root) if root else fp.parent
    html = fp.read_text(encoding="utf-8", errors="ignore")
    soup = BeautifulSoup(html, "html.parser")

    # title
    title = None
    if soup.title and soup.title.string:
        title = _strip_title_suffix(soup.title.string.strip())
    if not title:
        h1 = soup.find("h1")
        title = h1.get_text(strip=True).rstrip("#") if h1 else fp.stem

    # deboilerplate for prose + structure
    for tag in soup(_BOILER):
        tag.decompose()

    sections = _extract_sections(soup)
    lists = _extract_lists(soup)
    tables = _extract_tables(soup)
    prose = _clean(soup.get_text(" "))

    return {
        "concept_name": title.rstrip("#").strip(),
        "source": str(fp),
        "section": _section_from_path(fp, root),
        "prose": prose,
        "sections": sections,
        "lists": lists,
        "tables": tables,
    }


def read_html_dir(directory, pattern="*.html") -> list:
    """Recursively parse every HTML file under a directory."""
    root = Path(directory)
    files = sorted(root.rglob(pattern))
    files = [f for f in files if "__MACOSX" not in str(f)]
    return [read_html_file(f, root=root) for f in files]


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("usage: python -m readers.html_reader <file-or-dir> [out.json]")
        raise SystemExit(1)
    target = Path(sys.argv[1])
    if target.is_dir():
        recs = read_html_dir(target)
    else:
        recs = [read_html_file(target)]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    if out:
        Path(out).write_text(json.dumps(recs, indent=2), encoding="utf-8")
        print(f"wrote {len(recs)} records to {out}")
    else:
        # print the first record as a sample
        print(json.dumps(recs[0], indent=2)[:2000])
        print(f"\n... {len(recs)} record(s) total")
