"""
build_doc_index.py  —  Agent 3 (Documentation Miner) offline indexer.

VENDOR-AGNOSTIC. The engine doesn't care whose docs it reads — you pass a
vendor token and a docs folder, it writes <vendor>_index.json. Same engine,
one index per vendor (mirrors the TTL overlay pattern: neutral code, per-vendor
data).

FREE-TIER FRIENDLY: embedding quota is ~1000 requests/day. Large doc sets need
more than that, so this CHECKPOINTS and RESUMES:
  - progress saved after every file
  - already-embedded files skipped on the next run
  - clean stop + save when the daily quota (429) is hit
Re-run the same command each day until it prints "ALL FILES DONE".

USAGE (PowerShell, from project root):
    $env:GEMINI_API_KEY="your_key_here"
    python build_doc_index.py --vendor cm  --docs docs/cm/business-data
    python build_doc_index.py --vendor opc --docs docs/opc
    # optional: --title-suffix " - Critical Manufacturing Documentation Portal"

OUTPUT:
    <vendor>_index.json         the vector index (grows across runs)
    <vendor>_index.progress     checkpoint of files already embedded

Reads .html recursively. (For .docx/.pdf sources, run the intake readers first
or extend this with those readers — the record shape is shared.)
"""

import os
import re
import json
import time
import argparse
from pathlib import Path

import google.generativeai as genai
from bs4 import BeautifulSoup

BASE_DIR = Path(__file__).parent
EMBED_MODEL = "models/gemini-embedding-001"
CHUNK_WORDS = 220
CHUNK_OVERLAP = 40
_NOISE_RE = re.compile(r"(graph TD;|classDef |stroke-width:|fill:#)")


class QuotaExceeded(Exception):
    pass


def get_api_key() -> str:
    key = os.getenv("GEMINI_API_KEY")
    if not key:
        raise SystemExit(
            'GEMINI_API_KEY not set. In PowerShell:\n'
            '   $env:GEMINI_API_KEY="your_key_here"\n'
        )
    return key


def citation_from_path(fp: Path, docs_root: Path) -> str:
    """Stable citation id from the file's location under the docs root."""
    try:
        rel = fp.relative_to(docs_root)
    except ValueError:
        rel = Path(fp.name)
    parts = list(rel.parts)
    if parts and parts[-1] == "index.html":
        parts = parts[:-1]
    elif parts:
        parts[-1] = parts[-1].replace(".html", "")
    # prefix with the docs-root folder name for readability
    return "/".join([docs_root.name] + parts) if parts else docs_root.name


def extract_text_and_title(html: str, fallback: str, title_suffix: str = ""):
    soup = BeautifulSoup(html, "html.parser")
    title = None
    if soup.title and soup.title.string:
        title = soup.title.string.strip()
    if title and title_suffix and title.endswith(title_suffix):
        title = title[: -len(title_suffix)].strip()
    if not title:
        h1 = soup.find("h1")
        title = h1.get_text(strip=True).rstrip("#") if h1 else fallback
    for tag in soup(["script", "style", "nav", "header", "footer", "noscript"]):
        tag.decompose()
    text = re.sub(r"\s+", " ", soup.get_text(" ")).strip()
    return text, title.rstrip("#").strip()


def chunk_words(text: str, size: int, overlap: int):
    words = text.split()
    if not words:
        return []
    chunks, start, step = [], 0, max(1, size - overlap)
    while start < len(words):
        window = words[start:start + size]
        if window:
            ch = " ".join(window)
            if not _NOISE_RE.search(ch) or len(ch.split()) > 60:
                chunks.append(ch)
        start += step
    return chunks


def embed(text: str):
    for attempt in range(2):
        try:
            r = genai.embed_content(model=EMBED_MODEL, content=text,
                                    task_type="retrieval_document")
            return r["embedding"]
        except Exception as e:
            msg = str(e)
            if "429" in msg or "quota" in msg.lower() or "ResourceExhausted" in msg:
                raise QuotaExceeded(msg)
            if attempt == 0:
                time.sleep(2); continue
            raise
    return None


def load_progress(out_path, prog_path):
    records, done = [], set()
    if out_path.exists():
        try: records = json.loads(out_path.read_text(encoding="utf-8"))
        except Exception: records = []
    if prog_path.exists():
        try: done = set(json.loads(prog_path.read_text(encoding="utf-8")))
        except Exception: done = set()
    return records, done


def save_progress(out_path, prog_path, records, done):
    out_path.write_text(json.dumps(records), encoding="utf-8")
    prog_path.write_text(json.dumps(sorted(done)), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description="Vendor-agnostic doc indexer for Agent 3.")
    ap.add_argument("--vendor", required=True, help="vendor token, e.g. cm, opc")
    ap.add_argument("--docs", required=True, help="path to the vendor docs folder")
    ap.add_argument("--title-suffix", default="",
                    help="optional boilerplate title suffix to strip")
    args = ap.parse_args()

    genai.configure(api_key=get_api_key())

    docs_root = Path(args.docs)
    if not docs_root.exists():
        raise SystemExit(f"Docs folder not found: {docs_root}")

    out_path = BASE_DIR / f"{args.vendor}_index.json"
    prog_path = BASE_DIR / f"{args.vendor}_index.progress"

    html_files = sorted(docs_root.rglob("*.html"))
    html_files = [f for f in html_files if "__MACOSX" not in str(f)]
    if not html_files:
        raise SystemExit(f"No .html files under {docs_root}")

    records, done = load_progress(out_path, prog_path)
    next_id = max((r["id"] for r in records), default=-1) + 1
    todo = [f for f in html_files
            if citation_from_path(f, docs_root) not in done]

    print(f"[{args.vendor}] total={len(html_files)} done={len(done)} "
          f"remaining={len(todo)}")
    if not todo:
        print("ALL FILES DONE.")
        return

    processed = 0
    try:
        for fp in todo:
            html = fp.read_text(encoding="utf-8", errors="ignore")
            cite = citation_from_path(fp, docs_root)
            text, title = extract_text_and_title(html, cite, args.title_suffix)
            for ch in chunk_words(text, CHUNK_WORDS, CHUNK_OVERLAP):
                vec = embed(ch)
                if vec is None:
                    continue
                records.append({"id": next_id, "text": ch, "title": title,
                                "source": cite, "vendor": args.vendor,
                                "vector": vec})
                next_id += 1
                time.sleep(0.05)
            done.add(cite)
            processed += 1
            save_progress(out_path, prog_path, records, done)
            if processed % 25 == 0:
                print(f"  +{processed} files ({len(done)}/{len(html_files)}) "
                      f"-> {len(records)} chunks")
    except QuotaExceeded:
        save_progress(out_path, prog_path, records, done)
        print(f"\nDaily quota hit. Saved {len(done)}/{len(html_files)} files, "
              f"{len(records)} chunks. Re-run tomorrow to continue.")
        return
    except KeyboardInterrupt:
        save_progress(out_path, prog_path, records, done)
        print(f"\nStopped. Saved {len(done)}/{len(html_files)} files.")
        return

    save_progress(out_path, prog_path, records, done)
    print(f"\nALL FILES DONE. {len(records)} chunks in {out_path.name}")
    print(f"You can delete {prog_path.name}. Commit {out_path.name} if desired.")


if __name__ == "__main__":
    main()
