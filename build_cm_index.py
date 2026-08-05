"""
build_cm_index.py  —  Agent 3 (Documentation Miner) offline indexer, CM only.

Indexes the Critical Manufacturing 'business-data' user-guide section into a
committed vector index (cm_index.json). Run locally with your Gemini key.

FREE-TIER FRIENDLY: the free embedding quota is ~1000 requests/day. This full
build needs more than that, so the script CHECKPOINTS and RESUMES:
  - progress is saved to cm_index.json after every file
  - already-embedded files are skipped on the next run
  - when the daily quota (429) is hit, it saves and exits cleanly
Just re-run it each day until it prints "ALL FILES DONE".

USAGE (PowerShell, from project root):
    $env:GEMINI_API_KEY="your_key_here"
    python build_cm_index.py          # run again tomorrow to continue

INPUT  (nested, recursive):
    docs/cm/business-data/**/*.html
OUTPUT:
    cm_index.json         the vector index (grows across runs)
    cm_index.progress     list of files already embedded (checkpoint)
"""

import os
import re
import json
import time
from pathlib import Path

import google.generativeai as genai
from bs4 import BeautifulSoup

BASE_DIR = Path(__file__).parent
DOCS_ROOT = BASE_DIR / "docs" / "cm" / "business-data"
OUT_PATH = BASE_DIR / "cm_index.json"
PROGRESS_PATH = BASE_DIR / "cm_index.progress"

EMBED_MODEL = "models/gemini-embedding-001"
CHUNK_WORDS = 220
CHUNK_OVERLAP = 40
TITLE_SUFFIX = " - Critical Manufacturing Documentation Portal"

_NOISE_RE = re.compile(r"(graph TD;|classDef |stroke-width:|fill:#)")


class QuotaExceeded(Exception):
    """Raised when the daily free-tier embedding quota is hit."""


def get_api_key() -> str:
    key = os.getenv("GEMINI_API_KEY")
    if not key:
        raise SystemExit(
            "GEMINI_API_KEY not set. In PowerShell:\n"
            '   $env:GEMINI_API_KEY="your_key_here"\n'
            "then re-run: python build_cm_index.py"
        )
    return key


def citation_from_path(fp: Path):
    rel = fp.relative_to(DOCS_ROOT.parent)
    parts = list(rel.parts)
    if parts[-1] == "index.html":
        parts = parts[:-1]
    else:
        parts[-1] = parts[-1].replace(".html", "")
    return "/".join(parts)


def extract_text_and_title(html: str, fallback: str):
    soup = BeautifulSoup(html, "html.parser")
    title = None
    if soup.title and soup.title.string:
        title = soup.title.string.strip()
    if title and title.endswith(TITLE_SUFFIX):
        title = title[: -len(TITLE_SUFFIX)].strip()
    if not title:
        h1 = soup.find("h1")
        title = h1.get_text(strip=True) if h1 else fallback
    for tag in soup(["script", "style", "nav", "header", "footer", "noscript"]):
        tag.decompose()
    text = soup.get_text(separator=" ")
    text = re.sub(r"\s+", " ", text).strip()
    return text, title


def chunk_words(text: str, size: int, overlap: int):
    words = text.split()
    if not words:
        return []
    chunks, start, step = [], 0, max(1, size - overlap)
    while start < len(words):
        window = words[start:start + size]
        if window:
            chunk = " ".join(window)
            if not _NOISE_RE.search(chunk) or len(chunk.split()) > 60:
                chunks.append(chunk)
        start += step
    return chunks


def embed(text: str):
    """Embed one chunk. Raises QuotaExceeded on 429 so we can stop cleanly."""
    for attempt in range(2):
        try:
            r = genai.embed_content(
                model=EMBED_MODEL, content=text,
                task_type="retrieval_document",
            )
            return r["embedding"]
        except Exception as e:
            msg = str(e)
            if "429" in msg or "quota" in msg.lower() or "ResourceExhausted" in msg:
                raise QuotaExceeded(msg)
            if attempt == 0:
                time.sleep(2)
                continue
            raise
    return None


# ---- checkpoint helpers ---------------------------------------------------
def load_progress():
    """Return (records list, set of done url_paths)."""
    records = []
    done = set()
    if OUT_PATH.exists():
        try:
            with open(OUT_PATH, "r", encoding="utf-8") as f:
                records = json.load(f)
        except Exception:
            records = []
    if PROGRESS_PATH.exists():
        try:
            with open(PROGRESS_PATH, "r", encoding="utf-8") as f:
                done = set(json.load(f))
        except Exception:
            done = set()
    return records, done


def save_progress(records, done):
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(records, f)
    with open(PROGRESS_PATH, "w", encoding="utf-8") as f:
        json.dump(sorted(done), f)


# --------------------------------------------------------------------------
def main():
    genai.configure(api_key=get_api_key())

    if not DOCS_ROOT.exists():
        raise SystemExit(
            f"Folder not found: {DOCS_ROOT}\n"
            "Copy the CM 'business-data' folder to:\n"
            f"   {DOCS_ROOT}\n"
        )

    html_files = sorted(DOCS_ROOT.rglob("*.html"))
    html_files = [f for f in html_files if "__MACOSX" not in str(f)]
    if not html_files:
        raise SystemExit(f"No .html files under {DOCS_ROOT}")

    records, done = load_progress()
    next_id = (max((r["id"] for r in records), default=-1)) + 1
    todo = [f for f in html_files if citation_from_path(f) not in done]

    print(f"Total files: {len(html_files)}   "
          f"already done: {len(done)}   remaining: {len(todo)}")
    if not todo:
        print("ALL FILES DONE. cm_index.json is complete.")
        return
    print("Indexing... will stop cleanly if the daily quota is hit.\n")

    processed = 0
    try:
        for fp in todo:
            with open(fp, "r", encoding="utf-8", errors="ignore") as f:
                html = f.read()

            url_path = citation_from_path(fp)
            text, title = extract_text_and_title(html, url_path)
            chunks = chunk_words(text, CHUNK_WORDS, CHUNK_OVERLAP)

            for ch in chunks:
                vec = embed(ch)            # may raise QuotaExceeded
                if vec is None:
                    continue
                records.append({
                    "id": next_id, "text": ch, "title": title,
                    "source": url_path, "url_path": url_path, "vector": vec,
                })
                next_id += 1
                time.sleep(0.05)

            done.add(url_path)
            processed += 1

            # checkpoint after every file
            save_progress(records, done)
            if processed % 25 == 0:
                print(f"  +{processed} files this run  "
                      f"({len(done)}/{len(html_files)} total)  "
                      f"-> {len(records)} chunks")

    except QuotaExceeded:
        save_progress(records, done)
        print(f"\nDaily free-tier quota hit. Saved progress: "
              f"{len(done)}/{len(html_files)} files, {len(records)} chunks.")
        print("Re-run this SAME command tomorrow to continue where it stopped.")
        return
    except KeyboardInterrupt:
        save_progress(records, done)
        print(f"\nStopped by user. Saved {len(done)}/{len(html_files)} files.")
        return

    save_progress(records, done)
    print(f"\nALL FILES DONE. {len(records)} chunks in {OUT_PATH.name}")
    print("You can delete cm_index.progress now. Commit cm_index.json.")
    print("Do NOT commit your API key.")


if __name__ == "__main__":
    main()
