"""
doc_miner.py  —  Agent 3 (Documentation Miner) runtime retrieval.

VENDOR-AGNOSTIC. Loads every <vendor>_index.json present in the project and
answers "how does X work in <vendor>?" style questions by:
  1. embedding the question with Gemini (user's key)
  2. cosine-matching against the prebuilt chunk vectors (pure numpy)
  3. returning top passages WITH vendor + citation.

Same engine for all vendors; each vendor's knowledge is its own index file.
NO PyTorch / sentence-transformers — Render-safe.

Trust tier: 'retrieved' — always cite the source.
"""

import os
import glob
import json
from pathlib import Path

import numpy as np
import google.generativeai as genai

BASE_DIR = Path(__file__).parent
EMBED_MODEL = "models/gemini-embedding-001"

_LOADED = False
_INDEX = []       # list of {id, text, title, source, vendor}
_MATRIX = None    # (n_chunks, dim) L2-normalised
_VENDORS = []     # vendor tokens present


def _load_all():
    """Load and stack every <vendor>_index.json in the project (once)."""
    global _LOADED, _INDEX, _MATRIX, _VENDORS
    if _LOADED:
        return
    files = sorted(glob.glob(str(BASE_DIR / "*_index.json")))
    records = []
    for fp in files:
        try:
            recs = json.loads(Path(fp).read_text(encoding="utf-8"))
        except Exception:
            continue
        # infer vendor from filename if a record lacks it
        vendor = Path(fp).name.replace("_index.json", "")
        for r in recs:
            r.setdefault("vendor", vendor)
        records.extend(recs)

    if not records:
        _INDEX, _MATRIX, _VENDORS, _LOADED = [], np.zeros((0, 1)), [], True
        return

    vectors = np.array([r["vector"] for r in records], dtype=np.float32)
    norms = np.linalg.norm(vectors, axis=1, keepdims=True)
    norms[norms == 0] = 1.0
    _MATRIX = vectors / norms
    _INDEX = [{k: r.get(k) for k in ("id", "text", "title", "source", "vendor")}
              for r in records]
    _VENDORS = sorted({r["vendor"] for r in _INDEX if r.get("vendor")})
    _LOADED = True


def available_vendors() -> list:
    _load_all()
    return list(_VENDORS)


def index_is_ready() -> bool:
    _load_all()
    return len(_INDEX) > 0


def _embed_question(question: str, api_key: str):
    key = api_key or os.getenv("GEMINI_API_KEY")
    if not key:
        raise ValueError("No Gemini API key supplied for the Documentation Miner.")
    genai.configure(api_key=key)
    r = genai.embed_content(model=EMBED_MODEL, content=question,
                            task_type="retrieval_query")
    v = np.array(r["embedding"], dtype=np.float32)
    n = np.linalg.norm(v)
    return v / n if n else v


def search(question: str, api_key: str = None, top_k: int = 4,
           min_score: float = 0.35, vendor: str = None):
    """
    Return top-k passages. Optionally filter to one vendor.
    Each result: {rank, score, vendor, title, source, text}.
    """
    _load_all()
    if not _INDEX:
        return []

    qv = _embed_question(question, api_key)
    scores = _MATRIX @ qv

    order = np.argsort(scores)[::-1]
    results, rank = [], 1
    for idx in order:
        rec = _INDEX[int(idx)]
        if vendor and rec.get("vendor") != vendor:
            continue
        s = float(scores[idx])
        if s < min_score:
            continue
        results.append({"rank": rank, "score": round(s, 3),
                        "vendor": rec.get("vendor"),
                        "title": rec.get("title"),
                        "source": rec.get("source"),
                        "text": rec.get("text")})
        rank += 1
        if len(results) >= top_k:
            break
    return results


def format_citation(result: dict) -> str:
    v = result.get("vendor")
    tag = f"{v}: " if v else ""
    return f'{tag}{result.get("title")} ({result.get("source")})'


if __name__ == "__main__":
    import sys
    q = " ".join(sys.argv[1:]) or "how do I create a flow"
    if not index_is_ready():
        print("No <vendor>_index.json found. Run build_doc_index.py first.")
        raise SystemExit(1)
    print("vendors indexed:", available_vendors())
    for r in search(q):
        print(f'[{r["score"]}] {format_citation(r)}')
        print(f'    {r["text"][:160]}...')
