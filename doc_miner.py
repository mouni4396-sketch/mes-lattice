"""
doc_miner.py  —  Agent 3 (Documentation Miner) runtime retrieval, CM only.

Loads the prebuilt cm_index.json (committed to the repo) and answers
"how does X work in CM?" style questions by:
    1. embedding the user's question with Gemini (user's API key)
    2. cosine-matching against the prebuilt CM chunk vectors (pure numpy)
    3. returning the top passages WITH citations (title + source file)

This module is intentionally light: NO PyTorch, NO sentence-transformers.
The only runtime dependency beyond numpy is google-generativeai, which the
app already uses for chat. That keeps it deployable on Render's free tier.

Trust tier: results are 'retrieved' — always cite the source.
"""

import os
import json
from pathlib import Path

import numpy as np
import google.generativeai as genai

BASE_DIR = Path(__file__).parent
INDEX_PATH = BASE_DIR / "cm_index.json"
EMBED_MODEL = "models/gemini-embedding-001"

# module-level cache so we load + stack vectors only once
_INDEX = None          # list of {id, text, title, source}
_MATRIX = None         # np.ndarray (n_chunks, dim), L2-normalised


def _load_index():
    """Load cm_index.json once and pre-normalise vectors for cosine."""
    global _INDEX, _MATRIX
    if _INDEX is not None:
        return
    if not INDEX_PATH.exists():
        _INDEX = []
        _MATRIX = np.zeros((0, 1), dtype=np.float32)
        return

    with open(INDEX_PATH, "r", encoding="utf-8") as f:
        records = json.load(f)

    vectors = np.array([r["vector"] for r in records], dtype=np.float32)
    # L2 normalise so dot product == cosine similarity
    norms = np.linalg.norm(vectors, axis=1, keepdims=True)
    norms[norms == 0] = 1.0
    _MATRIX = vectors / norms

    # keep everything except the raw vector in the lightweight index
    _INDEX = [{k: r[k] for k in ("id", "text", "title", "source")}
              for r in records]


def index_is_ready() -> bool:
    _load_index()
    return len(_INDEX) > 0


def _embed_question(question: str, api_key: str):
    """Embed the query with Gemini. Returns a normalised 1-D vector or None."""
    key = api_key or os.getenv("GEMINI_API_KEY")
    if not key:
        raise ValueError("No Gemini API key supplied for Documentation Miner.")
    genai.configure(api_key=key)
    r = genai.embed_content(
        model=EMBED_MODEL,
        content=question,
        task_type="retrieval_query",
    )
    v = np.array(r["embedding"], dtype=np.float32)
    n = np.linalg.norm(v)
    return v / n if n else v


def search(question: str, api_key: str = None, top_k: int = 4,
           min_score: float = 0.35):
    """
    Return the top-k CM passages for a question.

    Output: list of dicts:
        {rank, score, title, source, text}
    Empty list if the index is missing/empty or nothing clears min_score.
    """
    _load_index()
    if not _INDEX:
        return []

    qv = _embed_question(question, api_key)
    scores = _MATRIX @ qv                      # cosine sim, shape (n_chunks,)

    order = np.argsort(scores)[::-1][:top_k]
    results = []
    rank = 1
    for idx in order:
        s = float(scores[idx])
        if s < min_score:
            continue
        rec = _INDEX[int(idx)]
        results.append({
            "rank": rank,
            "score": round(s, 3),
            "title": rec["title"],
            "source": rec["source"],
            "text": rec["text"],
        })
        rank += 1
    return results


def format_citation(result: dict) -> str:
    """Human-readable citation line for a single passage."""
    return f'{result["title"]} ({result["source"]})'


if __name__ == "__main__":
    # quick manual smoke test:
    #   $env:GEMINI_API_KEY="..."; python doc_miner.py "how do I create a flow"
    import sys
    q = " ".join(sys.argv[1:]) or "how do I create a flow"
    if not index_is_ready():
        print("cm_index.json missing or empty. Run build_cm_index.py first.")
        raise SystemExit(1)
    for r in search(q):
        print(f'[{r["score"]}] {format_citation(r)}')
        print(f'    {r["text"][:160]}...')
