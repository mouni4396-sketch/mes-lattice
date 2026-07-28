"""
graph_engine.py — the Graph Engineer (the ONLY authoritative, computed tier).

Loads every vetted .rq file in a directory, discovers each query's %PARAMS%
by scanning its text, and runs them against the OntologyStore. The query
BODIES are never generated or edited at runtime — only their parameters are
filled. That is what keeps comparison answers exact and reproducible.

Usage:
    from store import OntologyStore
    from graph_engine import GraphEngine

    store  = OntologyStore([("ref","mes-neutral-reference-4.ttl"),
                            ("cm","mes-cm-overlay.ttl")])
    engine = GraphEngine(store, query_dir=".")

    engine.list_queries()                    # what's available + params each needs
    engine.run("q1_coverage_by_dataobject")  # no params
    engine.run("q3_gap_by_vendor", {"VENDOR": store.layers["cm"]["ns"]})
"""

import os
import re
import glob


class GraphEngine:
    def __init__(self, store, query_dir="."):
        self.store = store
        self.queries = {}   # name -> {"text": ..., "params": [...], "path": ...}
        self._load_query_files(query_dir)

    def _load_query_files(self, query_dir):
        paths = sorted(glob.glob(os.path.join(query_dir, "*.rq")))
        for path in paths:
            name = os.path.splitext(os.path.basename(path))[0]
            with open(path, "r", encoding="utf-8") as f:
                text = f.read()
            params = sorted(set(re.findall(r"%([A-Z_]+)%", text)))
            self.queries[name] = {"text": text, "params": params, "path": path}
        print(f"GraphEngine: loaded {len(self.queries)} vetted queries from {query_dir}")

    def list_queries(self):
        """Print each query name and the parameters it expects."""
        for name in sorted(self.queries):
            p = self.queries[name]["params"]
            need = f"  params: {', '.join(p)}" if p else "  (no params)"
            print(f"  {name}{need}")

    def run(self, name, params=None):
        """
        Run a vetted query by name. `params` is a dict like {"VENDOR": "<iri>"}.
        Raises if a required %PARAM% is missing, so a half-filled query can never
        run and return a misleading result.
        """
        if name not in self.queries:
            raise KeyError(f"No query named '{name}'. Available: {sorted(self.queries)}")

        q = self.queries[name]
        required = set(q["params"])
        given = set((params or {}).keys())
        missing = required - given
        if missing:
            raise ValueError(
                f"Query '{name}' needs parameter(s) {sorted(missing)} but they were not given. "
                f"Example: engine.run('{name}', {{'{sorted(missing)[0]}': '<value>'}})"
            )

        return self.store.sparql(q["text"], params)

    # ----- convenience helpers so callers use vendor PREFIXES, not raw IRIs -----

    def vendor_ns(self, prefix):
        """Resolve a vendor prefix ('cm') to its full namespace IRI, for %VENDOR% params."""
        info = self.store.layers.get(prefix)
        if not info:
            raise KeyError(f"No layer '{prefix}' loaded. Loaded: {sorted(self.store.layers)}")
        return info["ns"]
