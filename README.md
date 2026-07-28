# MES Comparison Platform

Vendor-neutral semiconductor MES comparison, **computed** from an RDF ontology.
Comparison numbers come from vetted SPARQL over the graph — never from a model
reading the ontology. Vendors are discovered from the data; a new overlay slots
in with no code change.

## Project layout

```
Oxigraph-chat/
├─ api.py                      FastAPI backend (HTTP endpoints)
├─ store.py                    Vendor-agnostic ontology store (load + SPARQL)
├─ graph_engine.py            The Graph Engineer (runs the vetted query library)
├─ requirements.txt
├─ mes-neutral-reference-4.ttl   frozen reference layer (ref:)
├─ mes-cm-overlay.ttl            Critical Manufacturing overlay (cm:)
├─ queries/                    the 8 vetted .rq files (q1–q8)
└─ web/
   └─ index.html              frontend (calls the API, no build step)
```

## Run

```
pip install -r requirements.txt
uvicorn api:app --reload
```

Then open http://127.0.0.1:8000

- Sidebar shows loaded graphs and the vendors discovered from data.
- Click a query card, fill any parameters, hit Run. Results are computed rows.

## Add a vendor overlay

1. Drop its TTL in the folder (e.g. `mes-opcenter-overlay.ttl`).
2. Add one line to `LAYERS` in `api.py`:
   ```python
   ("opc", "mes-opcenter-overlay.ttl"),
   ```
3. Restart. The new vendor appears everywhere automatically — no other change.

## The three tiers (design)

- **Computed** (this app today): vetted SPARQL, exact and reproducible. Authoritative.
- **Retrieved** (later): passages from vendor docs, cited.
- **Generated** (later): model-authored drafts — always human-approved before they count.

Only the computed tier is authoritative on its own. Keep the tiers visually distinct.

## Endpoints

| Method | Path | Purpose |
|---|---|---|
| GET  | `/api/status`          | loaded layers, triple counts, vendors, query count |
| GET  | `/api/queries`         | vetted query names + params each needs |
| GET  | `/api/vendors`         | loaded vendor prefixes + namespaces |
| POST | `/api/query/{name}`    | run a vetted query; body `{"params": {...}}` |

## Notes / guardrails

- Query bodies are never generated or edited at runtime — only `%PARAM%` values filled.
- `q3` (gap) is only meaningful alongside `q8` (completeness guard): a low q8 count
  means an overlay is thin, so its "gaps" are "not yet modelled", not proven absence.
- Vendors are counted by the `#verdict` predicate (ref: never carries it), which keeps
  ref: out of the vendor set. This is why `q4` (convergence) was fixed to count that way.
```
