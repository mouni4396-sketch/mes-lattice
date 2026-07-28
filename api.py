"""
api.py — FastAPI backend for the MES comparison platform.

Wraps the proven vendor-agnostic core (store.py + graph_engine.py) as HTTP
endpoints returning JSON. The frontend calls these; nothing here generates
SPARQL — it only runs the vetted query library and reports what's loaded.

Run:
    uvicorn api:app --reload
Then open http://127.0.0.1:8000

Add a vendor overlay by adding one line to LAYERS below — no other change.
"""

from fastapi import FastAPI, HTTPException, UploadFile, File, Form, Header
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, PlainTextResponse
from pydantic import BaseModel
import os
import tempfile

from store import OntologyStore
from graph_engine import GraphEngine
from generate_ttl import Generator
import chatbot
import reasoner

# ---------------------------------------------------------------------------
# CONFIG — the only place layers are named. Vendor identity flows from data
# after this; endpoints below never hard-code "cm"/"opc".
# ---------------------------------------------------------------------------
LAYERS = [
    ("ref", "mes-neutral-reference-4.ttl"),
    ("cm",  "mes-cm-overlay.ttl"),
    # ("opc", "mes-opcenter-overlay.ttl"),   # <- a third overlay: just add here
]
QUERY_DIR = "queries"
WEB_DIR = "web"

# ---------------------------------------------------------------------------
# Load once at startup.
# ---------------------------------------------------------------------------
store = OntologyStore(LAYERS)
engine = GraphEngine(store, query_dir=QUERY_DIR)

app = FastAPI(title="MES Comparison Platform", version="0.1.0")


# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------
class QueryRequest(BaseModel):
    params: dict = {}


class ChatRequest(BaseModel):
    question: str


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------
@app.get("/api/status")
def status():
    """What's loaded: layers, triple counts, and vendors discovered from data."""
    layers = []
    for prefix, info in store.layers.items():
        count = sum(1 for _ in store.store.quads_for_pattern(
            None, None, None, info["graph"]))
        layers.append({
            "prefix": prefix,
            "namespace": info["ns"],
            "triples": count,
            "is_reference": prefix == "ref",   # UI hint only; logic never assumes this
        })
    return {
        "layers": layers,
        "vendors": store.vendors(),            # discovered via #verdict predicate
        "query_count": len(engine.queries),
    }


@app.get("/api/queries")
def list_queries():
    """The vetted query library: each name + the parameters it needs."""
    out = []
    for name in sorted(engine.queries):
        q = engine.queries[name]
        out.append({"name": name, "params": q["params"]})
    return out


@app.get("/api/vendors")
def vendors():
    """Loaded vendor prefixes + their namespaces, for filling %VENDOR%/%FROM%/%TO%."""
    out = []
    for prefix, info in store.layers.items():
        if prefix == "ref":
            continue
        out.append({"prefix": prefix, "namespace": info["ns"]})
    return out


@app.post("/api/query/{name}")
def run_query(name: str, req: QueryRequest):
    """
    Run a vetted query by name with parameters. Returns computed rows (JSON).
    This is the authoritative 'computed' tier — exact, reproducible.
    """
    try:
        rows = engine.run(name, req.params)
    except KeyError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    # prettify IRIs for display while keeping raw values too
    pretty = []
    for r in rows:
        pretty.append({k: store.pfx(v) if str(v).startswith("http") else v
                       for k, v in r.items()})
    return {"query": name, "params": req.params, "count": len(pretty), "rows": pretty}


@app.post("/api/chat")
def chat(req: ChatRequest, x_user_api_key: str = Header(None)):
    """
    Natural-language question -> routed to one vetted query -> computed rows ->
    Gemini narrates ONLY those rows. The LLM never writes SPARQL or invents numbers.
    Returns the prose answer plus the exact query name, params, and rows for audit.
    x_user_api_key: optional per-request Gemini key from the X-User-API-Key header;
    falls back to GEMINI_API_KEY when absent or empty (handled in chatbot._gemini_client).
    """
    try:
        return chatbot.answer(req.question, engine, store, api_key=x_user_api_key)
    except RuntimeError as e:
        # e.g. missing GEMINI_API_KEY
        raise HTTPException(status_code=503, detail=str(e))


@app.post("/api/generate")
async def generate(workbook: UploadFile = File(...), prefix: str = Form(...)):
    """
    Excel -> Ontology. Takes a filled capture workbook + a vendor prefix, runs the
    proven generator, and returns the TTL as text for the human to review and
    download. It does NOT load the result into the running store — approval and
    loading stay a deliberate, separate act (human gate).
    """
    prefix = prefix.strip().lower()
    if not prefix.isalnum():
        raise HTTPException(status_code=400, detail="Prefix must be letters/digits only, e.g. 'opc'.")

    # save upload to a temp file the generator can open
    suffix = os.path.splitext(workbook.filename or "")[1] or ".xlsx"
    fd, tmp_path = tempfile.mkstemp(suffix=suffix)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(await workbook.read())
        gen = Generator(tmp_path, prefix)
        ttl = gen.generate()
        gen.wb.close()   # release the file handle so Windows can delete it
    except KeyError as e:
        raise HTTPException(status_code=400,
                            detail=f"Workbook is missing an expected sheet: {e}. "
                                   f"It must match the standard capture template.")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Generation failed: {e}")
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass   # Windows may still hold the handle briefly; leftover temp is harmless

    # quick stats for the review screen
    stats = {
        "classes":   ttl.count(" a owl:Class"),
        "attributes": ttl.count(" a owl:DatatypeProperty"),
        "objectProps": ttl.count(" a owl:ObjectProperty"),
        "operations": ttl.count(f" a {prefix}:Operation"),
        "mappings":  sum(ttl.count(f"skos:{m}") for m in
                         ("closeMatch", "broadMatch", "narrowMatch", "relatedMatch")),
        "lines": ttl.count("\n"),
    }
    filename = f"mes-{prefix}-overlay.ttl"
    return {"prefix": prefix, "filename": filename, "stats": stats, "ttl": ttl,
            "tier": "generated", "gated": True}


@app.post("/api/reason")
def reason(req: ChatRequest, x_user_api_key: str = Header(None)):
    """
    Ontology Reasoner (retrieved-from-graph). Answers relationship / how-it-works
    questions by tracing the graph, then narrating only the traced facts. Distinct
    from the computed comparison tier — it explains structure, not counts.
    """
    try:
        return reasoner.answer(req.question, store, api_key=x_user_api_key)
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


@app.post("/api/load-ttl")
async def load_ttl(ttl: UploadFile = File(...), prefix: str = Form(...)):
    """
    Upload a ready-made .ttl overlay and load it into the RUNNING store under the
    given prefix. This is an in-memory load only — it does NOT edit LAYERS, so a
    restart reverts to the configured layers. To make an overlay permanent, add it
    to LAYERS in api.py (a deliberate, human act).
    """
    prefix = prefix.strip().lower()
    if not prefix.isalnum():
        raise HTTPException(status_code=400, detail="Prefix must be letters/digits only.")
    if prefix == "ref":
        raise HTTPException(status_code=400, detail="'ref' is the frozen reference layer and cannot be replaced.")

    import pyoxigraph as ox
    data = await ttl.read()
    graph = ox.NamedNode(f"urn:graph:{prefix}")
    try:
        # if this prefix is already loaded, clear it first (replace semantics)
        for quad in list(store.store.quads_for_pattern(None, None, None, graph)):
            store.store.remove(quad)
        store.store.load(data, format=ox.RdfFormat.TURTLE, to_graph=graph)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid Turtle: {e}")

    # register/refresh the layer so discovery + queries see it
    ns = store._discover_namespace(graph, prefix)
    store.layers[prefix] = {"graph": graph, "ns": ns, "path": f"(uploaded: {ttl.filename})"}
    if ns:
        store.ns_to_prefix[ns] = prefix
    count = sum(1 for _ in store.store.quads_for_pattern(None, None, None, graph))
    return {"prefix": prefix, "namespace": ns, "triples": count,
            "vendors_now": store.vendors(), "note": "loaded in memory; restart reverts to LAYERS config"}


# ---------------------------------------------------------------------------
# Frontend (served last so /api/* wins)
# ---------------------------------------------------------------------------
@app.get("/")
def index():
    return FileResponse(os.path.join(WEB_DIR, "index.html"))

if os.path.isdir(WEB_DIR):
    app.mount("/web", StaticFiles(directory=WEB_DIR), name="web")
