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
from pathlib import Path

from store import OntologyStore
from graph_engine import GraphEngine
from generate_ttl import Generator
import chatbot
import reasoner

# ---------------------------------------------------------------------------
# CONFIG — the only place layers are named. Vendor identity flows from data
# after this; endpoints below never hard-code "cm"/"opc".
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent

LAYERS = [
    ("ref", BASE_DIR / "mes-neutral-reference-4.ttl"),
    ("cm",  BASE_DIR / "mes-cm-overlay.ttl"),
    # ("opc", BASE_DIR / "mes-opcenter-overlay.ttl"),   # <- a third overlay: just add here
]
QUERY_DIR = BASE_DIR / "queries"
WEB_DIR = BASE_DIR / "web"

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


@app.post("/api/docs")
async def api_docs(payload: dict, x_user_api_key: str = Header(default=None)):
    """
    Documentation Miner (Agent 3), retrieved tier. Embeds the question once
    (Gemini, user's key) and cosine-matches it against the prebuilt per-vendor
    doc indexes - no LLM generation, every result carries its source citation.

    doc_miner is imported LAZILY here, not at module top-level, same reasoning
    as mapper/html_reader/docx_reader/pdf_reader in /api/intake: it imports
    numpy, which is listed in requirements.txt but - like anthropic/
    beautifulsoup4/python-docx/pdfplumber - isn't guaranteed installed in every
    environment. A missing numpy must only break /api/docs, not app startup.
    """
    try:
        import doc_miner
    except ImportError as e:
        raise HTTPException(status_code=503, detail=f"Doc search is not available: {e}. pip install numpy.")

    question = (payload.get("question") or "").strip()
    vendor = payload.get("vendor")          # optional; None = all vendors
    if not question:
        raise HTTPException(status_code=400, detail="No question provided.")
    key = x_user_api_key or os.getenv("GEMINI_API_KEY")
    if not key:
        raise HTTPException(status_code=400, detail="No Gemini API key. Add it in Settings.")
    try:
        results = doc_miner.search(question, api_key=key, top_k=4, vendor=vendor)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Doc search failed: {e}")
    return {"question": question,
            "vendors": doc_miner.available_vendors(),
            "results": results}   # each: {rank,score,vendor,title,source,text}


# ---------------------------------------------------------------------------
# Graph visualizer (/api/graph, /api/graph/layers)
#
# Merged from api_graph_endpoint.py. Rewired to this project's real shape:
#   - all reads go through store.sparql(), which already prepends the prefix
#     block (ref:/cm:/... + owl:/rdfs:/skos:/...) and runs with
#     use_default_graph_as_union=True (see store.py) - the placeholder's raw
#     store.query() call and its _val()/row["s"] unwrapping are gone; graph_engine.py
#     never does manual pyoxigraph unwrapping either, it just uses store.sparql().
#   - verdict is VENDOR-namespaced (cm:verdict, a future opc:verdict, ...), not
#     ref:verdict - there is no single VERDICT_PRED IRI to hard-code, so this
#     discovers it the same way store.vendors() does: by the "#verdict" suffix.
#   - the comparison-story mapping edges are direct skos:closeMatch/broadMatch/
#     narrowMatch/relatedMatch triples from a vendor term to a ref: concept -
#     the exact pattern every queries/*.rq file uses. ref:productSource /
#     ref:referenceTarget / ref:matchType (the placeholder's assumption) don't
#     exist under ref: at all; cm:productSource etc. exist only as a reified
#     cm:OperationMapping construct for lateral cm->peer-vendor Operations
#     mapping, a different and much narrower thing.
#   - structural edges are anything named operatesOn in ANY namespace, found by
#     the same "#operatesOn" suffix trick - covers both ref:Capability
#     ->ref:DataObject (what queries/q2 uses) and every vendor's
#     {prefix}:PortalOperation->{prefix}:DataObject verb axis, with no vendor
#     namespace hard-coded. ref:affectsDataObject / ref:partOfCapability (the
#     placeholder's assumption) don't exist anywhere in this schema.
#   - layer-of-an-IRI uses store.ns_to_prefix (the same map store.pfx() uses),
#     not a "mes-" string-split heuristic.
# ---------------------------------------------------------------------------
def _graph_layer_of(iri: str) -> str:
    for ns, prefix in store.ns_to_prefix.items():
        if iri.startswith(ns):
            return prefix
    return "unknown"


def _graph_local(iri: str) -> str:
    return iri.rsplit("#", 1)[1] if "#" in iri else iri.rstrip("/").rsplit("/", 1)[-1]


def discover_graph_layers():
    """All layers with at least one owl:Class, ref: first (mirrors store.vendors()'s
    ref-first-then-sorted convention)."""
    rows = store.sparql("SELECT DISTINCT ?s WHERE { ?s a owl:Class }")
    layers = {_graph_layer_of(r["s"]) for r in rows if r.get("s", "").startswith("http")}
    layers.discard("unknown")
    return (["ref"] if "ref" in layers else []) + sorted(layers - {"ref"})


def build_graph(layer: str = "all", attrs: bool = False):
    """
    {nodes, edges} for the graph visualizer.
      nodes: {id, label, layer, type, verdict?, kind?}
      edges: {from, to, label, matchType?, directed}
    `layer="all"` returns everything; a specific prefix returns that vendor's
    nodes/edges plus the ref: concepts they map to (so the ref spine stays visible).
    `attrs=True` also adds datatype-property (literal attribute) nodes/edges -
    off by default since they roughly double node count on a real overlay.
    """
    want = None if layer in (None, "", "all") else layer
    nodes, edges = {}, []

    def add_node(iri, ntype="Class", force=False, kind=None):
        """force=True bypasses the vendor-scope check for a confirmed edge endpoint
        (an operatesOn/mapping target) - a scoped view must show those even when they
        belong to another layer (typically ref:), or the edge that reaches them would
        have to be dropped too. Without force, section 1's blanket class sweep stays
        strictly scoped, so e.g. layer=cm doesn't pull in every unrelated ref: concept.
        kind, when given, is carried onto the node (e.g. "datatype" so the frontend
        renders it as a box instead of a dot - see graph.html's n.kind check)."""
        if not iri or not iri.startswith("http"):
            return
        lay = _graph_layer_of(iri)
        if want and not force and lay != want:
            return
        if iri not in nodes:
            node = {"id": iri, "label": _graph_local(iri), "layer": lay, "type": ntype}
            if kind:
                node["kind"] = kind
            nodes[iri] = node

    # 1) core classes: ref: Data Objects / Capabilities, plus every owl:Class
    #    (covers vendor DataObject/ExternalReference/PortalOperation/etc. subclasses).
    #    Strictly scoped - a ref: concept only reappears via sections 3/5 below if this
    #    vendor actually connects to it.
    for r in store.sparql("""
        SELECT ?s ?label ?type WHERE {
          { ?s rdfs:subClassOf+ ref:DataObject . BIND('DataObject' AS ?type) }
          UNION { ?s rdfs:subClassOf+ ref:Capability . BIND('Capability' AS ?type) }
          UNION { ?s a owl:Class . BIND('Class' AS ?type) }
          OPTIONAL { ?s rdfs:label ?label }
        }"""):
        s = r.get("s")
        add_node(s, r.get("type", "Class"))
        if s in nodes and r.get("label"):
            nodes[s]["label"] = r["label"]

    # 2) verdict badges - discovered by the "#verdict" suffix, same as store.vendors()
    for r in store.sparql(
            'SELECT ?s ?v WHERE { ?s ?vp ?v . FILTER(STRENDS(STR(?vp), "#verdict")) }'):
        s = r.get("s")
        if s in nodes:
            nodes[s]["verdict"] = r.get("v")

    # 3) structural edges: anything named operatesOn, in any namespace - covers both
    #    ref:Capability->ref:DataObject and every vendor's own verb axis. Scoped by the
    #    SOURCE's layer (ref: sources always shown, so the neutral skeleton stays
    #    visible); force-adds both ends since Operation instances aren't classes and
    #    wouldn't otherwise be on the graph at all.
    for r in store.sparql(
            'SELECT ?s ?p ?o WHERE { ?s ?p ?o . FILTER(STRENDS(STR(?p), "#operatesOn")) }'):
        s, o = r.get("s"), r.get("o")
        if not s or not o:
            continue
        if want and _graph_layer_of(s) not in (want, "ref"):
            continue
        add_node(s, force=True)
        add_node(o, force=True)
        edges.append({"from": s, "to": o, "label": "operatesOn", "directed": True})

    # 4) object properties between two classes already on the graph
    for r in store.sparql("""
        SELECT ?p ?dom ?ran ?label WHERE {
          ?p a owl:ObjectProperty ; rdfs:domain ?dom ; rdfs:range ?ran .
          OPTIONAL { ?p rdfs:label ?label }
        }"""):
        dom, ran = r.get("dom"), r.get("ran")
        if dom in nodes and ran in nodes:
            edges.append({"from": dom, "to": ran,
                          "label": r.get("label") or _graph_local(r["p"]), "directed": True})

    # 5) mapping edges - the comparison story. Direct skos:*Match triples from a
    #    vendor term to a ref: concept, exactly as every queries/*.rq file reads them.
    #    force-add: a mapped ATTRIBUTE (owl:DatatypeProperty) isn't a class, and the
    #    ref: target is out-of-scope by layer - section 1 alone would miss both.
    for r in store.sparql("""
        SELECT ?src ?match ?tgt WHERE {
          ?src ?match ?tgt .
          FILTER(STRSTARTS(STR(?match), STR(skos:)) && CONTAINS(STR(?match), "Match"))
        }"""):
        src, tgt, match = r.get("src"), r.get("tgt"), r.get("match")
        if not src or not tgt:
            continue
        if want and _graph_layer_of(src) != want:
            continue
        add_node(src, force=True); add_node(tgt, force=True)
        mt = _graph_local(match) if match else ""
        edges.append({"from": src, "to": tgt, "label": mt, "matchType": mt, "directed": True})

    # 6) datatype properties (literal attributes: stepName, processingType, ...) -
    #    opt-in via attrs=True. Domain must already be a node in this view (respects
    #    the layer filter the same way section 4's object-property block does); the
    #    property node itself is force-added since relevance is already established
    #    via its domain, exactly like the mapping-edge endpoints in section 5.
    if attrs:
        for r in store.sparql("""
            SELECT ?p ?dom ?label ?range WHERE {
              ?p a owl:DatatypeProperty ; rdfs:domain ?dom .
              OPTIONAL { ?p rdfs:label ?label }
              OPTIONAL { ?p rdfs:range ?range }
            }"""):
            p, dom = r.get("p"), r.get("dom")
            if not p or not dom or dom not in nodes:
                continue
            add_node(p, ntype="DatatypeProperty", force=True, kind="datatype")
            if r.get("label"):
                nodes[p]["label"] = r["label"]
            rng = r.get("range")
            edges.append({"from": dom, "to": p,
                          "label": _graph_local(rng) if rng else "literal", "directed": True})

    seen, uniq = set(), []
    for e in edges:
        k = (e["from"], e["to"], e.get("matchType", e.get("label", "")))
        if k not in seen:
            seen.add(k)
            uniq.append(e)

    return {"nodes": list(nodes.values()), "edges": uniq}


@app.get("/api/graph/layers")
def api_graph_layers():
    """Layer prefixes present in the graph (ref first), for a UI layer filter."""
    return discover_graph_layers()


@app.get("/api/graph")
def api_graph(layer: str = "all", attrs: bool = False):
    """Ontology graph shaped for a visualizer: {nodes, edges}. layer='all' or a vendor
    prefix; attrs=true adds datatype-property (literal attribute) nodes/edges."""
    return build_graph(layer, attrs)


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
# Intake (Agent 6): upload HTML -> AI-mapped draft overlay .xlsx for review.
# Merged from api_intake_endpoint.py. mapper/draft_writer/html_reader are
# imported LAZILY inside the route, not at module top-level like chatbot/
# reasoner - mapper.py imports the `anthropic` SDK at its own top level, and
# html_reader.py needs `beautifulsoup4` (not currently in requirements.txt).
# A deployment missing either package must not take the whole app down; it
# should only break /api/intake, with a clear message, exactly like
# chatbot._gemini_client() already defers `import google.generativeai` so a
# missing Gemini SDK doesn't block startup either.
# Nothing here loads TTL or writes to the graph - it produces a DRAFT only,
# for human review before any TTL generation.
# ---------------------------------------------------------------------------
TEMPLATE_PATH = BASE_DIR / "MES-Overlay-TEMPLATE.xlsx"


def load_ref_vocab():
    """
    Pull the frozen reference vocabulary from the LIVE store so the mapper maps
    ONLY to real ref: terms. Returns
      {"data_objects":[...], "capabilities":[...], "operations":[...]}
    Uses rdfs:label where present, else the local name.

    Goes through store.sparql() - the same pattern graph_engine.py uses (it
    never unwraps raw pyoxigraph QuerySolutions by hand, just calls
    store.sparql() and reads plain dicts). store.sparql() already prepends the
    ref:/rdfs:/owl: prefix block and runs with use_default_graph_as_union=True
    (see store.py), so no manual PREFIX lines or raw store.query() are needed.

    NOTE: ref: is a two-axis model (DataObject x Capability only - see
    mes-neutral-reference-4.ttl's own dct:description). There is no
    ref:Operation class, so "operations" will always come back empty;
    Operations only exist per-vendor (generate_ttl.py's PortalOperation verb
    axis). mapper.py's NODE_TYPES still allows the model to classify a concept
    as "Operation" - it will just never get a maps_to_neutral value for that
    type, which is a real, unavoidable gap the human reviewer needs to know
    about, not a bug in this wiring.
    """
    def names(subclass_of):
        rows = store.sparql(f"""
            SELECT ?s ?l WHERE {{
              ?s rdfs:subClassOf+ {subclass_of} .
              OPTIONAL {{ ?s rdfs:label ?l }}
            }}""")
        out = []
        for r in rows:
            label, iri = r.get("l"), r.get("s")
            if label:
                out.append(label)
            elif iri:
                out.append(iri.rsplit("#", 1)[-1])
        return sorted(set(out))

    return {
        "data_objects": names("ref:DataObject"),
        "capabilities": names("ref:Capability"),
        "operations": names("ref:Operation"),
    }


@app.post("/api/intake")
async def api_intake(
    files: list[UploadFile] = File(...),
    vendor: str = Form(...),                          # layer token, e.g. "opc"
    model: str = Form(default=None),                  # None -> mapper.MODEL
    x_anthropic_api_key: str = Header(default=None),
):
    """Upload HTML or Word (.docx) -> draft overlay .xlsx (for human review)."""
    try:
        import mapper
        import draft_writer
        from html_reader import read_html_file
        from docx_reader import read_docx_file
        from pdf_reader import read_pdf_file
    except ImportError as e:
        raise HTTPException(status_code=503,
                            detail=f"Intake is not available: {e}. "
                                   f"pip install anthropic beautifulsoup4 python-docx pdfplumber.")

    api_key = x_anthropic_api_key or os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise HTTPException(status_code=400,
                            detail="No Anthropic API key. Add it in the sidebar "
                                   "(an API key from console.anthropic.com, not a "
                                   "Claude Pro login).")
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded.")
    if len(files) > 40:
        raise HTTPException(status_code=400,
                            detail=f"{len(files)} files is a lot for one batch. "
                                   f"Upload up to ~40 pages (one section) at a time.")
    if not TEMPLATE_PATH.exists():
        raise HTTPException(status_code=503,
                            detail=f"{TEMPLATE_PATH.name} is not on the server - "
                                   f"add it at the project root before intake can run.")

    model = model or mapper.MODEL

    # 1) read uploaded HTML/Word/PDF -> concept records (deterministic)
    records = []
    with tempfile.TemporaryDirectory() as td:
        tdp = Path(td)
        for uf in files:
            name_lower = uf.filename.lower()
            if name_lower.startswith("~$"):
                continue   # Word lock file
            is_html = name_lower.endswith((".html", ".htm"))
            is_docx = name_lower.endswith(".docx")
            is_pdf = name_lower.endswith(".pdf")
            if not (is_html or is_docx or is_pdf):
                continue
            dest = tdp / Path(uf.filename).name
            dest.write_bytes(await uf.read())
            def _read_error(reason):
                return {"concept_name": uf.filename, "source": uf.filename,
                        "section": uf.filename, "prose": "",
                        "sections": [], "lists": [], "tables": [],
                        "_read_error": reason}
            try:
                if is_html:
                    rec = read_html_file(dest, root=tdp)
                elif is_docx:
                    rec = read_docx_file(dest)
                else:
                    rec = read_pdf_file(dest)
                    if not (rec.get("prose") or "").strip():
                        rec = _read_error("no text layer (scanned?) — needs OCR")
                records.append(rec)
            except Exception as e:
                # skip unreadable file, keep going
                records.append(_read_error(str(e)))

    bad = [r for r in records if r.get("_read_error")]
    good = [r for r in records if not r.get("_read_error") and r.get("concept_name")]
    if not good:
        raise HTTPException(status_code=400, detail="No readable concepts found in the upload.")

    # 2) ref: vocabulary from the live store (real mapping targets)
    ref_vocab = load_ref_vocab()
    if not any(ref_vocab.values()):
        raise HTTPException(status_code=500,
                            detail="Reference vocabulary is empty - is the ref: "
                                   "layer loaded in the store?")

    # 3) map (Claude) -> draft rows
    result = mapper.map_records(good, ref_vocab, vendor, api_key=api_key, model=model)

    # 4) write the draft workbook
    out_dir = Path(tempfile.mkdtemp())
    out_path = out_dir / f"MES-{vendor.upper()}-Overlay-DRAFT.xlsx"
    draft_writer.write_draft(result["rows"], vendor=vendor,
                             template_path=str(TEMPLATE_PATH),
                             out_path=str(out_path))

    # 5) return the draft for download. Failures are surfaced in a header so the
    #    UI can warn "N concepts need manual capture" without blocking download.
    n_ok = len(result["rows"])
    n_fail = len(result["failures"]) + len(bad)   # + files dropped before mapping (unreadable, scanned PDFs, ...)
    return FileResponse(
        str(out_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=out_path.name,
        headers={
            "X-Intake-Mapped": str(n_ok),
            "X-Intake-Failed": str(n_fail),
            "X-Intake-Vendor": vendor,
        },
    )


# ---------------------------------------------------------------------------
# Frontend (served last so /api/* wins)
# ---------------------------------------------------------------------------
@app.get("/")
def index():
    return FileResponse(os.path.join(WEB_DIR, "index.html"))

if os.path.isdir(WEB_DIR):
    app.mount("/", StaticFiles(directory=WEB_DIR), name="web")
