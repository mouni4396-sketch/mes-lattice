"""
reasoner.py — Ontology Reasoner (RETRIEVED-from-graph tier).

Answers "how does X work / how is X related to Y" by tracing the graph, then
narrating ONLY the traced facts. Distinct from the computed comparison tier.

Fixes proven against the CM overlay:
  1. VENDOR-FIRST RESOLUTION — resolve cm:Step, not the ref: stub; Data Objects
     preferred over Operations and over Context tables.
  2. MULTI-HOP PATHS — direct + 2-hop object-property paths.
  3. CONTEXT-RESOLUTION BRIDGE (structural signature). In CM, Step/Resource connect
     to most objects by *context resolution*, not static association. A context table
     is identified STRUCTURALLY — a class carrying object properties whose label
     starts with "key" (keyStep, keyProduct, ...). It is NOT reliable to require the
     table to be rdfs:subClassOf SmartTableContext: in the data these tables are
     rdfs:subClassOf cm:DataObject. So the bridge finds any class that has an object
     property reaching A and one reaching B, where at least one side is a key* field —
     that is a context table connecting the two by runtime resolution.

The LLM only extracts concept phrases and narrates. It never writes SPARQL and never
invents structure.
"""

import json
import re
from chatbot import _llm_json, _llm_text


# ---------------------------------------------------------------------------
# Concept extraction (LLM) + robust fallback.
# ---------------------------------------------------------------------------
def extract_concepts(question, api_key=None):
    sys = ("Extract the ontology concept name(s) the user asks about. Return JSON "
           '{"concepts": ["..."]}. One for "how does X work"; two for "how is X '
           'related to Y". Noun phrases only, drop any vendor prefix like cm:/ref:.')
    try:
        ex = _llm_json(sys, question, api_key)
        concepts = [c.strip() for c in ex.get("concepts", []) if c.strip()]
        if concepts:
            return concepts
    except Exception:
        pass
    q = question.strip().rstrip("?")
    for sep in [" related to ", " connected to ", " linked to ", " and ", " vs ", " versus "]:
        if sep in q.lower():
            idx = q.lower().index(sep)
            left = re.sub(r"^(how is|how are|how does|what is|explain|show me|tell me about)\s+",
                          "", q[:idx], flags=re.I).strip()
            right = q[idx + len(sep):].strip()
            return [c for c in (left, right) if c]
    single = re.sub(r"^(how does|how is|what is|explain|show me|tell me about|describe)\s+",
                    "", q, flags=re.I).strip()
    single = re.sub(r"\s+(work|function|behave)$", "", single, flags=re.I).strip()
    return [single] if single else []


# ---------------------------------------------------------------------------
# Concept resolution — vendor-first, Data-Objects before Operations/Contexts.
# ---------------------------------------------------------------------------
def resolve_concept(store, phrase, prefer_vendor=None):
    lit = '"' + phrase.replace('"', '') + '"'
    rows = store.sparql(f"""
    SELECT ?c ?label WHERE {{
      ?c rdfs:label ?label .
      FILTER( CONTAINS(LCASE(STR(?label)), LCASE({lit})) )
    }} LIMIT 200""")
    if not rows:
        return None
    pv_ns = store.layers.get(prefer_vendor, {}).get("ns") if prefer_vendor else None
    ref_ns = store.layers.get("ref", {}).get("ns")
    target = phrase.strip().lower()

    def ns_of(iri):
        return iri.rsplit("#", 1)[0] + "#" if "#" in iri else iri

    def score(r):
        iri, label, ns = r["c"], r.get("label", "").strip().lower(), ns_of(r["c"])
        exact = (label == target)
        is_op = "operation" in iri.lower() or "manage" in iri.lower()
        is_ctx = "context" in iri.lower()
        if exact and ns == pv_ns and not is_ctx:     return 0
        if exact and ns != ref_ns and not is_ctx:    return 1
        if exact and ns == ref_ns:                    return 2
        if not is_op and not is_ctx and ns == pv_ns:  return 3
        if not is_op and not is_ctx:                  return 4
        return 6

    best = min(rows, key=score)
    ns = ns_of(best["c"])
    return {"iri": best["c"], "label": best.get("label", ""),
            "short": store.pfx(best["c"]),
            "prefix": store.ns_to_prefix.get(ns, ns)}


# ---------------------------------------------------------------------------
# Trace one concept.
# ---------------------------------------------------------------------------
def trace_concept(store, iri):
    facts = {"concept": store.pfx(iri), "iri": iri}
    facts["mappings"] = [
        {"dir": r.get("dir"), "other": store.pfx(r["other"]),
         "label": r.get("otherLabel", ""), "match": store.pfx(r.get("match", ""))}
        for r in store.sparql("""
        SELECT ?dir ?other ?otherLabel ?match WHERE {
          { <%IRI%> ?match ?other . BIND("out" AS ?dir) }
          UNION { ?other ?match <%IRI%> . BIND("in" AS ?dir) }
          FILTER( STRSTARTS(STR(?match), STR(skos:)) )
          OPTIONAL { ?other rdfs:label ?otherLabel } }""", {"IRI": iri})]
    facts["attributes"] = [
        {"prop": store.pfx(r["p"]), "label": r.get("label", ""),
         "range": store.pfx(r.get("range", ""))}
        for r in store.sparql("""
        SELECT ?p ?label ?range WHERE {
          ?p rdfs:domain <%IRI%> ; a owl:DatatypeProperty .
          OPTIONAL { ?p rdfs:label ?label } OPTIONAL { ?p rdfs:range ?range } }""",
        {"IRI": iri})]
    facts["links_out"] = [
        {"prop": store.pfx(r["p"]), "label": r.get("plabel", ""),
         "target": store.pfx(r["target"]), "targetLabel": r.get("tlabel", "")}
        for r in store.sparql("""
        SELECT ?p ?plabel ?target ?tlabel WHERE {
          ?p rdfs:domain <%IRI%> ; a owl:ObjectProperty ; rdfs:range ?target .
          OPTIONAL { ?p rdfs:label ?plabel } OPTIONAL { ?target rdfs:label ?tlabel } }""",
        {"IRI": iri})]
    facts["links_in"] = [
        {"prop": store.pfx(r["p"]), "label": r.get("plabel", ""),
         "source": store.pfx(r["src"]), "sourceLabel": r.get("slabel", "")}
        for r in store.sparql("""
        SELECT ?p ?plabel ?src ?slabel WHERE {
          ?p rdfs:range <%IRI%> ; a owl:ObjectProperty ; rdfs:domain ?src .
          OPTIONAL { ?p rdfs:label ?plabel } OPTIONAL { ?src rdfs:label ?slabel } }""",
        {"IRI": iri})]
    return facts


# ---------------------------------------------------------------------------
# Direct + 2-hop object-property paths.
# ---------------------------------------------------------------------------
def find_paths(store, iri_a, iri_b):
    direct = store.sparql("""
    SELECT ?p ?plabel ?dir WHERE {
      { ?p rdfs:domain <%A%> ; rdfs:range <%B%> . BIND("a_to_b" AS ?dir) }
      UNION { ?p rdfs:domain <%B%> ; rdfs:range <%A%> . BIND("b_to_a" AS ?dir) }
      ?p a owl:ObjectProperty . OPTIONAL { ?p rdfs:label ?plabel } }""",
    {"A": iri_a, "B": iri_b})
    if direct:
        return {"kind": "direct",
                "hops": [{"prop": store.pfx(r["p"]), "label": r.get("plabel", ""),
                          "dir": r.get("dir")} for r in direct]}
    two = store.sparql("""
    SELECT ?p1 ?mid ?midLabel ?p2 WHERE {
      ?p1 rdfs:domain <%A%> ; rdfs:range ?mid ; a owl:ObjectProperty .
      ?p2 rdfs:domain ?mid ; rdfs:range <%B%> ; a owl:ObjectProperty .
      OPTIONAL { ?mid rdfs:label ?midLabel } } LIMIT 10""",
    {"A": iri_a, "B": iri_b})
    if two:
        return {"kind": "two_hop",
                "hops": [{"via": store.pfx(r["mid"]), "label": r.get("midLabel", ""),
                          "p1": store.pfx(r["p1"]), "p2": store.pfx(r["p2"])} for r in two]}
    return {"kind": "none", "hops": []}


# ---------------------------------------------------------------------------
# Context-resolution bridge (structural).
# A "context table" is any class C that has an object property to A and one to B.
# We do NOT require C rdfs:subClassOf SmartTableContext (in the data these tables are
# subclasses of cm:DataObject). Instead we surface, for each connecting property,
# whether it is a KEY field (label starts 'key') or a RESOLVED value — which is what
# makes this a runtime context resolution rather than a static link.
# ---------------------------------------------------------------------------
def context_bridge(store, iri_a, iri_b):
    rows = store.sparql("""
    SELECT ?ctx ?ctxLabel ?comment ?pa ?paLabel ?pb ?pbLabel WHERE {
      ?pa rdfs:domain ?ctx ; a owl:ObjectProperty ; rdfs:range <%A%> .
      ?pb rdfs:domain ?ctx ; a owl:ObjectProperty ; rdfs:range <%B%> .
      OPTIONAL { ?ctx rdfs:label ?ctxLabel }
      OPTIONAL { ?ctx rdfs:comment ?comment }
      OPTIONAL { ?pa rdfs:label ?paLabel }
      OPTIONAL { ?pb rdfs:label ?pbLabel }
    } LIMIT 40""", {"A": iri_a, "B": iri_b})

    seen, out = set(), []
    for r in rows:
        c = r["ctx"]
        if c == iri_a or c == iri_b:      # a concept linking to itself isn't a bridge
            continue
        key = (c, r["pa"], r["pb"])
        if key in seen:
            continue
        seen.add(key)
        pa_l = r.get("paLabel", "") or ""
        pb_l = r.get("pbLabel", "") or ""
        looks_context = ("context" in c.lower()
                         or pa_l.lower().startswith("key") or pb_l.lower().startswith("key"))
        out.append({
            "context": store.pfx(c), "label": r.get("ctxLabel", ""),
            "is_context_table": looks_context,
            "reachesA": {"prop": store.pfx(r["pa"]), "label": pa_l,
                         "is_key": pa_l.lower().startswith("key")},
            "reachesB": {"prop": store.pfx(r["pb"]), "label": pb_l,
                         "is_key": pb_l.lower().startswith("key")},
            "comment": r.get("comment", ""),
        })
    # rank real context tables first
    out.sort(key=lambda h: (not h["is_context_table"]))
    return out


# fallback for a single concept: which classes key on it (context tables)?
def contexts_keying_on(store, iri):
    rows = store.sparql("""
    SELECT ?ctx ?ctxLabel ?p ?pLabel WHERE {
      ?p rdfs:domain ?ctx ; a owl:ObjectProperty ; rdfs:range <%IRI%> .
      OPTIONAL { ?ctx rdfs:label ?ctxLabel } OPTIONAL { ?p rdfs:label ?pLabel }
    } LIMIT 60""", {"IRI": iri})
    out = []
    for r in rows:
        pl = r.get("pLabel", "") or ""
        if pl.lower().startswith("key") or "context" in r["ctx"].lower():
            out.append({"context": store.pfx(r["ctx"]), "label": r.get("ctxLabel", ""),
                        "via": store.pfx(r["p"]), "viaLabel": pl})
    return out


# ---------------------------------------------------------------------------
# Narration.
# ---------------------------------------------------------------------------
REASONER_SYSTEM = """You explain how concepts in a semiconductor MES ontology work
and relate, using ONLY the traced graph facts provided.

Rules:
1. Use only the provided facts. Never invent a relationship, attribute, or concept.
2. Explain plainly what each concept is, maps to, holds, and links to.
3. For a two-concept question, report HOW they connect, in this priority:
   - CONTEXT RESOLUTION (field 'context_connection' with is_context_table true): this
     is the REAL answer when present. It means the two concepts are NOT statically
     linked but connected by a runtime context table that keys on both (properties
     labelled keyStep, keyProduct, etc.) and resolves one from the other. Name the
     context class and the two connecting properties, and say it is runtime resolution
     keyed by fields like Step/Product/Material. If the context's comment names the
     ref: equivalent (e.g. 'ref:DataCollectionPlan collectsAtStep ref:Step') or a
     competitor differentiator, state it.
   - else a direct object-property link, or a 2-hop path, or
   - genuinely unconnected in the modelled graph.
4. If facts are empty, say the concept wasn't found - do not guess.
5. This is retrieved structure, not a comparison count."""


def narrate(question, facts, api_key=None):
    user = (f"QUESTION:\n{question}\n\nTRACED GRAPH FACTS:\n"
            f"{json.dumps(facts, indent=2)[:12000]}\n\nExplain using only these facts.")
    try:
        return _llm_text(REASONER_SYSTEM, user, api_key)
    except Exception as e:
        return f"(reasoner narration error: {e})"


# ---------------------------------------------------------------------------
# Entry point.
# ---------------------------------------------------------------------------
def answer(question, store, api_key=None):
    vendors = [p for p in store.layers if p != "ref"]
    prefer = None
    ql = question.lower()
    for p in vendors:
        if p in ql or f"{p}:" in ql:
            prefer = p
            break
    if prefer is None and vendors:
        prefer = vendors[0]

    concepts = extract_concepts(question, api_key)
    if not concepts:
        return {"answer": "I couldn't identify a concept to trace in that question.",
                "tier": "retrieved", "facts": None}

    resolved = []
    for phrase in concepts[:2]:
        r = resolve_concept(store, phrase, prefer_vendor=prefer)
        if r:
            resolved.append(r)
    if not resolved:
        return {"answer": f"I couldn't find '{', '.join(concepts)}' in the loaded ontology.",
                "tier": "retrieved", "facts": None}

    if len(resolved) == 1:
        facts = trace_concept(store, resolved[0]["iri"])
        facts["contexts_keying_on_it"] = contexts_keying_on(store, resolved[0]["iri"])
    else:
        a, b = resolved[0]["iri"], resolved[1]["iri"]
        facts = {
            "a": resolved[0]["short"], "b": resolved[1]["short"],
            "context_connection": context_bridge(store, a, b),
            "path": find_paths(store, a, b),
            "trace_a": trace_concept(store, a),
            "trace_b": trace_concept(store, b),
        }

    prose = narrate(question, facts, api_key)
    return {"answer": prose, "tier": "retrieved",
            "concepts": [r["short"] for r in resolved], "facts": facts}
