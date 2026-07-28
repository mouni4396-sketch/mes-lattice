"""
chatbot.py — natural-language front door to the computed tier.

Flow (this is the whole trick):
  1. ROUTE   — the LLM reads the question and the query catalog, and returns JSON:
               which of the 8 vetted queries to run + parameter values. It does
               NOT write SPARQL. If nothing fits, it says so.
  2. RUN     — the chosen vetted query runs via the GraphEngine (exact, computed).
  3. NARRATE — the LLM writes prose using ONLY the returned rows. It may not add,
               infer, or round any number that isn't in the rows.

Provider is swappable with ONE setting. Set PROVIDER below (or the LLM_PROVIDER
env var) to "gemini" or "deepseek". Everything provider-specific lives in the
plumbing section; the router/narrator logic never changes.

Keys (set the one you use as an environment variable):
  Gemini   -> GEMINI_API_KEY
  DeepSeek -> DEEPSEEK_API_KEY
"""

import os
import json
import re

# ---------------------------------------------------------------------------
# PROVIDER SWITCH — the only line you change to swap LLMs.
# Overridable at runtime with the LLM_PROVIDER environment variable.
# ---------------------------------------------------------------------------
PROVIDER = os.environ.get("LLM_PROVIDER", "gemini")   # "gemini" or "deepseek"

GEMINI_MODEL   = "gemini-flash-latest"   # alias -> newest stable Flash
DEEPSEEK_MODEL = "deepseek-v4-flash"     # current cheap DeepSeek model


# ---------------------------------------------------------------------------
# Provider plumbing — the only provider-specific code.
# Each provider exposes _<name>_json / _<name>_text; the dispatchers pick one.
# ---------------------------------------------------------------------------

# ----- Gemini -----
def _gemini_client(api_key=None):
    import google.generativeai as genai
    key = api_key or os.environ.get("GEMINI_API_KEY", "")
    if not key:
        raise RuntimeError("PROVIDER is 'gemini' but GEMINI_API_KEY is not set.")
    genai.configure(api_key=key)
    return genai


def _gemini_json(system, user, api_key=None):
    genai = _gemini_client(api_key)
    model = genai.GenerativeModel(GEMINI_MODEL, system_instruction=system)
    text = model.generate_content(user).text.strip()
    return _parse_json(text)


def _gemini_text(system, user, api_key=None):
    genai = _gemini_client(api_key)
    model = genai.GenerativeModel(GEMINI_MODEL, system_instruction=system)
    return model.generate_content(user).text.strip()


# ----- DeepSeek (OpenAI-compatible endpoint) -----
def _deepseek_client():
    from openai import OpenAI
    key = os.environ.get("DEEPSEEK_API_KEY", "")
    if not key:
        raise RuntimeError("PROVIDER is 'deepseek' but DEEPSEEK_API_KEY is not set.")
    return OpenAI(api_key=key, base_url="https://api.deepseek.com")


def _deepseek_json(system, user):
    client = _deepseek_client()
    resp = client.chat.completions.create(
        model=DEEPSEEK_MODEL,
        messages=[{"role": "system", "content": system},
                  {"role": "user", "content": user}],
        response_format={"type": "json_object"},   # DeepSeek supports JSON mode
        temperature=0,
    )
    return _parse_json(resp.choices[0].message.content)


def _deepseek_text(system, user):
    client = _deepseek_client()
    resp = client.chat.completions.create(
        model=DEEPSEEK_MODEL,
        messages=[{"role": "system", "content": system},
                  {"role": "user", "content": user}],
        temperature=0,
    )
    return resp.choices[0].message.content.strip()


# ----- shared helpers + dispatchers -----
def _parse_json(text):
    """Strip ```json fences if present, then parse."""
    text = text.strip()
    text = re.sub(r"^```(json)?|```$", "", text, flags=re.MULTILINE).strip()
    return json.loads(text)


def _llm_json(system, user, api_key=None):
    if PROVIDER == "gemini":
        return _gemini_json(system, user, api_key)
    if PROVIDER == "deepseek":
        return _deepseek_json(system, user)
    raise RuntimeError(f"Unknown PROVIDER '{PROVIDER}'. Use 'gemini' or 'deepseek'.")


def _llm_text(system, user, api_key=None):
    if PROVIDER == "gemini":
        return _gemini_text(system, user, api_key)
    if PROVIDER == "deepseek":
        return _deepseek_text(system, user)
    raise RuntimeError(f"Unknown PROVIDER '{PROVIDER}'. Use 'gemini' or 'deepseek'.")


# ---------------------------------------------------------------------------
# ROUTER
# ---------------------------------------------------------------------------
ROUTER_SYSTEM = """You are the router for a semiconductor MES comparison tool.
You do NOT answer questions and you do NOT write SPARQL. Your only job is to choose
ONE vetted query from the catalog and supply its parameters, based on the user's
question. Respond with a single JSON object and nothing else.

JSON shape:
  {"query": "<query_name or null>",
   "params": {"PARAM": "value", ...},
   "reason": "<one short sentence>"}

Rules:
- Choose the single best-fitting query. If none fit, set "query" to null.
- Parameter values for VENDOR / FROM / TO must be one of the vendor namespace IRIs
  given in the catalog (use the full IRI, not the prefix).
- REF_CONCEPT must be a full ref: IRI. If the user names a concept but you cannot
  form its exact IRI, set "query" to null and explain in "reason".
- Never invent a query name that is not in the catalog."""


def route(question, catalog, vendors, api_key=None):
    """
    catalog: list of {"name":..., "params":[...], "desc":...}
    vendors: list of {"prefix":..., "namespace":...}
    Returns {"query":..., "params":{...}, "reason":...}
    """
    cat_lines = []
    for q in catalog:
        p = ", ".join(q["params"]) if q["params"] else "none"
        cat_lines.append(f'- {q["name"]}  (params: {p})  — {q.get("desc","")}')
    vend_lines = [f'- {v["prefix"]} = {v["namespace"]}' for v in vendors]

    user = (
        f"QUESTION:\n{question}\n\n"
        f"QUERY CATALOG:\n" + "\n".join(cat_lines) + "\n\n"
        f"LOADED VENDORS (use these exact namespace IRIs for VENDOR/FROM/TO):\n"
        + "\n".join(vend_lines) + "\n\n"
        "Return the routing JSON now."
    )
    try:
        return _llm_json(ROUTER_SYSTEM, user, api_key)
    except Exception as e:
        return {"query": None, "params": {}, "reason": f"router error: {e}"}


# ---------------------------------------------------------------------------
# NARRATOR
# ---------------------------------------------------------------------------
NARRATOR_SYSTEM = """You explain the result of ONE computed MES comparison query.

NON-NEGOTIABLE grounding:
1. Use ONLY the rows provided. Do not add concepts, counts, vendors, or verdicts
   that are not in the rows.
2. Never invent, infer, estimate, or round a number. Every count you state must be
   literally derivable from the rows given.
3. If the rows are empty, say so plainly and explain what an empty result means for
   THIS query (e.g. for a gap query, empty means full coverage; for convergence,
   empty means no concept is shared by all vendors).
4. Lead with the direct answer. Keep it tight. Use the verdict meanings:
   covered/partial = the vendor maps to the neutral reference; vendor-extension =
   vendor-specific with no neutral equivalent.
5. This is the COMPUTED tier — the numbers are exact. Do not hedge them with
   'approximately' or 'around'."""


def narrate(question, query_name, rows, api_key=None):
    row_json = json.dumps(rows[:200], indent=2)  # cap payload
    note = "" if len(rows) <= 200 else f"\n(NOTE: {len(rows)} rows total; first 200 shown. The count to state is {len(rows)}.)"
    user = (
        f"USER QUESTION:\n{question}\n\n"
        f"QUERY RUN: {query_name}\n"
        f"ROWS RETURNED ({len(rows)} total):\n{row_json}{note}\n\n"
        "Explain the answer using only these rows."
    )
    try:
        return _llm_text(NARRATOR_SYSTEM, user, api_key)
    except Exception as e:
        return f"(narration error: {e})"


# ---------------------------------------------------------------------------
# ORCHESTRATION — the single entry point the API calls
# ---------------------------------------------------------------------------
def answer(question, engine, store, api_key=None):
    """
    Full pipeline: route -> run vetted query -> narrate.
    Returns a dict the API serialises. Always reports which query ran (traceable).
    api_key: optional per-request Gemini key; falls back to GEMINI_API_KEY if None.
    """
    catalog = []
    for name in sorted(engine.queries):
        q = engine.queries[name]
        catalog.append({"name": name, "params": q["params"],
                        "desc": _DESC.get(name, "")})
    vendors = [{"prefix": p, "namespace": info["ns"]}
               for p, info in store.layers.items() if p != "ref"]

    routing = route(question, catalog, vendors, api_key)
    qname = routing.get("query")

    if not qname or qname not in engine.queries:
        return {
            "answer": ("I couldn't map that to one of the vetted comparison queries. "
                       + routing.get("reason", "")),
            "query": None, "params": {}, "rows": [], "tier": "computed",
            "routing": routing,
        }

    params = routing.get("params", {}) or {}
    try:
        rows = engine.run(qname, params)
    except Exception as e:
        return {"answer": f"That query needs different parameters: {e}",
                "query": qname, "params": params, "rows": [], "tier": "computed",
                "routing": routing}

    pretty = [{k: store.pfx(v) if str(v).startswith("http") else v
               for k, v in r.items()} for r in rows]
    prose = narrate(question, qname, pretty, api_key)

    return {"answer": prose, "query": qname, "params": params,
            "rows": pretty, "count": len(pretty), "tier": "computed",
            "routing": routing}


# descriptions handed to the router so it can choose well
_DESC = {
    "q1_coverage_by_dataobject": "For every reference data object, which vendors map to it and how.",
    "q2_coverage_by_capability": "Coverage rolled up to the capability axis.",
    "q3_gap_by_vendor": "Reference objects a given vendor does not map at all. Needs VENDOR.",
    "q4_convergence": "Reference concepts every loaded vendor closeMatches.",
    "q5_migration_diff": "Concepts FROM-vendor covers that TO-vendor does not. Needs FROM and TO.",
    "q6_vendor_extension_surface": "A vendor's proprietary concepts with no reference mapping. Needs VENDOR.",
    "q7_drilldown": "Every vendor mapping row for one reference concept. Needs REF_CONCEPT (full ref: IRI).",
    "q8_completeness_guard": "How much of the reference each overlay has engaged with.",
}
