"""
mapper.py  —  Agent 6 (Intake) stage 2: the AI mapper.

This is the ONE place AI runs in the intake pipeline. It takes concept records
(from readers/html_reader.py etc.) and drafts overlay rows that match
MES-Overlay-TEMPLATE.xlsx, mapping each vendor concept to the FROZEN ref:
vocabulary via SKOS match types, with per-row confidence + source.

Model: Anthropic Claude (Sonnet by default; swap MODEL to an Opus string for
maximum mapping judgment). Uses a user-supplied Anthropic API key, mirroring
the Gemini X-User-API-Key pattern already in the app.

IMPORTANT: "Claude Pro" (the claude.ai subscription) does NOT grant API access.
Users need an Anthropic API key from console.anthropic.com (separate, token-
billed). The app should label the field "Anthropic API key" accordingly.

Design guarantees:
  - The model chooses neutral targets ONLY from the ref: vocabulary you pass in
    (no invented terms) -> this is what makes ONE mapper work for ALL vendors.
  - Output is validated JSON matching the template columns; malformed output is
    retried once, then the concept is flagged for manual capture (never silently
    dropped or hallucinated into the graph).
  - The mapper DRAFTS ONLY. A human reviews the resulting .xlsx before any TTL.
"""

import os
import json
from typing import Optional

import anthropic   # pip install anthropic

MODEL = "claude-sonnet-4-6"        # swap to "claude-opus-4-8" for max judgment
MAX_TOKENS = 4096

# controlled vocab (mirrors the template Guide sheet)
MATCH_TYPES = ["closeMatch", "broadMatch", "narrowMatch", "relatedMatch"]
VERDICTS = ["covered", "partial", "vendor-extension", "absent", "needs-verification"]
NODE_TYPES = ["Data Object", "Operation", "Capability"]


SYSTEM_PROMPT = """You map ONE vendor MES concept to a FROZEN neutral reference \
ontology. You are drafting rows a human will review; be accurate and calibrated, \
never confident-but-wrong.

STRICT RULES:
1. For "maps_to_neutral" you MUST choose an EXACT name from the REFERENCE \
VOCABULARY provided in the user message. Never invent a neutral term. If nothing \
fits, set maps_to_neutral to "" (empty) and verdict to "vendor-extension".
2. "match_type" MUST be one of: closeMatch, broadMatch, narrowMatch, relatedMatch.
   - closeMatch: same concept and scope
   - broadMatch: the vendor term is NARROWER than the neutral one
   - narrowMatch: the vendor term is BROADER than the neutral one
   - relatedMatch: associated, but not a subsumption match
3. "verdict" MUST be one of: covered, partial, vendor-extension, absent, \
needs-verification. Use needs-verification when the documentation is ambiguous.
4. "confidence" is a number 0.0-1.0 = YOUR certainty in THIS mapping. Low \
confidence is expected and USEFUL. Do NOT inflate it. A wrong 0.9 is worse than \
an honest 0.5.
5. "type" MUST be one of: Data Object, Operation, Capability.
6. Ground every field in the provided text. If the text does not state something, \
do not assert it. Prefer an empty cell over a guess.
7. Output ONLY valid JSON. No markdown, no commentary, no code fences. A single \
JSON object of the exact shape described. Nothing before or after it."""


USER_TEMPLATE = """REFERENCE VOCABULARY (choose neutral targets ONLY from these exact names):

DATA OBJECTS:
{ref_data_objects}

CAPABILITIES:
{ref_capabilities}

OPERATIONS:
{ref_operations}

---
VENDOR: {vendor}
CONCEPT NAME: {concept_name}
SOURCE: {source}

CONCEPT TEXT (documentation):
{concept_text}

---
Produce a JSON object with this exact shape:

{{
  "concept": {{
    "name": "<vendor concept name>",
    "type": "<Data Object | Operation | Capability>",
    "maps_to_neutral": "<exact ref name or empty>",
    "match_type": "<closeMatch|broadMatch|narrowMatch|relatedMatch or empty>",
    "verdict": "<one of the allowed verdicts>",
    "confidence": <0.0-1.0>,
    "source": "{source}",
    "notes": "<short, optional>"
  }},
  "attributes": [
    {{
      "name": "<attribute/field name>",
      "datatype": "<string|decimal|integer|boolean|dateTime or empty>",
      "value_kind": "<enum|lookup|plain or empty>",
      "maps_to_neutral": "<exact ref attribute or empty>",
      "match_type": "<... or empty>",
      "verdict": "<...>",
      "confidence": <0.0-1.0>,
      "source": "{source}",
      "notes": "<short, optional>"
    }}
  ],
  "relationships": [
    {{
      "name": "<object property name>",
      "target": "<the entity it points to>",
      "maps_to_neutral": "<exact ref property or empty>",
      "match_type": "<... or empty>",
      "verdict": "<...>",
      "confidence": <0.0-1.0>,
      "source": "{source}",
      "notes": "<short, optional>"
    }}
  ]
}}

If the concept has no attributes or no relationships evident in the text, use an \
empty list for that key. Output ONLY the JSON object."""


def _concept_text(record: dict, max_chars: int = 12000) -> str:
    """Flatten a concept record into readable text for the model."""
    parts = []
    for s in record.get("sections", []):
        h = s.get("heading", "")
        t = s.get("text", "")
        if t:
            parts.append(f"## {h}\n{t}")
    if not parts and record.get("prose"):
        parts.append(record["prose"])
    for i, lst in enumerate(record.get("lists", [])):
        if lst:
            parts.append("Constraints/List:\n- " + "\n- ".join(lst))
    for tb in record.get("tables", []):
        if tb.get("headers"):
            parts.append("Table headers: " + " | ".join(tb["headers"]))
    text = "\n\n".join(parts)
    return text[:max_chars]


def _strip_fences(s: str) -> str:
    s = s.strip()
    if s.startswith("```"):
        s = s.split("\n", 1)[1] if "\n" in s else s
        if s.endswith("```"):
            s = s[: -3]
        # drop a leading 'json' language tag if present
        if s.lstrip().startswith("json"):
            s = s.lstrip()[4:]
    return s.strip()


def _validate(obj: dict) -> list:
    """Return a list of problem strings; empty means valid."""
    problems = []
    c = obj.get("concept")
    if not isinstance(c, dict):
        return ["missing 'concept' object"]
    if c.get("type") and c["type"] not in NODE_TYPES:
        problems.append(f"concept.type '{c.get('type')}' not allowed")
    def check_rowset(key):
        for i, r in enumerate(obj.get(key, []) or []):
            mt = r.get("match_type", "")
            if mt and mt not in MATCH_TYPES:
                problems.append(f"{key}[{i}].match_type '{mt}' not allowed")
            vd = r.get("verdict", "")
            if vd and vd not in VERDICTS:
                problems.append(f"{key}[{i}].verdict '{vd}' not allowed")
            cf = r.get("confidence", None)
            if cf is not None and not (isinstance(cf, (int, float)) and 0 <= cf <= 1):
                problems.append(f"{key}[{i}].confidence '{cf}' out of range")
    for k in ("attributes", "relationships"):
        check_rowset(k)
    # concept-level checks
    if c.get("match_type") and c["match_type"] not in MATCH_TYPES:
        problems.append(f"concept.match_type '{c['match_type']}' not allowed")
    if c.get("verdict") and c["verdict"] not in VERDICTS:
        problems.append(f"concept.verdict '{c['verdict']}' not allowed")
    return problems


def map_concept(record: dict,
                ref_vocab: dict,
                vendor: str,
                api_key: Optional[str] = None,
                model: str = MODEL) -> dict:
    """
    Map one concept record -> draft rows (dict). Returns:
      {"ok": True,  "data": <validated obj>, "concept": name}
      {"ok": False, "error": <reason>, "raw": <raw text>, "concept": name}

    ref_vocab = {"data_objects":[...], "capabilities":[...], "operations":[...]}
    """
    key = api_key or os.getenv("ANTHROPIC_API_KEY")
    if not key:
        raise ValueError("No Anthropic API key supplied for the intake mapper.")

    client = anthropic.Anthropic(api_key=key)

    user_msg = USER_TEMPLATE.format(
        ref_data_objects="\n".join(f"- {x}" for x in ref_vocab.get("data_objects", [])),
        ref_capabilities="\n".join(f"- {x}" for x in ref_vocab.get("capabilities", [])),
        ref_operations="\n".join(f"- {x}" for x in ref_vocab.get("operations", [])),
        vendor=vendor,
        concept_name=record.get("concept_name", "?"),
        source=record.get("section") or record.get("source", ""),
        concept_text=_concept_text(record),
    )

    name = record.get("concept_name", "?")
    last_raw = ""
    for attempt in range(2):                 # one retry on malformed JSON
        resp = client.messages.create(
            model=model,
            max_tokens=MAX_TOKENS,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_msg}],
        )
        raw = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
        last_raw = raw
        try:
            obj = json.loads(_strip_fences(raw))
        except json.JSONDecodeError:
            if attempt == 0:
                user_msg += "\n\nREMINDER: Output ONLY the JSON object, no other text."
                continue
            return {"ok": False, "error": "invalid JSON", "raw": last_raw, "concept": name}

        problems = _validate(obj)
        if problems:
            if attempt == 0:
                user_msg += ("\n\nFIX THESE and re-output JSON only: "
                             + "; ".join(problems))
                continue
            return {"ok": False, "error": "; ".join(problems),
                    "raw": last_raw, "concept": name}

        return {"ok": True, "data": obj, "concept": name}

    return {"ok": False, "error": "exhausted retries", "raw": last_raw, "concept": name}


def map_records(records: list,
                ref_vocab: dict,
                vendor: str,
                api_key: Optional[str] = None,
                model: str = MODEL,
                progress=None) -> dict:
    """
    Map many records. Returns {"rows":[...ok data...], "failures":[...]}.
    `progress(i, total, concept, ok)` optional callback for UI/streaming.
    """
    rows, failures = [], []
    total = len(records)
    for i, rec in enumerate(records, 1):
        res = map_concept(rec, ref_vocab, vendor, api_key=api_key, model=model)
        if res["ok"]:
            rows.append(res["data"])
        else:
            failures.append(res)
        if progress:
            progress(i, total, res["concept"], res["ok"])
    return {"rows": rows, "failures": failures}


if __name__ == "__main__":
    print("mapper.py is a library. Wire it via /api/intake or call map_records().")
    print("Needs: pip install anthropic ; an Anthropic API key ; ref_vocab dict.")
