"""
generate_ttl.py — vendor-agnostic Excel -> OWL/TTL generator.

Turns a filled capture workbook into a {vendor}: overlay TTL that imports the frozen
ref: layer and emits ONLY vendor deltas. Excel is master; TTL is generated.

DESIGN RULES (why this file looks the way it does)
  A. Columns are read BY HEADER NAME, never by position. Vendor masters drift
     (Opcenter dropped "Confidence"); positional reads silently shift every later
     column. A missing REQUIRED header fails loudly; optional headers are genuinely
     optional (absent = omit the triple, never shift).
  B. Every literal is quoted and escaped. Confidence is emitted only when present
     AND numeric.
  C. Attribute/property IRIs are auto-qualified by owning class ONLY when the
     hand-authored local name would collide. CM masters already hand-qualify
     (areaName, batchName) and those are preserved verbatim.
  D. The verb-axis meta-class is namespaced so it cannot collide with a vendor noun
     (Camstar has "Operation" as a data object). Build fails if any meta IRI collides
     with an IRI minted from the master.
  E. A "Promotions" sheet (vendor independently confirms a Derived ref: concept) is
     emitted; it used to be silently dropped.
  F. Lookup allowed values and object-property Notes are carried through.
  G. A validation pass runs after generation: re-parses the TTL, checks for duplicate
     IRI definitions, mapping targets outside ref:, ref: IRIs used as subjects, and
     datatype properties whose range is a class. Prints a per-section count table.

Usage:
    python generate_ttl.py MASTER.xlsx --prefix opc --out mes-opc-overlay.ttl
"""

import argparse
import re
import sys
import openpyxl

BASE = "https://ontology.yourorg.com/mes-"
REF_NS = f"{BASE}ref#"

DTYPE = {"string": "xsd:string", "decimal": "xsd:decimal", "integer": "xsd:integer",
         "boolean": "xsd:boolean", "datetime": "xsd:dateTime"}
MATCH_OK = {"closematch", "broadmatch", "narrowmatch", "relatedmatch"}
# canonical SKOS spelling from a lowercased key
MATCH_CANON = {"closematch": "closeMatch", "broadmatch": "broadMatch",
               "narrowmatch": "narrowMatch", "relatedmatch": "relatedMatch"}


# ---------------------------------------------------------------------------
# Header-name column access (fix A)
# ---------------------------------------------------------------------------
class Sheet:
    """Wraps a worksheet with case-insensitive, whitespace-tolerant header lookup."""

    def __init__(self, wb, name, required=(), optional=()):
        if name not in wb.sheetnames:
            raise BuildError(f"Workbook is missing required sheet '{name}'. "
                             f"Sheets present: {wb.sheetnames}")
        self.name = name
        self.ws = wb[name]
        raw = [c.value for c in next(self.ws.iter_rows(min_row=1, max_row=1))]
        self.headers = {}
        for i, h in enumerate(raw):
            if h is None:
                continue
            self.headers[str(h).strip().lower()] = i
        missing = [h for h in required if h.lower() not in self.headers]
        if missing:
            raise BuildError(
                f"Sheet '{name}' is missing required column(s): {missing}. "
                f"Found: {sorted(self.headers)}")
        self.optional = {o.lower() for o in optional}

    def has(self, header):
        return header.strip().lower() in self.headers

    def rows(self):
        for row in self.ws.iter_rows(min_row=2, values_only=True):
            yield RowView(row, self.headers)


class RowView:
    def __init__(self, row, headers):
        self._row, self._h = row, headers

    def get(self, header, default=""):
        i = self._h.get(header.strip().lower())
        if i is None or i >= len(self._row):
            return default
        v = self._row[i]
        return default if v is None else v

    def str(self, header):
        return str(self.get(header, "")).strip()


class BuildError(Exception):
    pass


# ---------------------------------------------------------------------------
# Literal handling (fix B)
# ---------------------------------------------------------------------------
def lit(value):
    """Quote+escape a Turtle string literal."""
    s = str(value).replace("\\", "\\\\").replace('"', '\\"')
    s = s.replace("\r", " ").replace("\n", " ").strip()
    return f'"{s}"'


def as_decimal(value):
    """Return a Turtle decimal if value parses as a number, else None."""
    try:
        f = float(str(value).strip())
    except (TypeError, ValueError):
        return None
    return repr(f) if f != int(f) else str(int(f))


def norm(value):
    """
    Normalise a cell value for MATCHING (not for output): strip whitespace, strip any
    wrapping quote characters, collapse internal whitespace, lowercase.
    Masters sometimes carry values like "'opc'" (literal quotes inside the cell), which
    silently failed an exact comparison and skipped every row.
    """
    s = str(value).strip()
    s = s.strip("'\"\u2018\u2019\u201c\u201d")
    s = re.sub(r"\s+", " ", s)
    return s.strip().lower()


def local(name):
    """Turn a display name or IRI into a valid IRI local part."""
    s = str(name).strip()
    if "#" in s:
        s = s.rsplit("#", 1)[1]
    if ":" in s and " " not in s:
        s = s.split(":", 1)[1]
    if " " in s:
        parts = s.split()
        s = parts[0][:1].lower() + parts[0][1:] + "".join(p[:1].upper() + p[1:] for p in parts[1:])
    return re.sub(r"[^A-Za-z0-9_\-]", "", s)


def cls_local(name):
    """Class-style local name (leading capital)."""
    s = local(name)
    return s[:1].upper() + s[1:] if s else s


def ref_iri(value):
    s = str(value).strip()
    if s.startswith("ref:"):
        return s
    return "ref:" + local(s)


# ---------------------------------------------------------------------------
# Generator
# ---------------------------------------------------------------------------
class Generator:
    def __init__(self, path, prefix):
        self.wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        self.p = prefix
        self.ns = f"{BASE}{prefix}#"
        self.out = []
        self.defined = {}          # local name -> section (duplicate detection, fix C/G)
        self.name_to_iri = {}      # display name (normalised) -> minted class IRI
        self.warnings = []
        self.counts = {}
        # fix D: verb-axis meta-class is namespaced so a vendor noun ("Operation")
        # can never collide with it.
        self.OP_CLASS = "PortalOperation"
        self.meta_iris = {
            "DataObject", "ExternalReference", self.OP_CLASS,
            "operatesOn", "verdict", "sourceDoc", "operationKind", "confidence",
            "confirmsRefConcept", "promotionNote", "promotionEvidence", "UnresolvedReference",
        }

    def w(self, s=""):
        self.out.append(s)

    def define(self, name, section):
        """
        Register an IRI local name.

        Hard failure (fix C/G) for collisions that corrupt meaning: two properties, or
        a property colliding with a class. Multiple rdfs:domain on one property is
        CONJUNCTIVE in RDFS, so a duplicated property IRI silently entails that anything
        carrying it belongs to every domain class at once.

        A class/enum-set name overlap (e.g. a TransferRequirementType data object and a
        TransferRequirementType enum) is recorded as a warning instead: it is a naming
        smell worth surfacing, but the two are different node types and do not corrupt
        entailment.
        """
        prior = self.defined.get(name)
        if prior:
            typed = {"1. Data Objects": "class", "4. Enums & Lookups": "class",
                     "5. Operations": "class"}
            kind_a, kind_b = typed.get(prior, "property"), typed.get(section, "property")
            if kind_a == "class" and kind_b == "class":
                self.warnings.append(
                    f"name reused across sections: {self.p}:{name} "
                    f"({prior} and {section})")
                return
            raise BuildError(
                f"IRI {self.p}:{name} is defined twice "
                f"(first in {prior}, again in {section}). "
                f"Qualify the 'IRI local name' in the master to disambiguate.")
        self.defined[name] = section

    # ---- meta-schema ----
    def header(self):
        p = self.p
        self.w(f"@prefix {p}:   <{self.ns}> .")
        self.w(f"@prefix ref:  <{REF_NS}> .")
        self.w("@prefix owl:  <http://www.w3.org/2002/07/owl#> .")
        self.w("@prefix rdf:  <http://www.w3.org/1999/02/22-rdf-syntax-ns#> .")
        self.w("@prefix rdfs: <http://www.w3.org/2000/01/rdf-schema#> .")
        self.w("@prefix xsd:  <http://www.w3.org/2001/XMLSchema#> .")
        self.w("@prefix skos: <http://www.w3.org/2004/02/skos/core#> .")
        self.w()
        self.w(f"<{BASE}{p}> a owl:Ontology ;")
        self.w(f"    owl:imports <{BASE}ref> ;")
        self.w(f"    rdfs:comment {lit(f'{p} vendor overlay. Deltas only: {p}: nodes plus SKOS mapping edges to the frozen ref: layer (imported, never restated). GENERATED from the capture workbook - do not hand-edit.')} .")
        self.w()
        self.w("# ============================================================")
        self.w(f"# {p}: META-SCHEMA")
        self.w("# ============================================================")
        self.w(f"{p}:DataObject a owl:Class ; rdfs:label {lit(p + ' Data Object')} .")
        self.w(f"{p}:ExternalReference a owl:Class ; rdfs:subClassOf {p}:DataObject ; "
               f"rdfs:label {lit(p + ' External Reference')} .")
        # fix D
        self.w(f"{p}:{self.OP_CLASS} a owl:Class ; rdfs:label {lit(p + ' Portal Operation (action verb)')} ; "
               f"rdfs:comment {lit('Verb axis. Named PortalOperation so it cannot collide with a vendor data object called Operation.')} .")
        self.w(f"{p}:operatesOn a owl:ObjectProperty ; rdfs:domain {p}:{self.OP_CLASS} ; "
               f"rdfs:range {p}:DataObject ; rdfs:label {lit('operatesOn')} .")
        self.w(f"{p}:verdict           a owl:AnnotationProperty .")
        self.w(f"{p}:sourceDoc         a owl:AnnotationProperty .")
        self.w(f"{p}:operationKind     a owl:AnnotationProperty .")
        self.w(f"{p}:confidence        a owl:DatatypeProperty ; rdfs:range xsd:decimal .")
        # fix E
        self.w(f"{p}:confirmsRefConcept a owl:AnnotationProperty ; "
               f"rdfs:comment {lit('Vendor independently confirms a ref: concept currently tagged Derived (Derived->Grounded signal).')} .")
        self.w(f"{p}:promotionNote      a owl:AnnotationProperty .")
        self.w(f"{p}:promotionEvidence  a owl:AnnotationProperty .")
        self.w(f"{p}:UnresolvedReference a owl:Class ; rdfs:subClassOf {p}:ExternalReference ; "
               f"rdfs:label {lit('Unresolved Reference')} ; "
               f"rdfs:comment {lit('Placeholder for a domain/range the master marks NEEDS VERIFICATION.')} .")
        self.w()

    # ---- data objects ----
    def data_objects(self):
        sh = Sheet(self.wb, "1. Data Objects",
                   required=["Layer", "Name", "IRI local name", "Type"],
                   optional=["Maps to neutral", "Match type", "Confidence",
                             "VERDICT", "Source doc", "Notes"])
        p = self.p
        self.w("# ============================================================")
        self.w("# DATA OBJECTS")
        self.w("# ============================================================")
        self.w()
        n = 0
        for r in sh.rows():
            if norm(r.get("Layer")) != norm(p):
                continue
            name = r.str("Name")
            if not name:
                continue
            iri = cls_local(r.get("IRI local name") or name)
            self.define(iri, "1. Data Objects")
            # remember BOTH the display name and the IRI so later sheets can refer to
            # an object by either ("Parametric Data Definition (CDO)" -> ParametricDataDefinition)
            self.name_to_iri[norm(name)] = iri
            self.name_to_iri[norm(iri)] = iri
            typ = r.str("Type")
            parent = f"{p}:ExternalReference" if "external" in typ.lower() else f"{p}:DataObject"

            parts = [f"{p}:{iri} a owl:Class ; rdfs:subClassOf {parent} ; rdfs:label {lit(name)}"]
            for col, prop in (("VERDICT", "verdict"), ("Source doc", "sourceDoc")):
                if sh.has(col) and r.str(col):
                    parts.append(f"    {p}:{prop} {lit(r.str(col))}")
            if sh.has("Confidence"):                      # fix A/B: optional, numeric only
                d = as_decimal(r.get("Confidence"))
                if d is not None:
                    parts.append(f"    {p}:confidence {d}")
            if sh.has("Notes") and r.str("Notes"):
                parts.append(f"    rdfs:comment {lit(r.str('Notes'))}")
            self.w(" ;\n".join(parts) + " .")

            if sh.has("Maps to neutral") and sh.has("Match type"):
                m = r.str("Match type").lower().replace(" ", "")
                tgt = r.str("Maps to neutral")
                if tgt and m in MATCH_OK:
                    self.w(f"{p}:{iri} skos:{MATCH_CANON[m]} {ref_iri(tgt)} .")
            self.w()
            n += 1
        self.counts["data objects"] = n

    def resolve_class(self, name):
        """
        Map a Domain/Range/operatesOn cell to a DEFINED class IRI.
        Sheets refer to objects by DISPLAY name ("Parametric Data Definition (CDO)")
        while the class is minted from the 'IRI local name' column
        ("ParametricDataDefinition"). Transforming the display name directly produced
        dangling IRIs, so resolve through the Data Objects map first.
        """
        raw = str(name).strip()
        # Masters annotate ranges parenthetically: "Area (self)", "Flow (self of parent)",
        # "NEEDS VERIFICATION (likely Step)". The parenthetical is a note, not part of the
        # name - strip it before resolving, or we mint dangling IRIs like Areaself.
        base = re.sub(r"\s*\([^)]*\)\s*$", "", raw).strip()
        for cand in (norm(raw), norm(base)):
            if cand in self.name_to_iri:
                return self.name_to_iri[cand]
        # A cell may list SEVERAL alternative targets ("Area | Material | Resource").
        # RDFS rdfs:range is conjunctive, so emitting a fused IRI is wrong and emitting
        # all of them would over-constrain. Resolve to the first known class and record
        # the alternatives as a warning.
        if re.search(r"[|/]| or ", base, re.I):
            alts = [a.strip() for a in re.split(r"[|/]| or ", base, flags=re.I) if a.strip()]
            for a in alts:
                if norm(a) in self.name_to_iri:
                    self.warnings.append(
                        f"multi-target range {raw!r} resolved to {a!r} "
                        f"(alternatives not emitted; rdfs:range is conjunctive)")
                    return self.name_to_iri[norm(a)]
        if base.lower() in ("string", "text", "integer", "boolean", "decimal", "datetime"):
            self.warnings.append(
                f"object-property range looks like a datatype: {raw!r} -> placeholder")
            return "UnresolvedReference"
        if base.upper().startswith("NEEDS VERIFICATION"):
            self.warnings.append(
                f"unresolved range/domain kept as placeholder: {raw!r}")
            return "UnresolvedReference"
        return cls_local(base or raw)

    # ---- enums & lookups (fix F: keep allowed values) ----
    def enums(self):
        sh = Sheet(self.wb, "4. Enums & Lookups", required=["Set name", "Kind"],
                   optional=["Allowed values", "Used by", "Layer", "Notes"])
        p = self.p
        self.w("# ============================================================")
        self.w("# ENUMS & LOOKUPS")
        self.w("# ============================================================")
        self.w()
        n = 0
        for r in sh.rows():
            setname = r.str("Set name")
            if not setname:
                continue
            iri = cls_local(setname)
            self.define(iri, "4. Enums & Lookups")
            kind = r.str("Kind").lower()
            values = r.str("Allowed values")
            notes = r.str("Notes")
            vals = [v.strip() for v in re.split(r"[;|]", values) if v.strip()] if values else []
            extensible = "extensible" in values.lower() or kind == "lookup"

            if kind == "enum" and vals and not extensible:
                vlist = " ".join(lit(v) for v in vals)
                parts = [f"{p}:{iri} a rdfs:Datatype ; rdfs:label {lit(setname)}",
                         f"    owl:equivalentClass [ a rdfs:Datatype ; owl:oneOf ( {vlist} ) ]"]
                if notes:
                    parts.append(f"    rdfs:comment {lit(notes)}")
                self.w(" ;\n".join(parts) + " .")
            else:
                # extensible lookup: ConceptScheme + known members, without implying closure
                parts = [f"{p}:{iri} a skos:ConceptScheme ; rdfs:label {lit(setname)}"]
                if notes:
                    parts.append(f"    rdfs:comment {lit(notes)}")
                self.w(" ;\n".join(parts) + " .")
                for v in vals:
                    if "extensible" in v.lower():
                        continue
                    member = f"{iri}_{local(v)}"
                    self.w(f"{p}:{member} a skos:Concept ; skos:inScheme {p}:{iri} ; "
                           f"skos:prefLabel {lit(v)} .")
            self.w()
            n += 1
        self.counts["enum/lookup sets"] = n

    # ---- attributes (fix C: qualify only on collision) ----
    def attributes(self):
        sh = Sheet(self.wb, "2. Attributes",
                   required=["Data object", "Attribute", "IRI local name", "Datatype"],
                   optional=["Value kind", "Enum/Lookup set", "Cardinality",
                             "Conditional on", "Source doc", "Notes"])
        p = self.p
        # pre-scan to find local names used by more than one owning class
        seen = {}
        for r in sh.rows():
            dobj, iri = r.str("Data object"), local(r.get("IRI local name") or r.str("Attribute"))
            if dobj and iri:
                seen.setdefault(iri, set()).add(dobj)
        collide = {k for k, v in seen.items() if len(v) > 1}

        self.w("# ============================================================")
        self.w("# ATTRIBUTES (data properties)")
        self.w("# ============================================================")
        self.w()
        n = 0
        for r in sh.rows():
            dobj = r.str("Data object")
            if not dobj:
                continue
            attr = r.str("Attribute")
            base = local(r.get("IRI local name") or attr)
            # only auto-qualify when the hand-authored name is ambiguous
            iri = (local(dobj) + base[:1].upper() + base[1:]) if base in collide else base
            self.define(iri, "2. Attributes")

            vkind = r.str("Value kind").lower()
            eset = r.str("Enum/Lookup set")
            # If the "lookup set" is actually a defined DATA OBJECT (e.g. Resource.Factory
            # -> the Factory class), this is an object reference, not a literal. Emitting it
            # as a DatatypeProperty with a class range is invalid; emit an ObjectProperty.
            as_object_ref = bool(eset) and norm(eset) in self.name_to_iri
            if as_object_ref:
                rng = f"{p}:{self.resolve_class(eset)}"
            elif vkind in ("enum", "lookup") and eset:
                rng = f"{p}:{cls_local(eset)}"
            else:
                rng = DTYPE.get(r.str("Datatype").lower(), "xsd:string")

            ptype = "owl:ObjectProperty" if as_object_ref else "owl:DatatypeProperty"
            parts = [f"{p}:{iri} a {ptype} ; rdfs:label {lit(attr)} ; "
                     f"rdfs:domain {p}:{self.resolve_class(dobj)}",
                     f"    rdfs:range {rng}"]
            bits = []
            if r.str("Cardinality"):
                bits.append(f"cardinality {r.str('Cardinality')}")
            if r.str("Conditional on"):
                bits.append(f"conditional on {r.str('Conditional on')}")
            if r.str("Notes"):
                bits.append(r.str("Notes"))
            if bits:
                parts.append(f"    rdfs:comment {lit('; '.join(bits))}")
            self.w(" ;\n".join(parts) + " .")
            self.w()
            n += 1
        self.counts["attributes"] = n

    # ---- object properties (fix F: carry Notes) ----
    def object_properties(self):
        sh = Sheet(self.wb, "3. Object Properties",
                   required=["Domain", "Property", "IRI local name", "Range"],
                   optional=["Cardinality", "Conditional on", "Maps to neutral",
                             "Match type", "Confidence", "VERDICT", "Source doc", "Notes"])
        p = self.p
        seen = {}
        for r in sh.rows():
            dom, iri = r.str("Domain"), local(r.get("IRI local name") or r.str("Property"))
            if dom and iri:
                seen.setdefault(iri, set()).add(dom)
        collide = {k for k, v in seen.items() if len(v) > 1}

        self.w("# ============================================================")
        self.w("# OBJECT PROPERTIES")
        self.w("# ============================================================")
        self.w()
        n = 0
        for r in sh.rows():
            dom = r.str("Domain")
            if not dom:
                continue
            prop = r.str("Property")
            base = local(r.get("IRI local name") or prop)
            iri = (local(dom) + base[:1].upper() + base[1:]) if base in collide else base
            self.define(iri, "3. Object Properties")

            parts = [f"{p}:{iri} a owl:ObjectProperty ; rdfs:label {lit(prop)} ; "
                     f"rdfs:domain {p}:{self.resolve_class(dom)}",
                     f"    rdfs:range {p}:{self.resolve_class(r.str('Range'))}"]
            bits = []
            if r.str("Cardinality"):
                bits.append(f"cardinality {r.str('Cardinality')}")
            if r.str("Conditional on"):
                bits.append(f"conditional on {r.str('Conditional on')}")
            if sh.has("Notes") and r.str("Notes"):
                bits.append(r.str("Notes"))
            if bits:
                parts.append(f"    rdfs:comment {lit('; '.join(bits))}")
            if sh.has("VERDICT") and r.str("VERDICT"):
                parts.append(f"    {p}:verdict {lit(r.str('VERDICT'))}")
            if sh.has("Confidence"):
                d = as_decimal(r.get("Confidence"))
                if d is not None:
                    parts.append(f"    {p}:confidence {d}")
            self.w(" ;\n".join(parts) + " .")

            if sh.has("Maps to neutral") and sh.has("Match type"):
                m = r.str("Match type").lower().replace(" ", "")
                tgt = r.str("Maps to neutral")
                if tgt and m in MATCH_OK:
                    self.w(f"{p}:{iri} skos:{MATCH_CANON[m]} {ref_iri(tgt)} .")
            self.w()
            n += 1
        self.counts["object properties"] = n

    # ---- operations (fix D: PortalOperation) ----
    def operations(self):
        if "5. Operations" not in self.wb.sheetnames:
            self.counts["operations"] = 0
            return
        sh = Sheet(self.wb, "5. Operations", required=["Label"],
                   optional=["Operation IRI", "Kind", "operatesOn", "Maps to (peer vendor)",
                             "Match type", "Confidence", "VERDICT", "Source doc", "Notes"])
        p = self.p
        self.w("# ============================================================")
        self.w("# OPERATIONS (verb axis; lateral vendor->peer mapping)")
        self.w("# ============================================================")
        self.w()
        # operatesOn column name varies by master (e.g. "operatesOn (cm:)")
        op_col = next((h for h in sh.headers if h.startswith("operateson")), None)
        n = 0
        for r in sh.rows():
            label = r.str("Label")
            if not label:
                continue
            src = r.str("Operation IRI") or label
            iri = cls_local(src)
            self.define(iri, "5. Operations")
            parts = [f"{p}:{iri} a {p}:{self.OP_CLASS} ; rdfs:label {lit(label)}"]
            if op_col:
                tgt = str(r._row[sh.headers[op_col]] or "").strip()
                if tgt:
                    parts.append(f"    {p}:operatesOn {p}:{self.resolve_class(tgt)}")
            for col, prop in (("Kind", "operationKind"), ("VERDICT", "verdict"),
                              ("Source doc", "sourceDoc")):
                if sh.has(col) and r.str(col):
                    parts.append(f"    {p}:{prop} {lit(r.str(col))}")
            if sh.has("Confidence"):
                d = as_decimal(r.get("Confidence"))
                if d is not None:
                    parts.append(f"    {p}:confidence {d}")
            if sh.has("Notes") and r.str("Notes"):
                parts.append(f"    rdfs:comment {lit(r.str('Notes'))}")
            self.w(" ;\n".join(parts) + " .")
            self.w()
            n += 1
        self.counts["operations"] = n

    # ---- promotions (fix E) ----
    def promotions(self):
        name = next((s for s in self.wb.sheetnames if "promotion" in s.lower()), None)
        if not name:
            self.counts["promotions"] = 0
            return
        sh = Sheet(self.wb, name, required=[],
                   optional=["Subject", "Vendor concept", "Ref concept", "Target",
                             "Evidence", "Note", "Notes"])
        p = self.p
        self.w("# ============================================================")
        self.w("# PROMOTIONS (vendor confirms a Derived ref: concept)")
        self.w("# ============================================================")
        self.w()
        # column synonyms - masters spell these differently
        SUBJ = ("subject", "vendor concept", "opcenter evidence", "vendor evidence",
                "evidence", "source screen")
        TGT  = ("ref concept", "target", "ref: target", "ref target")
        NOTE = ("notes", "note", "recommendation", "rationale")
        subj_col = next((h for h in SUBJ if h in sh.headers), None)
        tgt_col  = next((h for h in TGT  if h in sh.headers), None)
        note_col = next((h for h in NOTE if h in sh.headers), None)
        if not tgt_col:
            raise BuildError(
                f"Sheet '{name}' has no recognisable ref-target column. "
                f"Looked for {list(TGT)}; found {sorted(sh.headers)}.")
        n = 0
        if subj_col and tgt_col:
            for r in sh.rows():
                t = str(r._row[sh.headers[tgt_col]] or "").strip()
                if not t:
                    continue
                raw_s = str(r._row[sh.headers[subj_col]] or "").strip() if subj_col else ""
                # evidence text is often prose ("opc:Container.Level (Batch...)").
                # take a vendor IRI if one is present, else anchor on the ref target.
                m = re.search(rf"{p}:([A-Za-z0-9_.]+)", raw_s)
                subj = cls_local(m.group(1).split(".")[0]) if m else cls_local(local(t))
                parts = [f"{p}:{subj} {p}:confirmsRefConcept {ref_iri(t)}"]
                if raw_s:
                    parts.append(f"    {p}:promotionEvidence {lit(raw_s)}")
                if note_col:
                    note = str(r._row[sh.headers[note_col]] or "").strip()
                    if note:
                        parts.append(f"    {p}:promotionNote {lit(note)}")
                self.w(" ;\n".join(parts) + " .")
                n += 1
            self.w()
        self.counts["promotions"] = n

    def generate(self):
        self.header()
        self.data_objects()
        self.enums()
        self.attributes()
        self.object_properties()
        self.operations()
        self.promotions()
        # fix D: meta IRIs must not collide with minted IRIs
        clash = self.meta_iris & set(self.defined)
        if clash:
            raise BuildError(
                f"Meta-schema IRI(s) {sorted(clash)} collide with names minted from the "
                f"master. Rename in the master or adjust the meta-schema.")
        return "\n".join(self.out) + "\n"


# ---------------------------------------------------------------------------
# Validation pass (fix G)
# ---------------------------------------------------------------------------
def validate(ttl, prefix, counts, warnings=()):
    problems = []

    # 1. re-parse with pyoxigraph
    try:
        import pyoxigraph as ox
        store = ox.Store()
        store.load(ttl.encode("utf-8"), format=ox.RdfFormat.TURTLE,
                   to_graph=ox.NamedNode("urn:validate"))
        triples = sum(1 for _ in store.quads_for_pattern(
            None, None, None, ox.NamedNode("urn:validate")))
    except ImportError:
        triples = None
        problems.append("WARN: pyoxigraph not installed - skipped parse check.")
    except Exception as e:
        problems.append(f"FAIL: output is not valid Turtle: {e}")
        triples = None

    ns = f"{BASE}{prefix}#"

    # 2. mapping targets must be in ref:
    for m in re.finditer(r"skos:\w+Match\s+(\S+)\s*\.", ttl):
        tgt = m.group(1)
        if not tgt.startswith("ref:"):
            problems.append(f"FAIL: SKOS mapping target outside ref: -> {tgt}")

    # 3. ref: must never be a subject
    for line in ttl.splitlines():
        if re.match(r"^\s*ref:\S+\s+(a|rdfs:|owl:)", line):
            problems.append(f"FAIL: ref: IRI used as a subject -> {line.strip()[:70]}")

    # 4. datatype property ranges must be datatypes, not classes
    declared_datatypes = set(re.findall(rf"{prefix}:(\w+) a rdfs:Datatype", ttl))
    for m in re.finditer(rf"{prefix}:(\w+) a owl:DatatypeProperty[^.]*?rdfs:range\s+(\S+)", ttl, re.S):
        prop, rng = m.groups()
        if rng.startswith(f"{prefix}:"):
            local_rng = rng.split(":", 1)[1]
            if local_rng not in declared_datatypes and not rng.startswith("xsd:"):
                # ConceptScheme ranges are an accepted modelling choice for lookups
                if f"{prefix}:{local_rng} a skos:ConceptScheme" not in ttl:
                    problems.append(
                        f"FAIL: datatype property {prefix}:{prop} has class range {rng}")

    # 5. no bare (unquoted) annotation objects.
    #    Skip the meta-schema declarations themselves (":verdict a owl:AnnotationProperty"),
    #    which legitimately have a non-literal object.
    for m in re.finditer(rf"{prefix}:(verdict|sourceDoc|operationKind)\s+([^\s\"<][^\s;.]*)", ttl):
        obj = m.group(2)
        if obj == "a" or obj.startswith("owl:") or obj.startswith("rdfs:"):
            continue
        problems.append(f"FAIL: unquoted literal -> {prefix}:{m.group(1)} {obj}")

    # 6. NON-EMPTINESS (a silently empty section is the worst failure mode:
    #    the build "succeeds" and produces an overlay with no link to ref: at all).
    if counts.get("data objects", 0) == 0:
        problems.append("FAIL: zero data objects emitted - check the Layer column "
                        "values match the --prefix, and that sheet '1. Data Objects' "
                        "has rows for this vendor.")
    n_map = len(re.findall(r"skos:\w+Match", ttl))
    if n_map == 0:
        problems.append("FAIL: zero SKOS mapping edges emitted - the overlay would have "
                        "no link to ref: and is unusable for comparison.")

    # 7. DANGLING IRIs: anything referenced by domain/range/operatesOn must be defined.
    defined_iris = set(re.findall(rf"^{prefix}:(\w+) a ", ttl, re.M))
    referenced = set()
    for pat in (rf"rdfs:domain\s+{prefix}:(\w+)", rf"rdfs:range\s+{prefix}:(\w+)",
                rf"{prefix}:operatesOn\s+{prefix}:(\w+)"):
        referenced |= set(re.findall(pat, ttl))
    dangling = sorted(referenced - defined_iris)
    if dangling:
        problems.append(
            f"FAIL: {len(dangling)} IRI(s) referenced but never defined: "
            f"{dangling[:8]}{' ...' if len(dangling) > 8 else ''}")

    print("\n--- generation summary ---")
    for k, v in counts.items():
        print(f"  {k:22} {v}")
    print(f"  {'skos mapping edges':22} {n_map}")
    if triples is not None:
        print(f"  {'triples parsed':22} {triples}")
    if warnings:
        print("\n--- warnings ---")
        for w in warnings:
            print("  " + w)
    if problems:
        print("\n--- validation problems ---")
        for p in problems:
            print("  " + p)
    else:
        print("\n  validation: all checks passed")
    return [p for p in problems if p.startswith("FAIL")]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--prefix", required=True, help="vendor prefix, e.g. opc")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    try:
        gen = Generator(args.workbook, args.prefix)
        ttl = gen.generate()
        gen.wb.close()
    except BuildError as e:
        print(f"BUILD FAILED: {e}", file=sys.stderr)
        sys.exit(1)

    failures = validate(ttl, args.prefix, gen.counts, gen.warnings)
    if failures:
        print(f"\nBUILD FAILED: {len(failures)} validation failure(s); file not written.",
              file=sys.stderr)
        sys.exit(1)

    with open(args.out, "w", encoding="utf-8") as f:
        f.write(ttl)
    print(f"\nwrote {args.out}")


if __name__ == "__main__":
    main()
