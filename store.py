"""
store.py — vendor-agnostic ontology store for the MES comparison platform.

The shared core. Loads any number of TTL layers into named graphs, runs SPARQL,
and returns named bindings. Both the Graph Engineer (vetted queries) and any
later reasoner sit on top of this.

Design rules honoured here:
  - Vendor identity is DERIVED from data (namespaces discovered at load time),
    never hard-coded. A third overlay loads with zero code change.
  - ref: is just one loaded layer; it is not special-cased in code.
  - SPARQL results come back as dicts keyed by the SELECT variable name
    (row["refLabel"]), not fragile positional v0/v1/v2.

Written for pyoxigraph 0.5.x.

Usage:
    from store import OntologyStore
    store = OntologyStore([
        ("ref", "mes-neutral-reference-4.ttl"),
        ("cm",  "mes-cm-overlay.ttl"),
        # ("opc", "mes-opcenter-overlay.ttl"),   # <- a third slots in here, no other change
    ])
    rows = store.sparql("SELECT ?s WHERE { ?s a owl:Class } LIMIT 5")
"""

import pyoxigraph as ox


# Standard vocabulary prefixes, always available to every query.
STD_PREFIXES = {
    "owl":  "http://www.w3.org/2002/07/owl#",
    "rdf":  "http://www.w3.org/1999/02/22-rdf-syntax-ns#",
    "rdfs": "http://www.w3.org/2000/01/rdf-schema#",
    "skos": "http://www.w3.org/2004/02/skos/core#",
    "xsd":  "http://www.w3.org/2001/XMLSchema#",
    "dct":  "http://purl.org/dc/terms/",
}


class OntologyStore:
    def __init__(self, layers):
        """
        layers: list of (prefix, ttl_path) tuples, e.g. [("ref", "..."), ("cm", "...")].
        Each file loads into its own named graph urn:graph:<prefix>.
        The namespace for each prefix is discovered from the file, not assumed.
        """
        self.store = ox.Store()
        self.layers = {}          # prefix -> {"graph": NamedNode, "ns": namespace-iri, "path": path}
        self.ns_to_prefix = {}    # namespace-iri -> prefix  (for pretty-printing IRIs)

        for prefix, path in layers:
            graph = ox.NamedNode(f"urn:graph:{prefix}")
            with open(path, "rb") as f:
                self.store.load(f, format=ox.RdfFormat.TURTLE, to_graph=graph)
            ns = self._discover_namespace(graph, prefix)
            self.layers[prefix] = {"graph": graph, "ns": ns, "path": path}
            if ns:
                self.ns_to_prefix[ns] = prefix
            count = sum(1 for _ in self.store.quads_for_pattern(None, None, None, graph))
            print(f"  loaded {prefix}: {count} triples  (ns: {ns or 'unknown'})")

    def _discover_namespace(self, graph, prefix):
        """
        Find the layer's own namespace by looking at the subjects it declares.
        The most common '<ns>#localname' base among subjects in this graph wins.
        No namespace is hard-coded — a new vendor's namespace is learned here.
        """
        counts = {}
        for quad in self.store.quads_for_pattern(None, None, None, graph):
            s = quad.subject
            if isinstance(s, ox.NamedNode) and "#" in s.value:
                base = s.value.rsplit("#", 1)[0] + "#"
                counts[base] = counts.get(base, 0) + 1
        if not counts:
            return None
        return max(counts, key=counts.get)

    def _prefix_block(self):
        """Build the PREFIX header: standard vocab + every loaded layer's own prefix."""
        lines = [f"PREFIX {p}: <{ns}>" for p, ns in STD_PREFIXES.items()]
        for prefix, info in self.layers.items():
            if info["ns"]:
                lines.append(f"PREFIX {prefix}: <{info['ns']}>")
        return "\n".join(lines) + "\n"

    def sparql(self, query, params=None):
        """
        Run a SELECT query and return a list of dicts keyed by SELECT variable name.
        `params` optionally substitutes %PLACEHOLDER% tokens (used by the vetted
        query files, e.g. {"VENDOR": "https://.../mes-cm#"}).
        Prefixes are prepended automatically; do not repeat them in `query`.
        """
        if params:
            for key, val in params.items():
                query = query.replace(f"%{key}%", val)

        full = self._prefix_block() + query
        # Data is loaded into named graphs (one per layer). By default pyoxigraph's
        # query() only sees the default graph, so plain SPARQL would return nothing.
        # use_default_graph_as_union makes the whole store visible as one dataset,
        # which is what the vetted queries assume (they don't wrap everything in GRAPH).
        results = self.store.query(full, use_default_graph_as_union=True)

        rows = []
        var_names = [v.value for v in results.variables]  # e.g. ['refObj', 'refLabel', ...]
        for solution in results:
            row = {}
            for name in var_names:
                term = solution[ox.Variable(name)]
                if term is not None:
                    row[name] = term.value if hasattr(term, "value") else str(term)
            rows.append(row)
        return rows

    def pfx(self, iri):
        """Shorten a full IRI to prefix:local using the discovered namespace map."""
        for ns, prefix in self.ns_to_prefix.items():
            if iri.startswith(ns):
                return f"{prefix}:{iri[len(ns):]}"
        for prefix, ns in STD_PREFIXES.items():
            if iri.startswith(ns):
                return f"{prefix}:{iri[len(ns):]}"
        return iri

    def vendors(self):
        """
        Return the set of vendor prefixes actually present in the data, discovered
        by the #verdict predicate (vendor overlays carry it; ref: does not).
        This is the same vendor-discovery rule the vetted queries use.
        """
        q = """
        SELECT DISTINCT ?v WHERE {
          ?t ?p ?o . FILTER( STRENDS(STR(?p), "#verdict") )
          BIND( REPLACE(STR(?t), "#.*$", "#") AS ?v )
        }"""
        out = []
        for row in self.sparql(q):
            ns = row.get("v")
            out.append(self.ns_to_prefix.get(ns, ns))
        return sorted(out)
