"""The catalogue of the 52 regression items on Test_Expected (#352).

update_expected.py writes one row per expected value with the columns Area,
Description, Source and Source status, so a new item cannot lack a row: an item
that describe() does not know raises. The paragraphs come from
docs/methodology.md. Agreement with the independent reference shows that the
VBA and the reference read the regulation the same way; it does not prove the
reading, which the status says.
"""

from __future__ import annotations

SUMMARY: dict[str, tuple[str, str, str]] = {
    "SBM_LOW": ("SBM", "Sum over risk classes of delta, vega and curvature, low scenario", "MAR21.7"),
    "SBM_MEDIUM": ("SBM", "Sum over risk classes of delta, vega and curvature, medium scenario", "MAR21.7"),
    "SBM_HIGH": ("SBM", "Sum over risk classes of delta, vega and curvature, high scenario", "MAR21.7"),
    "SBM": ("SBM", "The largest of the three scenario totals", "MAR21.7"),
    "DRC_NONSEC": ("DRC", "Default risk capital of non-securitizations, sum of the buckets", "MAR22.22-MAR22.26"),
    "RRAO": ("RRAO", "Gross notional times 1.0% (exotic) and 0.1% (other)", "MAR23.8"),
    "TOTAL": ("Total", "SBM plus DRC plus RRAO", "docs/methodology.md"),
}

#: Paragraphs of the delta risk weights and correlations per risk class.
DELTA = {
    "GIRR": "MAR21.42-MAR21.50",
    "CSR_NS": "MAR21.51-MAR21.57",
    "EQ": "MAR21.72-MAR21.80",
    "COM": "MAR21.82-MAR21.85",
    "FX": "MAR21.87-MAR21.89",
}
VEGA = "MAR21.25, MAR21.92, MAR21.94"
CURVATURE = "MAR21.5, MAR21.98-MAR21.101"
COMPONENT = {"DELTA": "Delta", "VEGA": "Vega", "CURVATURE": "Curvature"}
SCENARIO = {"LOW": "low", "MEDIUM": "medium", "HIGH": "high"}
RISK_CLASS = {"GIRR": "GIRR", "CSR_NS": "CSR non-securitization", "EQ": "equity", "COM": "commodity", "FX": "FX"}

#: What agreement with the independent reference does and does not show.
STATUS = "Parameters checked against the BIS text; reading not independently reviewed (#29)"


def describe(item: str) -> tuple[str, str, str, str]:
    """(area, description, source, source status) of an expected-value item."""
    if item in SUMMARY:
        area, description, paragraphs = SUMMARY[item]
        return area, description, f"{paragraphs}; independent reference", STATUS
    risk_class, component, scenario = item.split("|")
    paragraphs = {"DELTA": DELTA[risk_class], "VEGA": VEGA, "CURVATURE": CURVATURE}[component]
    return (
        f"SBM | {risk_class} | {COMPONENT[component]}",
        f"{COMPONENT[component]} capital of {RISK_CLASS[risk_class]}, {SCENARIO[scenario]} scenario",
        f"{paragraphs}; independent reference",
        STATUS,
    )
