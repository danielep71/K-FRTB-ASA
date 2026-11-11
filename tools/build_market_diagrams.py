"""Render the five market-risk and CVA diagrams as SVG masters and PNGs.

The diagrams in docs/assets/market-risk/ map the regulatory architecture of
FRTB-SA and CVA risk (Basel MAR20-MAR23 and MAR50; the framework diagram also
cites MAR30-MAR33, MAR40, BCBS 158 and BCBS 189). They are method maps: they do
not describe the scope this prototype implements.

Text is real SVG text; "K_{SBM}" in a label is written as K with a subscript.
PNGs are rendered with CairoSVG 2.9.1 (and its Cairo system library), which is
not one of the pinned test dependencies. Fonts change text metrics: Arial is
preferred, DejaVu Sans is the fallback.

Usage:
    python3 tools/build_market_diagrams.py [--output DIR]
"""

from __future__ import annotations

import argparse
import re
import textwrap
from collections.abc import Iterable, Sequence
from html import escape
from pathlib import Path

import cairosvg

DEFAULT_OUTPUT = Path(__file__).resolve().parents[1] / "docs" / "assets" / "market-risk"
NAVY = "#17365F"
BLUE = "#24558D"
PALE = "#EAF1FA"
EDGE = "#B5CBE8"
INK = "#31445C"
GREEN = "#218452"
MINT = "#EAF5EF"
LINK = "#7999BF"
FONT = "Arial, DejaVu Sans, sans-serif"
SUBSCRIPT = re.compile(r"_\{([^}]*)\}")


def plain(text: str) -> str:
    """The text as it reads, without subscript markup (for wrapping)."""
    return SUBSCRIPT.sub(r"\1", text)


def markup(text: str, size: float) -> str:
    """Escaped SVG text content; _{...} becomes a lowered, smaller tspan."""
    out, last, shift = [], 0, round(size * 0.28)
    for m in SUBSCRIPT.finditer(text):
        out.append(escape(text[last:m.start()]))
        out.append(f'<tspan dy="{shift}" font-size="{round(size * 0.68)}">{escape(m.group(1))}</tspan>'
                   f'<tspan dy="-{shift}">')
        last = m.end()
    out.append(escape(text[last:]))
    return "".join(out) + "</tspan>" * len(SUBSCRIPT.findall(text))


class Diagram:
    def __init__(self, name: str, title: str, subtitle: str, h: int) -> None:
        self.name = name
        self.h = h
        self.p: list[str] = []
        self.p.append(
            f'<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="{h}" viewBox="0 0 1600 {h}" '
            f'role="img" aria-labelledby="title desc"><title id="title">{escape(title)}</title>'
            f'<desc id="desc">{escape(subtitle)}</desc><defs><marker id="arrow" viewBox="0 0 10 10" refX="9" '
            f'refY="5" markerWidth="8" markerHeight="8" orient="auto-start-reverse"><path d="M0 0L10 5L0 10Z" '
            f'fill="{LINK}"/></marker></defs><rect width="1600" height="{h}" fill="white"/>')
        self.text(800, 55, title, 38, NAVY, True, "middle")
        self.text(800, 96, subtitle, 23, BLUE, False, "middle")

    def text(self, x: float, y: float, t: str, size: float = 22, color: str = INK, bold: bool = False,
             anchor: str = "start") -> None:
        if anchor != "start" and SUBSCRIPT.search(t):
            # CairoSVG misplaces centered or end-anchored text whose tspans shift
            # the baseline, so such text is measured and drawn start-anchored.
            w = text_width(t, size, bold)
            x, anchor = round(x - (w / 2 if anchor == "middle" else w), 1), "start"
        self.p.append(f'<text x="{x}" y="{y}" font-family="{FONT}" font-size="{size}" fill="{color}" '
                      f'font-weight="{700 if bold else 400}" text-anchor="{anchor}">{markup(t, size)}</text>')

    def rect(self, x: float, y: float, w: float, h: float, fill: str = PALE, stroke: str = EDGE, r: float = 10,
             dash: bool = False) -> None:
        pattern = ' stroke-dasharray="7 5"' if dash else ""
        self.p.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{r}" fill="{fill}" '
                      f'stroke="{stroke}" stroke-width="2"{pattern}/>')

    def line(self, points: Iterable[tuple[float, float]]) -> None:
        path = " L".join(f"{x} {y}" for x, y in points)
        self.p.append(f'<path d="M{path}" stroke="{LINK}" stroke-width="3" fill="none"/>')

    def block(self, x: float, y: float, w: float, h: float, lines: Sequence[tuple[str, float, bool, str]],
              fill: str = PALE, stroke: str = EDGE, dash: bool = False) -> None:
        """A box with centered, wrapped lines; each line is (text, size, bold, color)."""
        self.rect(x, y, w, h, fill, stroke, dash=dash)
        rows = [(part, size, bold, color) for text, size, bold, color in lines
                for part in wrap_width(text, size, bold, w - 28)]
        height = sum(size * 1.3 for _, size, _, _ in rows)
        assert height < h - 8, (self.name, lines, height, h)
        yy = y + (h - height) / 2
        for part, size, bold, color in rows:
            yy += size * 1.3
            self.text(x + w / 2, round(yy - size * 0.32, 1), part, size, color, bold, "middle")

    def arrow(self, x1: float, y1: float, x2: float, y2: float) -> None:
        self.p.append(f'<path d="M{x1} {y1} L{x2} {y2}" stroke="{LINK}" stroke-width="3" fill="none" '
                      f'marker-end="url(#arrow)"/>')

    def card(self, x: float, y: float, w: float, h: float, title: str, body: Iterable[str] = (), dark: bool = False,
             green: bool = False, size: float = 22) -> None:
        self.rect(x, y, w, h, NAVY if dark else MINT if green else PALE, NAVY if dark else GREEN if green else EDGE)
        self.text(x + 22, y + 36, title, 25, "white" if dark else NAVY, True)
        yy = y + 70
        for paragraph in body:
            for line in wrap(paragraph, int((w - 46) / (size * .53))):
                assert yy < y + h - 12, (self.name, title, yy, y + h)
                self.text(x + 22, yy, line, size, "#FFFFFF" if dark else INK)
                yy += size * 1.36
            yy += 8

    def bar(self, y: float, n: int, title: str, note: str = "") -> None:
        self.rect(32, y, 1536, 42, BLUE, BLUE, 3)
        self.rect(32, y, 48, 42, NAVY, NAVY, 3)
        self.text(56, y + 29, str(n), 23, "white", True, "middle")
        self.text(96, y + 29, title, 24, "white", True)
        if note:
            self.text(1544, y + 29, note, 19, "#DBE7F7", False, "end")

    def foot(self, source: str, scope: str) -> None:
        self.text(36, self.h - 55, scope, 18, INK)
        self.text(36, self.h - 25, source, 17, BLUE)

    def save(self, out: Path) -> None:
        data = ("".join(self.p) + "</svg>\n").encode()
        (out / (self.name + ".svg")).write_bytes(data)
        cairosvg.svg2png(bytestring=data, write_to=str(out / (self.name + ".png")))


# Arial, or Liberation Sans, which has Arial's metrics, for measuring text.
FONT_FILES = {
    False: ["/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", "C:/Windows/Fonts/arial.ttf",
            "/Library/Fonts/Arial.ttf"],
    True: ["/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf", "C:/Windows/Fonts/arialbd.ttf",
           "/Library/Fonts/Arial Bold.ttf"],
}


def text_width(text: str, size: float, bold: bool) -> float:
    """Rendered width of text with subscript markup, from Arial metrics when available."""
    # Pillow is a CairoSVG dependency, not a pinned one; unused-ignore covers environments that have it.
    from PIL import ImageFont
    path = next((f for f in FONT_FILES[bold] if Path(f).exists()), None)
    width, last = 0.0, 0
    segments: list[tuple[str, float]] = []
    for m in SUBSCRIPT.finditer(text):
        segments += [(text[last:m.start()], size), (m.group(1), round(size * 0.68))]
        last = m.end()
    segments.append((text[last:], size))
    for segment, points in segments:
        if path:
            width += ImageFont.truetype(path, points).getlength(segment)
        else:
            width += 0.55 * points * len(segment)
    return width


def wrap(paragraph: str, width: int) -> list[str]:
    """Wrap on the visible text, then restore any subscript markup per word."""
    words = paragraph.split(" ")
    lines: list[str] = []
    current: list[str] = []
    length = 0
    for word in words:
        n = len(plain(word))
        if current and length + 1 + n > width:
            lines.append(" ".join(current))
            current, length = [], 0
        current.append(word)
        length += n + (1 if length else 0)
    if current:
        lines.append(" ".join(current))
    return lines


def wrap_width(paragraph: str, size: float, bold: bool, width: float) -> list[str]:
    """Wrap to a width in pixels, measured as rendered."""
    lines: list[str] = []
    current: list[str] = []
    for word in paragraph.split(" "):
        if current and text_width(" ".join(current + [word]), size, bold) > width:
            lines.append(" ".join(current))
            current = []
        current.append(word)
    return lines + [" ".join(current)]


def boundary(x: float, y: float, cx: float, cy: float, half_w: float, half_h: float) -> tuple[int, int]:
    """Where the segment from the node center (cx, cy) to (x, y) leaves the node."""
    dx, dy = x - cx, y - cy
    t = min(half_w / abs(dx) if dx else float("inf"), half_h / abs(dy) if dy else float("inf"))
    return round(cx + dx * t), round(cy + dy * t)


GREY = "#5A6B80"
FAINT = "#C3CEDC"


def approaches(
    d: Diagram,
    y: float,
    groups: Sequence[tuple[float, float, str, str, str, Sequence[tuple[str, str]]]],
    era: str,
    earlier: Iterable[tuple[float, float, str]],
    notes: Iterable[str],
) -> None:
    """One framework as a tree: approach headers, their components joined by a separator,
    the earlier rules aligned below and notes. Each group is (x, w, title, scope, separator,
    components); each component is (name, detail); each earlier rule is (x, w, text)."""
    for x, w, title, scope, separator, components in groups:
        d.block(x, y, w, 76, [(title, 22, True, "white"), (scope, 18, False, "#DBE7F7")], NAVY, NAVY)
        n = len(components)
        cw = (w - 44 * (n - 1)) / n
        centers = [round(x + cw / 2 + i * (cw + 44)) for i in range(n)]
        d.line([(round(x + w / 2), y + 76), (round(x + w / 2), y + 96)])
        if n > 1:
            d.line([(centers[0], y + 96), (centers[-1], y + 96)])
        for i, (name, detail) in enumerate(components):
            d.arrow(centers[i], y + 96, centers[i], y + 118)
            d.block(round(x + i * (cw + 44)), y + 122, round(cw), 128,
                    [(name, 21, True, NAVY), (detail, 18, False, INK)])
            if i:
                d.text(round(x + i * (cw + 44) - 22), y + 196, separator, 30 if separator == "+" else 22,
                       LINK, True, "middle")
    d.text(40, y + 286, era, 18, GREY, True)
    for x, w, text in earlier:
        d.block(x, y + 300, w, 96, [(text, 18, False, GREY)], "white", FAINT, dash=True)
    for i, note in enumerate(notes):
        d.text(40, y + 432 + i * 26, note, 18, INK)


def framework(out: Path) -> None:
    d = Diagram("Market_risk_framework", "MARKET RISK AND CVA RISK",
                "Two capital requirements, their approaches and the earlier Basel rules", 1530)
    d.text(800, 137, "K = own funds requirement       |       RWA = risk-weighted assets", 22, NAVY, False, "middle")
    d.bar(170, 1, "MARKET RISK | MAR20–33", "the IMA needs approval desk by desk; other desks use the SA")
    approaches(d, 232, [
        (40, 730, "INTERNAL MODELS APPROACH (IMA)", "Approved trading desks (MAR30–33)", "+", [
            ("Expected shortfall", "Modellable risk factors (IMCC)"),
            ("Non-modellable risk factors", "Stressed expected shortfall (SES)"),
            ("Default risk charge", "Internal model: 99.9% VaR, one-year horizon")]),
        (830, 730, "STANDARDIZED APPROACH (SA)", "All other desks (MAR20–23)", "+", [
            ("Sensitivities-based method", "Delta, vega and curvature (MAR21)"),
            ("Default risk charge", "Jump-to-default (MAR22)"),
            ("Residual risk add-on", "Exotic and other residual risks (MAR23)")])],
        "BEFORE THE REVISION: BASEL 2.5 (2009)", [
            (40, 214, "VaR and stressed VaR"),
            (298, 214, "No general charge; risks not in VaR (RNIV) in some jurisdictions"),
            (556, 214, "Incremental risk charge: default and migration"),
            (830, 472, "Standardized measurement method: general and specific market risk charges"),
            (1346, 214, "No Pillar 1 charge")], [
            "Securitization positions are out of IMA scope and use the SA (MAR30.4). IMA and SA desks combine "
            "in the aggregate requirement of MAR33.43–33.45.",
            "Not drawn: the simplified standardized approach for eligible banks (MAR40)."])
    d.bar(740, 2, "CVA RISK | MAR50", "BA-CVA unless the supervisor approves SA-CVA (MAR50.7)")
    approaches(d, 802, [
        (40, 564, "BASIC APPROACH (BA-CVA)", "Exposure-based; the default route", "or", [
            ("Reduced version", "No hedge recognition"),
            ("Full version", "Recognizes eligible credit hedges")]),
        (668, 564, "STANDARDIZED APPROACH (SA-CVA)", "Sensitivity-based; supervisory approval", "+", [
            ("Delta risk", "Six risk classes"),
            ("Vega risk", "Five risk classes; none for counterparty credit spread")]),
        (1296, 264, "ALTERNATIVE", "MAR50.9; whole portfolio", "", [
            ("100% of CCR capital", "Up to EUR 100 billion of non-centrally cleared notional; no hedges")])],
        "BEFORE THE REVISION: BASEL III (2011)", [
            (40, 564, "Standardized CVA risk capital charge"),
            (668, 564, "Advanced CVA risk capital charge: VaR and stressed VaR of credit spreads, for banks with "
                       "IMM approval"),
            (1296, 264, "No Basel equivalent")], [
            "Netting sets carved out of SA-CVA use BA-CVA (MAR50.8); no netting set is capitalized under both. "
            "The revised framework has no internal model approach for CVA."])
    d.bar(1290, 3, "TWO SEPARATE CAPITAL REQUIREMENTS", "neither includes the other")
    d.text(800, 1384, "RWA = 12.5 × K_{market} + 12.5 × K_{CVA}", 34, NAVY, True, "middle")
    d.text(800, 1428, "Each requirement is converted to RWA separately (MAR20.1, MAR33.46, MAR50.1).",
           22, BLUE, False, "middle")
    d.foot("Sources: Basel MAR20–23, MAR30–33, MAR40 and MAR50; BCBS, Revisions to the Basel II market risk "
           "framework (2009); Basel III (2011).",
           "Conceptual architecture; not a statement of implemented coverage. Local rules, exemptions and "
           "transitions must be checked separately.")
    d.save(out)


def hub(out: Path, name: str, title: str, subtitle: str, center: str, caption: str,
        cards: Sequence[tuple[str, Sequence[str], str]], source: str, scope: str) -> None:
    """Six cards around a central node. Each card is (title, body, flow): 'in' feeds the
    node, 'calc' is a step the node performs, 'out' is a result (green)."""
    d = Diagram(name, title, subtitle, 920)
    cx, cy, hw, hh = 800, 460, 140, 105
    positions = [(40, 150), (1000, 150), (40, 380), (1000, 380), (40, 610), (1000, 610)]
    anchors = [(600, 270), (1000, 270), (600, 460), (1000, 460), (600, 650), (1000, 650)]
    for (x, y), (t, b, flow) in zip(positions, cards):
        d.card(x, y, 560, 160, t, b, green=flow == "out", size=21)
    d.rect(cx - hw, cy - hh, 2 * hw, 2 * hh, NAVY, NAVY)
    parts = [part for line in center.split(" | ") for part in textwrap.wrap(line, 16) + [""]][:-1]
    top = cy - (len(parts) - 1) * 15 + 9
    for i, part in enumerate(parts):
        if part:
            d.text(cx, top + i * 30, part, 25, "white", True, "middle")
    for (ax, ay), (_, _, flow) in zip(anchors, cards):
        bx, by = boundary(ax, ay, cx, cy, hw, hh)
        if flow == "in":
            d.arrow(ax, ay, bx, by)
        else:
            d.arrow(bx, by, ax, ay)
    d.text(800, 815, caption, 23, BLUE, False, "middle")
    d.foot(source, scope)
    d.save(out)


def frtb_uses(out: Path) -> None:
    hub(out, "FRTB_uses", "Where FRTB-SA fits in the capital framework",
        "A market-risk aggregation framework built around supplied risk measures",
        "K-FRTB-ASA | SENSITIVITIES TO CAPITAL",
        "K_{FRTB-SA} = K_{SBM} + K_{DRC} + K_{RRAO}", [
            ("1 | Regulatory perimeter", ["Apply the trading-book boundary and the scope rules for FX and "
                                          "commodity positions."], "in"),
            ("2 | Risk-measure inputs", ["Supply sensitivities and curvature inputs from upstream valuation "
                                         "systems; preserve units and signs."], "in"),
            ("3 | Sensitivities-based method", ["Aggregate delta, vega and curvature by risk class and bucket. "
                                                "Compare three correlation scenarios."], "calc"),
            ("4 | Complementary charges", ["DRC addresses jump-to-default. RRAO covers specified residual risks. "
                                           "Both supplement the SBM charge."], "calc"),
            ("5 | Capital and RWA", ["Add the SBM, DRC and RRAO capital requirements. Multiply the total by 12.5 "
                                     "for RWA (MAR20.1)."], "out"),
            ("6 | Benchmark and explain", ["Compare bank aggregation with a transparent reference; trace "
                                           "differences to inputs, parameters and aggregation."], "out")],
        "Sources: Basel MAR11 and MAR20–23.",
        "Framework illustration. Product scope, methodology profile and validation evidence must be stated "
        "separately.")


def frtb_aggregation(out: Path) -> None:
    d = Diagram("FRTB_aggregation", "How FRTB-SA aggregates risk inputs",
                "Three calculation branches; one standardized capital requirement", 1352)
    d.bar(140, 1, "ASSIGN INPUTS", "keep valuation and risk-input production upstream")
    for x, t, b in [(40, "SBM | RISK FACTORS", ["Delta and vega sensitivities; curvature inputs. Map to risk "
                                                "classes, factors and buckets."]),
                    (555, "DRC | POSITIONS", ["Jump-to-default inputs, issuer, seniority, maturity and bucket."]),
                    (1070, "RRAO | INSTRUMENTS", ["Gross notional and residual-risk category. Apply the "
                                                  "regulatory scope tests."])]:
        d.card(x, 206, 490, 170, t, b)
    for x in [285, 800, 1315]:
        d.arrow(x, 376, x, 419)
    d.bar(424, 2, "CALCULATE EACH BRANCH", "use the appropriate rules within each branch")
    d.card(40, 490, 490, 330, "SBM | MAR21", [
        "Net sensitivities by risk factor; apply regulatory risk weights.",
        "Aggregate within buckets, then across buckets, by risk class and risk measure.",
        "For curvature, apply its prescribed upward/downward shock and aggregation rules."])
    d.card(555, 490, 490, 330, "DRC | MAR22", [
        "Determine gross jump-to-default and apply permitted netting.",
        "Separate non-securitizations, securitizations and the correlation trading portfolio.",
        "Apply default risk weights by credit quality and the hedge benefit ratio per bucket."])
    d.card(1070, 490, 490, 330, "RRAO | MAR23", [
        "Identify instruments in the residual-risk scope.",
        "Multiply gross notional by the applicable residual-risk weight.",
        "Sum the instrument charges; no offset against SBM or DRC."])
    for x in [285, 800, 1315]:
        d.arrow(x, 820, x, 858)
    d.card(40, 868, 490, 180, "K_{SBM}", [
        "For each correlation scenario, sum delta, vega and curvature across risk classes.",
        "Take the maximum scenario total."], dark=True, size=21)
    d.card(555, 868, 490, 180, "K_{DRC}", ["Sum the applicable default-risk capital components."], dark=True)
    d.card(1070, 868, 490, 180, "K_{RRAO}", ["Sum residual-risk charges on gross notionals."], dark=True)
    for x in [285, 800, 1315]:
        d.arrow(x, 1048, x, 1086)
    d.bar(1092, 3, "ADD THE THREE CAPITAL COMPONENTS")
    d.text(800, 1182, "K_{FRTB-SA} = K_{SBM} + K_{DRC} + K_{RRAO}", 34, NAVY, True, "middle")
    d.text(800, 1226, "RWA_{FRTB-SA} = 12.5 × K_{FRTB-SA}", 27, BLUE, True, "middle")
    d.foot("Sources: Basel MAR20.1 and MAR20.4; MAR21, MAR22 and MAR23.",
           "Full framework shown. The K-FRTB-ASA prototype covers a subset; this diagram does not expand its "
           "implemented scope.")
    d.save(out)


def cva_uses(out: Path) -> None:
    hub(out, "CVA_uses", "Where CVA risk fits in the capital framework",
        "A separate charge for adverse changes in credit valuation adjustment",
        "K-CVA | REGULATORY CVA CAPITAL",
        "CVA valuation, CCR exposure and CVA capital are distinct measures", [
            ("1 | The risk being capitalized", ["CVA changes with counterparty credit quality and the market "
                                                "factors driving derivative exposure."], "in"),
            ("2 | Covered transactions", ["Covered derivatives (QCCP trades excluded) and fair-valued SFTs "
                                          "where CVA risk is material (MAR50.5)."], "in"),
            ("3 | BA-CVA route", ["Supplied exposure values, effective maturities and counterparty "
                                  "characteristics. Full BA-CVA recognizes eligible credit hedges."], "calc"),
            ("4 | SA-CVA route", ["CVA and hedge sensitivities; supervisory approval required. Carved-out "
                                  "netting sets use BA-CVA (MAR50.8)."], "calc"),
            ("5 | Capital and RWA", ["CVA own funds requirement times 12.5. Up to EUR 100 billion "
                                     "non-centrally cleared notional: 100% of CCR capital (MAR50.9)."], "out"),
            ("6 | Link to exposure engines", ["SA-CCR exposure can feed BA-CVA. SA-CVA needs CVA "
                                              "sensitivities; EAD alone is insufficient."], "in")],
        "Sources: Basel MAR50; CRR Part Three, Title VI; Bundesbank, CVA risk.",
        "Method map. Eligibility, exemptions and the simplified route depend on the applicable regulatory "
        "profile.")


def cva_aggregation(out: Path) -> None:
    d = Diagram("CVA_aggregation", "How CVA capital aggregates its inputs",
                "Choose the applicable approach for each covered netting set", 1340)
    d.bar(140, 1, "DEFINE COVERAGE AND METHOD", "identify exemptions and eligible hedges first")
    d.card(40, 206, 740, 140, "BA-CVA | EXPOSURE-BASED INPUTS", [
        "Netting-set exposure values, effective maturities, counterparty characteristics and eligible "
        "credit hedges."])
    d.card(820, 206, 740, 140, "SA-CVA | SENSITIVITY-BASED INPUTS", [
        "Supplied regulatory CVA sensitivities and eligible hedge sensitivities; supervisory approval "
        "required (MAR50.7)."])
    d.card(40, 362, 1520, 112, "ALTERNATIVES AND CARVE-OUTS", [
        "At or below EUR 100 billion of non-centrally cleared notional, CVA capital may equal 100% of CCR "
        "capital, with no hedge recognition (MAR50.9). With SA-CVA approval, carved-out netting sets use "
        "BA-CVA (MAR50.8)."], size=21)
    d.arrow(410, 474, 410, 515)
    d.arrow(1190, 474, 1190, 515)
    d.bar(520, 2, "AGGREGATE UNDER THE CHOSEN APPROACH", "BA-CVA and SA-CVA are alternative routes")
    d.card(40, 586, 740, 290, "BA-CVA", [
        "Build the stand-alone counterparty risk amounts from the relevant exposure and maturity inputs.",
        "Reduced version: aggregate counterparty amounts with prescribed dependence; no hedge recognition.",
        "Full version: recognize eligible credit hedges and retain the reduced-version component that "
        "limits hedge benefit."])
    d.card(820, 586, 740, 290, "SA-CVA", [
        "Map CVA and hedge sensitivities to risk classes, risk factors and buckets.",
        "Apply prescribed weights, hedge offsets and within-/across-bucket aggregation.",
        "Aggregate delta and vega requirements. This route has no curvature or default-risk component."])
    d.arrow(410, 876, 410, 915)
    d.arrow(1190, 876, 1190, 915)
    d.card(40, 925, 740, 120, "BA-CVA CAPITAL REQUIREMENT", [
        "Apply the selected version and relevant regulatory adjustments."], dark=True, size=21)
    d.card(820, 925, 740, 120, "SA-CVA CAPITAL REQUIREMENT", [
        "Apply the prescribed aggregation and regulatory adjustments."], dark=True, size=21)
    d.bar(1090, 3, "REPORT CVA CAPITAL SEPARATELY", "never capitalize a netting set twice")
    d.text(800, 1185, "RWA_{CVA} = 12.5 × K_{CVA}", 34, NAVY, True, "middle")
    d.text(800, 1235, "CVA capital supplements the bank-wide requirement; it is outside SBM + DRC + RRAO.",
           24, BLUE, True, "middle")
    d.foot("Sources: Basel MAR50; CRR Part Three, Title VI; Bundesbank, CVA risk.",
           "Conceptual calculation map. It does not assert that either approach is currently implemented "
           "in K-CVA.")
    d.save(out)


def main(argv: Sequence[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args(argv)
    out = args.output.resolve()
    out.mkdir(parents=True, exist_ok=True)
    for build in (framework, frtb_uses, frtb_aggregation, cva_uses, cva_aggregation):
        build(out)
    print(out)


if __name__ == "__main__":
    main()
