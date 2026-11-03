# 📐 Methodology

[![Framework: Basel MAR21-MAR23](https://img.shields.io/badge/framework-Basel%20MAR21--MAR23-1f3a5f)](#scope)
[![Parameters: checked against BIS](https://img.shields.io/badge/parameters-checked%20against%20BIS-2ea44f)](#bis-check)
[![Status: prototype](https://img.shields.io/badge/status-prototype-D93F0B)](#interpretations-and-simplifications-to-review)

This document is authoritative for **what K-FRTB-SA calculates**: the
formulas, the default parameters and the interpretations still open. It
follows the Basel Framework chapters MAR21 (sensitivities-based method),
MAR22 (default risk capital) and MAR23 (residual risk add-on), in force since
1 January 2023. Paragraph references below point to those chapters.

> [!IMPORTANT]
> Parameters are editable defaults on the `P_*` sheets. They were checked
> against the BIS text on 2026-10-17 ([BIS check](#bis-check)); verify them
> against the regulation that applies to you (Basel, CRR3 or another
> jurisdiction) before relying on any figure.

```text
Total = SBM + DRC (non-securitization) + RRAO
SBM   = max over {low, medium, high} correlation scenarios of
        sum over risk classes of (delta + vega + curvature)      (MAR21.7)
```

<a id="scope"></a>

## 🧭 Scope

| Component | Implemented | Not implemented |
| --- | --- | --- |
| SBM risk classes | GIRR, CSR non-securitization, equity, commodity, FX | CSR securitization (non-CTP), CSR securitization (CTP) |
| SBM measures | Delta, vega, curvature | Computation of sensitivities and CVRs from pricing models |
| DRC | Non-securitizations | Securitizations (non-CTP), CTP |
| RRAO | Exotic and other residual risks | Automatic identification of RRAO instruments |

Rows for unsupported risk classes are rejected and logged as errors, so a
missing component is never silently counted as zero.

## 📥 Inputs

Sensitivities are supplied in the reporting currency on `In_Delta`,
`In_Vega` and `In_Curvature`; rows with the same risk factor are netted
first (MAR21.4(2)). The workbook does not compute sensitivities: they must
follow MAR21.19-MAR21.26 (for example, GIRR delta is the PV change for a
1 bp shift divided by 0.0001; equity delta is the PV change for a 1% relative
shift divided by 0.01). Vega inputs are vega times implied volatility
(MAR21.25). Curvature inputs are CVR+ and CVR- (MAR21.5(2)), net across
instruments for each risk factor.

| Risk class | Bucket | Qualifier | Delta sub-type | Tenor / maturity |
| --- | --- | --- | --- | --- |
| GIRR | Currency | Curve name | `YIELD`, `INFLATION`, `XCCY` | Yield tenors 0.25-30y |
| CSR_NS | 1-18 | Issuer or index | `BOND`, `CDS` | 0.5, 1, 3, 5, 10y |
| EQ | 1-13 | Issuer or index | `SPOT`, `REPO` | n/a |
| COM | 1-11 | Commodity | Delivery location | 0-30y |
| FX | Currency | (= bucket) | n/a | n/a |

The FX bucket is the currency's ISO code. Onshore and offshore, and
deliverable and non-deliverable variants of a currency are one risk factor
(MAR21.14(4) and its FAQ): enter KRW, not KRO, for every Korean won
sensitivity.

Vega risk factors are the underlying and the option maturity (0.5, 1, 3, 5,
10y). GIRR vega has three sub-types (MAR21.8(4) and its FAQ4):

* `YIELD`, the default: one surface per currency, by option maturity and
  underlying residual maturity on the same grid;
* `INFLATION`, one risk factor per inflation curve (the qualifier), and
  `XCCY`, the cross-currency basis: the option maturity only. A row of either
  with an underlying maturity is rejected (#209).
Curvature risk factors are one per currency for GIRR (all curves shifted
together; no inflation or cross-currency basis curvature) and FX, one per
issuer for CSR (bond and CDS curves together), one per issuer for equity and
one per commodity (all delivery locations together) (MAR21.8-MAR21.14).

### Curvature input contract (MAR21.5, MAR21.98, MAR21.99)

The workbook does not price instruments and applies no curvature risk
weight, so the system that computes the CVRs must apply the prescribed
shocks (#29, item 6):

| Risk class | Shock to the curvature risk factor | Delta deducted for each instrument |
| --- | --- | --- |
| FX, EQ | Relative shift equal to the delta risk weight of the risk factor (MAR21.98) | Its delta sensitivity to the risk factor |
| GIRR, CSR_NS | Parallel shift of all tenors of every curve of the risk factor by the highest delta risk weight of the bucket; for GIRR, the 0.25-year weight (MAR21.99) | The sum of its delta sensitivities to all tenors of the relevant curves |
| COM | Relative shift of the constructed curve: every tenor price multiplied by (1 + RW) up and (1 - RW) down, RW the delta risk weight of the commodity's bucket, with no term-structure decomposition (MAR21.99 and its first FAQ, MAR21.23 FAQ1). Gold at 20%: each price times 1.2 and 0.8 | The sum of its delta sensitivities to all tenors and delivery locations |

- **Sign.** CVRs are positive for a loss (MAR21.5(1)):
  `CVR_k+ = -sum_i [V_i(x_k^(RW+)) - V_i(x_k) - RW_k * s_ik]` and
  `CVR_k- = -sum_i [V_i(x_k^(RW-)) - V_i(x_k) + RW_k * s_ik]`.
- **Granularity.** One row per curvature risk factor, as defined above, net
  across instruments. For FX, use up to two rows per risk factor: one for the
  options that reference the reporting currency, with `RefReportingCcy` TRUE,
  and one for the others; the rows are netted after the divisor.
- **FX divisor.** Enter FX CVRs before any division by 1.5: the workbook
  applies `FX_CURVATURE_DIVISOR` or `FX_CURVATURE_DIVISOR_ALL` itself.
  `RefReportingCcy` accepts TRUE, FALSE or blank (FALSE), as a Boolean cell
  or as text in any case; any other value rejects the row, including YES,
  NO, Y, N, 1 and 0.
- **CSR floor.** When the downward shock would make a credit spread
  negative, the bank may floor the CSR curvature risk factor at zero after
  the shock. It then caps the risk weight used in `CVR-` at the level of the
  risk factor, and leaves the risk weight in `CVR+` uncapped (MAR21.99 FAQ2).
- **sqrt(2) discretions.** MAR21.98 and MAR21.99 do not say whether the
  sqrt(2) divisions of MAR21.44 (GIRR) and MAR21.88 (FX) also reduce the
  curvature shock. The workbook takes no position: the upstream pricing must
  follow the reading of the bank's supervisor and record it.
- **Scenarios.** The shocks do not depend on the correlation scenario; only
  the aggregation does (MAR21.6, MAR21.100, MAR21.101).

## 🧮 Sensitivities-based method

### Aggregation (MAR21.4)

For delta and vega, with weighted sensitivity `WS_k = RW_k * s_k`:

```
K_b = sqrt( max(0, sum_k WS_k^2 + sum_k sum_{l!=k} rho_kl WS_k WS_l) )
S_b = sum_k WS_k
Charge = sqrt( sum_b K_b^2 + sum_b sum_{c!=b} gamma_bc S_b S_c )
```

If the cross-bucket radicand is negative, the alternative specification
`S_b = max(min(sum WS_k, K_b), -K_b)` is used for every bucket
(MAR21.4(5)(b)); `Out_Buckets` marks those rows `ALT_SB` and
`Out_RiskClass` sets `AltSbUsed = Y`.

The "other sector" buckets (CSR 16, equity 11) use `K_b = sum |WS_k|` and are
added to the risk-class charge outside the square root, without
diversification (MAR21.56, MAR21.79).

### Curvature (MAR21.5)

```
K_b+ = sqrt( max(0, sum_k max(CVR_k+, 0)^2 + sum_k sum_{l!=k} rho_kl CVR_k+ CVR_l+ psi(CVR_k+, CVR_l+)) )
K_b  = max(K_b+, K_b-)
Charge = sqrt( max(0, sum_b K_b^2 + sum_b sum_{c!=b} gamma_bc S_b S_c psi(S_b, S_c)) )
```

`psi(x, y) = 0` when both arguments are negative, otherwise 1. The bucket
uses the scenario with the larger `K_b`; on a tie, the upward scenario if
`sum CVR+ > sum CVR-`, otherwise the downward one. `S_b` is the sum of the
CVRs of the selected scenario. For the other-sector buckets,
`K_b = max(sum max(CVR+, 0), sum max(CVR-, 0))`.

### Correlations

Within a bucket, vega and curvature use only the dimension their risk
factors share with delta (MAR21.94, MAR21.100 and their FAQs): the name
correlation for CSR and equity, the commodity correlation for commodity.

```
vega:       rho_kl = min(rho_option_maturity x rho_underlying, 1)
            rho_option_maturity = exp(-1% |Tk - Tl| / min(Tk, Tl))
            GIRR, two YIELD surfaces: rho_underlying uses the underlying residual
              maturity, same formula, no floor
            GIRR, other pairs: rho_underlying is the delta correlation of the
              corresponding risk factors: 1 for the same curve, 99.9% between two
              inflation curves, 40% between yield and inflation, 0% with XCCY
              (MAR21.46-MAR21.49)
curvature:  rho_kl = (rho_underlying)^2,  gamma_bc = (gamma_bc)^2
```

### Correlation scenarios (MAR21.6)

Every correlation (`rho` and `gamma`, after squaring for curvature) is
transformed per scenario: high `min(1.25 rho, 1)`, medium `rho`, low
`max(2 rho - 1, 0.75 rho)`.

### Risk weights and correlations

| Risk class | Risk weights | Within-bucket correlation | Cross-bucket |
| --- | --- | --- | --- |
| GIRR | 1.7%-1.1% by tenor; inflation and cross-currency basis 1.6% (MAR21.42-43) | `max(exp(-3% abs(Tk-Tl) / min(Tk,Tl)), 40%)`; x 99.9% across curves (also between two inflation curves); inflation vs yield 40%; cross-currency basis 0% (MAR21.45-49) | 50% (MAR21.50) |
| CSR_NS | 0.5%-12% by bucket (MAR21.53 Table 4) | name 35% (indices 80%) x tenor 65% x bond/CDS basis 99.9% (MAR21.54-55) | rating factor 50% for IG vs HY/NR when both buckets are in 1-15, otherwise 100%, x sector (MAR21.57 Table 5) |
| EQ | Spot 15%-70%, repo spot/100 (MAR21.77 Table 10) | names 15% / 25% / 7.5% / 12.5% / 80% by bucket; spot vs repo 99.9% (MAR21.78) | 15%; 0% with bucket 11; 75% between 12 and 13; 45% between 12/13 and 1-10 (MAR21.80) |
| COM | 20%-80% by bucket (MAR21.82 Table 11) | commodity (by bucket) x tenor 99% x delivery location 99.9% (MAR21.83) | 20%; 0% with bucket 11 (MAR21.85) |
| FX | 15% (MAR21.87) | n/a (one risk factor per bucket) | 60% (MAR21.89) |

Vega risk weight: `min(55% x sqrt(LH / 10), 100%)` with liquidity horizons
GIRR 60, CSR 120, equity large cap and indices (buckets 1-8, 12, 13) 20,
equity small cap and other (9-11) 60, commodity 120, FX 40 days (MAR21.92
Table 13): 100% everywhere except 77.78% for equity large cap and indices.

### Discretions (Control sheet)

| Setting | Effect |
| --- | --- |
| `GIRR_SPECIFIED_CCY_DISCRETION` | GIRR delta risk weights, including the inflation and cross-currency basis weights, divided by sqrt(2) for EUR, USD, GBP, AUD, JPY, SEK, CAD and the reporting currency (MAR21.44; #136) |
| `FX_SPECIFIED_PAIR_DISCRETION` | FX risk weight divided by sqrt(2) for the specified pairs and their first-order crosses (MAR21.88, fn 23) |
| `FX_CURVATURE_DIVISOR` | FX CVRs of options that do not reference the reporting currency (`RefReportingCcy` not TRUE on `In_Curvature`) divided by 1.5 (MAR21.98) |
| `FX_CURVATURE_DIVISOR_ALL` | All FX CVRs divided by 1.5: the MAR21.98 alternative that needs supervisory approval and curvature sensitivities for all currencies, including shocks to the reporting currency; it overrides `FX_CURVATURE_DIVISOR` |

## 💳 Default risk capital, non-securitizations (MAR22)

1. Gross jump-to-default per position: `JTD = LGD x notional + (market value - notional)`,
   floored at zero for longs and capped at zero for shorts (signed amounts;
   MAR22.11-13). LGD defaults by seniority: equity and non-senior 100%,
   senior 75%, covered bonds 25% (MAR22.12); a per-row LGD overrides it.
2. Maturity scaling: `min(max(M, 0.25), 1)` (MAR22.15-18). Cash equity
   (`EQUITY` seniority) takes a maturity of three months or at least one
   year (MAR22.16 and its FAQ); any other maturity rejects the row (#134).
3. Netting per obligor: a short offsets a long only when it ranks the same as
   or junior to the long (MAR22.19). An obligor belongs to one bucket: one
   whose positions name different buckets is rejected with all its positions
   (MAR22.22). When its positions differ in credit quality, the worst one, with
   the highest risk weight, applies with a warning (MAR22.24; #142). Shorts are applied from the most junior
   rank upward, which maximizes the eligible offset.
4. Per bucket (corporates, sovereigns, local governments and municipalities):
   `HBR = sum net long / (sum net long + sum |net short|)` and
   `DRC_b = max(sum RW x net long - HBR x sum RW x |net short|, 0)`; the
   total is the simple sum of buckets (MAR22.22-26). Risk weights by credit
   quality: AAA 0.5%, AA 2%, A 3%, BBB 6%, BB 15%, B 30%, CCC 50%, unrated
   15%, defaulted 100% (MAR22.24 Table 2).

## 🧩 Residual risk add-on (MAR23)

`RRAO = sum gross notional x RW`, with RW 1.0% for exotic underlyings and
0.1% for other residual risks (MAR23.8).

<a id="bis-check"></a>

## ✅ BIS check

On 2026-10-17 every parameter and rule above was compared with the BIS
Basel Framework pages (MAR21 republished 23 March 2026 with FAQ changes only;
MAR22 and MAR23 as of 27 March 2020). The check confirmed the risk weights,
correlations, scenario formulas, curvature rules, DRC and RRAO parameters,
and led to these corrections in this version:

| Item | Correction |
| --- | --- |
| CSR sector table (MAR21.57 Table 5) | Column for bucket 16 added (0% with every bucket); buckets 1-8 vs 17 set to 45% |
| CSR rating factor | 50% only between two buckets of 1-15 with different credit quality; 17 vs 18 is 75% |
| Equity vega liquidity horizon | Index bucket 13 is large cap (20 days), as bucket 12 |
| Vega and curvature correlations | Name only for CSR (no bond/CDS basis), commodity only for commodity (no location) |
| CSR and commodity curvature risk factors | Bond and CDS curves of an issuer, and delivery locations of a commodity, form one risk factor |
| FX specified pairs | First-order crosses of the specified pairs qualify |

On 2026-11-01 I reviewed the 59 FAQs of MAR21, MAR22 and MAR23 in the Basel
Framework MAR standard with FAQs (bis.org, generated 31 October 2026); the
FAQ version of the January 2019 standard (BCBS d457) is no longer published
separately, and its FAQs are part of that text. The implementation already
follows the FAQs on the equity vega liquidity horizons (MAR21.92), the
correlation of inflation curves in one currency (MAR21.47), the bond-CDS
basis (MAR21.9, MAR21.54), the absence of a floor on GIRR and CSR risk
weights (MAR21.8, MAR21.9), the vega correlation (MAR21.94), the cash-equity
maturity (MAR22.16) and conflicting ratings (MAR22.24). The others concern
how a bank computes sensitivities or classifies instruments, which the
workbook takes as inputs. The review led to:

| Item | Change |
| --- | --- |
| Commodity curvature shock (MAR21.23 FAQ1, MAR21.99 FAQ1) | The [curvature input contract](#curvature-input-contract-mar215-mar2198-mar2199) states the relative shift (#210) |
| CSR curvature floor (MAR21.99 FAQ2) | The contract states the permitted floor at zero and its risk-weight cap (#210) |
| FX currency variants (MAR21.14 FAQ2) | One risk factor per currency, whatever the variant (#210) |
| GIRR vega for inflation and cross-currency basis (MAR21.8 FAQ4) | `INFLATION` and `XCCY` vega risk factors along the option maturity only (#209) |

<a id="interpretations-and-simplifications-to-review"></a>

## 🔎 Interpretations and simplifications to review

1. **GIRR sqrt(2) discretion scope.** Decided on October 27, 2026 (#29,
   item 1) and implemented in #136: MAR21.44 divides "the above risk
   weights", which follow the inflation and cross-currency basis weights of
   MAR21.43, so the discretion divides those weights too.
2. **Covered bonds.** The 1.5% risk weight for covered bonds rated AA- or
   better (MAR21.53 fn 17, bank's discretion) is not modeled: bucket 8 uses
   2.5%. Decided on October 27, 2026 (#29, item 2) and documented in #140:
   an override of the bucket 8 risk weight on `P_CSR` changes every covered
   bond in the bucket, whatever its rating, so use it only when all of them
   are rated AA- or better. Modeling the discretion per position is a v0.3.0
   scope item only if covered bonds are in the certified product scope.
3. **FX curvature divisor.** Decided on October 27, 2026 (#29, item 3) and
   implemented in #144: `FX_CURVATURE_DIVISOR` divides only the rows not
   flagged as referencing the reporting currency, and
   `FX_CURVATURE_DIVISOR_ALL`, labeled as needing supervisory approval,
   divides every FX row. The workbook does not check that the curvature
   sensitivities cover all currencies, a condition of that approval.
4. **DRC obligor attributes.** Decided on October 27, 2026 (#29, items 4a
   and 4b) and implemented in #142: an obligor split across buckets is an
   input error, and its positions are skipped; differing credit qualities
   resolve to the worst, with a warning. The full MAR22.24 FAQ procedure for
   several external ratings is not modeled, since the input carries one
   credit quality per position. Blank maturity defaults to one year. Cash
   equity: decided on October 27, 2026 (#29, item 4c) and implemented in
   #134; a maturity other than 0.25 or at least 1 rejects the row.
5. **Binding scenario ties** resolve in the order low, medium, high. Decided
   on October 27, 2026 (#29, item 5) as a reporting convention, not a reading
   of MAR21.7: the capital is the same, only the reported scenario and its
   breakdown depend on the order. Since #138, `Results` notes a tie and names
   the tied scenarios.
6. **Curvature risk weights** are not used: CVRs are inputs computed by the
   bank. Decided on October 27, 2026 (#29, item 6) and documented in #140:
   the bank's pricing follows the
   [curvature input contract](#curvature-input-contract-mar215-mar2198-mar2199),
   and the workbook takes no position on whether the sqrt(2) divisions
   reduce the curvature shock.
