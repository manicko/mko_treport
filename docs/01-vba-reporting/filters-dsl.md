---
id: treport-filters-dsl
domain: vba-reporting
tags:
  - vba
  - excel
  - filters
  - dsl
  - pivots
related:
  - treport-configuration-model
  - treport-pipeline-dataflow
  - vba-reporting-overview
---

## Purpose

Describes the filter domain-specific language (DSL) used in `data_files.filters` and the high-level resolution semantics for filter values and derived period placeholders. This is the SSOT for DSL syntax/semantics in governed docs.

## Main Concepts

- Filters are configured per `data_files` row and parsed/applied by `modTreportFilters` (with pivot operations delegated to `modTreportPivot`).
- Parsing and evaluation are deterministic and Trim-normalized.
- Resolution is dependency-aware for period bounds (e.g. deriving maxima in the context of selected year/period constraints).
- Validation is fail-fast for malformed expressions, duplicate fields, missing/unresolvable required values, and other contextual errors.

## Syntax

Filters consist of clauses separated by `|`.

- **Clause form:** `<field>=<expr>`
- **Clause separator:** `|` (multiple clauses)
- **List delimiter:** configured `general.delimeter` (value from configuration) used inside list expressions
- **Range form:** parenthesized bounds `(start-end)` (e.g. scalar/numeric/period bounds as supported by implementation). A bound token may be `min`, `max`, a numeric literal, or `max - N` (offset from the field's resolved maximum).
- **Special bounds:** `min` and `max` tokens in range/scalar contexts where applicable; `max - N` supported as an offset form (see below).
- **Whitespace:** trimmed/normalized around tokens. For `max - N`, optional whitespace around `-` is accepted in supported forms.

### Expression forms (high-level)

| Form | Example shape | Notes |
|---|---|---|
| Scalar | `Field=value` | Single value match/selection semantics. |
| List | `Field=v1&v2&v3` (delimiter per config) | Multiple explicit values. |
| Range | `Field=(min-max)`, `Field=(1-max-1)` | Inclusive bound semantics as implemented; supports `min`/`max` and `max-N` bounds (resolved against pre-resolved per-field numeric domain). |
| Scalar (max/offset) | `Field=max` or `Field=max-N` | Resolves to the field maximum or `max_value - N` (exact domain membership required). |
| Mixed clauses | `Year=2024|Week Number=(1-max-3)` | Clauses combined with `|`. |

> Exact matching/typing behavior depends on pivot field members and implementation normalization; runtime equivalence is not fully verified in this audit.

## Resolution Semantics

- **Field/value normalization:** values and field tokens are normalized/trimmed before resolution/matching.
- **`min` / `max` resolution:** resolved against available (typically numeric) pivot members for the relevant field/context, using pre-resolved per-field global numeric domain/state captured before clause application (deterministic ordering/selection semantics preserved).
- **`max - N` offsets:** supported wherever `max` is a valid bound/value token in `data_files.filters` (scalar and range contexts). `max_value` is the resolved maximum of the referenced field from pre-resolved global context.
  - **D1 (bounds):** `N` is a base-10 non-negative integer (`0 <= N < max_value`). `N == 0` is equivalent to `max`. `N >= max_value` fails fast with clear context.
  - **D2 (domain presence):** for every resolved `max-N` bound token, the computed exact value (`max_value - N`) must be present among the field's numeric domain items (exact membership consistent with domain semantics). If not present, fail fast with clear context (field, resolved max, N, computed value, clause/expression/token as applicable).
  - `min-N`, `+N`, and offsets on numeric literals are out of scope. `max-N` semantics are not introduced for list elements (list element tokens remain literal per existing behavior).
  - Malformed/unresolvable offset forms fail fast with clear context. Existing valid expressions remain supported (including `(1-max)`, `(min-max)`, numeric/signed ranges as applicable); offset parsing is deterministic and unambiguous with range separators.
- **Dependency-aware period resolution:** when `Year` context is present, year-related resolution/selection is applied before deriving dependent maxima such as `Month Number` and `Week Number` (as implemented in resolution logic).
- **Derived placeholders in `report_sheet`:** `<max_<field>>` and `<max-N_<field>>` are resolved separately for output naming in the report-sheet context (pre-filter, independent of `data_files.filters`) with aligned D1/D2 rules; see [Configuration model](./configuration-model.md).
- **Determinism:** given the same resolved context and members, resolution is intended to be deterministic.

## Application Contexts

- **Report-level filters**: report attributes (e.g. `articleLevel4Name_clean`, `targetAudience`) are applied per report/output context without leaking across reports.
- **Uniform selection rule (D9, all applicable filters):** configured values/tokens are resolved and matched against available pivot context (for multi-value `articleLevel4Name_clean`, values are split by `general.delimeter` and matched exactly after trim). If at least one configured value is matchable, the matched subset is applied and processing proceeds. If a required field is missing/unavailable in the applicable context, or none of the configured values are matchable, processing fails fast with contextual diagnostics (no warn-and-skip). This applies uniformly to data-file and report-level filters, including `articleLevel4Name_clean` and `targetAudience`.

## Validation and Fail-fast

Expected fail-fast conditions include (non-exhaustive, high-level):

- Empty or malformed clause/expression structure.
- Duplicate field clauses where prohibited.
- Required fields/members not found or not resolvable in the current pivot context.
- No matchable configured values for an applicable filter (D9), or a required field missing/unavailable in the applicable context.
- Unresolvable `min`/`max` or `max-N` offset (bounds/domain violations, missing numeric domain, malformed/ambiguous forms).
- Unresolvable/invalid derived placeholders in `report_sheet` (including offset/domain failures) where required.
- No valid selectable data/values where required by processing rules.

Errors are raised with contextual diagnostics in the implementation.

## Runtime Verification Status

- Parser/classification/resolution and application logic exist in current exported VBA modules, but end-to-end correctness is not fully runtime-verified.
- Dataset-specific behavior remains unverified, including `report_wclip` period/field identity cases involving `Year` vs grouped `Years` and all pivot-member-dependent `min/max` derivations.

## Related Docs

- [Configuration model](./configuration-model.md)
- [Pipeline and data flow](./pipeline-and-dataflow.md)
- [VBA reporting overview](./overview.md)
