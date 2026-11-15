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
- **Range form:** parenthesized bounds `(start-end)` (e.g. scalar/date/period bounds as supported by implementation)
- **Special bounds:** `min` and `max` tokens in range/scalar contexts where applicable
- **Whitespace:** trimmed/normalized around tokens

### Expression forms (high-level)

| Form | Example shape | Notes |
|---|---|---|
| Scalar | `Field=value` | Single value match/selection semantics. |
| List | `Field=v1&v2&v3` (delimiter per config) | Multiple explicit values. |
| Range | `Field=(min-max)` or bounded range | Inclusive/bound semantics as implemented; supports `min`/`max` resolution. |
| Mixed clauses | `Year=2024|Week Number=(1-max)` | Clauses combined with `|`. |

> Exact matching/typing behavior depends on pivot field members and implementation normalization; runtime equivalence is not fully verified in this audit.

## Resolution Semantics

- **Field/value normalization:** values and field tokens are normalized/trimmed before resolution/matching.
- **`min` / `max` resolution:** resolved against available (typically numeric) pivot members for the relevant field/context.
- **Dependency-aware period resolution:** when `Year` context is present, year-related resolution/selection is applied before deriving dependent maxima such as `Month Number` and `Week Number` (as implemented in `ResolveMinMaxAndYear`/related logic).
- **Derived placeholders:** `max_year`, `max_monthnum`, and `max_weeknum` are derived from resolved context and used for `report_sheet` placeholder substitution (see [Configuration model](./configuration-model.md)).
- **Determinism:** given the same resolved context and members, resolution is intended to be deterministic.

## Application Contexts

- **Data-file filters (`data_files.filters`)**: parsed/resolved and applied once per source data file context during processing.
- **Report-level filters**: report attributes (e.g. `articleLevel4Name_clean`, `targetAudience`) are applied per report/output context without leaking across reports.
- **Optional value handling**: implementation includes optional skip handling for the token `"ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ"` when absent, provided at least one valid value remains (exact runtime behavior unverified).

## Validation and Fail-fast

Expected fail-fast conditions include (non-exhaustive, high-level):

- Empty or malformed clause/expression structure.
- Duplicate field clauses where prohibited.
- Required fields/members not found or not resolvable in the current pivot context.
- Unresolvable `min`/`max` or derived placeholder references.
- No valid selectable data/values where required by processing rules.

Errors are raised with contextual diagnostics in the implementation.

## Runtime Verification Status

- Parser/classification/resolution and application logic exist in current exported VBA modules, but end-to-end correctness is not fully runtime-verified.
- Dataset-specific behavior remains unverified, including `report_wclip` period/field identity cases involving `Year` vs grouped `Years` and all pivot-member-dependent `min/max` derivations.

## Related Docs

- [Configuration model](./configuration-model.md)
- [Pipeline and data flow](./pipeline-and-dataflow.md)
- [VBA reporting overview](./overview.md)
