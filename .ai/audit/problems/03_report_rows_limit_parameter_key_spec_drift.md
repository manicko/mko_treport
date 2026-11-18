---
id: audit-report-rows-limit-parameter-key-spec-drift
domain: audit
tags:
  - discrepancy
  - configuration
  - p01
  - spec-drift
related:
  - treport-configuration-model
  - treport-pipeline-dataflow
---

## Purpose

Record a specification/plan-to-implementation discrepancy for the P01 general per-subreport row-limit parameter key. This report is separate from governed documentation updates; docs/* must describe implemented behavior because code is the source of truth.

## Discrepancy

- **Requirement/spec implication (planned/specified):** P01 planning/specification materials for combine support describe the required general parameter as eport_rows_limit_general (explicit positive integer, required, no fallback/default/override), applied per prepared subreport.
- **Current implementation (source of truth):** exported VBA configuration parsing accepts either eport_rows_limit or eport_rows_limit_general as the general parameter key. The parameter remains required as a single entry (duplicate/repeated key is rejected), must be an explicit positive integer (Long-compatible, no fractional/Boolean/error/non-numeric/blank/invalid values), and has no fallback/default or per-file override. Semantics/validation of the limit value and its application are otherwise consistent with the implemented P01 combine flow.
- **Relevant implementation area:** mko_treport/src/vbs/export/modTreportConfig.bas (ReadConfig and general-limit validation/parsing logic).

## Impact

- **Contract/documentation:** specification wording is narrower than implemented key acceptance. Configuration authors may use either accepted key in the current implementation, but external specs/tests that assume only eport_rows_limit_general are not fully aligned with code.
- **Documentation handling:** governed docs should state the actual accepted key form(s) and required semantics factually; this does not change runtime behavior.

## Verification status

- Based on exported VBA source review in this documentation pass. Workbook-embedded VBA parity and end-to-end runtime behavior remain subject to build/L2/L3 verification on working copies.
