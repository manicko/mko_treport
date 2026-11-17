# Execution Plan — P09: `insert_rows` anchor/count rework (sample row `P`, insert `H-2`, paste at `P-1`)

- **Plan ID:** 008-p09-insert-rows-anchor-rework  
- **Date:** 2026-11-17 (decisions updated 2026-11-17 per PO input)  
- **Source plan file (authoritative input):** `C:\py_exp\mko_treport\.ai\problems\P09.md` (read in full)  
- **Status:** Execution plan (Step 1 complete, Block A decisions locked; proceeding per workflow). **No VBA/code or prior audit phase files changed by this plan document.**  
- **Primary in-scope VBA exports (edit targets only if/when implementation proceeds later):** `mko_treport/src/vbs/export/modTreportExtract.bas`, `mko_treport/src/vbs/export/modTreportRun.bas` (plus any directly coupled helper calls). Docs/spec lineage to reconcile (if scope remains as P09 states): `docs/01-vba-reporting/configuration-model.md`, `docs/01-vba-reporting/pipeline-and-dataflow.md`, and amendment/supersession note for `.ai/problems/005_p05_insert_rows_template_row_spec.md` (cross-reference P09).  
- **Hard constraints (must be enforced):** VBA exports only as edit targets; never claim compile success without **L2 manual VBE `Debug -> Compile VBAProject`** on a working copy; L1 via `mko_treport/src/scripts/build_treport_vba.ps1` with working copy/backup/report recorded; L3 on copies only (no originals). No line-number-based reasoning in implementation guidance—use semantic anchors (modules/procedures/branches/contracts). Follow repo rules: English docs/comments style in changes should match repo conventions; no speculative redesign beyond P09.

---

## 0) `{code_context}` summary (from Auditor input) — referenced, not re-audited here

Current baseline (P05/P07) summarized in P09 §1–§4 and confirmed in repo inspection:
- `modTreportRun.ProcessDataFile` (`:152–215`) normalizes `pr/pc` from `cfg.DataFiles(...).pasteRow/pasteCol` (`<=0`/invalid → `1`), computes `templateSet = (Len(Norm(template))>0)`, calls `modTreportExtract.CopyPivotSnapshot(pt, outWs, pr, pc, repCtx, onlyValues, insertRows, skipHeader)`, then `ApplySheetFormatting(..., preserveTargetSettings = templateSet?)`, then `ApplyOutlineGrouping insertedRng, group_rows, rows_collapsed, repCtx+context`.  
- `modTreportExtract.CopyPivotSnapshot` (`:16–105`): `GetPresentedRange(pt,ctx)` then optional `skip_header` via `pt.DataBodyRange` offset rules; guards/clamps `pr/pc<=1?` code shows `pr<=0 Then pr=1; pc<=0 Then pc=1` (`:31–35`); if `h<=0 or w<=0` fail-fast 2704 with context (`:77–80`). **If `insertRows And h > 1` (current P05):** copy entire row `P` (`dstWs.Rows(pr).Copy`) and insert **`h-1`** copies starting at `P` (`Rows(pr).Resize(h-1).Insert Shift:=xlDown`) via `EHRowPrep`; clears `CutCopyMode`. No helper-row deletion (comment `:17–20,141`). Snapshot at `dstWs.Cells(pr,pc)`, paste via `PasteSnapshotData`, return `Range(Cells(pr,pc), Cells(pr+h-1,pc+w-1))` (`:91–95`).  
- Docs/specs currently state P05 contract: inserts `H-1` at `P`, paste at `P`, final `P..P+H-1` (`docs/01-vba-reporting/configuration-model.md`, `pipeline-and-dataflow.md`, `.ai/problems/005_p05_insert_rows_template_row_spec.md`).

P09 target (conflict): `insert_rows=TRUE` (in scope per OP-1 assumption unless clarified) with sample row `P` inserted **`H-2` times**, snapshot pasted at **`(P-1, pc)`**, final range expected **`P-1 .. P-1+H-1`** (to be verified against actual Excel shift behavior), **no deletions/no shrink**, with `only_values`/`skip_header` semantics preserved conceptually except anchor/count.

---

## 1) Execution blocks (logical), dependencies, execution order

Execution is **planning + controlled verification design** in this step; implementation is not performed here. Blocks are ordered to resolve decisions first, then define contract changes, then implementation steps, then verification/reporting.

### Block A — Decision & scope lock (OP-1..OP-4) + contract freeze
- **Purpose:** Convert P09 open points into explicit, testable decisions before any code edit. Prevents mixing incompatible behaviors.
- **Key activities (no code changes):**
  - Resolve **OP-1 (scope):** confirm whether new anchor/count applies **only when `insert_rows=TRUE AND template is set`** (assumed in P09 §4.1 phrasing “if required by scope…” and §5 OP-1) **or** to **all** `insert_rows=TRUE` targets (including non-template/plain prepared sheets). Document the chosen scope as a single contract line used by code, docs, and tests.
  - Resolve **OP-2 (H):** define deterministic behavior for `H = snapshot.Rows.Count` after `skip_header` slicing: `H>=3 => insert H-2`; `H==2 => insert 0` (paste `2×W` at `P-1` assumed valid unless disproven); `H==1 => H-2<0` choose **fail-fast** OR **treat as 0 inserts with paste at `P-1`** (current behavior with old anchor is 0 inserts at `P`, paste at `P`). Must be explicit and tested.
  - Resolve **OP-3 (P==1):** confirm **fail-fast** preferred (cannot paste at `P-1`) vs alternative anchor rule. If fail-fast, define error code/message contract (context fields).
  - Resolve **OP-4 (contract wording):** confirm: `paste_row` remains the configured value `P`; actual snapshot top is `P-1` (when in scope); returned range and downstream (`ApplySheetFormatting`, `ApplyOutlineGrouping`) follow actual placement; sample row remains entire row `P` (copy source semantics unchanged in identity of sample row number as configured).
- **Dependencies:** None (start). Consumes P09, current code/docs, P05 lineage.
- **Outputs:** A short “P09 Decision Record” section to embed in implementation notes / plan appendix (in the saved plan file) listing OP-1..OP-4 final choices, rationale (business vs technical), and explicit non-goals.
- **Required agents (by risk):** **Planner** (to structure decisions and trace impacts). **Researcher** (only if Excel/VBA behavioral uncertainty for a specific choice is material, e.g. `P==1` alternative anchor rejected but we need to be precise about messages; usually low). **Auditor** not required unless discovery reveals hidden consumers/paths. **Validator** not required for decisions alone unless combined with high-risk design sign-off (see Block D).
- **Uncertainty / alternatives (do not pick final impl here except choices above):** OP-1 scope is the main behavioral fork (template-only vs global). OP-2 `H==1` is a semantic fork (fail-fast vs degraded behavior) with different downstream expectations.

### Block B — Contract & documentation delta design (spec-first, no code edits)
- **Purpose:** Define exact, semantic contract for code and docs so implementation is unambiguous and P05 lineage is reconciled.
- **Key activities:**
  - Specify semantic behavior for **in-scope** cases only (per OP-1):
    - Anchors: `P_cfg = paste_row` (unchanged config meaning), `P_anchor = P_cfg-1` for snapshot paste/destination top; validate `P_cfg >= 2` (or chosen rule) → fail-fast with contextual diagnostics including `operation='copy/insert sample row'` / parameter context aligned to current error patterns (e.g. mention `insert_rows`, `paste_row`, `H`, `W`, `skip_header`, `only_values`, `template`/scope).
    - Sample prep: copy entire row `P_cfg`; insert `max(H-2, 0)` times starting at `P_cfg` when in scope **and** rule allows (define the exact rule for `H==1` per OP-2). For `H==2`, insert `0`. **No deletions** (explicit).
    - Snapshot: `PasteSnapshotData srcRg, dstWs, dstWs.Cells(P_anchor, pc), onlyValues, ctx` with unchanged `only_values`/`skip_header` effects on `srcRg/H/W` derivation (R5). Final returned range = actual bounds after insertion + paste at `P_anchor` over `H×W` → **expected** `Cells(P_anchor,pc) .. Cells(P_anchor + H - 1, pc + W - 1)` unless verified otherwise (A1).
    - Error handling: extend/clarify messages and context keys consistent with `modTreportExtract` style (`Err.Raise 2705/2702/2704` patterns) for new cases (P_cfg<2, invalid H-2 policy outcomes). Decide error codes/messages now in plan text (A2) to avoid drift.
  - List required doc/spec updates (semantic, not line numbers):
    - `docs/01-vba-reporting/configuration-model.md` (`data_files.insert_rows` row): update description of insert count, anchor (`P-1` in scope), final range (`P-1 .. P-1+H-1` conceptually relative to configured P/top anchor rule), H edge rules, P==1 rule, and explicit “no deletion”.
    - `docs/01-vba-reporting/pipeline-and-dataflow.md`: update P04/P05 unified sequence bullets to reflect new anchor/count and note scope (OP-1) and verification note about Excel shift behavior (A1).
    - `.ai/problems/005_p05_insert_rows_template_row_spec.md`: add **amendment/supersession note** referencing P09 (do not silently overwrite history); state what changes (anchor, count, final range) and link to P09 for current contract. (If team prefers to record final resolved contract in P09 only, state that explicitly as an alternative—do not choose implementation file edits now except documentation plan.)
  - Define **downstream impact contract**: callers treat returned `Range` as authoritative final band (R4). No other callers known from P09; document “verify call sites” in Block C.
- **Dependencies:** Block A (decisions). 
- **Outputs:** “Contract delta” table in plan: (area, current statement, new statement, edge rules) and explicit list of files to touch later.
- **Required agents:** **Planner** (primary). **Auditor** recommended if any ambiguity remains about other code paths that may assume `P` as top row of inserted band beyond `modTreportRun` (search/confirm list). **Researcher** helpful only if documentation style/patterns need alignment (low). **Validator** optional at this stage.
- **Alternatives (multiple paths, no final choice):**
  1. **Strict template-only scope (OP-1 = yes):** code branches on `templateSet` (or passed indicator) to apply P09 rules only when template present; `insert_rows` without template keeps legacy P05 behavior. *Pros:* minimal behavioral surface, aligns with P09 phrasing. *Cons:* two behaviors to document/test; risk of confusion.
  2. **Global for all `insert_rows=TRUE`:** apply P09 anchor/count whenever `insert_rows=TRUE` regardless of template. *Pros:* single behavior, simpler mental model. *Cons:* breaks existing non-template workbooks relying on old semantics (higher regression/compat risk).
- **Open decisions embedded:** OP-1 drives Branch 1 vs 2; OP-2/3 drive numeric/error rules.

### Block C — Implementation design (VBA changes, semantic plan only — no coding in this step)
- **Purpose:** Specify concrete, semantic edits to achieve Block B contract safely (for a later implementer).
- **Key activities (design, not code):**
  - **`modTreportExtract.CopyPivotSnapshot` (semantic anchors):**
    - After deriving `pr, pc, h, w, srcRg, ctx` (and after `skip_header` path), compute in-scope decision using **Block A OP-1** (and whether to pass an explicit indicator from caller). If out-of-scope by OP-1, preserve current P05 behavior exactly.
    - **Anchor validation:** enforce `pr >= 2` (i.e. `paste_row >= 2`) when in-scope P09 path applies; on violation fail-fast with enriched context (include `paste_row`, `H/W`, `insert_rows`, `only_values`, `skip_header`, `template`/scope note, `operation='copy/insert sample row'` or new clear operation label). Preserve existing clamping only if explicitly decided otherwise (current clamps `pr/pc<=0→1`; P09 requires `P>=2` → conflict to resolve in Block A: does clamp still happen? Must be stated.)
    - **Row preparation count:** let `insCount = H - 2` (in-scope). Define handling of `H==2 => 0`, `H==1` per OP-2. If `insCount > 0`, copy entire row `P` and insert `insCount` copies starting at `P` (same Excel API intent as today but count/semantics changed). Handle `CutCopyMode`/error paths consistently with `EHRowPrep`.
    - **Paste destination:** `dstWs.Cells(pr-1, pc)` (i.e. `P-1`) in-scope; unchanged `PasteSnapshotData` call shape, pass same `onlyValues, ctx`.
    - **Return range:** compute final band as the actual snapshot area at anchor `P-1` over `H×W`: `dstWs.Range(dstWs.Cells(pr-1,pc), dstWs.Cells(pr-1 + h-1, pc + w-1))`. Note A1: if Excel shifts produce unexpected final coordinates in practice, this is a *verification finding* to be captured in L3; plan should still state intended coordinates and require proof.
    - **Error messages/codes:** specify new/adjusted fail-fast points (A2): e.g. error when `pr < 2` for in-scope P09 path (distinct message: “paste_row must be >= 2 for insert_rows with P09 anchor…”), and when `H==1` under chosen OP-2 rule if fail-fast.
    - Preserve `CutCopyMode` cleanup and existing error wrapping patterns.
  - **`modTreportRun.ProcessDataFile` (semantic anchors):**
    - If OP-1 requires template-only, current `templateSet` is the natural indicator; if global, remove/ignore template gating for P09 behavior. Pass explicit indicator to `CopyPivotSnapshot` only if needed by design (prefer minimal signature change vs overloading; note alternative: infer inside `CopyPivotSnapshot` by adding a parameter like `templateContext As Boolean` or by deriving from context—no final choice now).
    - Continue to pass returned `insertedRng` to `ApplySheetFormatting` and `ApplyOutlineGrouping` (R4) — no change to call pattern, only semantics of range.
  - **Call-site / impact audit:** confirm no other in-repo callers of `CopyPivotSnapshot` beyond `modTreportRun` (P09 states only observed there). If any found, list them as additional touch points (Auditor/Planner task).
  - **No deletions:** explicitly guard against reintroducing helper-row deletion logic (forbidden by P09).
- **Dependencies:** Block A, Block B.
- **Outputs:** “Implementation delta checklist” (by module/procedure, semantic steps) + “Decision-gated branches” table (OP-1/2/3 mapped to code paths).
- **Required agents:** **Planner** (primary). **Auditor** if call-site audit or hidden coupling suspected (recommended to run a quick repo grep for `CopyPivotSnapshot` usages before implementation). **Researcher** only if multiple viable VBA patterns for row insertion in this exact context create tradeoffs (rare). **Validator** recommended before coding starts when high risk (see below).
- **Alternatives (multiple paths, no final choice):**
  1. **Add a boolean parameter to `CopyPivotSnapshot` (e.g. `useP09Anchor As Boolean`)** passed from `Run` based on OP-1. *Pros:* explicit, testable. *Cons:* signature change + caller update.
  2. **Infer from `template` context already known to caller** by extending behavior without new param (e.g. rely on `insertRows` only in global scope). *Pros:* fewer signature changes. *Cons:* couples logic across modules or hides intent.
  3. **Handle H==1 as fail-fast vs as 0-insert+paste at P-1** (OP-2) — impacts error surface and user guidance; must be chosen in Block A.
- **Excel/version/environment notes (to carry into design):** `Rows.Insert Shift:=xlDown` behavior with merges/objects/outlines is Excel-version-sensitive in practice; templates may contain formulas/CF that shift. Plan must require L3 on working copies with real templates (D3, R3). Do not assume cross-locale behavior.

### Block D — Risk assessment & verification plan (L1/L2/L3 + gates)
- **Purpose:** Make risks explicit and bind verification to each risk/edge case; ensure project workflow compliance.
- **Key activities:**
  - **Risk register (tie to P09 D1,R1–R9 + environment):**
    - **R1 (P==1, paste at P-1):** blocked if OP-3 = fail-fast; if alternative chosen, needs precise definition. *Mitigation:* fail-fast + clear diagnostic; L3 test P==1.
    - **R2 (H edges):** H==1 (OP-2 decision), H==2 (insert 0, paste 2×W at P-1) — verify space created by shift of row P down by 0 in effect? (Excel behavior) and that sample row P becomes row P in the shifted sheet? Must validate: copying/inserting 0 times starting at P means no inserts, row P remains in place shifted? No—if you don’t insert rows, structure unchanged; pasting at P-1 means writing into row above original P — dangerous/confusing unless sheet prepared. This is a **high-sensitivity** edge: requires L3 proof.
    - **R3 (Excel structural side effects):** merges, formulas, conditional formatting, data validation, comments/notes, shapes, outlines interaction. *Mitigation:* L3 scenarios on template copies.
    - **R4 (downstream consumers):** `insertedRng` used for formatting/grouping — must be correct. *Mitigation:* L3 verify `ApplySheetFormatting`, `ApplyOutlineGrouping` on new ranges.
    - **R5 (only_values/skip_header):** `skip_header` changes `H,W,srcRg` before this logic; `only_values` affects paste only. *Mitigation:* L3 matrix includes `skip_header True/False`, `only_values True/False`.
    - **R6 (scope):** template-only vs global (OP-1). *Mitigation:* decision lock + regression tests for the non-chosen scope behavior.
    - **R7 (PQ/pivot runtime variability):** presented range varies by refresh/state; `GetPresentedRange`/`DataBodyRange` edge cases already handled? P09 builds on existing behavior — re-validate in L3 after changes if anchors change error paths.
    - **R8 (data safety):** only working copies, no originals; report paths recorded.
    - **R9 (doc/spec drift):** enforce Block B outputs before “done”.
  - **Environment/compat risks (explicit):** Excel 2021+/Windows per assumptions; macro security/VBOM trust required for L1 automation in practice but L2 is manual in VBE — document both; Power Query refresh timing is out of scope for this change but L3 should use refreshed source copies consistent with existing workflow.
  - **Verification plan (must be followed):**
    - **L1:** `.\mko_treport\src\scripts\build_treport_vba.ps1` with `-SourceWorkbook "data\theiss\treport.xlsm"` (or appropriate canonical) + `-WorkCopyPath ".ai\tmp\treport_vba_check_<YYYYMMDD_HHMMSS>.xlsm"` + `-BackupDir ".ai\tmp"` + `-ReportPath ".ai\tmp\build_treport_vba_<YYYYMMDD_HHMMSS>.txt"`. Record exact command and report summary (SUCCESS/ISSUES/ERROR; VBOM notes). Script compile note is best-effort only.
    - **L2 (required to claim compilation):** open working copy in Excel → VBA Editor → **Debug → Compile VBAProject**. Record file path, date/time, pass/fail and errors if any.
    - **L3 (on copies only), scenarios (minimum set) driven by OP-1..OP-3 and A5:**
      - In-scope (per OP-1): `insert_rows=TRUE`, template set (if template-only) with `H >= 3` (e.g. 3,4) → verify sample row `P` inserted `H-2`, paste at `P-1`, final range `P-1..P-1+H-1`, no deletions, `ApplySheetFormatting`/`ApplyOutlineGrouping` correct.
      - `H == 2` (OP-2): verify behavior matches decision (insert 0, paste 2×W at `P-1`) and downstream still correct.
      - `H == 1` per OP-2 (fail-fast or allow) — test accordingly.
      - `P == 1` per OP-3 — expect fail-fast (or defined alternative) with actionable message.
      - Variants: `only_values=TRUE` and `FALSE`; `skip_header` True/False if used in tested datasets; at least one case with grouping (`group_rows`) to confirm bounds; compare against out-of-scope cases unchanged (`insert_rows=FALSE`, and non-template if template-only scope) for regression.
      - Use copies of `data/theiss/*` style inputs only; originals untouched (R8).
  - **Report template (required):** list changed export files; exact `build_treport_vba.ps1` command; report path + short summary; L2 result (done with file/time or not done + limitation); L3 results per scenario or “not performed”.
- **Dependencies:** Block A–C defined (so scenarios are precise).
- **Outputs:** Risk register table (in plan), verification checklist, and explicit “gates” (e.g. do not proceed to implementation beyond Block C design until OP-1..OP-3 decided).
- **Required agents:** **Planner** (primary). **Validator** **required** when implementation risk is assessed as high (see gate below). **Auditor** optional if new risks discovered. **Researcher** optional for Excel behavior edge-case literature only if team wants external guidance (not required if L3 on real files is planned).
- **Gate (high-risk criteria triggers):** treat as **high risk** if any of: (a) OP-1 = global (changes all `insert_rows=TRUE`), (b) OP-2 allows `H==1` to proceed with `P-1` in a non-obvious way, (c) real templates have merges/formulas/CF/outlines in affected rows, (d) uncertainty about Excel shift returning expected `insertedRng` bounds (A1). If **high risk**, **Validator agent is required** to review this plan (and later implementation plan) before coding; also ensure **Auditor** has confirmed no hidden callers/assumptions (Block C).
- **Alternatives (verification):** none materially different; must follow repo workflow. Alternative test data sources only if `data/theiss` not available — must still be copies and documented.

### Block E — Rollout, sequencing, and “done” criteria
- **Purpose:** Define safe rollout order and acceptance.
- **Key activities:**
  - Recommended order (if implementation approved later): A → B → C (design review + Validator if high risk) → code edits in `modTreportExtract.bas` (+ minimal `modTreportRun.bas` if needed) → doc/spec updates (Block B list) → L1 → fix loop until clean import → L2 on working copy → L3 scenarios → produce final report.
  - **Done criteria:** all OP-1..OP-4 reflected in code/docs; no helper/deletion behavior reintroduced; L2 passed on recorded working copy; L3 scenarios executed for in-scope rules with recorded results; docs/spec lineage updated (or explicitly deferred with recorded reason—prefer not deferred).
- **Dependencies:** A–D.
- **Outputs:** Sequenced checklist in plan.
- **Required agents:** **Planner**. **Validator** sign-off if high risk before starting edits.

---

## 2) Dependencies and execution order (summary DAG)

- A (decisions) → B (contract/docs delta) → C (impl design) → **(gate: Validator if high risk)** → implementation → D verification execution (L1→L2→L3) → E report/done.  
- B can be drafted in parallel with finalizing A textually, but **must be reconciled to A before C**.  
- Docs updates (part of B/E) should be applied **after** code behavior is finalized in implementation, but contract is defined in B first.

---

## 3) Risks (implementation, rollout, regression, compatibility) — consolidated

- **Implementation:** wrong anchor/count mapping after `skip_header` slicing (uses effective `H,W,srcRg` as today — must remain consistent); off-by-one for `H-2`; forgetting to change return range computation.  
- **Rollout:** partial doc updates cause operator confusion (P05 vs P09).  
- **Regression:** non-template targets if OP-1 becomes global; any `insert_rows=FALSE` path accidentally touched (must not).  
- **Compatibility:** Excel row insert/copy quirks (merges, CF, formulas, shapes, tables/lists not expected but possible) — VBA/COM behavior varies by Excel build; Windows-only environment assumed. Macro security/VBOM trust affects L1 tooling; L2 is manual. Power Query refresh timing not changed but L3 should run on refreshed copies to reflect realistic state (R7). Workbook state/locking: run on local copies, close Excel as needed for script reliability. Output fidelity: visual differences possible if sample formatting differs from expectations — validate with human-reviewed L3 on representative templates.

---

## 4) Multiple implementation path alternatives (for maintainability/evolution/conventions) — **no final choice made**

1. **Scope branch inside `CopyPivotSnapshot` only (OP-1 = template-only):** use existing `templateSet`/indicator from `Run`; minimal cross-module change. *Tradeoff:* dual behavior to maintain/test.
2. **Scope branch driven by a new explicit parameter from `Run` to `CopyPivotSnapshot` (OP-1 global or template-only):** clearer API intent. *Tradeoff:* signature change; need to audit all callers (currently only `modTreportRun` per P09).
3. **Compute `insCount = H-2` with explicit branches for OP-2 (`H==1`, `H==2`):** centralized in one place. *Tradeoff:* must document all cases; alternative is to reject `H<3` broadly (stricter) — **not proposed** without decision.
4. **Error strategy:** reuse existing `Err.Raise` numbers/context style vs introduce new codes for P09-specific failures. *Recommendation direction (not final code):* prefer consistent context + clear messages; codes only if conventions require — **decision deferred to implementation design review**.
5. **Anchor validation timing:** validate `P>=2` after clamps? Current code clamps `pr/pc<=0→1` then later may fail for `H<=0`. P09 says “require `P >= 2` so `P-1 >= 1`”. This conflicts with silent clamp-to-1 behavior — **must be resolved in Block A** (treat invalid configured `paste_row<=1` as error in scope, or still clamp then require >=2 meaning clamp makes it 1 → error). Alternatives: (a) fail-fast if `pr<2` before any structural changes; (b) change clamp policy only in-scope. **No final choice.**

---

## 5) Agent requirements by block (and overall)

- **Block A (Decision & scope lock):** **Planner** required. **Researcher** optional (only if Excel behavior nuance forces a decision). **Auditor** not required. **Validator** not required.
- **Block B (Contract & doc delta design):** **Planner** required. **Auditor** recommended (verify no other assumptions about `P` as top of band outside `modTreportRun`/`CopyPivotSnapshot`). **Researcher** optional. **Validator** optional unless high-risk scope/global chosen early.
- **Block C (Implementation design):** **Planner** required. **Auditor** recommended for call-site/coupling audit. **Researcher** optional. **Validator** recommended before coding if **high risk** (see Block D gate).
- **Block D (Risk & verification):** **Planner** required. **Validator** **required if high risk** (gate). **Auditor** optional if new findings. **Researcher** optional for external Excel refs only.
- **Block E (Rollout/done):** **Planner** required. **Validator** sign-off if high risk.

**Overall:** At minimum **Planner** for all blocks. Expect **Validator** involvement before implementation if OP-1 is global or other high-risk triggers hold. **Auditor** should do a quick confirmation of call sites and any code/comments implying old `P..P+H-1` paste-at-`P` contract remains outside the listed files before edits.

---

## 6) Block A — Decision record (LOCKED, PO input 2026-11-17)

User decisions (Step 2): **Q1=A, Q2=B, Q3=A, Q4=B**.

- **OP-1 (scope) = A (template-only):** Apply new anchor/count rule **only when `insert_rows=TRUE` AND a template is set** (`Len(Norm(template)) > 0` context from caller). Non-template `insert_rows=TRUE` keeps current/P05 behavior (`H-1` inserts at `P`, paste at `P`, final `P..P+H-1`). *Rationale:* lower regression risk; aligns with Q4=B (preserve existing layouts where possible). Out-of-scope cases must remain unchanged.
- **OP-2 (H values) = B (allow H==1 with 0 inserts at P-1):** Deterministic rules for in-scope cases (after `skip_header` slicing, if applicable):
  - `H >= 3`: copy entire row `P` and insert **`H-2`** sample rows starting at `P`; paste snapshot `H x W` at **`(P-1, pc)`**.
  - `H == 2`: insert **0** sample rows; paste snapshot `2 x W` at **`(P-1, pc)`**.
  - `H == 1`: **allow** — insert **0** sample rows; paste snapshot `1 x W` at **`(P-1, pc)`** (no fail-fast for H==1 by this decision). *Note:* this is a behavior change vs current H==1 (which pastes at `P`); acceptance/layout must be validated in L3 (Q4=B means preserve out-of-scope, but in-scope template behavior follows this rule).
  - `H <= 0`: remain fail-fast as today (invalid snapshot size).
- **OP-3 (P==1) = A (fail-fast):** For in-scope P09 path, if configured/effective `paste_row P < 2` (so `P-1 < 1`), **fail fast** with clear, actionable diagnostic requiring `paste_row >= 2` for this mode. Must include context (`data_sheet`, template/scope, `paste_row`, `rows/cols`/H/W, `insert_rows`, `only_values`, `skip_header`, operation). *Constraint:* current code clamps `pr <= 0` to `1`; implementation must not silently allow `P==1` in-scope — validation must occur before in-scope structural changes.
- **OP-4 (contract):** `paste_row` remains the configured value `P`; sample row remains entire row `P`. Actual snapshot top becomes `P-1` in-scope; returned range and downstream (`ApplySheetFormatting`, `ApplyOutlineGrouping`) must follow **actual placement/bounds**. Final range is **expected** `P-1 .. P-1+H-1` over `pc..pc+W-1` but **must be proven/verified against Excel shift behavior** in L3 (A1), especially for `H==2` and `H==1` where no rows are inserted.
- **Q4 (priority) = B:** Preserve existing layouts/outputs where possible (reinforces OP-1=A and strict out-of-scope regression checks).

**Ambiguities still to resolve in Block B/C (technical, not PO):** A1 (prove final coordinates under Excel row-insert/shift semantics for `H>=3`, `H==2`, `H==1` — design must compute returned range from actual anchor `P-1` and L3 must verify), A2 (exact error codes/messages for P<2 in-scope consistent with existing 2702/2704/2705 patterns), A3 (`H,W` derived after `skip_header` slicing as today), A4 (no deletions), A5 (concrete L3 scenarios), A6 (broader discrepancies out of scope). OP-1/A means `CopySnapshot` must know template scope (caller passes indicator or infers via new parameter) — technical choice in Block C, not PO.

---

## 7) Step 2 (Ask user as Product Owner) — is it needed?

**Decision: Step 2 is needed (business/product uncertainty exists) — but limited to non-technical choices.**

P09 OP-1..OP-4 are framed as decisions to confirm in the problem text; some are technical/contract choices, but they have **business/operational consequences** (which workbooks/configs are affected, whether existing reports break, what error users see). Per instructions: “do not ask about technical implementation/architecture. Only if there is business uncertainty…” Ask only about business goals, expected behavior/outcomes, user impact, compatibility expectations, acceptance scenarios, priorities.

### Business uncertainties identified (from OP-1, OP-2 `H==1`, OP-3, and impact)

1. **Compatibility / blast radius (tied to OP-1):** Do you want the new behavior **only for templates that set `insert_rows=TRUE` with a template** (safer, smaller change), or **for all `insert_rows=TRUE` reports** even if no template is used (single behavior, may change existing non-template outputs)?
2. **Edge case `H == 1` with `insert_rows=TRUE` (tied to OP-2):** When the snapshot has only **1 data row** (`H==1`), what should happen under the new rule (paste at `P-1`)?
3. **Edge case `paste_row == 1` (tied to OP-3):** When configured `paste_row` is **1**, pasting at `P-1` is impossible. What is the expected business outcome?
4. **Acceptance priority / risk tolerance (optional but useful):** If the new behavior could change layout in existing reports, do you prioritize **matching the new template requirement exactly** over **minimizing surprises in current outputs**?

### If no business uncertainty existed
Not applicable here; the above are present. If the PO has already decided OP-1..OP-4 as business policy, state those decisions explicitly and Step 2 can be satisfied by recording them.

### Suggested concise multiple-choice questions (with recommended options)

**Q1 (OP-1, business/compatibility impact) — required if not already decided**
> **Where should the new anchor/count rule apply?** (This affects which reports change.)
> - **A (Recommended, lower risk):** Only when `insert_rows=TRUE` **and** a template is set (non-template `insert_rows=TRUE` keeps current behavior).  
> - **B:** For **all** `insert_rows=TRUE` cases (template or not).  
> - **C:** Other (describe business reason).

**Q2 (OP-2, `H == 1`) — required if not already decided**
> **If a snapshot has only 1 row (`H==1`) with the new rule (`paste at P-1`)**, what should happen?
> - **A (Recommended, explicit failure):** **Fail with a clear error** (do not produce output) so the configuration/data issue is fixed intentionally.  
> - **B:** **Allow it** by treating as “no extra inserted rows” and placing the 1-row snapshot at `P-1` (may be surprising; needs validation).  
> - **C:** Other (describe expected business behavior).

**Q3 (OP-3, `paste_row == 1`) — required if not already decided**
> **If `paste_row` is set to 1 (so `P-1` does not exist)**, what should happen?
> - **A (Recommended):** **Fail with a clear error** explaining that `paste_row` must be **≥ 2** for this mode.  
> - **B:** Use a different anchor rule (please specify the desired row/behavior in business terms).  
> - **C:** Other.

**Q4 (priority, optional)**
> **If changing behavior could affect existing reports**, what is preferred?
> - **A (Recommended):** **Correctness for the new template requirement first**; accept that affected reports may need review/adjustment.  
> - **B:** **Preserve existing report layouts** where possible; only change when strictly necessary.  
> - **C:** Decide case-by-case (describe).

**Note:** Technical alternatives in Block C (API shape, exact error codes beyond clarity) are intentionally **not** asked in Step 2.

---

## 8) Deliverable reference (this file)

- **Saved plan path:** `C:\py_exp\mko_treport\.ai\plans\008-p09-insert-rows-anchor-rework.md`  
- **`plan_file` reference (input):** `C:\py_exp\mko_treport\.ai\problems\P09.md`  

---

## 9) Checklist (for implementer, when proceeding)

- [ ] Block A: OP-1..OP-4 decisions recorded (and Step 2 answers incorporated if provided).  
- [ ] Block B: Contract delta applied to docs/spec lineage plan; no code edits yet in that block.  
- [ ] Block C: Semantic edits planned for `modTreportExtract.bas` (+ `modTreportRun.bas` if needed); call sites audited.  
- [ ] Validator review completed if high-risk gate triggered.  
- [ ] L1 run with recorded command/report (`.ai/tmp/build_treport_vba_*.txt`).  
- [ ] L2 manual VBE Compile passed (recorded) OR limitation stated (no false claim).  
- [ ] L3 scenarios executed on copies; results recorded (including `H==1,2`, `P==1`, `only_values`, `skip_header`, grouping/regressions per OP-1).  

Project rule reminder: **never** claim “compiled successfully” from script best-effort compile alone; L2 is required.
## 10) Block B/C — Implementor task (locked decisions applied)

This section is the **Implementor-ready task description** for P09 (`insert_rows` anchor/count rework), derived from `\.ai\problems\P09.md`, the locked Block A decisions (Q1=A, Q2=B, Q3=A, Q4=B / OP-1..OP-4), and the current semantic implementation in `mko_treport/src/vbs/export/modTreportExtract.bas` (`CopyPivotSnapshot`, `PasteSnapshotData`) and `mko_treport/src/vbs/export/modTreportRun.bas` (`ProcessDataFile`).

### 10.1 Task summary (task_template-style)

```yaml
id: p09_insert_rows_anchor_rework
title: "P09: template-only insert_rows anchor/count rework (sample row P, insert H-2, paste at P-1)"
priority: high

depends_on: []

source_reference: .ai\plans\008-p09-insert-rows-anchor-rework.md
source_section: "Block A Decision record (locked) + Block B/C implementor task"
source_file: .ai\problems\P09.md

description: >
  Rework insert_rows=TRUE behavior for template-scoped subreports (OP-1=A) so that
  the configured sample row P is inserted H-2 times (OP-2=B), the snapshot is
  pasted at (P-1, pc), and the returned/downstream range follows actual placement,
  without deletions or snapshot shrink. Preserve legacy behavior for all out-of-scope
  cases (insert_rows=FALSE and non-template insert_rows=TRUE). Align docs/spec lineage.

goals:
  - Implement OP-1=A template-only P09 path with deterministic OP-2=B H rules.
  - Enforce OP-3=A fail-fast when P < 2 for in-scope path with actionable diagnostics.
  - Return actual post-placement range (A1) for downstream formatting/grouping.
  - Preserve out-of-scope behavior and outputs where possible (Q4=B).
  - Update VBA comments/docs/spec notes to reflect reconciled contract.

files:
  - path: mko_treport/src/vbs/export/modTreportExtract.bas
    targets:
      - type: function
        name: CopyPivotSnapshot
      - type: procedure/comment
        name: module/procedure behavior comments (P05/P07/P09 lineage)
    semantic_anchors:
      - after: "Public Function CopyPivotSnapshot(...)" signature and parameter normalization/guards
      - around: "If insertRows And h > 1 Then" sample-row preparation branch
      - around: "Set dstCell = dstWs.Cells(pr, pc)" / paste and return-range computation
      - around: "EHRowPrep" / "EH" fail-fast context construction

  - path: mko_treport/src/vbs/export/modTreportRun.bas
    targets:
      - type: procedure
        name: ProcessDataFile
    semantic_anchors:
      - around: "templateSet = (Len(Norm(cfg.DataFiles(dfIdx).template)) > 0)"
      - around: "Set insertedRng = modTreportExtract.CopyPivotSnapshot(...)" call site
      - keep: downstream "ApplySheetFormatting" / "ApplyOutlineGrouping" consume returned insertedRng

  - path: docs/01-vba-reporting/configuration-model.md
    targets:
      - type: section
        name: "data_files.insert_rows (and related anchor/final-range semantics)"
  - path: docs/01-vba-reporting/pipeline-and-dataflow.md
    targets:
      - type: section
        name: "P04/P05 unified insertion/formatting sequence (update to P09-scoped contract)"
  - path: .ai/problems/005_p05_insert_rows_template_row_spec.md
    targets:
      - type: note
        name: "Amendment/supersession note referencing P09 and locked decisions"

files_not_to_change_as_primary_scope:
  - modTreportConfig.bas (no schema/parsing changes required by P09)
  - modTreportOutput.bas (unless a coupled semantic issue is proven; prefer none)
  - workbook binaries, canonical data, prior audit phase files

acceptance_criteria:
  - In-scope (insert_rows=TRUE AND template set): H>=3 inserts H-2 at P and pastes HxW at (P-1,pc); H==2 inserts 0 and pastes 2xW at (P-1,pc); H==1 inserts 0 and pastes 1xW at (P-1,pc). No deletions, no snapshot shrink.
  - In-scope P<2 fails fast before structural changes with clear actionable diagnostic including required context.
  - Returned range is actual destination range at P-1 anchor (P-1 .. P-1+H-1, pc .. pc+W-1) and used by ApplySheetFormatting/ApplyOutlineGrouping.
  - Out-of-scope unchanged: insert_rows=FALSE; non-template insert_rows=TRUE retains legacy P05 behavior (H-1 inserts at P when H>1, paste at P, return P..P+H-1).
  - only_values TRUE/FALSE and skip_header semantics preserved (H/W after skip_header slicing as today).
  - Docs/spec amendment updated; L1/L2/L3 executed per project workflow with required records.
```

### 10.2 Exact implementation scope

**In-scope (P09 path, OP-1=A):** applies only when `insertRows = True` **AND** the subreport has a template set (`Len(Norm(cfg.DataFiles(dfIdx).template)) > 0`, i.e. `templateSet=True` in `ProcessDataFile`). This is a template-only insertion-mode rework.

- Sample row: entire target-sheet row `P = configured/effective paste_row` (identity unchanged).
- Anchor validation (OP-3=A): require `P >= 2` so `P-1 >= 1`. If not, **fail-fast before any row copy/insert/paste/structural mutation** with actionable diagnostic.
- Row preparation count (OP-2=B), using effective snapshot height `H = srcRg.Rows.Count` after existing `skip_header` slicing (A3): `H >= 3 => insCount = H-2`; `H == 2 => insCount = 0`; `H == 1 => insCount = 0` (allowed, no fail-fast for H==1). `H <= 0` remains existing fail-fast (2704 invalid snapshot size) — unchanged.
- If `insCount > 0`: copy entire row `P` and insert `insCount` copies starting at `P` (`Shift:=xlDown`) with same Excel API intent/error-handling discipline as today. If `insCount <= 0`: perform no row insertions.
- Snapshot destination anchor (OP-4): `dstTop = P - 1`; paste snapshot `H x W` at `(dstTop, pc)` via existing `PasteSnapshotData srcRg, dstWs, dstWs.Cells(dstTop, pc), onlyValues, ctx` (no snapshot shrink, no deletions).
- Returned range (A1): actual destination bounds `dstWs.Range(dstWs.Cells(dstTop, pc), dstWs.Cells(dstTop + H - 1, pc + W - 1))` (expected `P-1 .. P-1+H-1`). L3 must prove this matches Excel reality for `H>=3/H==2/H==1`.
- Downstream: `modTreportRun.ProcessDataFile` must continue passing the returned range to `ApplySheetFormatting` and `ApplyOutlineGrouping` (no call-pattern change; semantics follow actual placement).

**Out-of-scope / must remain unchanged (Q4=B, OP-1=A):**
- `insert_rows = False` path: no P09 anchor/count changes, no new P<2 validation, no behavior change.
- `insert_rows = True` **without** template set (non-template): preserve current/P05 behavior exactly — if `H > 1` insert `H-1` copies at `P`, paste at `P`, return `P .. P+H-1`. Do not apply P-1 anchor, do not apply H-2 count, do not add P>=2 P09 fail-fast.
- `skip_header` slicing logic and its effect on effective `H/W/srcRg` unchanged (A3).
- `only_values` transfer semantics unchanged (A4 context: no deletions/shrink).
- No helper-row deletion behavior reintroduced.
- No changes to config schema/parsing, placeholder/filter/pivot algorithms, or unrelated modules.

### 10.3 Implementation sequence (semantic, dependency-safe)

1. **Baseline & working copy discipline (before edits in repo/workbook verification):** confirm target exports and docs to change; plan L1/L2/L3 on working copies only (see 10.5). Do not claim compilation.
2. **`modTreportExtract.CopyPivotSnapshot` — signature and scope flag:** add an explicit template-scope indicator parameter (minimal, maintainable choice): `Optional ByVal templateInsertMode As Boolean = False` (name is clear; alternatively `inScopeP09` acceptable if used consistently). Update procedure comment to document P09 template-only rules and that legacy P05 applies when `insertRows=True` but `templateInsertMode=False`. Compute `inScopeP09 = (insertRows And templateInsertMode)` after parameters are known.
3. **Anchor/parameter normalization (semantic):** keep existing `pr/pc <= 0 => 1` normalization for backward compatibility, but for `inScopeP09` enforce `pr >= 2` **after** deriving effective `h/w/srcRg` (so diagnostics include rows/cols/skip_header context) and **before** any sample-row copy/insert or paste. On violation raise fail-fast 2705 (see 10.6). This means configured `P==1` (or clamped-to-1) is rejected in-scope with actionable message.
4. **In-scope row preparation (OP-2=B):** for `inScopeP09`, compute `insCount`: `If h >= 3 Then insCount = h - 2 Else insCount = 0` (covers H==2/H==1). If `insCount > 0`, perform sample-row duplication at `P` for `insCount` rows (reuse `EHRowPrep` error handling; clear `CutCopyMode`). For `insCount <= 0`, skip insertion. Do not apply legacy `h > 1 => h-1` branch in-scope.
5. **Legacy branch preservation (OP-1=A):** for `insertRows=True And Not inScopeP09`, retain existing behavior exactly: insert `h-1` copies at `pr` only when `h > 1` (P05). Paste/return remain at `pr` as today.
6. **Paste destination and return range (A1, OP-4):** if `inScopeP09` set `dstTop = pr - 1` else `dstTop = pr`; paste at `dstWs.Cells(dstTop, pc)`; return `dstWs.Range(dstWs.Cells(dstTop, pc), dstWs.Cells(dstTop + h - 1, pc + w - 1))`. Ensure fail-fast contexts (`EH`, `EHRowPrep`, 2704 message text/context) reflect the correct anchor/scope fields (include `template_insert_mode`/scope indicator where useful) without breaking existing diagnostic usefulness for legacy path.
7. **`modTreportRun.ProcessDataFile` — call site:** pass template scope explicitly: `modTreportExtract.CopyPivotSnapshot(pt, outWs, pr, pc, repCtx, onlyValues, insertRows, skipHeader, templateSet)`. Keep `templateSet` computation and `ApplySheetFormatting`/`ApplyOutlineGrouping` flow unchanged (formatting already receives template-preserving intent via `templateSet`; grouping must receive actual returned range).
8. **Documentation/spec alignment:** update `CopyPivotSnapshot` header comments (P09 lineage), `configuration-model.md` (`data_files.insert_rows`), `pipeline-and-dataflow.md` (unified sequence for scoped case + legacy note), and append amendment/supersession note to `.ai/problems/005_p05_insert_rows_template_row_spec.md` referencing P09 and OP-1..OP-4 (old vs new for template-scope).
9. **Verification (L1 → fix loop → L2 → L3)** on working copies; record evidence as specified in 10.5.

### 10.4 Architectural and workbook-state constraints

- **VBA-only runtime; export-source edits only:** modify `mko_treport/src/vbs/export/*.bas` (and docs/spec notes) as specified; no non-VBA runtime dependencies. Do not edit canonical workbook binaries as source-of-truth artifacts except through controlled L1/L2 workflow on **working copies**.
- **Semantic targeting only:** no line-number-based instructions in implementation; use procedure names, branches, and contracts as anchors.
- **No deletions, no snapshot shrink (A4):** do not add or restore helper-row deletion; snapshot remains `H x W`.
- **Stage/order invariants:** row preparation/structural changes (if any) occur before paste at final anchor and before final-range return; downstream formatting/grouping operate on returned actual range (GA-4-compatible intent).
- **Workbook/application safety:** L1/L3 run on local working copies; originals unchanged (R8). Expect Excel/VBOM/macro-security sensitivity; close Excel if needed for script reliability. L2 is manual in VBE on the working copy.
- **Template/source immutability:** P09 does not change template discovery/copy rules; `ThisWorkbook` templates remain immutable by existing contract.
- **Q4=B regression discipline:** out-of-scope behavior/layouts must be preserved as much as possible; any in-scope template behavior follows P09 even where it differs from prior anchor (e.g. H==1 now at P-1).

### 10.5 Required validation (L1 / L2 / L3)

#### L1 — automated build/sync/import (script, best-effort compile note)
- Run `mko_treport/src/scripts/build_treport_vba.ps1` on a **working copy** with backup/report paths recorded. Example (adjust date/time/paths as needed):
  ```powershell
  .\mko_treport\src\scripts\build_treport_vba.ps1 `
    -SourceWorkbook "data\theiss\treport.xlsm" `
    -WorkCopyPath ".ai\tmp\treport_vba_check_<YYYYMMDD_HHMMSS>.xlsm" `
    -BackupDir ".ai\tmp" `
    -ReportPath ".ai\tmp\build_treport_vba_<YYYYMMDD_HHMMSS>.txt"
  ```
- Review generated report for import/legacy/duplicate issues and VBOM/compile notes. If VBOM access missing, state limitation and enable "Trust access to the VBA project object model" in Excel Trust Center, then retry on a new working copy.
- **Record:** exact command, report path, short summary (SUCCESS/ISSUES/ERROR), any VBOM/compile-note limitations. Script compile is **not** sufficient to claim compilation.

#### L2 — manual VBE compile (required to claim compilation)
- Open the **working copy** `.xlsm` in Excel → VBA Editor → **Debug → Compile VBAProject**.
- **If clean:** record working-copy path, date/time, and "L2 passed (manual VBE Compile VBAProject)".
- **If errors:** fix in exports, re-run L1 on a **new** working copy, repeat until resolved or report precise blocker. Never claim compiled from L1 alone.

#### L3 — functional checks on copies only (behavior changed)
- Test on **copies** of data/workbooks only; **originals must remain unchanged**. Use deterministic/representative template+insert_rows scenarios (prefer `data/theiss` copies where available) with refreshed state consistent with existing workflow.
- Record per-scenario: inputs/config (template set?, `insert_rows`, `only_values`, `skip_header`, `paste_row`, expected `H`), observed row-prep count, paste anchor, final range, formatting/grouping outcome, and pass/fail.

**Concrete L3 scenario checklist (minimum):**

| ID | Scope/config | Scenario focus | Expected result |
|---|---|---|---|
| L3-01 | In-scope (template set, insert_rows=TRUE) | `H >= 3`, `P >= 2` | Insert `H-2` copies of row `P` at `P`; paste `H x W` at `(P-1,pc)`; final range `P-1..P-1+H-1` over `pc..pc+W-1`; no deletions/shrink; formatting/grouping correct on actual range |
| L3-02 | In-scope | `H == 2`, `P >= 2` | Insert 0; paste `2 x W` at `(P-1,pc)`; final range `P-1..P` correct; downstream correct |
| L3-03 | In-scope (OP-2=B allows) | `H == 1`, `P >= 2` | Insert 0; paste `1 x W` at `(P-1,pc)`; final range single row at `P-1`; no fail-fast; layout/grouping bounds correct (A1 verification critical) |
| L3-04 | In-scope (OP-3=A) | `P == 1` (or effective < 2) | Fail-fast before structural changes with clear actionable diagnostic (requires `paste_row >= 2` for this mode) including context fields; no partial mutation |
| L3-05 | In-scope variants | `only_values = TRUE` and `FALSE` (at least one case each, include H>=3) | Value-only preserves prepared/target formatting as per existing semantics; FALSE applies existing snapshot transfer semantics; anchors/counts correct |
| L3-06 | In-scope variant | `skip_header` relevant case(s) (True and/or False as applicable) | Effective `H/W/srcRg` after slicing drive count/anchor/final range correctly; P07 semantics preserved |
| L3-07 | In-scope with grouping/layout | `group_rows` set (valid) and/or freeze/layout-sensitive template | `ApplyOutlineGrouping` and `ApplySheetFormatting` operate correctly on actual returned range/bounds |
| L3-08 | Regression (out-of-scope) | `insert_rows = FALSE` (template and/or ordinary as available) | Behavior/layout unchanged from baseline |
| L3-09 | Regression (out-of-scope, OP-1=A) | Non-template `insert_rows = TRUE`, `H > 1` | Legacy P05 behavior preserved exactly: `H-1` inserts at `P`, paste at `P`, final `P..P+H-1` |
| L3-10 | Safety | Originals check | Confirm canonical/source workbooks and original data files unchanged after all L3 runs |

**Notes:** if a scenario is not feasible with available data, record "not performed" with reason and residual risk; do not claim coverage. L3 must specifically probe **A1** (returned range matches Excel reality) for H==1/H==2/no-insert cases and H>=3 inserted cases.

#### Reporting (always, after VBA/export or behavior-affecting edits)
- List changed files under `mko_treport/src/vbs/export/` and docs/spec files changed.
- Exact `build_treport_vba.ps1` command used (or why not run).
- Build report path and short summary.
- L2 result: done (file/time/pass) or not done with explicit limitation (not proof of compilation).
- L3 results per checklist item or "not performed" with reason; confirm originals unchanged.

### 10.6 Technical design choices (consistent with locked decisions)

#### 10.6.1 Template-scope communication (OP-1=A)
- **Chosen approach (minimal, explicit):** extend `CopyPivotSnapshot` with `Optional ByVal templateInsertMode As Boolean = False` as the last parameter and pass `templateSet` from `ProcessDataFile` (`templateSet` already computed as `Len(Norm(cfg.DataFiles(dfIdx).template)) > 0`).
- **In-scope predicate:** `inScopeP09 = (insertRows And templateInsertMode)` inside `CopyPivotSnapshot`. This applies P09 only for template + insert_rows, preserving legacy non-template insert_rows behavior exactly.
- **Rationale:** explicit intent, minimal signature/call-site change, no hidden inference from unrelated state, easy to test/review. Update all relevant declarations/comments/call sites.

#### 10.6.2 OP-3=A anchor validation and error handling (A2, consistent with 2704/2705/2702)
- **When:** for `inScopeP09` only, validate `pr >= 2` **before** row preparation/copy-insert and **before** paste (no structural mutation on failure). Perform after effective `h/w/srcRg` and skip_header handling so diagnostics are complete.
- **Error number (chosen):** `2705` (row-preparation/precondition failure in insert_rows flow), consistent with existing `EHRowPrep` usage for sample-row prep failures.
- **Context/message style (chosen):** include `data_sheet`, `parameter='insert_rows/paste_row'`, `paste_row=<pr>`, `paste_col=<pc>`, `rows=<h>`, `cols=<w>`, `insert_rows=<insertRows>`, `only_values=<onlyValues>`, `skip_header=<skipHeader>`, `template_insert_mode=True` (or `scope='template_insert_rows_p09'`), `operation='validate_p09_anchor'`, and a clear description: e.g. "Invalid paste_row for template insert_rows (P09): paste_row must be >= 2 so snapshot can be placed at paste_row-1; configured/effective paste_row=<pr>."
- **Clamp interaction:** retain existing `If pr <= 0 Then pr = 1` normalization globally; if in-scope this results in `pr < 2` and must fail-fast with the above diagnostic (do not silently proceed).

#### 10.6.3 A1 — returned/final range and Excel shift behavior (design requirement)
- **Chosen computation:** use actual destination top anchor for the path taken:
  - `inScopeP09`: `dstTop = pr - 1`, `insCount = IIf(h >= 3, h - 2, 0)` (0 for H==2/H==1), paste at `(dstTop, pc)`, return `Range(Cells(dstTop, pc), Cells(dstTop + h - 1, pc + w - 1))`.
  - legacy/out-of-scope: `dstTop = pr`, existing insert/paste/return semantics unchanged.
- **Rationale:** returned range must represent actual snapshot bounds for downstream `ApplySheetFormatting`/`ApplyOutlineGrouping`. For no-insert cases (H==1/H==2 in-scope) this is straightforward; for H>=3, inserting `H-2` rows at `P` shifts original `P` downward but `P-1` remains the intended snapshot top — **L3 must verify** the computed range matches on-sheet reality (merges/objects/outlines/templates can be sensitive). If L3 disproves, capture as verification finding (do not guess in code without evidence).

#### 10.6.4 Error/context consistency details
- Keep `2704` for invalid snapshot size (`h <= 0 Or w <= 0`) unchanged.
- Keep `2702`/`EH` and `2705`/`EHRowPrep` patterns; enrich contexts with `template_insert_mode` (or scope flag) where it aids diagnosis, but avoid changing legacy meaning for out-of-scope cases.
- `skip_header`-related failures (2708/2709) remain unchanged.

### 10.7 Doc/spec updates required (in scope)

- **`mko_treport/src/vbs/export/modTreportExtract.bas`:** update `CopyPivotSnapshot` header/comments to describe P09 template-only behavior (OP-1=A), OP-2=B counts/anchors, OP-3=A P>=2 precondition, A1 actual-range return, and preserved legacy path; note P05 retained for non-template insert_rows.
- **`docs/01-vba-reporting/configuration-model.md`:** revise `data_files.insert_rows` description to state template-set in-scope behavior (`H-2` inserts at `P` for `H>=3`, 0 for `H==2/H==1`, paste at `P-1`, final/actual range `P-1..P-1+H-1` subject to L3 verification) and explicitly state legacy/non-template `insert_rows=TRUE` behavior unchanged (P05). Keep `only_values`/`skip_header`/`paste_row` semantics aligned (P remains configured sample row).
- **`docs/01-vba-reporting/pipeline-and-dataflow.md`:** update unified insertion/formatting sequence bullets to reflect P09-scoped anchor/count/final range, OP-1 scope, and note A1 Excel-shift verification; retain legacy description for out-of-scope cases.
- **`.ai/problems/005_p05_insert_rows_template_row_spec.md`:** append an amendment/supersession note (do not rewrite accepted history) titled e.g. "Amendment — P09 (2026-11-17, locked OP-1..OP-4)" referencing `.ai/problems/P09.md` and `.ai/plans/008-p09-insert-rows-anchor-rework.md`; state old vs new for template-scope `insert_rows=TRUE` (H-1@P,paste@P,final P..P+H-1 → H-2@P for H>=3 else 0, paste@P-1, final P-1..P-1+H-1, P>=2 required, H==1 allowed) and that non-template/insert_rows=FALSE remain per prior contract.

### 10.8 Residual high risks / Validator review recommendations

- **A1 (high, Excel behavior):** for `H==1` and `H==2` (0 inserts) and `H>=3` (`H-2` inserts at `P`) with paste at `P-1`, actual on-sheet bounds/shift side effects (merges, formulas, CF, data validation, shapes, outlines, tables) may differ by template/Excel build. **Validator should review design and L3 evidence** before accepting implementation as complete.
- **OP-3 clamp interaction:** legacy clamp-to-1 plus new in-scope P>=2 fail-fast changes effective behavior for invalid/configured-1 inputs in template mode — intentional per OP-3=A but must be clearly messaged and regression-checked.
- **Out-of-scope regression guard (Q4=B):** ensure non-template `insert_rows=TRUE` and `insert_rows=FALSE` paths are byte-for-byte equivalent in behavior; recommend Validator spot-check branch separation (`templateInsertMode` gating).
- **Call-site completeness:** `CopySnapshot` signature changes require all callers updated (currently only `ProcessDataFile` per audit) — quick grep confirmation recommended before/after.
- **L2 evidence:** compilation cannot be claimed without manual VBE compile; any blocked environment must be reported as limitation, not success.

### 10.9 Implementor checklist (execution)

- [ ] Confirm OP-1..OP-4 locked decisions reflected in code/comments/docs (no scope drift). Call-site audit evidence recorded: `CopyPivotSnapshot` callers are only `modTreportRun.ProcessDataFile` call site and the definition/return in `modTreportExtract.bas` (grep-verified 2026-11-17).
- [ ] Implement `CopyPivotSnapshot` in-scope branch (`templateInsertMode`), OP-3 validation, OP-2 counts, A1 return range; preserve legacy branch exactly.
- [ ] **Validator correction (critical):** P09 `paste_row < 2` (2705) validation applies **only** when `inScopeP09` is true; legacy non-template `insert_rows=TRUE` and `insert_rows=FALSE` retain pre-P09 normalization/behavior unchanged (do not fail-fast P<2 for out-of-scope paths).
- [ ] **Validator correction (A1 semantics):** for in-scope `H==1` or `H==2`, `insCount==0` means no `Rows(...).Insert` call and **no sheet shift**; snapshot overwrites at existing row `P-1` over `H x W`. Returned range must still be computed as `Cells(P-1,pc)..Cells(P-1+H-1,pc+w-1)` and L3 must confirm actual Excel bounds/overwrites.
- [ ] Update `ProcessDataFile` call site to pass `templateSet`.
- [ ] Update docs: `configuration-model.md`, `pipeline-and-dataflow.md` (scope-conditional: template-only P09 vs legacy non-template P05; insert_rows=FALSE unchanged); add amendment in `005_p05_insert_rows_template_row_spec.md`.
- [ ] L1 run on working copy with recorded command/report (`.ai/tmp/...`); fix issues and repeat as needed.
- [ ] L2 manual VBE `Debug -> Compile VBAProject` on working copy recorded (path/date/time/pass) or limitation stated (do not claim compile without L2).
- [ ] L3 scenarios L3-01..L3-10 executed on copies with results recorded; originals unchanged. **Record observed final range address and row-prep count per scenario** (A1 evidence). If any scenario not feasible, record "not performed" with reason and residual risk (seek representative `H==1` after `skip_header` if possible).
- [ ] L3 regressions explicitly verify: `insert_rows=FALSE` unchanged; non-template `insert_rows=TRUE` legacy behavior unchanged (including not subject to P09 P<2 fail-fast/P-1 anchor; test H>1 and H==1 if relevant).
- [ ] Produce final report: changed files, exact script command, report path+summary, L2 status, L3 status.

### 10.10 Validator corrections/additions (must apply, pre-implementation review 2026-11-17)
- **Verdict:** Acceptable to proceed with corrections. No blocking code/spec contradiction found if corrections applied.
- **Clamp vs in-scope fail-fast precedence (normative):** normalize `pr/pc` as today; **only if `inScopeP09` and effective `pr < 2` then fail-fast (2705) before any structural mutation/paste**, regardless of whether `1` came from config or clamp; do **not** apply this new fail-fast to out-of-scope paths. Include `data_sheet`, template/scope indicator, `paste_row` (effective/config context as available), `rows/cols`, `insert_rows`, `only_values`, `skip_header`, `operation='validate_p09_anchor'` in diagnostics.
- **No-insert in-scope semantics:** for `inScopeP09` and `H==1|2`, `insCount==0` => no `Rows(...).Insert`, no sheet shift; destination is existing row `P-1` (overwrite in snapshot bounds). Returned range computed as `Cells(P-1,pc)..Cells(P-1+H-1,pc+w-1)`; L3 must confirm actual bounds/overwrites.
- **Diagnostics:** in-scope P<2 error recommended as 2705 consistent with row-prep failures; include `parameter='insert_rows/paste_row'`, `template_insert_mode=True` or `scope='template_insert_rows_p09'`, clear message requiring `paste_row >= 2` so snapshot can be placed at `paste_row-1`. Enrich EH/EHRowPrep contexts with scope/anchor where helpful without changing legacy meaning.
- **Evidence/gates:** record call-site audit evidence; L3 must record observed final `Address`/row-prep count per scenario; explicitly verify legacy `insert_rows=FALSE` and non-template `insert_rows=TRUE` unchanged (H>1 and H==1 if relevant). If L3 item skipped, document residual risk.
- **Docs:** scope-conditional wording in `configuration-model.md` and `pipeline-and-dataflow.md` (template-only P09 vs legacy non-template P05; insert_rows=FALSE unchanged; H==1 allowed in-scope per OP-2=B). P05 spec gets amendment/supersession note (do not rewrite history) with before/after for in-scope and statement that non-template/insert_rows=FALSE remain per prior contract.
- **Residual high risks:** A1 Excel behavior for H==1/H==2/H>=3 (merges/CF/formulas/objects/templates) — L3 proof required; legacy path regression guard; L2 cannot be claimed without manual VBE compile.

