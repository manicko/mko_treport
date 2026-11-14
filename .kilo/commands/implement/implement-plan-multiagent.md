---
name: implement-plan-multiagent

description: Execute development plan incrementally using adaptive multi-agent workflow for an Excel/VBA project

alwaysApply: false

---

# Task: Multi-Agent Plan Implementation (Excel/VBA)

# Workflow
Do not argue on the workflow, just follow.
Max allowed parallel subagents = 2
Do not launch agents in background.
If stop or break prefer resume old session not launching new agent.

## 0. Inspect relevant current architecture and implementation 

Launch an `Auditor` to inspect the current workbook/VBA implementation and related artifacts relevant to the plan.

Provide:
{plan_file} +
<original_prompt>
- What is already implemented
- Whether the plan matches the current implementation
- Relevant architecture, dependencies, and constraints (VBA modules `.bas`/`.cls`/`.frm`, Workbook/Worksheet code-behind, exported sources such as `mko_treport/src/vbs/export/`, `vbaProject.bin`, macros/controls/entry points, ListObjects/tables/settings, Power Query queries/connections, pivot tables/caches, and gitignored local artifacts like `data/theiss/*.xlsm/*.xlsx` or outputs when relevant)
- Important discrepancies or risks (including VBA opacity/binary-vs-export drift, PQ/pivot behavior, Excel/Windows compatibility, and local-only artifact constraints)
**Important: Never change any VBA/code or existing audit phase files.**
Output as `{code_context}`.
<original_prompt>


## 1. Decompose Plan
Phase may be skipped if plan is already sufficiently detailed and clear to proceed with implementation if not launch a `Planner` with the {plan_file} and `{code_context}`.

Provide:
<original_prompt>
- Decompose the plan into logical execution blocks
- Identify dependencies and execution order
- Assess implementation, rollout, regression, and compatibility risks (including Excel version/Windows environment, macro security, Power Query refresh timing, pivot/object-model behavior, workbook state/locking, and fidelity of outputs)
- Identify multiple implementation paths alternatives when relevant, for maintainability, future evolution, and project conventions
- **Not choose the final implementation approach when technical uncertainty exists**
- Keep the scope minimal and avoid speculative redesign
- For each block, determine whether the following agents are required:
  - **High risks** - all agents below 
  - **Auditor** — deeper VBA/workbook/architecture investigation due to uncertainty or complexity
  - **Researcher** — modern/best practices, Excel/VBA/Power Query/pivot behavior, multiple viable approaches, support implications
  - **Planner** — detailed pre-implementation design, workbook/module structure, validation, or complex execution
  - **Validator** — independent plan review when implementation risk is high
**Important: Never change any VBA/code or existing audit phase files.**
<original_prompt>

Save the result as `Execution plan`: `.ai\plans\{next-number}-{plan-name}.md`.
----------------------------------

## 2. Ask the user as the Product Owner

Skip this step unless there is business uncertainty.

Do not ask about technical implementation or architecture — that is your task.

Focus questions on:

* **User goals and problems** — what users need to achieve and why
* **User journey and experience** — how users expect the flow to work, including key scenarios and edge cases
* **Product behavior** — what the product should do from the user's perspective
* **Business rules and policies** — constraints, conditions, and exceptions
* **Expected outcomes and success criteria** — what should happen when the feature works correctly
* **Priorities and trade-offs** — what matters most when requirements conflict
* **Permissions and responsibilities** — who can do what and who owns each decision
* **Data and content expectations** — what information users need to see, enter, or receive
* **Failure and recovery experience** — what users should see and be able to do when something goes wrong
* **Examples and acceptance scenarios** — concrete examples of expected behavior

Provide concise multiple-choice options based on best practices, **highlight the recommended option**, and ask only questions needed to resolve business uncertainty.

Use the answers to complete the specification.



## 3. Execute Blocks

For each block, execute the required agents sequentially (following the order below):

---
### 3.1 Auditor — if required

Launch `Auditor`  with:

`{plan_context_exec_block}`

<original_prompt>
Inspect the current implementation and architecture relevant to the block:
* VBA code and workbook architecture (modules, procedures, controls, ListObjects, PQ/pivots, connections)
* Existing patterns and constraints
* Current implementation behavior and artifacts
* Risks and compatibility concerns
**Important: Never change any VBA/code or existing audit phase files.**
Update `Execution plan` and Return `{context_a}`.
<original_prompt>

### 3.2 Researcher — if required

Launch a `Researcher` with:

`{plan_context_exec_block} + {context_a}`

<original_prompt>
- Identify viable alternatives when relevant
- Research the relevant modern/best practices for Excel VBA, Power Query, pivot/object-model behavior, and maintainability
- Evaluate viable implementation approaches
- Assess architectural, implementation, rollout, regression, and compatibility risks
- Select the **best implementation path** for maintainability, future evolution, and project conventions
- Avoid speculative redesign
**Important: Never change any VBA/code or existing audit phase files.**
Update `Execution plan` and Return `{context_r}`.
<original_prompt>

### 3.3 Planner — if required

Launch a `Planner` with:

`{plan_context_exec_block} + {context_a} + {context_r}`


<original_prompt>
Create the implementation task for the `Implementor`.
Use **semantic code units only**:
- Files (e.g. `.bas`, `.cls`, `.frm`, workbook artifacts)
- Standard Modules, Class Modules, and UserForms
- Workbook/ThisWorkbook and Worksheet code-behind modules
- Procedures (Sub/Function/Property)
- Controls, ListObjects/tables, named ranges, and Power Query queries/connections where relevant

**Never use line numbers.**

Define:
- Exact implementation scope
- Implementation sequence
- Architectural and workbook-state constraints
- Required validation
Validation should verify **logic and component interaction**, not trivial implementation details. Include VBA compile checks in the VBA IDE when feasible, deterministic in-Excel/manual verification steps, golden-output or structural workbook checks where available, and any existing automated harness/tests relevant to the changed scope.
- Use template `.ai\tasks\templates\task_template.yaml` to organize `{task_description}` in the `Execution plan` if present; otherwise define `{task_description}` clearly in the plan
**Important: Never change any VBA/code or existing audit phase files.**
<original_prompt>

Update `Execution plan` and Return `{task_description}`.


### 3.4 Implementor
Never parallel implementor agents only one at a time is allowed.

Launch one **`Implementor`** with the required context:
`{context_a} + {context_r} + {task_description}`
<original_prompt>
Implementor owns the local cycle:

- Implementation in versioned VBA sources (`.bas`/`.cls`/`.frm`) and related tracked artifacts as appropriate; keep exported VBA sources in sync with workbook behavior when feasible (e.g. `mko_treport/src/vbs/export/` and `mko_treport/src/vbs/vbaProject.bin` when applicable)
- Do not commit gitignored local data/workbook/output artifacts (`data/theiss/*`, generated outputs) unless explicitly tracked; record controlled local changes/repro steps in plan/evidence notes when required
- Required validation and relevant checks: VBA compile checks when feasible, deterministic in-Excel/manual verification for affected flows, and relevant automated checks for any non-VBA helper code touched (e.g. `uv run ruff check`/`uv run mypy`/`uv run pytest` for Python modules)
- Fixing failures
- Local validation against plan acceptance criteria and available golden/structural baselines
- Commit only the current versioned work item
```bash
git add <specific-files>
git commit -m "{type}({scope}): {description}"
```

You are working with other agents in parallel if you see changes not done by you - it is normal.
Never run `git reset`, `git checkout`
Never rewrite history.
- Return only when locally validated and committed
**Important: Never change plans or existing audit phase files.**
<original_prompt>

---


## 4. Documentation

After all implementation blocks are complete:

Launch `Doc-specialist`.

Ask it to update only the documentation affected by the implementation, following:

<original_prompt>
Project documentation conventions in `docs/` and `.ai/` (if present). Do not introduce unrelated documentation changes.

Commit documentation changes separately.
```bash
git add <specific-files>
git commit -m "{type}({scope}): {description}"
```
You are working with other agents in parallel if you see changes not done by you - it is normal.
Never run `git reset`, `git checkout`
Never rewrite history.
**Important: Never change any VBA/code or plans or existing audit phase files.**
<original_prompt>
---

## 5. Final Validation

Launch `Validator` for the completed implementation.

<original_prompt>
Check:

* Plan completeness
* Cross-block integration
* Architectural and workbook/VBA consistency
* Regressions
* Validation, tests, and quality gates (including VBA compile/manual/golden checks as applicable)
* Documentation
* Unrelated changes
* Important: Never change any VBA/code or existing audit phase or plan files.
<original_prompt>

## 6. If issues are found:
Launch one `Implementor` to fix and commit
Validate locally

---

# Constraints
The Tech Lead is **orchestrator only**:
- Coordinate agents
- Do **not** rewrite, summarize, replace, weaken <original_prompt>
- You may add relevant context, findings, decisions, risks, file/module/workbook information, or execution instructions 
- If you need to study the workbook/VBA implementation or artifacts, you launch researchers or auditors agents. You do not study yourself.

The Tech Lead understands project rules and general architecture but never implements code/VBA. 

---

# Expected Result

* Current workbook/VBA implementation and artifacts understood before execution planning
* Plan decomposed into dependency-safe work blocks
* Agent usage adapted to block complexity and risk
* Each block implemented, locally validated (with VBA-appropriate checks), and committed separately where versioned
* Documentation updated according to repository conventions
* Final validation passed
* Architecture and project conventions preserved

---

# Plan File

