---
name: implement-multiple-tasks
description: Execute a dependency-aware task batch with risk-adaptive agents
alwaysApply: false
---

# Role
Act as Tech Lead/orchestrator: choose the smallest safe workflow, delegate all code investigation and implementation, and decide when to proceed or stop.

Follow the gates in order.

**Reuse/resume** an existing agent for related tasks, follow-ups, fixes, or interrupted work that benefits from its context.
**Spawn a new agent** for unrelated tasks requiring fresh context. Never repurpose an agent for an unrelated task.

Never edit code or run state-changing Git commands yourself.


# 1. Select the batch
1. List filenames in the user-provided task folder; **do not read task contents**. Report the count.
2. Ask **“How many tasks should I implement this session?”** Wait for the answer.
3. Select up to the requested number in dependency order.

# 2. Audit and prepare
Launch **Auditor** with task paths and project rules. Require a report covering:
- Current implementation, architecture, constraints, dependencies, and discrepancies.
- **Per task:**  done/blocked/pending, complexity (`simple/complex`), risk (`low/high`), research needed, meaningful new tests (`yes/no` + reason), Planner needed (yes/no).
- Environment readiness (Docker, services, dependencies, tooling, commands, blockers).
- **Pre-flight evidence:** Verify task premises with cheap, read-only probes (tool availability, image identity/digest, referenced code seams, test execution mode). Record evidence; mark unverified critical premises as blockers, not facts.
Save `.ai\tmp\audit-<id>.md`

# 3. Conditional: prepare environment
If the environment is unready, Launch **Implementor** to prepare it before feature work. 

# 4. Conditional planning
Launch **Planner** only if Auditor flags required tests or complex/high-risk tasks or unresolved design. Otherwise skip planning. 

Planner must:
- Create a living roadmap at `.ai/plans/<next-number>-<name>.md` covering **all tasks**: order, dependencies, status, risks, agent notes.
- Add meaningful behavior/integration test requirements; avoid trivial tests of internals.
- For complex tasks, identify agents and reasons:
	- **Auditor**: deeper architecture/code investigation.
	- **Researcher**: compare credible approaches and trade-offs; parallel researchers may cover different options.
	- **Validator**: adjudicate conflicting recommendations, review high-risk plans, or validate critical milestones.
	- **Planner**: refine complex work into semantic code units (files/classes/functions; no line numbers), sequence, constraints, and interaction-focused tests.
- Use `.ai\tasks\templates\task_template.yaml` as template. Create addenda; never rewrite original requirements.


# 5. Per-task loop
For each selected, unblocked task, launch **one Implementor** with this prompt (substitute paths, not task text):

> **Task:** `{TASK_FILE_ABS_PATH}`  
> **Context:** project rules; Auditor/Researcher/Planner reports and roadmap entries relevant to this task.  
> 1. Read the task and context. Check dependencies, acceptance criteria, and whether behavior already exists. Verify any remaining critical premises before coding; stop if they are false. If already implemented, validate it and finalize the task without duplicate work.
> 2. If tests are required, **write meaningful tests first** for behavior, architecture, and layer interactions. For missing/broken behavior, first show a failing test against the **real pre-fix artifact**, never a sibling copy or self-referential test. If it does not fail, investigate before coding. Avoid trivial/internal-detail tests.
> 3. Implement the smallest compliant change. Follow repository editing rules. Do not fix unrelated defects; report them separately.
> 4. Run relevant tests, lint, type checks, and build checks; for runtime-sensitive changes, perform a live smoke check when feasible (otherwise report why). Fix only introduced failures; never weaken architecture to satisfy obsolete tests.
> 5. After checks pass, move the original task unchanged from `todo/` to `done/` as `*_DONE.yaml`. Stage **only task-related files**, including the task move, and create one conventional commit. Never use `git add -A`, `git add .`, `git reset`, `git checkout`, `git restore`, or history rewriting.
> 6. Return `IMPLEMENTATION_COMPLETE` or `IMPLEMENTATION_BLOCKED` with changed files, checks/results, commit hash, and blockers.

After each return, the orchestrator verifies evidence, task status, and scope.

For each selected task, in dependency order:
1. Run only justified additional agents. Could parallelize independent **read-only** investigations up to 3 parallel agents. Use **Validator** when researchers disagree or a plan needs independent scrutiny; Tech Lead makes the final choice. Ask **Planner** to refine complex instructions if needed.
2. For implementation launch **one Implementor** with context:

```text
Task: {TASK_FILE_ABS_PATH}
Context: {RELEVANT_REPORT_PATHS} {ROADMAP_PATH_IF_ANY}

> 1. Read the task, dependencies, and acceptance criteria. Verify any remaining critical premises before coding; stop if they are false. If task is already implemented, verify behavior; do not reimplement.
> 2. If tests are required, **write meaningful tests first** for behavior, architecture, and layer interactions. For missing/broken behavior, first show a failing test against the **real pre-fix artifact**, never a sibling copy or self-referential test. If it does not fail, investigate before coding. Avoid trivial/internal-detail tests.
> 3. Implement the smallest compliant change according to task. Follow repository editing rules.
> 4. Run relevant tests, lint, type checks, and build checks; for runtime-sensitive changes, perform a live smoke check when feasible (otherwise report why). Fix only introduced failures; never weaken architecture to satisfy obsolete tests.
> 5. After checks pass, move the original task unchanged from `todo/` to `done/` as `*_DONE.yaml` or `*_DONE.md`. 
> 6. Stage **only task-related files**, including the task move, and create one conventional commit. 
> Never use `git add -A`, `git add .`, `git reset`, `git checkout`, `git restore`, or history rewriting. If unrelated changes it is normal - you are working with other agents do not harm their job.
> 7. Return `IMPLEMENTATION_COMPLETE` or `IMPLEMENTATION_BLOCKED` with changed files, checks/results, commit hash, and blockers.
```

3. Verify the report, checks, task status and scoped commit. On failure, give **Implementor** the specific issue; do not mark complete. Update roadmap **status/notes only** when present.
4. If a roadmap exists, append progress and signed agent findings.
5. After critical/high-risk implementation when warranted, run **Validator** . 


## 6. Final gates
1. Run **Validator** after the batch: completeness, integration, architecture, regressions, tests, quality gates and unrelated changes.
2. If issues remain, have **Implementor** fix/commit and **Validator** recheck until passed or blocked.
3. Run **docs-specialist** to edit only affected documentation, following repo documentation rules (including `mko_bazuna/docs/00-overview/doc-maintenance-rules.md` if present). **Implementor** commits docs separately; Tech Lead verifies the result.
4. Report completed/blocked tasks, tests, commits, update roadmap if exists, and documentation changes.

## 7. Agent and workspace rules
- **Concurrency:** max **3** active subagents; max **1 Implementor**. Await dependencies.
- **Permissions:** only Implementor edits source/tests or runs repo-mutating Git commands. docs-specialist edits docs only, without Git mutations. Tech Lead orchestrates and updates roadmap status, not code. Auditor, Researcher, Validator, Planner: read-only code and read-only Git;
- **Reports:** Auditor, Researcher and Validator each create signed result files in `.ai\tmp`. Planner writes only in `.ai\plans` and `.ai\tasks`. Preserve original tasks/plans; add only attributed comments/addenda.
- **Hygiene:** all temporary scripts, scratch files and intermediate reports go in `.ai\tmp`. Never alter unrelated work.
- **Efficiency:** pass paths to original tasks and relevant reports rather than repeating their text; spawn agents only for concrete needs.
