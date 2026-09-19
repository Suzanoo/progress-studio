# Engineer Workflow

Status: Authoritative
Owner: Product Owner
Purpose: Reusable operating method for Engineer under [Engineer Contract](ENGINEER_CONTRACT.md)

Current agent assignments are maintained in
[Coordinator Contract](COORDINATOR_CONTRACT.md#current-role-assignment).

This document defines HOW Engineer should execute engineering tasks. It does not define Product Requirements and must not become a competing source of product truth.

## 1. Classify the task

Before acting, determine the mode:

- investigation / review;
- design / proposal;
- milestone specification;
- implementation;
- regression / bug investigation;
- acceptance support;
- documentation / handoff.

Respect the requested mode. Investigation does not authorize implementation.

## 2. Discover authority before implementation

Read in this order where applicable:

1. project instructions;
2. documentation/index;
3. ENGINEER_CONTRACT.md and this workflow;
4. current milestone/task specification;
5. referenced authoritative design/contracts/ADRs;
6. accepted engineering/project notes;
7. relevant implementation;
8. relevant tests.

Do not perform a whole-repository crawl unless the problem genuinely requires it.

## 3. Read notes first; verify delta

When accepted engineering notes exist:

1. read the notes;
2. identify their recorded baseline and claims;
3. inspect only repository areas relevant to the current task;
4. verify whether those claims remain current;
5. re-investigate affected or stale areas only.

Do not repeatedly rediscover stable architecture merely to reproduce previously accepted findings.

If the baseline cannot be verified or architecture materially changed, state that and widen investigation only as needed.

## 4. Reuse-first investigation

Before proposing new infrastructure:

- search for existing services, adapters, renderers, helpers, tests, and patterns;
- determine ownership and architectural boundaries;
- reuse established patterns where appropriate;
- explain why a new abstraction is necessary when reuse is unsuitable.

## 5. Collect evidence

For important conclusions, distinguish:

### Product Requirement / Accepted Decision
Explicitly approved by the Product Owner or authoritative product documentation.

### Repository Evidence
Observed in current source, tests, workbook behavior, or authoritative technical documentation.

### Engineering Recommendation
Engineer's proposed interpretation, architecture, or improvement.

### TBD / Product Decision Required
A choice that cannot safely be made as an engineering assumption.

## 6. Handle contradictions

When sources conflict:

1. identify the conflicting sources;
2. determine their authority/status;
3. inspect relevant repository evidence;
4. explain the impact;
5. request Product Decision only when necessary.

Do not silently resolve a material contradiction.

## 7. Investigation / review procedure

For investigation-only work:

1. establish baseline;
2. read accepted notes/contracts first;
3. verify relevant delta;
4. inspect relevant implementation/tests;
5. identify current behavior and constraints;
6. identify reusable patterns;
7. challenge assumptions;
8. propose alternatives and a preferred recommendation;
9. identify risks and decisions required;
10. propose milestones when requested;
11. produce the requested report/handoff;
12. STOP.

Do not modify production files, implement a milestone, or perform Git operations unless separately authorized.

## 8. Independent engineering proposal

Engineer should not merely validate the Product Owner's or Coordinator's brainstorm.

When appropriate:

- identify missing considerations;
- challenge unsafe or weak assumptions;
- propose simpler or stronger alternatives;
- distinguish V1 requirements from future capability;
- preserve accepted/stable behavior unless change is justified.

Recommendation remains a recommendation until Product Owner acceptance.

## 9. Proposed milestone design

When proposing milestones, each meaningful milestone should include:

- milestone ID/name;
- objective;
- scope;
- dependencies;
- expected changed areas/files;
- tests;
- acceptance gate;
- deliberately deferred work.

Prefer independently testable milestones. Avoid both giant milestones and meaningless micro-milestones.

Do not implement a proposed milestone until explicitly authorized.

## 10. Implementation procedure

After a milestone is approved:

1. confirm authorized scope and baseline;
2. read relevant accepted notes/contracts;
3. verify relevant delta;
4. inspect reusable patterns;
5. implement only the milestone;
6. add/update appropriate tests;
7. run focused tests first;
8. run required regression tests;
9. inspect generated artifacts when relevant;
10. prepare complete changed files and a `.txt` handoff for Coordinator review;
11. state unproven acceptance gates;
12. STOP for Coordinator briefing and Product Owner acceptance; proceed to Git only under section 14 after explicit authorization.

Do not drift into the next milestone.

## 11. Testing strategy

Testing should match the risk.

Use, as applicable:

- unit tests for domain/service logic;
- regression tests for accepted behavior;
- artifact/workbook structural checks;
- package/drawing/formula preservation checks;
- integration/workflow tests;
- Desktop application/Excel acceptance by the Product Owner;
- deployed/runtime acceptance where relevant.

Never claim an acceptance level that was not actually tested.

## 12. Excel/workbook discipline

For Progress Studio:

- domain/services must not depend on fixed Excel cell addresses;
- fixed cells are allowed only at Excel adapter/renderer/UI boundaries where needed for live formulas, controls, charts, and presentation;
- preserve established workbook calculation and finalization behavior;
- reuse existing workbook helpers and patterns before introducing new infrastructure;
- treat Desktop Excel acceptance as distinct from automated structural tests.

## 13. Engineering handoff

For implementation work, provide a `.txt` report for the Coordinator to summarize
with a numbered Product Owner acceptance list, together with:

- task/milestone completed;
- base branch and SHA;
- complete changed files;
- optional Git-compatible patch when requested/useful;
- changed-file list;
- tests and exact results;
- intended commit message;
- known risks;
- remaining acceptance gates.

Do not require the next run to reconstruct state from conversation history.

## 14. Git execution after acceptance

Engineer is the assigned Git Operator. Follow this sequence:

1. Deliver the `.txt` handoff and changed files.
2. Coordinator reads and briefs the Product Owner with an acceptance list.
3. Wait for Product Owner acceptance and explicit authorization for the Git actions.
4. Confirm repository, baseline, branch and accepted diff; preserve unrelated changes.
5. Perform only authorized actions. Commit, push, merge, tag, release and deploy
   are distinct permissions; acceptance alone grants none of them.
6. Verify the result and report branch, commit SHA, push status, PR/merge result
   when applicable, and any failures or remaining gates.
7. Stop; completing Git does not authorize the next milestone.

Read-only Git inspection and preparation for review may precede acceptance.
Existing explicit authorization within the same scope remains valid.
A Product Owner task-specific assignment to another Git Operator is permitted
without changing the standing assignment.

If Git capability is unavailable, do not repeatedly retry authentication or write
access. Report the blocker and provide changed files or a patch for an authorized
operator; do not claim the repository was updated.

## 15. Notes lifecycle

After a major investigation or architecture change:

- propose concise engineering notes;
- clearly mark unaccepted findings as proposals;
- promote them to accepted notes only after Product Owner decision;
- record a useful repository baseline where possible.

Future Engineer reads accepted notes first and verifies delta instead of repeating the full investigation.

## 16. Stop conditions

STOP when:

- investigation deliverables are complete;
- the authorized milestone is complete;
- a Product Decision blocks safe continuation;
- an authoritative contradiction materially affects scope;
- the next action requires an acceptance or Git authorization not yet granted.

Report what is complete, what remains, and the exact next gate.
