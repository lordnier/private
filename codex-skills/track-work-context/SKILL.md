---
name: track-work-context
description: Extract and maintain Ryo's work context from meeting minutes, Jira issues or pasted Jira snapshots, direct instructions, task lists, decisions, and work results while strictly separating explicit evidence, Ryo-confirmed facts, pending candidates, and AI inference. Use when Ryo asks what he is responsible for, shares Jira tasks or meeting notes, reports task progress or completion, asks what became possible after prior work, or wants to update and review his current work state without inventing context.
---

# Track Work Context

Build an evidence-backed view of Ryo's work. Treat helpful interpretation as a proposal, never as a stored fact.

## Non-negotiable rules

1. Never store or present an AI inference as something Ryo said, intended, agreed, or owns.
2. Preserve the source and source location for every extracted assignment, deadline, decision, dependency, and completion claim.
3. Separate `Ryo-confirmed`, `source-explicit`, `pending confirmation`, and `AI inference` in both reasoning and output.
4. Add a candidate to the confirmed task ledger only when Ryo directly instructs it or explicitly approves it.
5. Do not invent a purpose, owner, deadline, priority, dependency, completion condition, outcome, or causal relationship.
6. Record missing fields as `not stated`; do not silently complete them from convention or project knowledge.
7. Do not mutate a task store unless Ryo asks to update it. Extraction and review requests are read-only by default.
8. Never commit meeting minutes, task records, or project data into this skill directory.

## Evidence classes

Use these classes exactly:

- `Ryo-confirmed`: Ryo directly instructed, corrected, accepted, or reported the item.
- `source-explicit`: The supplied source explicitly states the item, including an explicit assignment to Ryo. This is strong evidence but still requires Ryo's approval before entering the confirmed ledger.
- `pending confirmation`: The source suggests an action, but ownership, scope, deadline, or commitment is ambiguous.
- `AI inference`: A proposed relationship, priority, implication, or next action not explicitly stated by Ryo or the source.

Do not promote an item merely because it appears in formal minutes. Formality affects evidence quality, not whether Ryo accepted the extracted interpretation.

## Process meeting minutes

1. Read the original minutes and retain document date, meeting name, and relevant section or timestamp.
2. Extract only explicit action language, assignments, deadlines, decisions, blockers, and follow-ups.
3. Classify each item using the evidence classes above.
4. Separate items into:
   - explicit assignments to Ryo;
   - possible assignments requiring confirmation;
   - team or other-person tasks that block or enable Ryo's work;
   - decisions that may affect existing tasks.
5. Quote or tightly paraphrase the smallest useful source passage and point to its location.
6. Present proposed additions and changes before writing them.
7. Apply only the items Ryo approves. Preserve rejected candidates only when Ryo explicitly asks to retain them.

Do not treat Ryo speaking about a topic, volunteering information, or being present in the meeting as task ownership.

## Process direct instructions

When Ryo directly states a task, record only what he states. Keep unspecified fields as `not stated`. If ambiguity is not blocking, register the task without interrogating him and surface the missing fields briefly.

## Process Jira tasks

Treat Jira as the source of truth for Jira-owned fields and the local work ledger as a read cache for planning. Follow [references/jira-cache.md](references/jira-cache.md).

When Ryo supplies a Jira issue or pasted Jira snapshot:

1. Preserve the issue key, URL when available, summary, status, assignee, priority, due date, description, acceptance conditions, blockers, and Jira update timestamp when stated.
2. Record the local observation time separately from Jira's update time.
3. Do not infer missing Jira fields or treat a pasted summary as a live Jira read.
4. Let Ryo's direct priority instruction override Jira priority only in the local planning view. Preserve the original Jira priority unchanged.
5. Compare requirement claims in Jira with authoritative project sources. Record a conflict or possible stale statement instead of silently rewriting either source.
6. Never write changes back to Jira unless Ryo explicitly requests the exact update.

## Process completion and history

Do not mark work complete only because code, a document, or a message exists. Accept completion when Ryo reports it or when an agreed completion condition can be directly verified.

For a completed task, record only evidenced fields:

- what was completed;
- resulting artifact or verifiable change;
- explicitly stated outcome;
- explicitly stated remaining issue;
- confirmed successor or newly unblocked task;
- source references.

Label proposed lessons, causal relationships, and successor tasks as `AI inference` until Ryo confirms them.

## Maintain current state and history

Keep two views when a task store is available:

- `current state`: confirmed active tasks, blockers, waiting items, and next actions;
- `work history`: confirmed completed, cancelled, superseded, or merged work with evidence.

Use the schema in [references/record-format.md](references/record-format.md). Keep project data outside the skill directory and outside the skill's Git repository.

## Output contract

For an intake or review, report in this order:

1. `Confirmed current state`
2. `Source-explicit candidates`
3. `Needs Ryo's confirmation`
4. `AI suggestions` when useful
5. `Proposed ledger changes`

Omit empty sections. Make the boundary between facts and suggestions visually obvious. Never use confident prose to blur an uncertain classification.
