# Work context record format

Use this structure for a task ledger. Store only fields that are evidenced, using `not stated` for required-but-missing values.

```yaml
id: stable-local-id
title: concise task title
status: candidate | confirmed | in_progress | blocked | completed | cancelled | superseded
evidence_class: ryo-confirmed | source-explicit | pending-confirmation
owner: Ryo | team | named-person | not-stated
source:
  type: ryo-instruction | meeting-minutes | task-system | repository | other
  reference: document, URL, file, or conversation reference
  location: section, page, timestamp, line, or not-stated
  observed_at: date-time or not-stated
purpose: stated purpose or not-stated
deadline: stated deadline or not-stated
priority: stated priority or not-stated
completion_condition: stated condition or not-stated
dependencies:
  - task id plus explicit relationship and evidence
blockers:
  - explicit blocker and evidence
result:
  summary: evidenced result or not-stated
  artifact: path, URL, PR, document, or not-stated
remaining:
  - evidenced remaining issue
supersedes: task id or not-stated
confirmed_by_ryo_at: date-time or not-confirmed
```

Keep AI proposals outside the record until confirmed:

```yaml
proposal:
  target_task: stable-local-id or new-candidate
  suggestion: proposed priority, dependency, next action, or interpretation
  reasoning: concise explanation
  evidence: supporting references
  status: ai-inference-awaiting-confirmation
```

Never rewrite the original evidence when Ryo corrects an extraction. Preserve the evidence reference and store Ryo's correction as the authoritative interpretation.
