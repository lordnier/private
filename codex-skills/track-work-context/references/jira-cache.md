# Jira cache policy

Keep Jira authoritative and use local records as a bounded cache.

## Cache layers

- Store a normalized record for planning: key, URL, summary, status, assignee, Jira priority, due date, acceptance conditions, blockers, and relevant relationships.
- Preserve the source observation time and Jira update time separately.
- Keep Ryo's local planning priority in a distinct field; never overwrite the Jira priority in the cached source record.
- Do not store credentials, session data, unrelated comments, attachments, or full project exports.
- Keep cached project data outside the skill Git repository.

## Freshness

- Answer from the cache when Ryo accepts cached context or live status is not material.
- Show `last observed` whenever status, assignee, deadline, or blocker may have changed.
- Refresh before making a consequential recommendation from a stale status or before claiming that Jira currently has a value.
- If live access is unavailable, use the cache and say so; do not imply a live read.

## Conflict handling

- Jira is authoritative for issue workflow fields such as status, assignee, Jira priority, and due date.
- Authoritative project sources govern agreed requirements and decisions.
- Ryo governs his local planning priority and confirmation of personal work context.
- Preserve conflicts between these authorities and identify the action needed to reconcile them.
