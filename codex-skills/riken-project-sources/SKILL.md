---
name: riken-project-sources
description: Ground RIKEN and SAT +理研 project answers, requirement checks, specification interpretation, and implementation decisions in the local project sources while applying the project source policy, authority, certainty, chronology, and supersession rules. Use when the user refers to 理研, SAT +理研, 仕様書, 提案書, キックオフ資料, 定例討議資料, 議事録, Mock, or related implementation work.
---

# 理研プロジェクト資料

Use the local originals under:

`/Users/ryosaito/Development/stockmarkteam/riken/spike/static-html-export/SAT +理研資料`

## Required workflow

1. Read `SAT +理研資料/SOURCE_POLICY.md` completely before selecting evidence.
2. List the available local sources and choose only those relevant to the question.
3. Read selected sources in their original formats. Inspect rendered pages or slides when tables, figures, wording placement, or visual structure matter.
4. Evaluate each material statement by:
   - statement status and certainty;
   - decision-making authority and meeting participants;
   - scope and topic;
   - chronology among statements of comparable authority;
   - originality of the source.
5. Apply later decisions only to the topics they explicitly change. Do not treat a newer document as replacing unrelated earlier requirements.
6. Separate `最新時点で確定している仕様`, `方針`, `未確定・検証予定`, `資料からの推論`, and `資料に記載なし` when relevant.
7. Cite source filenames for material claims. Include a page, slide, section, decision heading, or meeting timestamp when useful.
8. For implementation work, inspect the repository only after gathering source evidence. Treat repository behavior as implementation evidence, not automatically as an agreed requirement.

## Core decision rules

- Treat `決定事項`, `合意事項`, and `前回の決定・合意事項` as confirmed for that document's date, participants, scope, and authority.
- Allow a later source to supersede a confirmed item only when it has comparable or higher authority and clearly changes the same topic.
- Treat `方針`, `方向で合意`, and similar wording as adopted direction whose unresolved details must remain explicit.
- Treat `案`, `提案`, `たたき台`, `検討中`, `PoCで検証`, `可能性`, `今後議論`, examples, and hypothetical content as non-final.
- Prefer formal originals and approved records over AI-generated meeting summaries. Use Gemini-generated notes for discovery and provisional evidence; verify material claims against formal or direct sources when available.
- Report unresolved conflicts instead of silently selecting a winner.
- Do not fill missing requirements with web knowledge, conventions, repository behavior, or unstated assumptions. Label necessary choices as `実装上の仮定`.

## Maintenance boundary

Keep stable reasoning rules in this skill. Keep the changing source inventory, document relationships, known supersessions, and project-specific exceptions in `SAT +理研資料/SOURCE_POLICY.md`.

When a new source is added or its authority becomes clear, update `SOURCE_POLICY.md`; change this skill only when the evaluation method itself changes.
