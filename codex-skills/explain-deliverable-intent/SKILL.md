---
name: explain-deliverable-intent
description: Explain the intent behind a deliverable and draft natural Japanese review communications. Use when a user has created or changed a UI, prototype, document, image, spreadsheet, code change, pull request, or other artifact and needs to communicate what changed, why it was designed that way, when it appears or operates, what is intentionally unsupported, which alternatives were rejected, or what reviewers should confirm. Also use for review requests, Slack messages, emails, PR descriptions, handoff notes, and situations where an intentional limitation could be mistaken for a defect.
---

# Explain Deliverable Intent

Turn implementation decisions into an explanation that prevents reviewers from mistaking intentional behavior for omissions or defects. Prefer evidence from the artifact, requirements, diffs, review comments, and the user's statements.

## Workflow

1. Inspect the available artifact and surrounding evidence before drafting.
2. Separate confirmed facts from inference. Ask one focused question when a missing fact would materially change the explanation; otherwise label uncertainty briefly.
3. Classify each change or decision before drafting:
   - **Worth sharing**: non-obvious design intent, behavior or display conditions, tradeoffs, intentional limitations, rejected alternatives, or anything a reviewer could reasonably mistake for a defect or omission.
   - **Not necessary to share**: self-evident corrections, minor visual cleanup, mechanical maintenance, and implementation details that do not require reviewer awareness or judgment.
   - When the user asks to assess or summarize multiple changes, show these as two separate sections: 「共有した方がよい内容」 and 「共有しなくてよい内容」. Give a brief reason for the classification of each item.
   - Keep items in the second section out of the send-ready review message unless the user explicitly asks to include them.
4. Identify the following when relevant for items worth sharing:
   - What changed or was created
   - What problem or misunderstanding it addresses
   - Why this design was selected
   - When the behavior appears, runs, or remains hidden
   - What is intentionally excluded or deferred, and why
   - Alternatives considered and the reason they were not selected
   - What the reviewer should confirm
5. Call out intentional non-display and non-operation conditions explicitly. Never leave a reviewer to infer whether an absent result is a defect.
6. Draft the requested communication for the user's channel. Default to a natural internal Slack message when the channel is not specified.
7. Check that every important limitation has a reason and that no unsupported rationale was invented.

## Default Output

When the user does not specify a format, provide one natural review-request message ready to send. Add a separate internal decision summary only when the user asks for one.

When the user asks which changes should be communicated, first provide the two-section classification described above. If the user also asks for a send-ready message, draft it from 「共有した方がよい内容」 only.

Choose the structure by the number of distinct decisions:

- For one material decision, use a short natural paragraph.
- For two or more decisions, use one bullet per change or design decision.
- In every bullet, pair the concrete change with its intent or reason. Do not separate all changes from all reasons.
- Keep the opening and confirmation request as prose outside the bullets.
- Combine only tightly related decisions. Do not bury separate rationales in one long paragraph merely to make the message shorter.

## Natural Japanese Tone

Write like a thoughtful colleague communicating inside a Japanese company.

- Keep the message polite, conversational, and specific.
- Lead with what was changed, then explain the decision and its reason.
- Use natural phrases such as:
  - 「〜を反映しました」
  - 「今回は、〜という理由でこの形にしています」
  - 「〜を避けるため、現時点では表示していません」
  - 「この認識で問題ないか、確認いただけると助かります」
  - 「気になる点があれば教えてください」
- Use 「〜していただけますと幸いです」 only when its formality fits the recipient and channel.
- Prefer short sentences and paragraph breaks for a single decision.
- Use bullets when there are multiple changes, design decisions, display conditions, or review points.
- Give each bullet a short descriptive label, followed by one to three natural sentences explaining both the change and its intent.
- Apologize only when an actual mistake, delay, or communication gap occurred.
- Explain a tradeoff calmly; do not sound defensive or blame the reviewer.

Avoid stiff or artificial language unless the user explicitly requests a formal external email:

- 「ご査収くださいますようお願い申し上げます」
- 「何卒ご高覧賜りますよう」
- 「ご確認いただけますと幸甚に存じます」
- 「下記の通り実装いたしました」
- Repeated 「〜させていただきました」
- A mechanical list with no explanation of the judgment behind it
- 「仕様です」 by itself without the reason

## Content Rules

- Do not invent intent merely because it sounds plausible.
- Distinguish current prototype behavior from the intended production behavior.
- State visibility and execution conditions concretely, including exceptions.
- Describe intentional limitations as decisions with reasons, not as excuses.
- Keep a review request compact, but do not remove the context required to prevent misunderstanding.
- Do not repeat the same rationale in a second summary unless the user explicitly requests an internal memo.
- Preserve the user's terminology and level of technical detail where practical.
- If the user provides recipients or a prior communication style, match them without exaggerating politeness.

## Review-Request Shape

Use this as a flexible reasoning order, not a fixed form:

1. Brief greeting when appropriate
2. What was changed
3. Why this design was selected
4. Important display, execution, or non-display conditions
5. Difference between prototype and production behavior, if relevant
6. A specific and natural confirmation request

When several decisions exist, format the middle as:

- **Short item label**  
  State what changed, then explain why it was designed that way. Include any relevant visibility condition or intentional limitation in the same item.

Avoid changelog-only bullets such as 「ボタンを追加」「文言を変更」 with no rationale.

Example tone for multiple decisions:

> お疲れさまです。強制発想法の図解をStep 9に追加しました。
>
> 今回の変更内容と意図は以下です。
>
> - **図解に表示する内容**  
>   選択した解決方向性と研究シーズから、アイデアが生まれた理由を追える構成にしています。結果だけでは組み合わせの意図が伝わりにくいためです。
>
> - **既存フローでの表示**  
>   内容が一致するサンプルフローから実行した場合に表示します。画面上の選択内容と図解が食い違うのを避けるためです。
>
> - **新規フローで表示しない理由**  
>   固定サンプルでは実際の入力内容とずれる可能性があるため、現時点では図解を表示していません。本番では案件ごとの内容から生成する想定です。
>
> この表示条件と内容で問題ないか、確認いただけると助かります。

Adapt wording to the artifact and evidence. Never copy this example when its facts do not apply.
