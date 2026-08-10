---
name: format-readable-message
description: Format supplied content into a polished, copy-ready plain-text message with deliberate line breaks, compact sections, restrained bullets, and medium-appropriate emphasis. Use only when Ryo explicitly invokes this skill or asks to apply this formatter to a Slack post, chat message, review request, status update, handoff, announcement, or other paste-ready communication.
---

# Format Readable Message

Format the content without changing its meaning. Optimize for the message as it will appear after pasting, not for Markdown source aesthetics.

## Preserve content

1. Preserve facts, intent, uncertainty, terminology, links, names, requests, and tone.
2. Do not add a rationale, promise, priority, deadline, apology, or conclusion that the source does not support.
3. Resolve only harmless wording friction. Ask one focused question only when ambiguity would materially change the message.
4. Remove repetition and mechanical labels when doing so does not remove necessary context.

## Select the target format

- For Slack, use Slack-native plain text with minimal `mrkdwn` only when it improves scanning. Allow `*short label*`, ordinary bullets, and direct URLs. Do not use Markdown headings, tables, fenced code blocks, or decorative separators unless the content itself requires code.
- For an unspecified chat medium, default to portable plain text. Do not emit Markdown heading markers, tables, HTML, or Slack-only syntax.
- For a named medium, use only formatting that renders reliably in that medium. When uncertain, fall back to portable plain text.

Treat “plain text” as paste-ready text without document-style Markdown structure. Minimal Slack emphasis is allowed only for a Slack target.

## Shape the message

1. Open with the purpose or outcome in one or two short sentences. Add a greeting only when natural for the recipient and channel.
2. Keep one idea cluster per paragraph.
3. Insert exactly one blank line between meaningful sections. Do not insert blank lines between every sentence or every bullet.
4. Use a short section label only when the message has at least two distinct information groups or needs a clear action boundary.
5. Use bullets only for two or more parallel items, decisions, changes, conditions, or checks. Keep a single item as prose.
6. Keep bullets grammatically parallel and at a similar level of detail.
7. Use one indentation level only when a child item genuinely depends on a parent item. Avoid decorative or deeply nested indentation.
8. End with the requested action, confirmation point, or next step. Do not repeat the opening.

## Control density

- Prefer two to five compact sections for a substantial internal message.
- Prefer paragraphs of one to three sentences.
- Split a paragraph when its purpose changes, not merely because it looks long.
- Combine adjacent one-line sections when their separation creates visual noise.
- Avoid more than six bullets in one group. Regroup only when the groups have meaningful labels.
- Avoid consecutive headings with little or no content between them.

## Use emphasis sparingly

For Slack, bold only short labels, decisions, warnings, deadlines, or the exact review point. Do not bold full sentences or every bullet label. For portable plain text, create hierarchy through wording and spacing rather than markup.

Do not use emoji as structural bullets. Retain an emoji from the source when it fits the tone; otherwise add none unless Ryo requests it.

## Output contract

Return one final message that can be pasted directly into the target medium.

- Output the message body only.
- Do not introduce it with “formatted version,” “draft,” or an explanation.
- Do not wrap it in a Markdown code fence.
- Do not provide multiple variants unless Ryo asks for alternatives.
- Preserve URLs as pasteable text.

Before returning, inspect the message visually and remove excessive blank lines, unnecessary labels, uneven indentation, isolated bullets, and markup that will not render in the target medium.
