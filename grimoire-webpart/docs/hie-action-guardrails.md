# HIE Header Action Guardrails

This document defines how header actions (`Focus`, `Summarize`, `Chat`) apply to right-pane blocks and selected items.

## Core rules

- Header actions are scoped to the **latest actionable block** in the right pane.
- Selection is block-local and uses left-side checkboxes where the block has row items.
- `Focus` supports single and multi selection.
- `Summarize` supports only summarizable kinds and runs on at most 5 items.
- `Chat` requires exactly 1 selected item.
- Follow-up mutation intents (reply/forward/delete email, update/cancel calendar event) require exactly 1 focused target.

## Block matrix

| Block type | Selection model | Focus | Summarize | Chat | Notes |
| --- | --- | --- | --- | --- | --- |
| `search-results` | Row checkboxes | Yes (multi) | Yes (doc/file) | Yes (single) | Main search result flow |
| `document-library` | Row checkboxes | Yes (multi) | Yes (files only) | Yes (single) | Folders are not summarizable |
| `markdown` | Numbered-card checkboxes | Yes (multi) | Yes (email/event/message/doc) | Yes (single) | Item kind inferred from fields |
| `list-items` | Row checkboxes | Yes (multi) | No by default | Yes (single) | Keeps summarize constrained |
| `permissions-view` | Row checkboxes | Yes (multi) | No | Yes (single) | Context/discussion only |
| `activity-feed` | Row checkboxes | Yes (multi) | No | Yes (single) | Context/discussion only |
| `file-preview` | Single implicit item | Yes | Yes | Yes | Auto-selected single item |
| `site-info` | Single implicit item | Yes | No | Yes | Context/discussion only |
| `user-card` | Single implicit item | Yes | No | Yes | Context/discussion only |
| `info-card` | Single implicit item | Yes | Yes | Yes | Useful for summary follow-up |
| `chart` | Single implicit item | Yes | No | Yes | Context/discussion only |
| `progress-tracker` | Single implicit item | Yes | No | Yes | Context/discussion only |

## Focus payload guardrails

When `Focus` is used, a silent structured context payload is injected with:

- block identity (`id`, `type`, `title`)
- selected items (`index`, `kind`, key fields like subject/from/itemId/date)
- guardrail hints (`singleSelectionRequiredFor`, notes)

Guardrail note examples:

- `Email reply/forward/delete must target exactly one focused email.`
- `Calendar update/cancel must target exactly one focused event.`

This is intended to make follow-up commands like “reply” or “update the event” deterministic and safe.

## Runtime notes

- `show_selection_list` is suppressed when current actionable results are already visible and the prompt is only generic chooser language (pick/open/preview/summarize), unless the user explicitly asks for options/radio buttons.
- User cards should degrade gracefully if `photoUrl` fails to load (placeholder icon fallback).
- Browser-console `MutationObserver` stack traces from `VM...` scripts are treated as external unless reproduced from bundled Grimoire sources.
