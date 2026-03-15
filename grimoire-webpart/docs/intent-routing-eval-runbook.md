# Intent Routing Eval Runbook

## Purpose

Validate first-turn capability routing after prompt changes. This runbook is required because prompt-only routing changes cannot be trusted from mocks alone.

## Scope

Check parity across:

- typed text with voice disconnected
- typed text while voice is connected
- spoken voice first turn

## Preconditions

- Frontend running against the deployed backend that serves the `grimoire-reasoning` deployment
- Browser console open
- Relevant logs visible from `LogService`
- Voice connected for the second and third passes

## Regression Utterances

Use these first-turn prompts exactly:

| Utterance | Expected route | Expected first tool |
| --- | --- | --- |
| `i am searching for informations about animals` | internal-first generic enterprise search | `search_sharepoint` |
| `i am searching for documents about animals` | internal content search | `search_sharepoint` |
| `search for nova marketing` | internal-first generic enterprise search | `search_sharepoint` |
| `find emails about animals` | email search | `search_emails` |
| `find people working on animals` | people search | `search_people` |
| `find sites about animals` | site search | `search_sites` |
| `look up animal migration on the web` | external web research via Copilot | `use_m365_capability` with `copilot_chat` |

## What To Check

For each utterance and modality:

1. Confirm the first meaningful capability call matches the expected tool.
2. Confirm the assistant does not ask the user to choose between Microsoft 365 and the web for the generic internal-first cases.
3. Confirm the right pane and left pane stay aligned with the chosen capability.

## Expected Logs

Look for one first-turn routing log per user turn:

- `First-turn routing: tool_call`
- `First-turn routing: clarification`
- `First-turn routing: answer_only`

The log detail should include:

- `channel`: `text` or `voice`
- `expectedRoute`
- `expectedToolName`
- `actualToolName` when a tool call happened
- `genericEnterpriseSearch`

## Regression Signal

Treat this as a failure:

- `Generic enterprise search fell back to clarification`

If it appears for `i am searching for informations about animals`, the internal-first default regressed.

## Manual Acceptance

The change is acceptable only if all of the following are true:

1. Typed text with voice disconnected routes the seven prompts as expected.
2. Typed text while voice is connected routes the same seven prompts the same way.
3. Spoken voice first turns route the same seven prompts the same way.
4. The generic internal-first prompt no longer triggers a domain-choice clarification.
5. No text-only pre-router or Nano-only classifier is required to obtain the behavior above.
