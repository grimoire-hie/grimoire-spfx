# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Grimoire is a visual AI assistant for Microsoft 365 — an SPFx web part with an SVG avatar, dual-path voice/text interaction, tri-source search, MCP server integration, and a rich UI block system.

## Build & Development Commands

All commands run from `grimoire-webpart/`:

```bash
cd grimoire-webpart

# Build (includes tests)
npx heft build --clean

# Build for production (build + test + package .sppkg)
npm run build    # = heft test --clean --production && heft package-solution --production

# Run tests only
npx heft test --clean

# Run a single test file (use heft's jest passthrough)
npx heft test --clean -- --testPathPattern="McpResultMapper"

# Local dev server (SharePoint workbench)
npm run start    # = heft start --clean

# Lint (eslint is invoked automatically during heft build; no standalone script)
npx eslint src/ --ext .ts,.tsx
```

Backend (from `grimoire-backend/`):

```bash
cd grimoire-backend
npm install
func start              # Run locally (requires Azure Functions Core Tools)
npm run test            # Smoke tests
npm run deploy          # Deploy to Azure
```

Documentation site: separate repo [grimoire-hie.github.io](https://github.com/grimoire-hie/grimoire-hie.github.io).

## Architecture

### Two Packages

- **`grimoire-webpart/`** — SPFx 1.22 web part (TypeScript, React 17, Fluent UI v8, Zustand)
- **`grimoire-backend/`** — Azure Functions v4 (plain JS, ESM) — LLM proxy, MCP session management, user persistence

### Frontend Source Layout

All frontend code lives under `grimoire-webpart/src/webparts/grimoireAssistant/`:

- **`components/`** — React UI: `GrimoireAssistant.tsx` (root), `action-panel/blocks/` (block types defined in `IBlock.ts`), `action-panel/interactionSchemas.ts` (declarative block action → HIE event mapping), `avatar/` (SVG avatar + `SvgAvatarMotionTuning`), `layout/`, `widget/`
- **`services/`** — Business logic (no React). Key modules:
  - `realtime/` — `RealtimeAudioService` (WebRTC voice), `TextChatService` (HTTP SSE text), `SystemPrompt`, `RealtimeVoiceCatalog`
  - `tools/` — `handleFunctionCall.ts` (shared voice+text dispatch), runtime handler partitions by category, `ToolAcknowledgment` (deterministic localized acks)
  - `hie/` — Hybrid Interaction Engine (event-driven state projector): HieStateReducer, HieContextProjector, HieInteractionFormatter, HieArtifactLinkage, HieTurnStartPolicy, BlockTracker, ContextInjector, FlowOrchestrator, VerbosityDirector, DataExpressionDirector, HieArtifactKindResolver, HiePromptProtocol, HAEContracts, HIETypes, HybridInteractionEngine
  - `mcp/` — `McpClientService` (HTTP client to backend), `McpResultMapper`, `EntityParsers`, `GenericBlockBuilder`
  - `search/` — `CopilotSearchService`, `CopilotRetrievalService`, `SharePointSearchService`, `RRFFusionService`, `SearchIntentPlanner`
  - `avatar/` — `ExpressionEngine`, `PersonalityEngine`, `SpeechMouthAnalyzer`, `IdleAnimationController`, `AmbientSoundAnalyzer` (legacy particle system files retained but not active)
  - `context/` — `ContextService` (Graph user profile), `PersistenceService` (backend table storage), `ConversationLanguage` (5-language detection + sticky switching)
  - `sharing/` — `ContextualComposeScope` (NL scope parsing), `SessionShareFormatter`, `ShareSubmissionService`
  - `forms/` — `FormPresets` (14 templates), `FormSubmissionService`
- **`config/`** — `toolCatalog.ts` (tool definitions — single source of truth for tool count), `promptCatalog.ts`, `webPartDefaults.ts`
- **`store/`** — `useGrimoireStore.ts` — single Zustand store
- **`hooks/`** — `useVoiceSession.ts` (voice orchestration + HIE integration), `useMcpServers.ts`
- **`models/`** — `IBlock.ts` (block type definitions), `IMcpTypes.ts`, `ISearchTypes.ts`, `McpServerCatalog.ts`, `catalog/` (server definitions)

### Dual-Path Voice/Text

- Voice connected → text and voice go through WebRTC data channel (`RealtimeAudioService`)
- Voice NOT connected → text goes through HTTP streaming SSE (`TextChatService`)
- Both paths share `handleFunctionCall()` from `services/tools/handleFunctionCall.ts`

### Backend Routes

All traffic from the browser to LLM/MCP services routes through the Azure Functions backend. Key endpoints:
- `GET /api/health` — health check
- `POST /api/realtime/token` — ephemeral WebRTC token
- `*/api/{backend}/openai/deployments/{deployment}/chat/completions` — LLM proxy (`backend` = `reasoning` or `fast`)
- `POST /api/mcp/connect|execute|disconnect` — MCP session lifecycle
- `POST /api/mcp/discover` — discover single MCP server tools
- `POST /api/mcp/discover-all` — discover all MCP server tools
- `GET /api/mcp/sessions` — list active MCP sessions
- `POST /api/user/notes|preferences` — Table Storage persistence

## TypeScript / SPFx Constraints

The SPFx build rig targets **ES5** (via `@microsoft/spfx-web-build-rig/profiles/default/tsconfig-base.json` → `"target": "es5"`). This causes several restrictions:

- **No `for...of` on `Map`/`Set`** — use `.forEach()` instead
- **No `u` flag on RegExp** — for emoji removal, use surrogate pair ranges: `\uD83C[\uDF00-\uDFFF]`
- **No `\.` or `\)` inside character classes** — ESLint `no-useless-escape` fires; use `[.)]` instead
- **No `new RegExp(...)` with user input** — ESLint `@rushstack/security/no-unsafe-regexp`; use `indexOf`-based search
- **React 17** — no `useId`, no automatic JSX transform; explicit `React.createElement` is the compiled output
- **Fluent UI v8** — no `dropdownItemHovered` style key; use `selectors: { '&:hover': {...} }` inside `dropdownItem`

## Logging API

Use `logService` singleton (not `console.log` directly in services):

```typescript
import { logService } from '../logging/LogService';
logService.info('mcp', 'Connected to server', serverName);
logService.warning('search', 'Empty results');
logService.error('voice', 'WebRTC failed', error.message);
logService.debug('graph', 'Response received');
```

Methods: `.info()`, `.debug()`, `.warning()`, `.error()` — NOT `.log()`.
Categories: `'mcp' | 'llm' | 'search' | 'graph' | 'system' | 'voice'`

## Key Types

- **Expression**: `'idle' | 'listening' | 'thinking' | 'speaking' | 'surprised' | 'happy' | 'confused'`
- **BlockType**: see `IBlock.ts` — `'search-results' | 'document-library' | 'file-preview' | 'site-info' | 'user-card' | 'list-items' | 'permissions-view' | 'activity-feed' | 'chart' | 'confirmation-dialog' | 'selection-list' | 'progress-tracker' | 'error' | 'info-card' | 'markdown' | 'form'`
- **IFunctionCallStore**: minimal store interface for tool dispatch (defined in `services/tools/ToolRuntimeContracts.ts`)

## Testing

Tests use **Jest via Heft** (`@types/heft-jest`). No standalone jest config file — the SPFx build rig provides it.

```bash
# All tests
npx heft test --clean

# Single test file by pattern
npx heft test --clean -- --testPathPattern="McpResultMapper"

# Single test by name
npx heft test --clean -- --testNamePattern="should parse email"
```

Test files are co-located with source (`*.test.ts` next to `*.ts`). The pattern uses `jest.mock()` for service dependencies and `jest.fn()` stores implementing `IFunctionCallStore`.

## ESLint

Config: `grimoire-webpart/.eslintrc.js` extends `@microsoft/eslint-config-spfx/lib/profiles/react`. Key enforced rules:
- `@typescript-eslint/no-floating-promises: error` — all promises must be awaited or voided
- `@typescript-eslint/no-for-in-array: error`
- `guard-for-in: error`
- `max-lines: warn (2000)`
- `@rushstack/security/no-unsafe-regexp: warn`
- `@typescript-eslint/explicit-function-return-type: warn` (with `allowExpressions: true`)

## HIE (Hybrid Interaction Engine)

The HIE is an event-driven state projector that orchestrates context between visual UI blocks and the LLM. Located in `services/hie/`:

**Core state machine:**
- **HieStateReducer** — pure reducer: 30+ typed events → derived state (task context, artifacts, shell state)
- **HieContextProjector** — converts derived state + events into narrative prompt text for the LLM
- **HieTurnStartPolicy** — decides whether a new user turn inherits the current thread or starts a new root

**Block & context tracking:**
- **BlockTracker** — tracks active blocks, lifecycle, per-type summarization with numbered references
- **ContextInjector** — injects `[Visual context:]`, `[User interaction:]`, `[Flow update:]` messages to LLM via `sendContextMessage()`
- **HieInteractionFormatter** — per-action prompt templates separating trusted action text from untrusted payload data

**Artifact lineage:**
- **HieArtifactLinkage** — resolves artifact chains (e.g., recap ← summary ← search) and current artifact context
- **HieArtifactKindResolver** — maps block types to artifact kinds (block, summary, preview, lookup, recap, form, share, etc.)

**Advisory layers:**
- **FlowOrchestrator** — multi-step flow detection (search-then-drill, browse-then-open, compose-then-submit)
- **VerbosityDirector** — computes `minimal|brief|normal|detailed` from block density
- **DataExpressionDirector** — priority-based avatar expression rules from data events

**Integration:** `useVoiceSession.ts` calls `hie.onBlockCreated/Updated/ToolComplete`; `ActionPanel.tsx` uses `interactionSchemas.ts` to declaratively map block actions to HIE events. HIE is page-session scoped, initialized lazily on voice connect or first text message.

## MCP Tool Flow

1. LLM calls `use_m365_capability` or `call_mcp_tool`
2. `handleFunctionCall` → `connectToM365Server` (auto-connects if needed) → `McpClientService.execute()`
3. Result → `McpResultMapper.mapMcpResultToBlock()` → `EntityParsers.tryParseEntityReply()` or `GenericBlockBuilder` fallback
4. Block pushed to Zustand store → ActionPanel renders it → HIE injects context back to LLM
