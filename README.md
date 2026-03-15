# GriMoire

An AI assistant for Microsoft 365 with reflexive UI awareness.

GriMoire is a SharePoint Framework web part that combines an LLM assistant with a visual block system, intent-based search, and Microsoft 365 actions. The Hybrid Interaction Engine (HIE) keeps the model grounded in the UI it creates — interactions with blocks feed back into the conversation as structured events. All Microsoft 365 calls use delegated access under the signed-in user's identity.

## Repository structure

```
grimoire-webpart/   SPFx 1.22 web part (TypeScript, React 17, Fluent UI v8)
grimoire-backend/   Azure Functions v4 backend (LLM proxy, MCP routing, persistence)
```

Documentation site: [grimoire-hie.github.io](https://grimoire-hie.github.io) ([repo](https://github.com/grimoire-hie/grimoire-hie.github.io))

## Quick start

```bash
# Clone
git clone git@github.com:grimoire-hie/grimoire-spfx.git
cd grimoire-spfx

# Deploy the backend
cd grimoire-backend
npm install
npm run deploy

# Build the web part
cd ../grimoire-webpart
npm install
npm run build

# Output: sharepoint/solution/grimoire-spfx.sppkg
```

Upload the `.sppkg` to your tenant app catalog, approve API permissions, and configure the web part property pane with the backend outputs.

## Documentation

Full setup guide, architecture, and concepts: **https://grimoire-hie.github.io**

## Prerequisites

- Microsoft 365 tenant with a tenant app catalog
- Microsoft Frontier AI Program enrollment
- Azure subscription (Function App, AI resource, storage)
- Node.js 22+, Azure CLI, Azure Functions Core Tools
- PowerShell 7+ (for the Agent 365 service principal script)

See [Prerequisites](https://grimoire-hie.github.io/docs/getting-started/prerequisites) for the full checklist.

## Acknowledgements

- Lip sync mouth shape classification adapted from [lipsync-engine](https://github.com/Amoner/lipsync-engine) (MIT License, Beer Digital LLC)

## License

[MIT](LICENSE)

## Authors

- [Nello D'Andrea](https://www.linkedin.com/in/nello-d-andrea/) ([GitHub](https://github.com/ferrarirosso))
- [Nicole Beck Dekkara](https://www.linkedin.com/in/nicole-beck-dekkara/) ([GitHub](https://github.com/NicoolB))
