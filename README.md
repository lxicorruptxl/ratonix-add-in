# Ratonix Outlook Add-in

This repository contains an Outlook add-in surface for Ratonix AI communication intelligence. It brings the Chrome extension's publicly described flows into Outlook:

- rewrite the current draft in Balanced, Softer, or Firmer tones
- draft replies from a message being read
- insert suggestions back into compose or reply windows
- remember local tone/style examples and frequent relationships on the device
- optionally call a hosted Ratonix API for adaptive AI output

The original Chrome extension source was not present in this repository, so this implementation is a new Office.js add-in foundation with matching feature surfaces and an API seam for the production Ratonix intelligence service.

## Project layout

```text
manifest.xml          Office add-in manifest for Outlook
index.html            Task pane shell
commands.html         Ribbon command function file
src/main.ts           Task pane UI and interaction flow
src/outlook.ts        Office.js mailbox adapter
src/ratonixClient.ts  Ratonix API client plus local fallback suggestions
src/styleMemory.ts    Local-only style and relationship memory
```

## Development

Install dependencies:

```bash
npm install
```

Run the local task pane:

```bash
npm run dev
```

The checked-in manifest points to `https://ratonix.ai` for production use. For local sideloading, temporarily change the manifest URLs to your HTTPS dev tunnel or a local Office add-in certificate setup.

Build and typecheck:

```bash
npm run build
```

Validate the Office manifest:

```bash
npm run validate:manifest
```

## Backend configuration

The add-in works without a backend by using deterministic local tone guidance. To connect the adaptive Ratonix service, create `.env` from `.env.example`:

```bash
VITE_RATONIX_API_ENDPOINT=https://api.ratonix.ai/outlook
VITE_RATONIX_API_KEY=your-development-token
```

Expected API endpoints:

- `POST /rewrite`
- `POST /draft-reply`

Both endpoints receive:

```json
{
  "message": {
    "bodyText": "Message or draft text",
    "subject": "Subject",
    "from": "sender@example.com",
    "to": ["recipient@example.com"],
    "conversationId": "...",
    "mode": "compose"
  },
  "tone": "balanced",
  "memory": {
    "examples": [],
    "interactions": [],
    "keyRelationships": [],
    "relationships": {}
  }
}
```

## Vercel deployment and Microsoft Store submission

Deploy this project to the Ratonix Vercel project that serves `https://ratonix.ai`:

1. Configure the Vercel project root to this repository.
2. Use `npm run build` as the build command.
3. Use `dist` as the output directory.
4. Add `VITE_RATONIX_API_ENDPOINT` and `VITE_RATONIX_API_KEY` in Vercel if the hosted adaptive API is ready.
5. Confirm these URLs load over HTTPS after deploy:
   - `https://ratonix.ai/index.html`
   - `https://ratonix.ai/commands.html`
   - `https://ratonix.ai/assets/icon-64.png`
   - `https://ratonix.ai/assets/icon-128.png`

Before submitting to Microsoft Partner Center:

```bash
npm run validate:manifest
```

Then upload `manifest.xml` in Partner Center as a Microsoft 365 / Office Add-in offer. Microsoft Store validation requires the production URLs in the manifest to be public HTTPS URLs.

Both endpoints should return:

```json
{
  "text": "Suggested rewrite or reply",
  "rationale": "Short explanation shown in the task pane",
  "source": "api"
}
```
