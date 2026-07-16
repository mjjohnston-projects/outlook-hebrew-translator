# Outlook Hebrew Translator

An Outlook compose add-in that turns an English draft into a copy-ready Hebrew email. It asks for the recipient's gender so gendered Hebrew wording is correct, suggests an English subject, supplies the Hebrew subject, and preserves line breaks and links in the body.

## What it does

1. Open a new email or reply in Outlook and select **Hebrew Translator**.
2. Choose **Man**, **Woman**, **Men**, or **Women** for the person or people being addressed. Hebrew agreement follows the selected singular/plural gender.
3. Select **Translate draft**.
4. Review the result. Each response starts with the selected gender, then shows the English subject suggestion, Hebrew subject, and the Hebrew body.
5. Use **Apply Hebrew subject & body** to replace the draft, or copy either field.

`Regards, Michael` is explicitly rendered as `בברכה, מייקל`.

## Run locally

Prerequisites: Node.js 20+ and an OpenAI API key.

```powershell
cd C:\Users\michaeljo1\Documents\Codex\2026-07-16\i-wou\outputs\outlook-hebrew-translator
Copy-Item .env.example .env
# Edit .env and set OPENAI_API_KEY and OPENAI_MODEL.
npm install
npm run dev
```

The add-in must be hosted over HTTPS for Outlook. For this personal Windows setup, use `https://localhost:3000` with a trusted development certificate, then sideload `manifest.xml` in Outlook.

To avoid missing a URL in the manifest, configure it once with:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\set-manifest-url.ps1 -Origin https://localhost:3000 -ReplaceExisting
```

## Deploy

Deploy this app to any HTTPS-capable Node host. Set `OPENAI_API_KEY` and `OPENAI_MODEL` as server secrets, set `PUBLIC_ORIGIN` to the deployed HTTPS origin, replace `YOUR-HTTPS-DOMAIN` in the manifest, then distribute the manifest through Microsoft 365 admin integrated apps or sideload it for individual users.

## Start with Windows

To start the local translation server automatically whenever you sign in to Windows, run the following once from PowerShell:

```powershell
cd C:\Users\michaeljo1\Documents\Codex\2026-07-16\i-wou\outputs\outlook-hebrew-translator
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install-startup.ps1
```

This creates a user-level **Outlook Hebrew Translator** scheduled task and starts it immediately. It does not require administrator access. To remove it later, run `./scripts/uninstall-startup.ps1`. Outlook loads the service from `https://localhost:3000` on this Windows PC only.

## Privacy

The current draft's subject and HTML body are sent only to this app's `/api/translate` endpoint and then to the configured LLM provider. Do not use the add-in with content your organization does not permit you to process through that provider.
