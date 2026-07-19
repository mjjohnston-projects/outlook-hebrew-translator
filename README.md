# Outlook Email Language Assistant

An Outlook compose add-in that translates English drafts into Hebrew or Russian, and Hebrew drafts into English. It asks for the recipient form when translating into Hebrew or Russian so gendered wording is correct, suggests an English subject, supplies a copy-ready translated subject, and preserves line breaks and links in the body.

## What it does

1. Open a new email or reply in Outlook and select **Email Language Assistant**.
2. Choose **English → Hebrew**, **English → Russian**, or **Hebrew → English**.
3. When translating to Hebrew or Russian, choose **Man**, **Woman**, **Men**, or **Women** for the person or people being addressed. The target-language agreement follows the selection.
4. Select **Translate draft**.
4. Review the result. Each response starts with the selected gender, then shows the English subject suggestion, Hebrew subject, and the Hebrew body.
5. Use **Apply Hebrew subject & body** to replace the draft, or copy either field.

`Regards, Michael` is explicitly rendered as `בברכה, מייקל`.

The add-in can also proofread an English, Hebrew, or Russian draft. Choose the proofread language, then select **Check spelling & grammar**. It corrects spelling, grammar, punctuation, capitalization, and clear typos without translating or rewriting the email.

## Run locally

Prerequisites: Node.js 20+ and an OpenAI API key. The default model is `gpt-4.1-mini`, selected for fast, capable translation and proofreading.

```powershell
cd C:\Users\michaeljo1\Documents\Codex\2026-07-16\i-wou\outputs\outlook-email-language-assistant
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
cd C:\Users\michaeljo1\Documents\Codex\2026-07-16\i-wou\outputs\outlook-email-language-assistant
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install-startup.ps1
```

This creates a user-level **Outlook Email Language Assistant** background service and starts it immediately. Task Scheduler restarts it up to three times if it exits unexpectedly. Outlook loads the service from `https://localhost:3000` on this Windows PC only.

Restart the service at any time with:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\restart-service.ps1
```

To remove it later, run `./scripts/uninstall-startup.ps1`.

## Native Windows service

For an entry in `services.msc`, build and install the included native Windows service wrapper. This replaces the scheduled-task service and requires administrator confirmation.

```powershell
dotnet publish .\service\EmailLanguageAssistantService.csproj -c Release -r win-x64 -o .\service\publish
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install-native-service.ps1
```

After installation, open `services.msc` and manage **Outlook Email Language Assistant** like any other Windows service. To uninstall it, run `./scripts/uninstall-native-service.ps1` from an elevated PowerShell window.

## Privacy

The current draft's subject and HTML body are sent only to this app's `/api/translate` endpoint and then to the configured LLM provider. Do not use the add-in with content your organization does not permit you to process through that provider.
