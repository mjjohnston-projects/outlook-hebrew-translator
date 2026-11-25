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

The add-in can also proofread an English, Hebrew, or Russian draft. Choose the proofread language, then select **Check spelling & grammar**. It corrects spelling, grammar, punctuation, capitalization, and clear typos without translating or rewriting the email. Like translation, it works only on your own text: the signature and quoted reply or forward history are left unchanged and are not sent to the AI provider.

If Hebrew was typed while the keyboard was accidentally set to English, select **Fix Hebrew keyboard layout**. This locally converts the editable subject and body using the standard English-to-Hebrew keyboard mapping; it does not use the AI provider and leaves links, signatures, and quoted email history unchanged. Only words that look like mistyped Hebrew are converted: correctly typed Hebrew, capitalised words (names, acronyms), numbers, file names, email addresses, and web addresses stay as they are. Short words such as `do` or `to` are converted only when the words next to them are mistyped Hebrew too. Review the result before applying it.

## Run locally

Prerequisites: Node.js 20+ and a GitHub personal access token with **Models: Read** permission. The default provider is GitHub Models with `openai/gpt-4.1-mini`, selected for fast, capable translation and proofreading. GitHub Models has a free, rate-limited tier; paid usage is optional.

```powershell
# Run this from the folder that contains package.json.
$HOME_DIR = (Resolve-Path .).Path
Set-Location $HOME_DIR
Copy-Item .env.example .env
# Edit .env and set GITHUB_MODELS_TOKEN. Do not share this value.
npm install
npm run dev
```

To create the token, open GitHub **Settings** → **Developer settings** → **Personal access tokens** → **Fine-grained tokens**, create a token with **Models: Read**, then paste it only into `GITHUB_MODELS_TOKEN` in your local `.env` file. The sample configuration uses `openai/gpt-4.1-mini`, which is in GitHub Models' low free-rate-limit tier.

To revert to your existing OpenAI API account later, set `LLM_PROVIDER=openai` and configure `OPENAI_API_KEY` and `OPENAI_MODEL` in `.env`.

The add-in must be hosted over HTTPS for Outlook. For this personal Windows setup, use `https://localhost:3000` with a trusted development certificate, then sideload `manifest.xml` in Outlook.

To avoid missing a URL in the manifest, configure it once with:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\set-manifest-url.ps1 -Origin https://localhost:3000 -ReplaceExisting
```

## Tests

```powershell
npm test
```

The tests run the task pane in a simulated Outlook draft and the server against a fake AI provider on this PC, so they need no token and send nothing externally.

`sanitize-html` is pinned to 2.17.5 because later versions require Node.js 22.12 or newer. That version has a known moderate advisory, but the add-in's sanitiser settings do not allow the tags it relies on (`svg`, `textarea`, `noscript`, `meta`). After upgrading Node.js to 22.12 or newer, run `npm install sanitize-html@latest` and `npm test` to take the fix.

## Deploy

Deploy this app to any HTTPS-capable Node host. Set `LLM_PROVIDER`, `GITHUB_MODELS_TOKEN`, and `GITHUB_MODELS_MODEL` as server secrets (or use the documented OpenAI fallback), set `HOST=0.0.0.0` so the server accepts connections from other machines (by default it listens only on this PC, `127.0.0.1`, so nobody else on the network can use your API key) and put it behind your host's access control, replace `YOUR-HTTPS-DOMAIN` in the manifest, then distribute the manifest through Microsoft 365 admin integrated apps or sideload it for individual users.

## Start with Windows

To start the local translation server automatically whenever you sign in to Windows, run the following once from PowerShell:

```powershell
# Run this from the folder that contains package.json.
$HOME_DIR = (Resolve-Path .).Path
Set-Location $HOME_DIR
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

The current draft's subject and the HTML of your own text (not the signature, quoted history, or images) are sent only to this app's `/api/translate` or `/api/proofread` endpoint and then to the configured LLM provider. Images are replaced by `[[IMG-1]]`-style placeholders before sending and restored afterwards. To make re-clicking Translate instant, the server keeps up to 50 recent translations in memory for ten minutes; nothing is written to disk, and restarting the service clears them. Do not use the add-in with content your organization does not permit you to process through that provider.
