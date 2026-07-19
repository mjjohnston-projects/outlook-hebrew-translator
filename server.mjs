import "dotenv/config";
import fs from "node:fs";
import https from "node:https";
import express from "express";
import OpenAI from "openai";
import sanitizeHtml from "sanitize-html";
import devCerts from "office-addin-dev-certs";

const app = express();
const port = Number(process.env.PORT || 3000);
app.use(express.json({ limit: "2mb" }));
app.use(express.static("public"));
const createOpenAiClient = () => new OpenAI({ apiKey: process.env.OPENAI_API_KEY, timeout: 85000, maxRetries: 0 });

const responseSchema = {
  name: "outlook_hebrew_translation",
  strict: true,
  schema: {
    type: "object",
    additionalProperties: false,
    required: ["englishSubject", "translatedSubject", "translatedBodyHtml"],
    properties: {
      englishSubject: { type: "string" },
      translatedSubject: { type: "string" },
      translatedBodyHtml: { type: "string" }
    }
  }
};

const proofreadSchema = {
  name: "outlook_proofreading",
  strict: true,
  schema: {
    type: "object",
    additionalProperties: false,
    required: ["subject", "bodyHtml"],
    properties: {
      subject: { type: "string" },
      bodyHtml: { type: "string" }
    }
  }
};

const sanitizeEmailHtml = (html) => sanitizeHtml(html, {
  allowedTags: ["a", "b", "strong", "i", "em", "u", "s", "span", "p", "br", "div", "ul", "ol", "li", "table", "thead", "tbody", "tr", "td", "th", "blockquote", "hr"],
  allowedAttributes: { "a": ["href", "title"], "span": ["style"], "p": ["style"], "div": ["style"], "td": ["style", "colspan", "rowspan"], "th": ["style", "colspan", "rowspan"] },
  allowedSchemes: ["http", "https", "mailto"],
  allowedStyles: { "*": { "text-align": [/^left$/, /^right$/, /^center$/], "direction": [/^ltr$/, /^rtl$/] } }
});

app.post("/api/translate", async (req, res) => {
  const { subject, bodyHtml, recipientGender, direction } = req.body || {};
  if (!subject && !bodyHtml) return res.status(400).json({ error: "Write an email before translating it." });
  if (!["en-he", "en-ru", "he-en"].includes(direction)) return res.status(400).json({ error: "Choose a supported translation direction." });
  if (direction !== "he-en" && !["man", "woman", "men", "women"].includes(recipientGender)) {
    return res.status(400).json({ error: "Select who is being addressed." });
  }
  if (!process.env.OPENAI_API_KEY || !process.env.OPENAI_MODEL) {
    return res.status(500).json({ error: "The server is missing OPENAI_API_KEY or OPENAI_MODEL." });
  }

  const client = createOpenAiClient();
  const recipientInstructions = {
    man: "The email addresses one man. Use masculine singular Hebrew forms whenever addressing the recipient.",
    woman: "The email addresses one woman. Use feminine singular Hebrew forms whenever addressing the recipient.",
    men: "The email addresses multiple men. Use masculine plural Hebrew forms whenever addressing the recipients.",
    women: "The email addresses multiple women. Use feminine plural Hebrew forms whenever addressing the recipients."
  };
  const target = direction === "en-he" ? "Hebrew" : direction === "en-ru" ? "Russian" : "English";
  const source = direction === "he-en" ? "Hebrew" : "English";
  const genderInstruction = direction === "he-en" ? "No recipient gender selection is needed for an English translation." : recipientInstructions[recipientGender];

  try {
    const completion = await client.chat.completions.create({
      model: process.env.OPENAI_MODEL,
      response_format: { type: "json_schema", json_schema: responseSchema },
      messages: [
        {
          role: "system",
          content: `You are an expert business email translator. Translate from ${source} to ${target}. ${genderInstruction}
Return only JSON that satisfies the schema. englishSubject must be a concise, appropriate English subject based on the email's intent. If translating Hebrew to English, it is the translated English subject; otherwise preserve the existing English subject when it is appropriate. translatedSubject is the subject in ${target}.
Translate the email body into ${target}. Preserve the input HTML structure exactly where possible: keep every hyperlink href unchanged, keep link text translated, keep all paragraph, line-break, list, and table boundaries so the translated message has the same visual breaks. Do not invent, remove, or move URLs. Keep email addresses, numbers, dates, file names, identifiers, and code unchanged unless target-language punctuation needs surrounding spacing. When translating English to Hebrew, if the sign-off is exactly or substantively 'Regards, Michael', translate it exactly as 'בברכה, מייקל'. Do not add a gender label, headings, Markdown, or commentary to translatedBodyHtml.`
        },
        {
          role: "user",
          content: JSON.stringify({ subject: subject || "", bodyHtml: bodyHtml || "" })
        }
      ]
    });
    const content = completion.choices[0]?.message?.content;
    if (!content) throw new Error("The translation service returned no content.");
    const translated = JSON.parse(content);
    // The model is asked to retain the email's HTML; still, never render active content in the task pane.
    translated.translatedBodyHtml = sanitizeEmailHtml(translated.translatedBodyHtml);
    res.json(translated);
  } catch (error) {
    console.error(error);
    const status = Number(error?.status) || 502;
    const providerMessage = typeof error?.message === "string" ? error.message : "Unknown provider error.";
    const guidance = status === 401
      ? "OpenAI authentication failed. Check OPENAI_API_KEY in .env."
      : status === 404
        ? "The configured OPENAI_MODEL is unavailable to this API project."
        : status === 429
          ? "The OpenAI project has reached a usage limit or needs billing enabled."
          : `Translation provider error: ${providerMessage}`;
    res.status(status >= 400 && status < 600 ? status : 502).json({ error: guidance });
  }
});

app.post("/api/proofread", async (req, res) => {
  const { subject, bodyHtml, language } = req.body || {};
  if (!subject && !bodyHtml) return res.status(400).json({ error: "Write an email before proofreading it." });
  if (!["English", "Hebrew", "Russian"].includes(language)) return res.status(400).json({ error: "Choose English, Hebrew, or Russian." });
  if (!process.env.OPENAI_API_KEY || !process.env.OPENAI_MODEL) return res.status(500).json({ error: "The server is missing OPENAI_API_KEY or OPENAI_MODEL." });

  try {
    const client = createOpenAiClient();
    const completion = await client.chat.completions.create({
      model: process.env.OPENAI_MODEL,
      response_format: { type: "json_schema", json_schema: proofreadSchema },
      messages: [
        { role: "system", content: `You are a meticulous ${language} business-email proofreader. Return only JSON that satisfies the schema. Correct only spelling, grammar, punctuation, capitalization, and obvious typographical mistakes in the subject and body. Do not translate, change the email's meaning or tone, rewrite sentences for style, add content, remove content, or add commentary. Preserve the input HTML structure exactly where possible: keep every hyperlink href unchanged and preserve all paragraph, line-break, list, and table boundaries.` },
        { role: "user", content: JSON.stringify({ subject: subject || "", bodyHtml: bodyHtml || "" }) }
      ]
    });
    const content = completion.choices[0]?.message?.content;
    if (!content) throw new Error("The proofreading service returned no content.");
    const corrected = JSON.parse(content);
    corrected.bodyHtml = sanitizeEmailHtml(corrected.bodyHtml);
    res.json(corrected);
  } catch (error) {
    console.error(error);
    const status = Number(error?.status) || 502;
    const providerMessage = typeof error?.message === "string" ? error.message : "Unknown provider error.";
    const guidance = status === 401 ? "OpenAI authentication failed. Check OPENAI_API_KEY in .env."
      : status === 404 ? "The configured OPENAI_MODEL is unavailable to this API project."
      : status === 429 ? "The OpenAI project has reached a usage limit or needs billing enabled."
      : `Proofreading provider error: ${providerMessage}`;
    res.status(status >= 400 && status < 600 ? status : 502).json({ error: guidance });
  }
});

const httpsOptions = process.env.TLS_CERT_PATH && process.env.TLS_KEY_PATH
  ? { cert: fs.readFileSync(process.env.TLS_CERT_PATH), key: fs.readFileSync(process.env.TLS_KEY_PATH) }
  : await devCerts.getHttpsServerOptions();
https.createServer(httpsOptions, app).listen(port, () => {
  console.log(`Outlook Email Language Assistant listening on https://localhost:${port}`);
});
