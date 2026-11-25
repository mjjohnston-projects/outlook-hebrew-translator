import "dotenv/config";
import { fileURLToPath } from "node:url";
import express from "express";
import OpenAI from "openai";
import sanitizeHtml from "sanitize-html";

export const app = express();
app.use((req, res, next) => {
  if (!req.path.startsWith("/api/")) return next();
  const started = Date.now();
  res.on("finish", () => console.log(`${req.method} ${req.path} ${res.statusCode} ${Date.now() - started}ms`));
  next();
});
app.use(express.json({ limit: "2mb" }));
app.use(express.static(fileURLToPath(new URL("./public", import.meta.url))));

// One retry absorbs a transient network timeout without requiring the Outlook user to repeat the action.
const createLlmProvider = () => {
  const selectedProvider = (process.env.LLM_PROVIDER || "openai").trim().toLowerCase();

  if (selectedProvider === "github-models") {
    const apiKey = process.env.GITHUB_MODELS_TOKEN;
    const model = process.env.GITHUB_MODELS_MODEL;
    if (!apiKey || !model) {
      return { error: "The server is missing GITHUB_MODELS_TOKEN or GITHUB_MODELS_MODEL." };
    }

    return {
      name: "GitHub Models",
      apiKeyName: "GITHUB_MODELS_TOKEN",
      model,
      client: new OpenAI({
        apiKey,
        baseURL: "https://models.github.ai/inference",
        defaultHeaders: {
          Accept: "application/vnd.github+json",
          "X-GitHub-Api-Version": "2026-03-10"
        },
        timeout: 60000,
        maxRetries: 1
      })
    };
  }

  if (selectedProvider === "openai") {
    const apiKey = process.env.OPENAI_API_KEY;
    const model = process.env.OPENAI_MODEL;
    if (!apiKey || !model) {
      return { error: "The server is missing OPENAI_API_KEY or OPENAI_MODEL." };
    }

    return {
      name: "OpenAI",
      apiKeyName: "OPENAI_API_KEY",
      model,
      client: new OpenAI({ apiKey, timeout: 60000, maxRetries: 1 })
    };
  }

  return { error: "LLM_PROVIDER must be either 'github-models' or 'openai'." };
};

const providerErrorMessage = (error, provider) => {
  const status = Number(error?.status) || 502;
  const providerMessage = typeof error?.message === "string" ? error.message : "Unknown provider error.";
  const label = provider?.name || "Translation provider";
  const keyName = provider?.apiKeyName || "provider API key";
  const modelName = provider?.name === "GitHub Models" ? "GITHUB_MODELS_MODEL" : "OPENAI_MODEL";
  const guidance = status === 401
    ? `${label} authentication failed. Check ${keyName} in .env.`
    : status === 404
      ? `The configured ${modelName} is unavailable to ${label}.`
      : status === 429
        ? `${label} has reached a rate or usage limit. Try again later, or review its billing settings.`
        : `Translation provider error: ${providerMessage}`;
  return { status: status >= 400 && status < 600 ? status : 502, guidance };
};

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

// The task pane replaces each image with a [[IMG-n]] token before sending, and restores it afterwards.
const imagePlaceholderInstruction = "Tokens of the form [[IMG-1]], [[IMG-2]] and so on stand for images: copy each one into the output exactly once, unchanged, in the same position.";

const directionAttribute = ["dir"];
export const sanitizeEmailHtml = (html) => sanitizeHtml(html, {
  allowedTags: ["a", "b", "strong", "i", "em", "u", "s", "strike", "sub", "sup", "font", "span", "p", "br", "div", "ul", "ol", "li", "table", "thead", "tbody", "tr", "td", "th", "blockquote", "hr"],
  allowedAttributes: {
    "a": ["href", "title", ...directionAttribute],
    "font": ["face", "size", "color"],
    "span": ["style", ...directionAttribute],
    "p": ["style", ...directionAttribute],
    "div": ["style", ...directionAttribute],
    "ul": directionAttribute, "ol": directionAttribute, "li": directionAttribute,
    "table": directionAttribute, "blockquote": directionAttribute,
    "td": ["style", "colspan", "rowspan", ...directionAttribute],
    "th": ["style", "colspan", "rowspan", ...directionAttribute]
  },
  allowedSchemes: ["http", "https", "mailto"],
  allowedStyles: {
    "*": {
      "text-align": [/^left$/, /^right$/, /^center$/],
      "direction": [/^ltr$/, /^rtl$/],
      "font-family": [/^[\w\s,.'"-]+$/],
      "font-size": [/^\d+(?:\.\d+)?(?:px|pt|em|rem|%)$/],
      "font-weight": [/^(?:normal|bold|[1-9]00)$/],
      "font-style": [/^(?:normal|italic|oblique)$/],
      "text-decoration": [/^(?:none|underline|line-through)$/],
      "color": [/^(?:#[0-9a-fA-F]{3,8}|rgb\([\d\s,]+\)|[a-zA-Z]+)$/],
      "background-color": [/^(?:#[0-9a-fA-F]{3,8}|rgb\([\d\s,]+\)|[a-zA-Z]+)$/]
    }
  },
  // Outlook writes dir="RTL"; normalise the case so the attribute is kept.
  transformTags: {
    "*": (tagName, attribs) => {
      if (attribs.dir) attribs.dir = attribs.dir.toLowerCase();
      if (attribs.dir && !["rtl", "ltr", "auto"].includes(attribs.dir)) delete attribs.dir;
      return { tagName, attribs };
    }
  }
});

const targetLanguages = { "en-he": "Hebrew", "en-ru": "Russian", "he-en": "English" };
const recipientDescriptions = {
  man: ["one man", "masculine singular"],
  woman: ["one woman", "feminine singular"],
  men: ["multiple men", "masculine plural"],
  women: ["multiple women", "feminine plural"]
};
export const recipientInstruction = (recipientGender, target) => {
  const [who, form] = recipientDescriptions[recipientGender];
  const noun = recipientGender === "man" || recipientGender === "woman" ? "recipient" : "recipients";
  return `The email addresses ${who}. Use ${form} ${target} forms whenever addressing the ${noun}.`;
};

app.post("/api/translate", async (req, res) => {
  const { subject, bodyHtml, recipientGender, direction } = req.body || {};
  if (!subject && !bodyHtml) return res.status(400).json({ error: "Write an email before translating it." });
  if (!Object.hasOwn(targetLanguages, direction)) return res.status(400).json({ error: "Choose a supported translation direction." });
  if (direction !== "he-en" && !Object.hasOwn(recipientDescriptions, recipientGender)) {
    return res.status(400).json({ error: "Select who is being addressed." });
  }
  const provider = createLlmProvider();
  if (provider.error) return res.status(500).json({ error: provider.error });
  const target = targetLanguages[direction];
  const source = direction === "he-en" ? "Hebrew" : "English";
  const genderInstruction = direction === "he-en" ? "No recipient gender selection is needed for an English translation." : recipientInstruction(recipientGender, target);
  // Outlook can add a large amount of Word/Office markup. Remove that noise before calling the model,
  // while retaining the email content and common visual formatting that the add-in supports.
  const cleanedBodyHtml = sanitizeEmailHtml(bodyHtml || "");
  console.log(`Translation input: ${String(bodyHtml || "").length} HTML chars, reduced to ${cleanedBodyHtml.length} chars.`);

  try {
    const completion = await provider.client.chat.completions.create({
      model: provider.model,
      response_format: { type: "json_schema", json_schema: responseSchema },
      messages: [
        {
          role: "system",
          content: `You are an expert business email translator. Translate from ${source} to ${target}. ${genderInstruction}
Return only JSON that satisfies the schema. englishSubject must be a concise, appropriate English subject based on the email's intent. If translating Hebrew to English, it is the translated English subject; otherwise preserve the existing English subject when it is appropriate. translatedSubject is the subject in ${target}.
Translate the email body into ${target}. Preserve the input HTML structure exactly where possible: keep every hyperlink href unchanged, keep link text translated, keep all paragraph, line-break, list, and table boundaries so the translated message has the same visual breaks. Retain each element's font family, font size, color, bold/italic/underline styling, and text alignment. Do not invent, remove, or move URLs. ${imagePlaceholderInstruction} Keep email addresses, numbers, dates, file names, identifiers, and code unchanged unless target-language punctuation needs surrounding spacing. When translating English to Hebrew, if the sign-off is exactly or substantively 'Regards, Michael', translate it exactly as 'בברכה, מייקל'. Do not add a gender label, headings, Markdown, or commentary to translatedBodyHtml.`
        },
        {
          role: "user",
          content: JSON.stringify({ subject: subject || "", bodyHtml: cleanedBodyHtml })
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
    const { status, guidance } = providerErrorMessage(error, provider);
    res.status(status).json({ error: guidance });
  }
});

app.post("/api/proofread", async (req, res) => {
  const { subject, bodyHtml, language } = req.body || {};
  if (!subject && !bodyHtml) return res.status(400).json({ error: "Write an email before proofreading it." });
  if (!["English", "Hebrew", "Russian"].includes(language)) return res.status(400).json({ error: "Choose English, Hebrew, or Russian." });
  const provider = createLlmProvider();
  if (provider.error) return res.status(500).json({ error: provider.error });
  const cleanedBodyHtml = sanitizeEmailHtml(bodyHtml || "");
  console.log(`Proofreading input: ${String(bodyHtml || "").length} HTML chars, reduced to ${cleanedBodyHtml.length} chars.`);

  try {
    const completion = await provider.client.chat.completions.create({
      model: provider.model,
      response_format: { type: "json_schema", json_schema: proofreadSchema },
      messages: [
        { role: "system", content: `You are a meticulous ${language} business-email proofreader. Return only JSON that satisfies the schema. Correct only spelling, grammar, punctuation, capitalization, and obvious typographical mistakes in the subject and body. Do not translate, change the email's meaning or tone, rewrite sentences for style, add content, remove content, or add commentary. Preserve the input HTML structure exactly where possible: keep every hyperlink href, all font family, font size, color, bold/italic/underline styling, text direction, and all paragraph, line-break, list, and table boundaries. ${imagePlaceholderInstruction}` },
        { role: "user", content: JSON.stringify({ subject: subject || "", bodyHtml: cleanedBodyHtml }) }
      ]
    });
    const content = completion.choices[0]?.message?.content;
    if (!content) throw new Error("The proofreading service returned no content.");
    const corrected = JSON.parse(content);
    corrected.bodyHtml = sanitizeEmailHtml(corrected.bodyHtml);
    res.json(corrected);
  } catch (error) {
    console.error(error);
    const { status, guidance } = providerErrorMessage(error, provider);
    res.status(status).json({ error: guidance.replace("Translation provider error", "Proofreading provider error") });
  }
});

// Express's default error page is HTML, which the task pane cannot show. Answer API errors in JSON.
app.use((error, req, res, next) => {
  if (res.headersSent) return next(error);
  console.error(error);
  if (error.type === "entity.too.large") {
    return res.status(413).json({ error: "This email is too large to process (over 2 MB of text). Shorten it, or translate it in parts." });
  }
  if (error.type === "entity.parse.failed") {
    return res.status(400).json({ error: "The add-in sent a request the server could not read. Reload the add-in and try again." });
  }
  res.status(error.status || 500).json({ error: "The Email Language Assistant server hit an unexpected error." });
});
