import "dotenv/config";
import https from "node:https";
import express from "express";
import OpenAI from "openai";
import sanitizeHtml from "sanitize-html";
import devCerts from "office-addin-dev-certs";

const app = express();
const port = Number(process.env.PORT || 3000);
app.use(express.json({ limit: "2mb" }));
app.use(express.static("public"));

const responseSchema = {
  name: "outlook_hebrew_translation",
  strict: true,
  schema: {
    type: "object",
    additionalProperties: false,
    required: ["englishSubject", "hebrewSubject", "hebrewBodyHtml"],
    properties: {
      englishSubject: { type: "string" },
      hebrewSubject: { type: "string" },
      hebrewBodyHtml: { type: "string" }
    }
  }
};

app.post("/api/translate", async (req, res) => {
  const { subject, bodyHtml, recipientGender } = req.body || {};
  if (!subject && !bodyHtml) return res.status(400).json({ error: "Write an email before translating it." });
  if (!["man", "woman", "men", "women"].includes(recipientGender)) {
    return res.status(400).json({ error: "Select who is being addressed." });
  }
  if (!process.env.OPENAI_API_KEY || !process.env.OPENAI_MODEL) {
    return res.status(500).json({ error: "The server is missing OPENAI_API_KEY or OPENAI_MODEL." });
  }

  const client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
  const recipientInstructions = {
    man: "The email addresses one man. Use masculine singular Hebrew forms whenever addressing the recipient.",
    woman: "The email addresses one woman. Use feminine singular Hebrew forms whenever addressing the recipient.",
    men: "The email addresses multiple men. Use masculine plural Hebrew forms whenever addressing the recipients.",
    women: "The email addresses multiple women. Use feminine plural Hebrew forms whenever addressing the recipients."
  };
  const genderInstruction = recipientInstructions[recipientGender];

  try {
    const completion = await client.chat.completions.create({
      model: process.env.OPENAI_MODEL,
      response_format: { type: "json_schema", json_schema: responseSchema },
      messages: [
        {
          role: "system",
          content: `You are an expert English-to-Hebrew business email translator. ${genderInstruction}
Return only JSON that satisfies the schema. Create a concise, appropriate English subject based on the email's intent; preserve the existing subject when it is already appropriate. Translate it into natural Hebrew.
Translate the email body into Hebrew. Preserve the input HTML structure exactly where possible: keep every hyperlink href unchanged, keep link text translated, keep all paragraph, line-break, list, and table boundaries so the translated message has the same visual breaks. Do not invent, remove, or move URLs. Keep email addresses, numbers, dates, file names, identifiers, and code unchanged unless Hebrew punctuation needs surrounding spacing. If the sign-off is exactly or substantively 'Regards, Michael', translate it exactly as 'בברכה, מייקל'. Do not add a gender label, headings, Markdown, or commentary to hebrewBodyHtml.`
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
    translated.hebrewBodyHtml = sanitizeHtml(translated.hebrewBodyHtml, {
      allowedTags: ["a", "b", "strong", "i", "em", "u", "s", "span", "p", "br", "div", "ul", "ol", "li", "table", "thead", "tbody", "tr", "td", "th", "blockquote", "hr"],
      allowedAttributes: { "a": ["href", "title"], "span": ["style"], "p": ["style"], "div": ["style"], "td": ["style", "colspan", "rowspan"], "th": ["style", "colspan", "rowspan"] },
      allowedSchemes: ["http", "https", "mailto"],
      allowedStyles: { "*": { "text-align": [/^left$/, /^right$/, /^center$/], "direction": [/^ltr$/, /^rtl$/] } }
    });
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

const httpsOptions = await devCerts.getHttpsServerOptions();
https.createServer(httpsOptions, app).listen(port, () => {
  console.log(`Outlook Hebrew Translator listening on https://localhost:${port}`);
});
