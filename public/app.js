let translated = null;
let activeRequest = null;
const $ = (selector) => document.querySelector(selector);

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    showStatus("Open this add-in from an Outlook email draft.");
    return;
  }
  $("#translate").addEventListener("click", translateDraft);
  $("#keyboard-hebrew").addEventListener("click", fixHebrewKeyboardLayout);
  $("#proofread").addEventListener("click", proofreadDraft);
  $("#direction").addEventListener("change", updateDirectionUi);
  $("#copy-subject").addEventListener("click", () => copyText($("#translated-subject").value));
  $("#copy-body").addEventListener("click", copyBody);
  $("#apply").addEventListener("click", applyTranslation);
});

function showStatus(message = "") { $("#status").textContent = message; }
// While any request runs, every action is disabled, including Apply, so a stale result cannot be applied.
function setRequestState(busy) {
  ["#translate", "#keyboard-hebrew", "#proofread", "#apply"].forEach((selector) => { $(selector).disabled = busy; });
}

async function postJson(url, payload, actionName) {
  const controller = new AbortController();
  activeRequest = controller;
  const timeout = window.setTimeout(() => controller.abort(), 130000);
  try {
    const response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      signal: controller.signal
    });
    const text = await response.text();
    let data;
    try { data = text ? JSON.parse(text) : {}; }
    catch {
      if (response.status === 413) throw new Error("This email is too large to process. Shorten it, or translate it in parts.");
      throw new Error(`${actionName} returned an unexpected response from the local server (HTTP ${response.status}).`);
    }
    if (!response.ok) throw new Error(data.error || `${actionName} failed.`);
    return data;
  } catch (error) {
    if (error.name === "AbortError") throw new Error(`${actionName} took longer than two minutes. Check that the Email Language Assistant service is running, then try again.`);
    throw error;
  } finally {
    window.clearTimeout(timeout);
    if (activeRequest === controller) activeRequest = null;
  }
}
function getGender() { return document.querySelector('input[name="gender"]:checked')?.value; }
const recipientLabels = { man: "Man", woman: "Woman", men: "Men", women: "Women" };
const directions = {
  "en-he": { label: "Hebrew", language: "he", dir: "rtl", needsRecipientForm: true },
  "en-ru": { label: "Russian", language: "ru", dir: "ltr", needsRecipientForm: true },
  "he-en": { label: "English", language: "en", dir: "ltr", needsRecipientForm: false }
};

function currentDirection() { return directions[$("#direction").value]; }
function updateDirectionUi() {
  const direction = currentDirection();
  $("#recipient-form").hidden = !direction.needsRecipientForm;
  if (!direction.needsRecipientForm) document.querySelectorAll('input[name="gender"]').forEach((input) => { input.checked = false; });
}

function getSubject() {
  return new Promise((resolve, reject) => Office.context.mailbox.item.subject.getAsync((result) => {
    result.status === Office.AsyncResultStatus.Succeeded ? resolve(result.value) : reject(result.error);
  }));
}
function getBody() {
  return new Promise((resolve, reject) => Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
    result.status === Office.AsyncResultStatus.Succeeded ? resolve(result.value) : reject(result.error);
  }));
}

// Outlook returns a full HTML document. Keep everything outside <body> (including the <head> styles)
// exactly as Outlook wrote it, and only work on what is inside <body>.
function splitDocument(html) {
  const open = /<body\b[^>]*>/i.exec(html);
  if (!open) return { documentStart: "", bodyInnerHtml: html, documentEnd: "" };
  const start = open.index + open[0].length;
  const closes = [...html.matchAll(/<\/body\s*>/gi)];
  const end = closes.length && closes.at(-1).index >= start ? closes.at(-1).index : html.length;
  return { documentStart: html.slice(0, start), bodyInnerHtml: html.slice(start, end), documentEnd: html.slice(end) };
}

function splitPreservedTail(bodyHtml) {
  // Preserve Outlook's signature and quoted reply/forward history exactly as Outlook created them.
  // The first matching marker begins the unchanged tail of the draft.
  const markers = [
    // Signatures, reply/forward headers and quoted history from Outlook and other mail clients.
    /<(?:div|p|table|span|td)\b[^>]*(?:\bid|\bclass|\bdata-[\w-]+)\s*=\s*["'][^"']*(?:signature|divRplyFwdMsg|reply[-_ ]?header|quoted|gmail_quote|yahoo_quoted|protonmail_quote)[^"']*["'][^>]*>/i,
    /<div\b[^>]*\bid\s*=\s*["']appendonsend["'][^>]*>/i,
    // A plain <blockquote> the user wrote is still translated; only cited ones are history.
    /<blockquote\b[^>]*\btype\s*=\s*["']?cite\b/i,
    /<!--\s*(?:original|forwarded)\s+message[^>]*-->/i,
    /<(?:div|p|table|span|td)\b[^>]*>\s*(?:<br\s*\/?>(?:\s*))*?(?:<(?:b|strong|span)\b[^>]*>\s*)*(?:-{2,}\s*(?:original|forwarded)\s+message\s*-{2,}|(?:From|Sent|To|Subject)\s*:)/i,
    /<hr\b[^>]*>[\s\S]{0,1200}?(?:<\/?\w+[^>]*>\s*)*(?:From|Sent|To|Subject)\s*:/i
  ];
  const starts = markers
    .map((marker) => marker.exec(bodyHtml)?.index)
    .filter((index) => index !== undefined);
  const tailStart = starts.length ? Math.min(...starts) : undefined;
  if (tailStart === undefined) return { translatableHtml: bodyHtml, preservedTailHtml: "" };
  return {
    translatableHtml: bodyHtml.slice(0, tailStart),
    preservedTailHtml: bodyHtml.slice(tailStart)
  };
}

// Reads the draft and separates the part the user is editing from everything that must stay unchanged.
async function readDraft() {
  const [subject, bodyHtml] = await Promise.all([getSubject(), getBody()]);
  const { documentStart, bodyInnerHtml, documentEnd } = splitDocument(bodyHtml || "");
  const { translatableHtml, preservedTailHtml } = splitPreservedTail(bodyInnerHtml);
  return {
    subject: subject || "",
    editableHtml: translatableHtml,
    rebuild: (newEditableHtml) => documentStart + newEditableHtml + preservedTailHtml + documentEnd
  };
}

// Images are not sent to the AI provider: they can be megabytes of data and the sanitiser would drop them.
// Each one is swapped for a [[IMG-n]] token and put back afterwards.
function protectImages(html) {
  const images = [];
  const protectedHtml = html.replace(/<img\b[^>]*>/gi, (tag) => {
    images.push(tag);
    return `[[IMG-${images.length}]]`;
  });
  return { protectedHtml, images };
}
function restoreImages(html, images) {
  let restored = html;
  const missing = [];
  images.forEach((tag, index) => {
    const token = `[[IMG-${index + 1}]]`;
    const first = restored.indexOf(token);
    if (first === -1) { missing.push(tag); return; }
    restored = restored.slice(0, first) + tag + restored.slice(first + token.length).split(token).join("");
  });
  // If the model dropped a token, keep the image rather than lose it.
  return missing.length ? `${restored}<p>${missing.join(" ")}</p>` : restored;
}

// Standard Israeli Hebrew keyboard mapping. This repairs text typed while the
// user intended Hebrew but Windows/Outlook was set to the English layout.
const hebrewKeyboardMap = Object.freeze({
  q: "/", w: "'", e: "ק", r: "ר", t: "א", y: "ט", u: "ו", i: "ן", o: "ם", p: "פ", "[": "]", "]": "[",
  a: "ש", s: "ד", d: "ג", f: "כ", g: "ע", h: "י", j: "ח", k: "ל", l: "ך", ";": "ף", "'": ",",
  z: "ז", x: "ס", c: "ב", v: "ה", b: "נ", n: "מ", m: "צ", ",": "ת", ".": "ץ", "/": "."
});
const protectedEmailTokens = /(?:https?:\/\/[^\s<]+|www\.[^\s<]+|[\w.+-]+@[\w.-]+\.[A-Za-z]{2,})/gi;
const hebrewFinalLetters = "ךםןףץ";
const hebrewLetter = /[א-ת]/;
// Short English words whose keys also spell plausible Hebrew (for example "do" -> "גם").
// They are converted only when the neighbouring word is mistyped Hebrew too.
const ambiguousEnglishWords = new Set(["a", "am", "an", "and", "are", "as", "at", "be", "but", "by", "can", "do", "for", "go", "had", "has", "have", "he", "hi", "me", "my", "no", "not", "ok", "re", "see", "send", "she", "so", "that", "the", "this", "to", "up", "us", "was", "we", "who", "yes", "you"]);

function convertKeyboardSegment(value) {
  return Array.from(value, (character) => hebrewKeyboardMap[character.toLowerCase()] ?? character).join("");
}

// "english" = leave unchanged, "hebrew" = convert, "ambiguous" = follow the neighbouring words.
function classifyKeyboardWord(word) {
  if (!/[a-z]/i.test(word) || hebrewLetter.test(word)) return "english";
  // Capitals, digits and non-layout characters indicate names, acronyms, numbers and real English.
  if (/[A-Z0-9]/.test(word)) return "english";
  const letters = word.replace(/^[^a-z]+|[^a-z]+$/g, "");
  if (ambiguousEnglishWords.has(letters)) return "ambiguous";
  // Hebrew final letters (ך ם ן ף ץ) only end words, so a final form in the middle means English, e.g. "file" -> "כןךק".
  const converted = convertKeyboardSegment(word);
  for (let index = 0; index < converted.length - 1; index++) {
    if (hebrewFinalLetters.includes(converted[index]) && hebrewLetter.test(converted[index + 1])) return "english";
  }
  return "hebrew";
}

function convertKeyboardWords(value) {
  const parts = value.split(/(\s+)/);
  const words = parts.map((part, index) => (index % 2 === 0 && part ? { index, kind: classifyKeyboardWord(part) } : null)).filter(Boolean);
  words.forEach((word, position) => {
    if (word.kind !== "ambiguous") return;
    const neighbour = [words[position - 1], words[position + 1]].find((candidate) => candidate && candidate.kind !== "ambiguous");
    word.kind = neighbour?.kind === "hebrew" ? "hebrew" : "english";
  });
  words.forEach((word) => { if (word.kind === "hebrew") parts[word.index] = convertKeyboardSegment(parts[word.index]); });
  return parts.join("");
}

function convertKeyboardText(value) {
  let converted = "";
  let cursor = 0;
  for (const match of value.matchAll(protectedEmailTokens)) {
    converted += convertKeyboardWords(value.slice(cursor, match.index));
    converted += match[0];
    cursor = match.index + match[0].length;
  }
  return converted + convertKeyboardWords(value.slice(cursor));
}

function convertKeyboardHtml(bodyHtml) {
  const documentFragment = new DOMParser().parseFromString(bodyHtml, "text/html");
  const walker = documentFragment.createTreeWalker(documentFragment.body, NodeFilter.SHOW_TEXT);
  const textNodes = [];
  while (walker.nextNode()) textNodes.push(walker.currentNode);
  let changed = false;
  textNodes.forEach((textNode) => {
    // URLs and their displayed text stay unchanged, just like normal translation.
    if (textNode.parentElement?.closest("a")) return;
    const original = textNode.nodeValue || "";
    const converted = convertKeyboardText(original);
    if (converted !== original) { textNode.nodeValue = converted; changed = true; }
  });
  return { html: documentFragment.body.innerHTML, changed };
}

function alignTranslatedBody(bodyHtml, direction) {
  if (direction.dir !== "rtl") return bodyHtml;
  const documentFragment = new DOMParser().parseFromString(bodyHtml, "text/html");
  // Set the font on every translated element so Outlook does not retain the original
  // English font from a nested span or paragraph when the HTML is pasted/applied.
  documentFragment.body.querySelectorAll("*").forEach((element) => {
    element.style.fontFamily = "Arial, sans-serif";
    if (element.tagName === "FONT") element.setAttribute("face", "Arial");
  });
  documentFragment.body.querySelectorAll("p, div, li, td, th, blockquote, table").forEach((element) => {
    element.dir = "rtl";
    const alignment = element.style.textAlign.toLowerCase();
    if (!alignment || alignment === "left") element.style.textAlign = "right";
  });
  // The wrapper also aligns unwrapped text nodes, which are common in simple Outlook drafts.
  return `<div dir="rtl" style="text-align: right; font-family: Arial, sans-serif;">${documentFragment.body.innerHTML}</div>`;
}
function setSubject(value) {
  return new Promise((resolve, reject) => Office.context.mailbox.item.subject.setAsync(value, (result) => {
    result.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(result.error);
  }));
}
function setBody(value) {
  return new Promise((resolve, reject) => Office.context.mailbox.item.body.setAsync(value, { coercionType: Office.CoercionType.Html }, (result) => {
    result.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(result.error);
  }));
}

async function fixHebrewKeyboardLayout() {
  if (activeRequest) return showStatus("Another request is already in progress.");
  setRequestState(true);
  showStatus("Reading and fixing the Hebrew keyboard layout…");
  try {
    const draft = await readDraft();
    const correctedSubject = convertKeyboardText(draft.subject);
    const correctedBody = convertKeyboardHtml(draft.editableHtml);
    if (correctedSubject === draft.subject && !correctedBody.changed) {
      return showStatus("No English-keyboard Hebrew text was found in the editable part of this draft.");
    }
    const correctedEditableHtml = alignTranslatedBody(correctedBody.html, directions["en-he"]);
    translated = { translatedSubject: correctedSubject, replacementEditableHtml: correctedEditableHtml };
    $("#gender-label").textContent = "Hebrew keyboard layout corrected";
    $("#english-subject-label").textContent = "Original subject (typed with English keyboard)";
    $("#english-subject").value = draft.subject;
    $("#translated-subject-group").hidden = false;
    $("#translated-subject-label").textContent = "Hebrew subject — ready to copy";
    $("#translated-body-label").textContent = "Hebrew email body — ready to copy";
    $("#translated-subject").value = correctedSubject;
    $("#translated-subject").dir = "rtl"; $("#translated-subject").lang = "he";
    $("#translated-body").innerHTML = correctedEditableHtml;
    $("#translated-body").dir = "rtl"; $("#translated-body").lang = "he";
    $("#result").hidden = false;
    showStatus("");
  } catch (error) {
    showStatus(error.message || "Unable to fix the Hebrew keyboard layout.");
  } finally { setRequestState(false); }
}

async function translateDraft() {
  const recipientGender = getGender();
  const direction = currentDirection();
  if (direction.needsRecipientForm && !recipientGender) return showStatus("Choose who is being addressed first.");
  if (activeRequest) return showStatus("Another request is already in progress.");
  setRequestState(true);
  showStatus("Reading and translating your draft…");
  try {
    const draft = await readDraft();
    const { protectedHtml, images } = protectImages(draft.editableHtml);
    const data = await postJson("/api/translate", { subject: draft.subject, bodyHtml: protectedHtml, recipientGender, direction: $("#direction").value }, "Translation");
    const translatedEditableHtml = alignTranslatedBody(restoreImages(data.translatedBodyHtml, images), direction);
    translated = { translatedSubject: data.translatedSubject, replacementEditableHtml: translatedEditableHtml };
    $("#english-subject-label").textContent = "Suggested subject (English)";
    $("#translated-subject-group").hidden = false;
    $("#gender-label").textContent = direction.needsRecipientForm ? `Recipient: ${recipientLabels[recipientGender]}` : "Translation: Hebrew to English";
    $("#english-subject").value = data.englishSubject;
    $("#translated-subject-label").textContent = `${direction.label} subject — ready to copy`;
    $("#translated-body-label").textContent = `${direction.label} email body — ready to copy`;
    $("#translated-subject").value = data.translatedSubject;
    $("#translated-subject").dir = direction.dir; $("#translated-subject").lang = direction.language;
    $("#translated-body").innerHTML = translatedEditableHtml;
    $("#translated-body").dir = direction.dir; $("#translated-body").lang = direction.language;
    $("#result").hidden = false;
    showStatus("");
  } catch (error) {
    showStatus(error.message || "Unable to translate this draft.");
  } finally { setRequestState(false); }
}

async function proofreadDraft() {
  const language = $("#proofread-language").value;
  if (activeRequest) return showStatus("Another request is already in progress.");
  setRequestState(true);
  showStatus("Checking spelling and grammar…");
  try {
    // Like translation, only the user's own text is proofread: the signature and quoted history
    // are neither sent to the AI provider nor changed.
    const draft = await readDraft();
    const { protectedHtml, images } = protectImages(draft.editableHtml);
    const data = await postJson("/api/proofread", { subject: draft.subject, bodyHtml: protectedHtml, language }, "Proofreading");
    const correctedEditableHtml = restoreImages(data.bodyHtml, images);
    translated = { translatedSubject: data.subject, replacementEditableHtml: correctedEditableHtml };
    $("#gender-label").textContent = `Spelling & grammar: ${language}`;
    $("#english-subject-label").textContent = "Corrected subject — ready to copy";
    $("#english-subject").value = data.subject;
    $("#translated-subject-group").hidden = true;
    $("#translated-body-label").textContent = "Corrected email body — ready to copy";
    $("#translated-body").innerHTML = correctedEditableHtml;
    const isHebrew = language === "Hebrew";
    $("#translated-body").dir = isHebrew ? "rtl" : "ltr";
    $("#translated-body").lang = isHebrew ? "he" : language === "Russian" ? "ru" : "en";
    $("#result").hidden = false;
    showStatus("");
  } catch (error) {
    showStatus(error.message || "Unable to proofread this draft.");
  } finally { setRequestState(false); }
}

async function copyText(value) {
  try {
    await navigator.clipboard.writeText(value);
    showStatus("Copied.");
  } catch {
    showStatus("Copying was blocked. Select the text and press Ctrl+C instead.");
  }
}
async function copyBody() {
  const range = document.createRange();
  range.selectNodeContents($("#translated-body"));
  const selection = window.getSelection();
  selection.removeAllRanges(); selection.addRange(range);
  const copied = document.execCommand("copy"); selection.removeAllRanges();
  showStatus(copied ? "Copied." : "Copying was blocked. Select the text and press Ctrl+C instead.");
}
async function applyTranslation() {
  if (!translated || activeRequest) return;
  const button = $("#apply"); button.disabled = true; showStatus("Applying translation to the draft…");
  try {
    // Read the draft again at the moment of Apply, so the signature, quoted thread and any changes
    // Outlook made to them since Translate was clicked are kept exactly as they are now.
    const draft = await readDraft();
    await Promise.all([setSubject(translated.translatedSubject), setBody(draft.rebuild(translated.replacementEditableHtml))]);
    showStatus("Applied. Review the email, then send it when ready.");
  } catch (error) { showStatus(error.message || "Unable to apply the translation."); }
  finally { button.disabled = false; }
}
