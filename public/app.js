let translated = null;
let activeRequest = null;
const $ = (selector) => document.querySelector(selector);

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    showStatus("Open this add-in from an Outlook email draft.");
    return;
  }
  $("#translate").addEventListener("click", translateDraft);
  $("#proofread").addEventListener("click", proofreadDraft);
  $("#direction").addEventListener("change", updateDirectionUi);
  $("#copy-subject").addEventListener("click", () => copyText($("#translated-subject").value));
  $("#copy-body").addEventListener("click", copyBody);
  $("#apply").addEventListener("click", applyTranslation);
});

function showStatus(message = "") { $("#status").textContent = message; }
function setRequestState(button, busy) {
  $("#translate").disabled = busy;
  $("#proofread").disabled = busy;
  if (!busy) button.disabled = false;
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
    catch { throw new Error(`${actionName} returned an invalid response from the local server.`); }
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

function splitPreservedTail(bodyHtml) {
  // Preserve Outlook's signature and quoted reply/forward history exactly as Outlook created them.
  // The first matching marker begins the unchanged tail of the draft.
  const markers = [
    /<(?:div|p|table|span)\b[^>]*(?:\bid|\bclass|\bdata-[\w-]*signature[\w-]*)\s*=\s*["'][^"']*signature[^"']*["'][^>]*>/i,
    /<div\b[^>]*\bid\s*=\s*["']divRplyFwdMsg["'][^>]*>/i,
    /<blockquote\b/i,
    /<hr\b[^>]*>[\s\S]{0,800}?(?:<\/?\w+[^>]*>\s*)*(?:From|Sent|To|Subject)\s*:/i
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

async function translateDraft() {
  const recipientGender = getGender();
  const direction = currentDirection();
  if (direction.needsRecipientForm && !recipientGender) return showStatus("Choose who is being addressed first.");
  const button = $("#translate");
  if (activeRequest) return showStatus("Another request is already in progress.");
  setRequestState(button, true);
  showStatus("Reading and translating your draft…");
  try {
    const [subject, bodyHtml] = await Promise.all([getSubject(), getBody()]);
    const { translatableHtml, preservedTailHtml } = splitPreservedTail(bodyHtml);
    const data = await postJson("/api/translate", { subject, bodyHtml: translatableHtml, recipientGender, direction: $("#direction").value }, "Translation");
    data.translatedBodyHtml = alignTranslatedBody(data.translatedBodyHtml, direction) + preservedTailHtml;
    translated = data;
    $("#english-subject-label").textContent = "Suggested subject (English)";
    $("#translated-subject-group").hidden = false;
    $("#gender-label").textContent = direction.needsRecipientForm ? `Recipient: ${recipientLabels[recipientGender]}` : "Translation: Hebrew to English";
    $("#english-subject").value = data.englishSubject;
    $("#translated-subject-label").textContent = `${direction.label} subject — ready to copy`;
    $("#translated-body-label").textContent = `${direction.label} email body — ready to copy`;
    $("#translated-subject").value = data.translatedSubject;
    $("#translated-subject").dir = direction.dir; $("#translated-subject").lang = direction.language;
    $("#translated-body").innerHTML = data.translatedBodyHtml;
    $("#translated-body").dir = direction.dir; $("#translated-body").lang = direction.language;
    $("#result").hidden = false;
    showStatus("");
  } catch (error) {
    showStatus(error.message || "Unable to translate this draft.");
  } finally { setRequestState(button, false); }
}

async function proofreadDraft() {
  const language = $("#proofread-language").value;
  const button = $("#proofread");
  if (activeRequest) return showStatus("Another request is already in progress.");
  setRequestState(button, true);
  showStatus("Checking spelling and grammar…");
  try {
    const [subject, bodyHtml] = await Promise.all([getSubject(), getBody()]);
    const data = await postJson("/api/proofread", { subject, bodyHtml, language }, "Proofreading");
    translated = { translatedSubject: data.subject, translatedBodyHtml: data.bodyHtml };
    $("#gender-label").textContent = `Spelling & grammar: ${language}`;
    $("#english-subject-label").textContent = "Corrected subject — ready to copy";
    $("#english-subject").value = data.subject;
    $("#translated-subject-group").hidden = true;
    $("#translated-body-label").textContent = "Corrected email body — ready to copy";
    $("#translated-body").innerHTML = data.bodyHtml;
    const isHebrew = language === "Hebrew";
    $("#translated-body").dir = isHebrew ? "rtl" : "ltr";
    $("#translated-body").lang = isHebrew ? "he" : language === "Russian" ? "ru" : "en";
    $("#result").hidden = false;
    showStatus("");
  } catch (error) {
    showStatus(error.message || "Unable to proofread this draft.");
  } finally { setRequestState(button, false); }
}

async function copyText(value) {
  await navigator.clipboard.writeText(value);
  showStatus("Copied.");
}
async function copyBody() {
  const range = document.createRange();
  range.selectNodeContents($("#translated-body"));
  const selection = window.getSelection();
  selection.removeAllRanges(); selection.addRange(range);
  document.execCommand("copy"); selection.removeAllRanges();
  showStatus("Copied.");
}
async function applyTranslation() {
  if (!translated) return;
  const button = $("#apply"); button.disabled = true; showStatus("Applying translation to the draft…");
  try {
    await Promise.all([setSubject(translated.translatedSubject), setBody(translated.translatedBodyHtml)]);
    showStatus("Applied. Review the email, then send it when ready.");
  } catch (error) { showStatus(error.message || "Unable to apply the translation."); }
  finally { button.disabled = false; }
}
