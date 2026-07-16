let translated = null;
const $ = (selector) => document.querySelector(selector);

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    showStatus("Open this add-in from an Outlook email draft.");
    return;
  }
  $("#translate").addEventListener("click", translateDraft);
  $("#direction").addEventListener("change", updateDirectionUi);
  $("#copy-subject").addEventListener("click", () => copyText($("#translated-subject").value));
  $("#copy-body").addEventListener("click", copyBody);
  $("#apply").addEventListener("click", applyTranslation);
});

function showStatus(message = "") { $("#status").textContent = message; }
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
  button.disabled = true;
  showStatus("Reading and translating your draft…");
  try {
    const [subject, bodyHtml] = await Promise.all([getSubject(), getBody()]);
    const response = await fetch("/api/translate", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ subject, bodyHtml, recipientGender, direction: $("#direction").value })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Translation failed.");
    translated = data;
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
  } finally { button.disabled = false; }
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
