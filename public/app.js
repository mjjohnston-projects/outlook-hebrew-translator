let translated = null;
const $ = (selector) => document.querySelector(selector);

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    showStatus("Open this add-in from an Outlook email draft.");
    return;
  }
  $("#translate").addEventListener("click", translateDraft);
  $("#copy-subject").addEventListener("click", () => copyText($("#hebrew-subject").value));
  $("#copy-body").addEventListener("click", copyBody);
  $("#apply").addEventListener("click", applyTranslation);
});

function showStatus(message = "") { $("#status").textContent = message; }
function getGender() { return document.querySelector('input[name="gender"]:checked')?.value; }
const recipientLabels = { man: "Man", woman: "Woman", men: "Men", women: "Women" };

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
  if (!recipientGender) return showStatus("Choose who is being addressed first.");
  const button = $("#translate");
  button.disabled = true;
  showStatus("Reading and translating your draft…");
  try {
    const [subject, bodyHtml] = await Promise.all([getSubject(), getBody()]);
    const response = await fetch("/api/translate", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ subject, bodyHtml, recipientGender })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Translation failed.");
    translated = data;
    $("#gender-label").textContent = `Recipient: ${recipientLabels[recipientGender]}`;
    $("#english-subject").value = data.englishSubject;
    $("#hebrew-subject").value = data.hebrewSubject;
    $("#hebrew-body").innerHTML = data.hebrewBodyHtml;
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
  range.selectNodeContents($("#hebrew-body"));
  const selection = window.getSelection();
  selection.removeAllRanges(); selection.addRange(range);
  document.execCommand("copy"); selection.removeAllRanges();
  showStatus("Copied.");
}
async function applyTranslation() {
  if (!translated) return;
  const button = $("#apply"); button.disabled = true; showStatus("Applying translation to the draft…");
  try {
    await Promise.all([setSubject(translated.hebrewSubject), setBody(translated.hebrewBodyHtml)]);
    showStatus("Applied. Review the email, then send it when ready.");
  } catch (error) { showStatus(error.message || "Unable to apply the translation."); }
  finally { button.disabled = false; }
}
