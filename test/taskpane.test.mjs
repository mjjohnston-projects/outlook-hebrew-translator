// Runs the real public/index.html and public/app.js in jsdom against a simulated Outlook draft.
import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import { JSDOM } from "jsdom";

const read = (file) => fs.readFileSync(new URL(`../${file}`, import.meta.url), "utf8");
const html = read("public/index.html").replace(/<script[^>]*><\/script>/g, "");
const appJs = read("public/app.js");
const internals = "convertKeyboardText, convertKeyboardHtml, splitPreservedTail, splitDocument, protectImages, restoreImages, alignTranslatedBody, directions";

function boot({ subject = "", body = "", fetch: fetchImpl } = {}) {
  const dom = new JSDOM(html, { runScripts: "outside-only", url: "https://localhost:3000/index.html" });
  const w = dom.window;
  const draft = { subject, body, requests: [] };
  const ok = (value) => ({ status: "succeeded", value });
  w.Office = {
    HostType: { Outlook: "Outlook" }, AsyncResultStatus: { Succeeded: "succeeded" }, CoercionType: { Html: "html" },
    onReady: (callback) => callback({ host: "Outlook" }),
    context: { mailbox: { item: {
      subject: { getAsync: (cb) => cb(ok(draft.subject)), setAsync: (value, cb) => { draft.subject = value; cb(ok()); } },
      body: { getAsync: (_type, cb) => cb(ok(draft.body)), setAsync: (value, _options, cb) => { draft.body = value; cb(ok()); } }
    } } }
  };
  w.fetch = async (url, options) => {
    draft.requests.push({ url, body: JSON.parse(options.body) });
    return fetchImpl(url, JSON.parse(options.body));
  };
  w.eval(`${appJs}\n;window.__internals = { ${internals} };`);
  return {
    w, draft, fn: w.__internals,
    $: (selector) => w.document.querySelector(selector),
    settle: () => new Promise((resolve) => setTimeout(resolve, 20))
  };
}
const json = (data, status = 200) => new Response(JSON.stringify(data), { status });
const outlookDoc = (inner) => `<html><head><style>p.MsoNormal{margin:0;font-family:Aptos}</style></head><body lang="EN-US"><div class="WordSection1">${inner}</div></body></html>`;
const replyTail = `<div id="Signature"><p>Michael</p><img src="cid:logo"></div><div id="divRplyFwdMsg"><b>From:</b> Someone else</div><div>their private text</div>`;

test("source files contain no double-encoded (mojibake) characters", () => {
  for (const file of ["public/app.js", "public/index.html", "README.md", "app.mjs"]) {
    assert.doesNotMatch(read(file), /×[\u0080-˿ -⁯]|â€|â†|Ã./, file);
  }
});

test("keyboard fix converts mistyped Hebrew", () => {
  const { fn } = boot();
  assert.equal(fn.convertKeyboardText("akuo"), "שלום");
  assert.equal(fn.convertKeyboardText("nv akunu'"), "מה שלומו,");
  assert.equal(fn.convertKeyboardText("tbh rumv kv,rtu,"), "אני רוצה להתראות");
});

test("keyboard fix leaves correct Hebrew, English words, names, numbers and links alone", () => {
  const { fn } = boot();
  for (const text of ["שלום, מה נשמע. 'בסדר'", "Hi David, the file is report.pdf", "Room 205", "TAU", "is it ok"]) {
    assert.equal(fn.convertKeyboardText(text), text);
  }
  const mixed = fn.convertKeyboardText("Hi David, akuo https://example.com/a?b=c me@tau.ac.il");
  assert.equal(mixed, "Hi David, שלום https://example.com/a?b=c me@tau.ac.il");
});

test("ambiguous short words follow their neighbours", () => {
  const { fn } = boot();
  assert.equal(fn.convertKeyboardText("do akuo"), "גם שלום");
  assert.equal(fn.convertKeyboardText("do it"), "do it");
});

test("keyboard fix reports when there is nothing to fix", async () => {
  const t = boot({ subject: "שלום", body: outlookDoc("<p class=MsoNormal>שלום רב, מה נשמע.</p>") });
  t.$("#keyboard-hebrew").click(); await t.settle();
  assert.match(t.$("#status").textContent, /No English-keyboard Hebrew text/);
  assert.equal(t.$("#result").hidden, true);
});

test("keyboard fix keeps the Outlook document head, signature and reply history", async () => {
  const t = boot({ subject: "akuo", body: outlookDoc(`<p>akuo</p>${replyTail}`) });
  t.$("#keyboard-hebrew").click(); await t.settle();
  t.$("#apply").click(); await t.settle();
  assert.equal(t.draft.subject, "שלום");
  assert.match(t.draft.body, /^<html><head><style>p\.MsoNormal/);
  assert.match(t.draft.body, /שלום/);
  assert.ok(t.draft.body.includes(replyTail));
  assert.match(t.draft.body, /<\/body><\/html>$/);
});

test("tail split keeps signatures and quoted history, but not the user's own blockquote", () => {
  const { fn } = boot();
  const reply = fn.splitPreservedTail(`<p>Hello</p>${replyTail}`);
  assert.equal(reply.translatableHtml, "<p>Hello</p>");
  const own = fn.splitPreservedTail("<p>Note:</p><blockquote>my quote</blockquote><p>Thanks</p>");
  assert.equal(own.preservedTailHtml, "");
  const cited = fn.splitPreservedTail('<p>Hi</p><blockquote type="cite">old</blockquote>');
  assert.equal(cited.translatableHtml, "<p>Hi</p>");
});

test("images are replaced by tokens and restored, even if the model drops one", () => {
  const { fn } = boot();
  const { protectedHtml, images } = fn.protectImages('<p>a<img src="cid:1"></p><p><img src="data:image/png;base64,AAAA"></p>');
  assert.equal(protectedHtml, "<p>a[[IMG-1]]</p><p>[[IMG-2]]</p>");
  assert.equal(fn.restoreImages("<p>b[[IMG-1]][[IMG-1]]</p>", images), '<p>b<img src="cid:1"></p><p><img src="data:image/png;base64,AAAA"></p>');
});

test("translate sends only the user's text (no images, signature or history) and applies right-to-left", async () => {
  const t = boot({
    subject: "Hello",
    body: outlookDoc(`<p>Hello<img src="data:image/png;base64,${"A".repeat(5000)}"></p>${replyTail}`),
    fetch: (_url, body) => json({ englishSubject: "Hello", translatedSubject: "שלום", translatedBodyHtml: body.bodyHtml.replace("Hello", "שלום") })
  });
  t.$('input[value="woman"]').checked = true;
  t.$("#translate").click(); await t.settle();
  const sent = t.draft.requests[0].body;
  assert.equal(sent.bodyHtml, '<div class="WordSection1"><p>Hello[[IMG-1]]</p>');
  assert.equal(sent.recipientGender, "woman");
  t.$("#apply").click(); await t.settle();
  assert.match(t.draft.body, /^<html><head><style>/);
  assert.match(t.draft.body, /<div dir="rtl"/);
  assert.match(t.draft.body, /שלום<img src="data:image\/png;base64,A{5000}"/);
  assert.ok(t.draft.body.includes(replyTail));
});

test("proofread sends only the user's text and keeps the signature and history", async () => {
  const t = boot({
    subject: "s",
    body: outlookDoc(`<p dir="rtl">שלם</p>${replyTail}`),
    fetch: (_url, body) => json({ subject: "s", bodyHtml: body.bodyHtml.replace("שלם", "שלום") })
  });
  t.$("#proofread-language").value = "Hebrew";
  t.$("#proofread").click(); await t.settle();
  assert.doesNotMatch(t.draft.requests[0].body.bodyHtml, /their private text|Signature/);
  t.$("#apply").click(); await t.settle();
  assert.match(t.draft.body, /<p dir="rtl">שלום<\/p>/);
  assert.ok(t.draft.body.includes(replyTail));
});

test("a non-JSON 413 response shows a helpful message", async () => {
  const t = boot({ subject: "s", body: "<p>x</p>", fetch: () => new Response("<html>Payload Too Large</html>", { status: 413 }) });
  t.$('input[value="man"]').checked = true;
  t.$("#translate").click(); await t.settle();
  assert.match(t.$("#status").textContent, /too large/);
});

test("Apply is disabled while another request is running", async () => {
  let release;
  let calls = 0;
  const t = boot({
    subject: "s", body: "<p>x</p>",
    fetch: async () => { if (++calls === 2) await new Promise((resolve) => (release = resolve)); return json({ englishSubject: "e", translatedSubject: `T${calls}`, translatedBodyHtml: "<p>x</p>" }); }
  });
  t.$('input[value="man"]').checked = true;
  t.$("#translate").click(); await t.settle();
  t.$("#translate").click(); await t.settle();
  assert.equal(t.$("#apply").disabled, true);
  release(); await t.settle();
  assert.equal(t.$("#apply").disabled, false);
});
