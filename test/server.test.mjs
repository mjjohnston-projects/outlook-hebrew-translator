// Exercises the API routes against a fake OpenAI-compatible provider on 127.0.0.1; nothing leaves this PC.
import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import http from "node:http";

const providerCalls = [];
let providerReply = () => ({ status: 500 });
const fakeProvider = http.createServer((req, res) => {
  let body = "";
  req.on("data", (chunk) => (body += chunk));
  req.on("end", () => {
    const request = JSON.parse(body || "{}");
    providerCalls.push(request);
    const reply = providerReply(request);
    res.writeHead(reply.status || 200, { "content-type": "application/json" });
    res.end(JSON.stringify(reply.status
      ? { error: { message: "fake provider error" } }
      : { id: "x", object: "chat.completion", created: 0, model: "fake", choices: [{ index: 0, finish_reason: "stop", message: { role: "assistant", content: reply.content } }] }));
  });
});
const replyWith = (data) => () => ({ content: typeof data === "string" ? data : JSON.stringify(data) });
const echoBody = (request) => JSON.parse(request.messages[1].content).bodyHtml;

let server;
let baseUrl;
before(async () => {
  await new Promise((resolve) => fakeProvider.listen(0, "127.0.0.1", resolve));
  // Set before importing app.mjs: dotenv never overrides variables that are already set.
  Object.assign(process.env, {
    LLM_PROVIDER: "openai", OPENAI_API_KEY: "test-key", OPENAI_MODEL: "fake-model",
    OPENAI_BASE_URL: `http://127.0.0.1:${fakeProvider.address().port}`
  });
  const { app } = await import("../app.mjs");
  server = http.createServer(app);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  baseUrl = `http://127.0.0.1:${server.address().port}`;
});
after(() => { server?.close(); fakeProvider.close(); });

const post = async (route, body, rawBody) => {
  const response = await fetch(baseUrl + route, { method: "POST", headers: { "content-type": "application/json" }, body: rawBody ?? JSON.stringify(body) });
  return { status: response.status, data: await response.json() };
};

test("rejects invalid requests with JSON errors", async () => {
  assert.equal((await post("/api/translate", { direction: "en-he", recipientGender: "man" })).status, 400);
  assert.equal((await post("/api/translate", { subject: "a", direction: "en-fr", recipientGender: "man" })).status, 400);
  assert.equal((await post("/api/translate", { subject: "a", direction: "en-he" })).status, 400);
  assert.equal((await post("/api/translate", { subject: "a", direction: "en-he", recipientGender: "toString" })).status, 400);
  assert.equal((await post("/api/proofread", { subject: "a", language: "French" })).status, 400);
  const malformed = await post("/api/translate", null, "{oops");
  assert.equal(malformed.status, 400);
  assert.match(malformed.data.error, /could not read/);
});

test("emails over the size limit get a readable JSON error", async () => {
  const result = await post("/api/translate", { subject: "a", bodyHtml: "x".repeat(2_200_000), direction: "he-en" });
  assert.equal(result.status, 413);
  assert.match(result.data.error, /too large/);
});

test("gender instructions name the target language", async () => {
  providerReply = replyWith({ englishSubject: "e", translatedSubject: "t", translatedBodyHtml: "<p>t</p>" });
  providerCalls.length = 0;
  await post("/api/translate", { subject: "Hello", direction: "en-ru", recipientGender: "women" });
  await post("/api/translate", { subject: "Hello", direction: "en-he", recipientGender: "man" });
  assert.match(providerCalls[0].messages[0].content, /multiple women\. Use feminine plural Russian forms whenever addressing the recipients\./);
  assert.doesNotMatch(providerCalls[0].messages[0].content, /Hebrew forms/);
  assert.match(providerCalls[1].messages[0].content, /one man\. Use masculine singular Hebrew forms whenever addressing the recipient\./);
});

test("active content from the model is removed", async () => {
  providerReply = replyWith({ englishSubject: "e", translatedSubject: "t", translatedBodyHtml: `<p onclick="x()">a</p><script>alert(1)</script><a href="javascript:alert(1)">l</a><img src=x onerror=alert(1)>` });
  const { data } = await post("/api/translate", { subject: "a", bodyHtml: "<p>a</p>", direction: "he-en" });
  assert.doesNotMatch(data.translatedBodyHtml, /script|onclick|onerror|javascript:|<img/i);
});

test("text direction and image tokens survive a round trip", async () => {
  providerReply = (request) => ({ content: JSON.stringify({ subject: "s", bodyHtml: echoBody(request) }) });
  const { data } = await post("/api/proofread", { subject: "s", bodyHtml: '<p dir="RTL">שלם [[IMG-1]]</p><div dir="bogus">x</div>', language: "Hebrew" });
  assert.equal(data.bodyHtml, '<p dir="rtl">שלם [[IMG-1]]</p><div>x</div>');
});

test("provider errors produce clear messages", async () => {
  providerReply = () => ({ status: 401 });
  let result = await post("/api/translate", { subject: "a", direction: "he-en" });
  assert.equal(result.status, 401);
  assert.match(result.data.error, /authentication failed\. Check OPENAI_API_KEY/);
  providerReply = () => ({ status: 429 });
  result = await post("/api/proofread", { subject: "a", language: "English" });
  assert.match(result.data.error, /rate or usage limit/);
  providerReply = replyWith("not json");
  result = await post("/api/translate", { subject: "a", direction: "he-en" });
  assert.equal(result.status, 502);
});

test("missing provider configuration is reported", async () => {
  const saved = process.env.OPENAI_API_KEY;
  delete process.env.OPENAI_API_KEY;
  try {
    const result = await post("/api/translate", { subject: "a", direction: "he-en" });
    assert.equal(result.status, 500);
    assert.match(result.data.error, /missing OPENAI_API_KEY/);
  } finally { process.env.OPENAI_API_KEY = saved; }
});
