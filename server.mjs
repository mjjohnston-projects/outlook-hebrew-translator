import fs from "node:fs";
import https from "node:https";
import devCerts from "office-addin-dev-certs";
import { app } from "./app.mjs";

const port = Number(process.env.PORT || 3000);
// Listen on this PC only by default, so other machines on the network cannot use the configured API key.
// Set HOST=0.0.0.0 only when deploying behind your own access control.
const host = process.env.HOST || "127.0.0.1";

const httpsOptions = process.env.TLS_CERT_PATH && process.env.TLS_KEY_PATH
  ? { cert: fs.readFileSync(process.env.TLS_CERT_PATH), key: fs.readFileSync(process.env.TLS_KEY_PATH) }
  : await devCerts.getHttpsServerOptions();
https.createServer(httpsOptions, app).listen(port, host, () => {
  console.log(`Outlook Email Language Assistant listening on https://localhost:${port} (bound to ${host})`);
});
