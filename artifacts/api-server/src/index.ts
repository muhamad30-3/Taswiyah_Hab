import dotenv from "dotenv";
import { createServer as createHttpServer } from "node:http";
import { createServer as createHttpsServer } from "node:https";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

dotenv.config({
  path: path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../../../.env"),
});

const [{ default: app }, { logger }] = await Promise.all([
  import("./app"),
  import("./lib/logger"),
]);

const rawPort = process.env["PORT"];

if (!rawPort) {
  throw new Error(
    "PORT environment variable is required but was not provided.",
  );
}

const port = Number(rawPort);

if (!Number.isInteger(port) || port <= 0 || port > 65535) {
  throw new Error(`Invalid PORT value: "${rawPort}"`);
}

const tlsKeyPath = process.env.API_TLS_KEY_PATH?.trim();
const tlsCertPath = process.env.API_TLS_CERT_PATH?.trim();
if (Boolean(tlsKeyPath) !== Boolean(tlsCertPath)) {
  throw new Error("Set both API_TLS_KEY_PATH and API_TLS_CERT_PATH to enable HTTPS.");
}

const server = tlsKeyPath && tlsCertPath
  ? createHttpsServer({
      key: readFileSync(tlsKeyPath),
      cert: readFileSync(tlsCertPath),
    }, app)
  : createHttpServer(app);

server.listen(port, "0.0.0.0", () => {
  const protocol = tlsKeyPath ? "https" : "http";
  logger.info({ port, protocol }, "Server listening");
});

server.on("error", (err) => {
  logger.error({ err }, "Error listening on port");
  process.exit(1);
});
