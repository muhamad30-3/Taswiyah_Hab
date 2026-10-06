import express, { type Express } from "express";
import cors from "cors";
import pinoHttp from "pino-http";
import router from "./routes";
import { logger } from "./lib/logger";

const app: Express = express();

function configuredAppOrigin(): string | false {
  const configured = process.env.APP_ORIGIN?.trim();
  if (!configured) return false;

  let origin: URL;
  try {
    origin = new URL(configured);
  } catch {
    throw new Error("APP_ORIGIN must be a valid absolute HTTP(S) origin.");
  }
  if (
    !["http:", "https:"].includes(origin.protocol) ||
    origin.username ||
    origin.password ||
    origin.pathname !== "/" ||
    origin.search ||
    origin.hash ||
    (process.env.NODE_ENV === "production" && origin.protocol !== "https:")
  ) {
    throw new Error("APP_ORIGIN must be an HTTPS origin in production, without a path, query, or credentials.");
  }
  return origin.origin;
}

app.set("trust proxy", process.env.TRUST_PROXY === "true");
app.disable("x-powered-by");
app.use((_request, response, next) => {
  response.setHeader("X-Content-Type-Options", "nosniff");
  response.setHeader("X-Frame-Options", "DENY");
  response.setHeader("Referrer-Policy", "no-referrer");
  if (process.env.NODE_ENV === "production") {
    response.setHeader("Strict-Transport-Security", "max-age=31536000");
  }
  next();
});
app.use(
  pinoHttp({
    logger,
    serializers: {
      req(req) {
        return {
          id: req.id,
          method: req.method,
          url: req.url?.split("?")[0],
        };
      },
      res(res) {
        return {
          statusCode: res.statusCode,
        };
      },
    },
  }),
);
app.use(cors({
  origin: configuredAppOrigin(),
  credentials: true,
  methods: ["GET", "HEAD", "POST", "PUT", "DELETE"],
  allowedHeaders: ["Content-Type", "X-File-Name", "X-File-Type"],
}));
app.use("/api/documents/ocr", express.json({ limit: "8mb" }));
app.use("/api/storage", express.json({ limit: "160mb" }));
app.use(express.json({ limit: "512kb" }));
app.use(express.urlencoded({ extended: true, limit: "100kb", parameterLimit: 100 }));

app.use("/api", router);

export default app;
