import { randomBytes, randomUUID, scrypt as scryptCallback, createHmac, timingSafeEqual } from "node:crypto";
import { gzip as gzipCallback, gunzip as gunzipCallback } from "node:zlib";
import { promisify } from "node:util";
import express, { Router, type NextFunction, type Request, type Response } from "express";
import { logger } from "../lib/logger";
import { passwordPolicyError } from "../../../../src/lib/password-policy.js";

const router = Router();
const scrypt = promisify(scryptCallback);
const gzip = promisify(gzipCallback);
const gunzip = promisify(gunzipCallback);
const sessionCookieName = "taswiyah_session";
const sessionLifetimeSeconds = 12 * 60 * 60;
const maximumStoredKeyLength = 120;
const maximumValueSize = 50 * 1024 * 1024;
const compressValueAboveBytes = 128 * 1024;
const compressedValueEncoding = "gzip-base64-v2";
const legacyCompressedValueEncoding = "gzip-base64-v1";
const maximumFileSize = 25 * 1024 * 1024;
const maximumLoginFailures = 10;
const loginWindowMilliseconds = 15 * 60 * 1000;
const loginFailures = new Map<string, { count: number; resetAt: number }>();
const allowedExtensions = new Set([
  "xlsx", "xls", "xlsm", "xlsb", "csv",
  "pdf", "doc", "docx", "docm", "dotx", "odt", "rtf",
  "ppt", "pptx", "pptm", "ppsx", "odp",
  "txt", "md", "html", "htm",
  "png", "jpg", "jpeg", "webp", "gif", "tif", "tiff", "bmp",
]);

type UserRole = "admin" | "user";
type SessionPayload = { username: string; expiresAt: number };
type AuthenticatedRequest = Request & { appUser?: { id: string; username: string; role: UserRole } };

function sessionSecret() {
  const secret = process.env.APP_SESSION_SECRET;
  return secret && Buffer.byteLength(secret, "utf8") >= 32 ? secret : null;
}

function authConfigurationError() {
  const missing: string[] = [];
  if (!process.env.DATABASE_URL) missing.push("DATABASE_URL");
  if (!process.env.APP_USERNAME?.trim()) missing.push("APP_USERNAME");
  if (!process.env.APP_PASSWORD || passwordPolicyError(process.env.APP_PASSWORD)) {
    missing.push("APP_PASSWORD (8 أحرف على الأقل، مع حرف ورقم)");
  }
  if (!sessionSecret()) missing.push("APP_SESSION_SECRET (32 بايت على الأقل)");
  return missing.length ? `إعدادات الخادم غير مكتملة: ${missing.join("، ")}.` : null;
}

function loginRateLimited(address: string) {
  const now = Date.now();
  const current = loginFailures.get(address);
  if (!current || current.resetAt <= now) {
    loginFailures.set(address, { count: 0, resetAt: now + loginWindowMilliseconds });
    return false;
  }
  return current.count >= maximumLoginFailures;
}

function recordLoginFailure(address: string) {
  if (loginFailures.size >= 10_000) {
    const now = Date.now();
    for (const [client, entry] of loginFailures) if (entry.resetAt <= now) loginFailures.delete(client);
    if (loginFailures.size >= 10_000) return;
  }
  const current = loginFailures.get(address);
  if (current) current.count++;
  else loginFailures.set(address, { count: 1, resetAt: Date.now() + loginWindowMilliseconds });
}

function normalizeUsername(username: string) {
  return username.normalize("NFKC").trim().toLocaleLowerCase("en-US");
}

function sign(value: string) {
  return createHmac("sha256", sessionSecret()!).update(value).digest("base64url");
}

function readCookie(request: Request, name: string) {
  const prefix = `${name}=`;
  const cookie = request.headers.cookie?.split(";").map(part => part.trim()).find(part => part.startsWith(prefix));
  return cookie ? cookie.slice(prefix.length) : null;
}

function readSession(request: Request): SessionPayload | null {
  const token = readCookie(request, sessionCookieName);
  if (!token) return null;
  const [payload, signature, extra] = token.split(".");
  if (!payload || !signature || extra !== undefined || !sessionSecret()) return null;
  const expected = sign(payload);
  const providedBytes = Buffer.from(signature);
  const expectedBytes = Buffer.from(expected);
  if (providedBytes.length !== expectedBytes.length || !timingSafeEqual(providedBytes, expectedBytes)) return null;
  try {
    const parsed = JSON.parse(Buffer.from(payload, "base64url").toString("utf8")) as SessionPayload;
    if (typeof parsed.username !== "string" || typeof parsed.expiresAt !== "number" || parsed.expiresAt <= Date.now()) return null;
    return parsed;
  } catch {
    return null;
  }
}

function setSessionCookie(response: Response, username: string) {
  const payload = Buffer.from(JSON.stringify({
    username,
    expiresAt: Date.now() + sessionLifetimeSeconds * 1000,
  })).toString("base64url");
  const token = `${payload}.${sign(payload)}`;
  const secure = process.env.NODE_ENV === "production" ? "; Secure" : "";
  response.setHeader("Set-Cookie", `${sessionCookieName}=${token}; Path=/api; HttpOnly; SameSite=Strict; Max-Age=${sessionLifetimeSeconds}${secure}`);
}

function clearSessionCookie(response: Response) {
  const secure = process.env.NODE_ENV === "production" ? "; Secure" : "";
  response.setHeader("Set-Cookie", `${sessionCookieName}=; Path=/api; HttpOnly; SameSite=Strict; Max-Age=0${secure}`);
}

async function getDataStore() {
  const { pool } = await import("@workspace/db");
  return pool;
}

async function compressStorageResponse(jsonValue: string) {
  const payload = await gzip(Buffer.from(`{"value":${jsonValue}}`, "utf8"));
  return {
    payload,
    storedValue: JSON.stringify({
      __taswiyah_encoding: compressedValueEncoding,
      data: payload.toString("base64"),
    }),
  };
}

async function hashPassword(password: string) {
  const salt = randomBytes(16);
  const derived = await scrypt(password, salt, 64) as Buffer;
  return `${salt.toString("hex")}:${derived.toString("hex")}`;
}

async function verifyPassword(password: string, stored: string) {
  const [saltHex, hashHex, extra] = stored.split(":");
  if (!saltHex || !hashHex || extra !== undefined) return false;
  const salt = Buffer.from(saltHex, "hex");
  const expected = Buffer.from(hashHex, "hex");
  if (salt.length !== 16 || expected.length !== 64) return false;
  const actual = await scrypt(password, salt, expected.length) as Buffer;
  return timingSafeEqual(actual, expected);
}

async function ensureBootstrapAdmin() {
  const username = normalizeUsername(process.env.APP_USERNAME!);
  const pool = await getDataStore();
  const existing = await pool.query<{ id: string; username: string; role: UserRole; status: string }>(
    "SELECT id, username, role, status FROM app_users WHERE username = $1 LIMIT 1",
    [username],
  );
  if (existing.rows[0]) {
    const admin = existing.rows[0];
    if (admin.role !== "admin" || admin.status !== "approved") {
      await pool.query("UPDATE app_users SET role = 'admin', status = 'approved' WHERE id = $1", [admin.id]);
    }
    await migrateLegacyAdminData(pool, username, admin.id);
    return { ...admin, role: "admin" as const, status: "approved" };
  }
  const passwordHash = await hashPassword(process.env.APP_PASSWORD!);
  const email = process.env.APP_EMAIL?.trim().toLowerCase() || null;
  await pool.query(
    `INSERT INTO app_users (id, username, email, password_hash, role, status)
     VALUES ($1, $2, $3, $4, 'admin', 'approved')
     ON CONFLICT (username) DO NOTHING`,
    [randomUUID(), username, email, passwordHash],
  );
  const { rows } = await pool.query<{ id: string; username: string; role: UserRole; status: string; password_hash: string }>(
    "SELECT id, username, role, status, password_hash FROM app_users WHERE username = $1 LIMIT 1",
    [username],
  );
  const admin = rows[0];
  if (!admin) throw new Error("تعذر إنشاء حساب المسؤول الأولي.");
  await migrateLegacyAdminData(pool, username, admin.id);
  return { ...admin, role: "admin" as const, status: "approved" };
}

async function migrateLegacyAdminData(pool: Awaited<ReturnType<typeof getDataStore>>, username: string, userId: string) {
  const legacyUsername = process.env.APP_USERNAME!.trim();
  await pool.query(
    `INSERT INTO app_data (owner, key, value, updated_at)
     SELECT $3, key, value, updated_at FROM app_data
     WHERE owner IN ($1, $2) AND owner <> $3
     ON CONFLICT (owner, key) DO NOTHING`,
    [legacyUsername, username, userId],
  );
}

function authReady(response: Response) {
  const error = authConfigurationError();
  if (error) {
    response.status(503).json({ error });
    return false;
  }
  if (process.env.CENTRAL_STORAGE_ENABLED !== "true") {
    response.status(404).json({ error: "تسجيل الدخول المركزي غير مفعّل." });
    return false;
  }
  return true;
}

async function requireSession(request: AuthenticatedRequest, response: Response, next: NextFunction) {
  if (!authReady(response)) return;
  const session = readSession(request);
  if (!session) {
    response.status(401).json({ error: "انتهت جلسة الدخول. سجّل الدخول مجدداً." });
    return;
  }
  try {
    const pool = await getDataStore();
    const { rows } = await pool.query<{ id: string; username: string; role: UserRole; status: string }>(
      "SELECT id, username, role, status FROM app_users WHERE username = $1 LIMIT 1",
      [session.username],
    );
    const user = rows[0];
    if (!user || user.status !== "approved") {
      clearSessionCookie(response);
      response.status(401).json({ error: "هذا الحساب غير معتمد أو لم يعد نشطاً." });
      return;
    }
    request.appUser = { id: user.id, username: user.username, role: user.role };
    next();
  } catch (error) {
    logger.error({ err: error }, "Central authentication lookup failed");
    response.status(503).json({ error: "تعذر التحقق من الحساب عبر قاعدة البيانات." });
  }
}

function requireAdmin(request: AuthenticatedRequest, response: Response, next: NextFunction) {
  if (request.appUser?.role !== "admin") {
    response.status(403).json({ error: "هذه الصفحة مخصصة للمسؤول." });
    return;
  }
  next();
}

router.get("/auth/config", (_request, response) => {
  response.json({ centralStorage: process.env.CENTRAL_STORAGE_ENABLED === "true" });
});

router.post("/auth/register", async (request, response) => {
  if (!authReady(response)) return;
  if (loginRateLimited(request.ip || "unknown")) {
    response.status(429).json({ error: "طلبات كثيرة من هذا الاتصال. انتظر 15 دقيقة ثم أعد المحاولة." });
    return;
  }
  recordLoginFailure(request.ip || "unknown");
  const { username, email, password } = request.body ?? {};
  if (typeof username !== "string" || typeof email !== "string" || typeof password !== "string") {
    response.status(400).json({ error: "أدخل اسم المستخدم والبريد الإلكتروني وكلمة المرور." });
    return;
  }
  const normalizedUsername = normalizeUsername(username);
  const normalizedEmail = email.trim().toLowerCase();
  if (!/^[\p{L}\p{N}._-]{3,32}$/u.test(normalizedUsername)) {
    response.status(400).json({ error: "اسم المستخدم يجب أن يتكون من 3 إلى 32 حرفاً أو رقماً." });
    return;
  }
  if (normalizedEmail.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizedEmail)) {
    response.status(400).json({ error: "أدخل بريداً إلكترونياً صالحاً." });
    return;
  }
  const passwordError = passwordPolicyError(password);
  if (passwordError) {
    response.status(400).json({ error: passwordError });
    return;
  }
  try {
    await ensureBootstrapAdmin();
    const pool = await getDataStore();
    await pool.query(
      `INSERT INTO app_users (id, username, email, password_hash, role, status)
       VALUES ($1, $2, $3, $4, 'user', 'pending')`,
      [randomUUID(), normalizedUsername, normalizedEmail, await hashPassword(password)],
    );
    response.status(201).json({ registered: true, status: "pending" });
  } catch (error) {
    const code = (error as { code?: string }).code;
    if (code === "23505") {
      response.status(409).json({ error: "اسم المستخدم أو البريد الإلكتروني مسجّل مسبقاً." });
      return;
    }
    logger.error({ err: error }, "Account registration failed");
    response.status(503).json({ error: "تعذر إنشاء الحساب في قاعدة البيانات. أعد المحاولة لاحقاً." });
  }
});

router.post("/auth/login", async (request, response) => {
  if (!authReady(response)) return;
  const address = request.ip || "unknown";
  if (loginRateLimited(address)) {
    response.status(429).json({ error: "محاولات دخول كثيرة. انتظر 15 دقيقة ثم أعد المحاولة." });
    return;
  }
  const { username, password } = request.body ?? {};
  if (typeof username !== "string" || typeof password !== "string" || username.length > 256 || password.length > 256) {
    response.status(400).json({ error: "صيغة طلب تسجيل الدخول غير صالحة." });
    return;
  }
  try {
    const bootstrap = await ensureBootstrapAdmin();
    const normalizedUsername = normalizeUsername(username);
    const pool = await getDataStore();
    const { rows } = await pool.query<{ id: string; username: string; role: UserRole; status: string; password_hash: string }>(
      "SELECT id, username, role, status, password_hash FROM app_users WHERE username = $1 LIMIT 1",
      [normalizedUsername],
    );
    const user = rows[0];
    if (!user || !(await verifyPassword(password, user.password_hash))) {
      recordLoginFailure(address);
      response.status(401).json({ error: "اسم المستخدم أو كلمة المرور غير صحيحة." });
      return;
    }
    if (user.status !== "approved") {
      recordLoginFailure(address);
      response.status(403).json({
        error: user.status === "pending" ? "تم استلام طلبك، وينتظر موافقة المسؤول قبل تسجيل الدخول." : "لم تتم الموافقة على هذا الحساب.",
      });
      return;
    }
    loginFailures.delete(address);
    setSessionCookie(response, user.username);
    response.json({ username: user.username, role: user.id === bootstrap.id ? "admin" : user.role });
  } catch (error) {
    logger.error({ err: error }, "Central login failed");
    response.status(503).json({ error: "تعذر التحقق من بيانات الدخول عبر قاعدة البيانات." });
  }
});

router.get("/auth/me", async (request, response) => {
  if (!authReady(response)) return;
  const session = readSession(request);
  if (!session) {
    response.status(401).json({ error: "لا توجد جلسة دخول." });
    return;
  }
  try {
    const bootstrap = await ensureBootstrapAdmin();
    const pool = await getDataStore();
    const { rows } = await pool.query<{ id: string; username: string; role: UserRole; status: string }>(
      "SELECT id, username, role, status FROM app_users WHERE username = $1 LIMIT 1",
      [session.username],
    );
    const user = rows[0];
    if (!user || user.status !== "approved") {
      clearSessionCookie(response);
      response.status(401).json({ error: "هذا الحساب غير معتمد أو لم يعد نشطاً." });
      return;
    }
    response.json({ username: user.username, role: user.id === bootstrap.id ? "admin" : user.role });
  } catch (error) {
    logger.error({ err: error }, "Central session lookup failed");
    response.status(503).json({ error: "تعذر التحقق من الجلسة عبر قاعدة البيانات." });
  }
});

router.post("/auth/change-password", requireSession, async (request: AuthenticatedRequest, response) => {
  const address = request.ip || "unknown";
  if (loginRateLimited(address)) {
    response.status(429).json({ error: "محاولات كثيرة. انتظر 15 دقيقة ثم أعد المحاولة." });
    return;
  }
  const { currentPassword, newPassword } = request.body ?? {};
  if (
    typeof currentPassword !== "string" ||
    currentPassword.length === 0 ||
    currentPassword.length > 256 ||
    typeof newPassword !== "string" ||
    newPassword.length === 0 ||
    newPassword.length > 256
  ) {
    response.status(400).json({ error: "أدخل كلمة المرور الحالية والجديدة (بحد أقصى 256 حرفاً)." });
    return;
  }
  const passwordError = passwordPolicyError(newPassword);
  if (passwordError) {
    response.status(400).json({ error: passwordError });
    return;
  }

  try {
    const pool = await getDataStore();
    const { rows } = await pool.query<{ password_hash: string }>(
      "SELECT password_hash FROM app_users WHERE id = $1 LIMIT 1",
      [request.appUser!.id],
    );
    const user = rows[0];
    if (!user || !(await verifyPassword(currentPassword, user.password_hash))) {
      recordLoginFailure(address);
      response.status(401).json({ error: "كلمة المرور الحالية غير صحيحة." });
      return;
    }
    loginFailures.delete(address);
    await pool.query(
      "UPDATE app_users SET password_hash = $1 WHERE id = $2",
      [await hashPassword(newPassword), request.appUser!.id],
    );
    response.json({ updated: true });
  } catch (error) {
    logger.error({ err: error }, "Account password change failed");
    response.status(503).json({ error: "تعذر تغيير كلمة المرور. حاول مجدداً." });
  }
});

router.post("/auth/logout", (_request, response) => {
  clearSessionCookie(response);
  response.status(204).end();
});

router.get("/admin/users", requireSession, requireAdmin, async (_request: AuthenticatedRequest, response) => {
  try {
    const pool = await getDataStore();
    const { rows } = await pool.query(
      "SELECT id, username, email, role, status, created_at AS \"createdAt\" FROM app_users ORDER BY created_at DESC",
    );
    response.json({ users: rows });
  } catch (error) {
    logger.error({ err: error }, "Admin user list failed");
    response.status(503).json({ error: "تعذر تحميل قائمة المستخدمين." });
  }
});

router.post("/admin/users/:id/approve", requireSession, requireAdmin, async (request: AuthenticatedRequest, response) => {
  try {
    const pool = await getDataStore();
    const result = await pool.query(
      "UPDATE app_users SET status = 'approved' WHERE id = $1 AND role = 'user' RETURNING id",
      [request.params.id],
    );
    if (!result.rowCount) {
      response.status(404).json({ error: "لم يتم العثور على حساب مستخدم قابل للاعتماد." });
      return;
    }
    response.json({ updated: true });
  } catch (error) {
    logger.error({ err: error }, "Admin user approval failed");
    response.status(503).json({ error: "تعذر اعتماد الحساب." });
  }
});

router.post("/admin/users/:id/reject", requireSession, requireAdmin, async (request: AuthenticatedRequest, response) => {
  try {
    const pool = await getDataStore();
    const result = await pool.query(
      "UPDATE app_users SET status = 'rejected' WHERE id = $1 AND role = 'user' RETURNING id",
      [request.params.id],
    );
    if (!result.rowCount) {
      response.status(404).json({ error: "لم يتم العثور على حساب مستخدم قابل للتحديث." });
      return;
    }
    response.json({ updated: true });
  } catch (error) {
    logger.error({ err: error }, "Admin user rejection failed");
    response.status(503).json({ error: "تعذر رفض الحساب." });
  }
});

router.get("/storage/:key", requireSession, async (request: AuthenticatedRequest, response) => {
  const key = request.params.key;
  if (key.length > maximumStoredKeyLength) {
    response.status(400).json({ error: "مفتاح التخزين طويل جداً." });
    return;
  }
  try {
    const pool = await getDataStore();
    const { rows } = await pool.query<{ payload: string; encoding: string | null; revision: string }>(
      `SELECT
         CASE WHEN value->>'__taswiyah_encoding' IN ($3, $4) THEN value->>'data' ELSE value::text END AS payload,
         value->>'__taswiyah_encoding' AS encoding,
         updated_at::text AS revision
       FROM app_data WHERE owner = $1 AND key = $2 LIMIT 1`,
      [request.appUser!.id, key, compressedValueEncoding, legacyCompressedValueEncoding],
    );
    const record = rows[0];
    if (!record) {
      response.status(404).json({ error: "عنصر التخزين غير موجود." });
      return;
    }
    const acceptsGzip = request.acceptsEncodings("gzip") === "gzip";
    let responseBody: string | Buffer;
    let responseIsCompressed = false;
    if (record.encoding === compressedValueEncoding) {
      const compressedResponse = Buffer.from(record.payload, "base64");
      if (acceptsGzip) {
        responseBody = compressedResponse;
        responseIsCompressed = true;
      } else {
        responseBody = await gunzip(compressedResponse);
      }
    } else {
      const jsonValue = record.encoding === legacyCompressedValueEncoding
        ? (await gunzip(Buffer.from(record.payload, "base64"))).toString("utf8")
        : record.payload;
      responseBody = `{"value":${jsonValue}}`;
      if (Buffer.byteLength(responseBody, "utf8") > compressValueAboveBytes) {
        try {
          const compressed = await compressStorageResponse(jsonValue);
          await pool.query(
            `UPDATE app_data SET value = $1::jsonb, updated_at = NOW()
             WHERE owner = $2 AND key = $3 AND updated_at::text = $4`,
            [compressed.storedValue, request.appUser!.id, key, record.revision],
          );
          if (acceptsGzip) {
            responseBody = compressed.payload;
            responseIsCompressed = true;
          }
        } catch (error) {
          logger.warn({ err: error, key }, "Could not compress legacy storage value; returning it without compression");
        }
      }
    }
    const result = response.status(200).type("application/json");
    if (responseIsCompressed) result.set("Content-Encoding", "gzip").vary("Accept-Encoding");
    result.send(responseBody);
  } catch (error) {
    logger.error({ err: error }, "Central storage read failed");
    response.status(503).json({ error: "تعذرت قراءة البيانات من قاعدة البيانات. أعد المحاولة." });
  }
});

router.put("/storage/:key", requireSession, async (request: AuthenticatedRequest, response) => {
  const key = request.params.key;
  if (key.length > maximumStoredKeyLength || !Object.prototype.hasOwnProperty.call(request.body ?? {}, "value")) {
    response.status(400).json({ error: "طلب التخزين غير صالح." });
    return;
  }
  const serialized = JSON.stringify(request.body.value);
  const serializedBytes = serialized ? Buffer.byteLength(serialized, "utf8") : 0;
  if (!serialized || serializedBytes > maximumValueSize) {
    response.status(413).json({ error: "حجم البيانات أكبر من الحد المسموح للتخزين." });
    return;
  }
  try {
    const pool = await getDataStore();
    const storedValue = serializedBytes > compressValueAboveBytes
      ? (await compressStorageResponse(serialized)).storedValue
      : serialized;
    await pool.query(
      `INSERT INTO app_data (owner, key, value, updated_at)
       VALUES ($1, $2, $3::jsonb, NOW())
       ON CONFLICT (owner, key)
       DO UPDATE SET value = EXCLUDED.value, updated_at = NOW()`,
      [request.appUser!.id, key, storedValue],
    );
    response.json({ saved: true });
  } catch (error) {
    logger.error({ err: error }, "Central storage write failed");
    response.status(503).json({ error: "تعذر حفظ البيانات في قاعدة البيانات. لم يتم تأكيد الحفظ." });
  }
});

router.delete("/storage/:key", requireSession, async (request: AuthenticatedRequest, response) => {
  const key = request.params.key;
  if (key.length > maximumStoredKeyLength) {
    response.status(400).json({ error: "مفتاح التخزين طويل جداً." });
    return;
  }
  try {
    const pool = await getDataStore();
    await pool.query("DELETE FROM app_data WHERE owner = $1 AND key = $2", [request.appUser!.id, key]);
    response.status(204).end();
  } catch (error) {
    logger.error({ err: error }, "Central storage delete failed");
    response.status(503).json({ error: "تعذر حذف البيانات من قاعدة البيانات." });
  }
});

router.get("/files", requireSession, async (request: AuthenticatedRequest, response) => {
  try {
    const pool = await getDataStore();
    const { rows } = await pool.query(
      `SELECT id, file_name AS "fileName", mime_type AS "mimeType", size,
              uploaded_at AS "uploadedAt"
       FROM user_files WHERE owner = $1 ORDER BY uploaded_at DESC`,
      [request.appUser!.id],
    );
    response.json({ files: rows });
  } catch (error) {
    logger.error({ err: error }, "User file listing failed");
    response.status(503).json({ error: "تعذر تحميل ملفاتك من قاعدة البيانات." });
  }
});

router.post(
  "/files",
  requireSession,
  (request, response, next) => {
    express.raw({ type: "*/*", limit: maximumFileSize })(request, response, error => {
      if (!error) {
        next();
        return;
      }
      if ((error as { status?: number }).status === 413) {
        response.status(413).json({ error: "حجم الملف أكبر من الحد المسموح (25 ميغابايت)." });
        return;
      }
      next(error);
    });
  },
  async (request: AuthenticatedRequest, response) => {
  const body = request.body;
  let fileName = "";
  try {
    fileName = decodeURIComponent(String(request.headers["x-file-name"] ?? ""));
  } catch {
    response.status(400).json({ error: "اسم الملف غير صالح." });
    return;
  }
  const requestedMimeType = String(request.headers["x-file-type"] ?? "application/octet-stream");
  const mimeType = /^[\w!#$&^_.+-]+\/[\w!#$&^_.+-]+$/.test(requestedMimeType)
    ? requestedMimeType.slice(0, 160)
    : "application/octet-stream";
  const extension = fileName.split(".").pop()?.toLowerCase();
  if (!Buffer.isBuffer(body) || body.length === 0 || !fileName.trim() || !extension || !allowedExtensions.has(extension)) {
    response.status(400).json({ error: "الملف فارغ أو نوعه غير مدعوم." });
    return;
  }
  if (Buffer.byteLength(fileName, "utf8") > 240) {
    response.status(400).json({ error: "اسم الملف طويل جداً." });
    return;
  }
  try {
    const pool = await getDataStore();
    const id = randomUUID();
    await pool.query(
      `INSERT INTO user_files (id, owner, file_name, mime_type, size, data)
       VALUES ($1, $2, $3, $4, $5, $6)`,
      [id, request.appUser!.id, fileName, mimeType, body.length, body],
    );
    response.status(201).json({ id, fileName, mimeType, size: body.length });
  } catch (error) {
    logger.error({ err: error }, "User file upload failed");
    response.status(503).json({ error: "تعذر حفظ الملف في قاعدة البيانات." });
  }
});

router.get("/files/:id", requireSession, async (request: AuthenticatedRequest, response) => {
  try {
    const pool = await getDataStore();
    const { rows } = await pool.query<{ file_name: string; mime_type: string; data: Buffer }>(
      "SELECT file_name, mime_type, data FROM user_files WHERE id = $1 AND owner = $2 LIMIT 1",
      [request.params.id, request.appUser!.id],
    );
    const file = rows[0];
    if (!file) {
      response.status(404).json({ error: "الملف غير موجود في مكتبتك." });
      return;
    }
    const encodedName = encodeURIComponent(file.file_name).replace(/['()]/g, escape);
    response.setHeader("Content-Type", file.mime_type);
    response.setHeader("Content-Length", file.data.length);
    response.setHeader("Content-Disposition", `attachment; filename="download"; filename*=UTF-8''${encodedName}`);
    response.setHeader("X-Content-Type-Options", "nosniff");
    response.send(file.data);
  } catch (error) {
    logger.error({ err: error }, "User file download failed");
    response.status(503).json({ error: "تعذر تنزيل الملف من قاعدة البيانات." });
  }
});

router.delete("/files/:id", requireSession, async (request: AuthenticatedRequest, response) => {
  try {
    const pool = await getDataStore();
    const result = await pool.query("DELETE FROM user_files WHERE id = $1 AND owner = $2", [
      request.params.id,
      request.appUser!.id,
    ]);
    if (!result.rowCount) {
      response.status(404).json({ error: "الملف غير موجود في مكتبتك." });
      return;
    }
    response.status(204).end();
  } catch (error) {
    logger.error({ err: error }, "User file delete failed");
    response.status(503).json({ error: "تعذر حذف الملف من قاعدة البيانات." });
  }
});

export default router;
