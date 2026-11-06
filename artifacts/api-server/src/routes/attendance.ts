import { randomBytes, randomUUID } from "node:crypto";
import express, { Router, type Request, type Response } from "express";
import { logger } from "../lib/logger";
import { getDataStore, hashPassword, requireSession, verifyPassword, type AuthenticatedRequest } from "./app-storage";

const router = Router();
const kioskConfigKey = "attendance:kiosk_config";
const staffKey = "attendance:staff";
const failedPinAttempts = new Map<string, { count: number; resetAt: number }>();
const pinFailureLimit = 10;
const pinFailureWindowMs = 15 * 60 * 1000;
router.use((_request, response, next) => {
  response.setHeader("Cache-Control", "no-store");
  next();
});

type AttendanceStaff = {
  employeeId: string
  employeeNumber: string
  name: string
  pinHash: string
}

type AttendancePunch = {
  id: string
  employeeId: string
  employeeNumber: string
  employeeName: string
  action: "in" | "out"
  timestamp: string
  date: string
  time: string
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function isValidDate(value: unknown): value is string {
  return typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value);
}

function localDateTime(timestamp: Date) {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Asia/Gaza",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hourCycle: "h23",
  }).formatToParts(timestamp);
  const get = (type: string) => parts.find(part => part.type === type)?.value ?? "";
  return {
    date: `${get("year")}-${get("month")}-${get("day")}`,
    time: `${get("hour")}:${get("minute")}:${get("second")}`,
  };
}

function recordsKey(ownerId: string, month: string) {
  return `attendance:records:${ownerId}:${month}`;
}

function normalizeDigits(value: string) {
  return value
    .replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
    .replace(/[۰-۹]/g, digit => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)))
    .normalize("NFKC")
    .trim();
}

function readFailedPinStatus(address: string) {
  const current = failedPinAttempts.get(address);
  if (!current || current.resetAt <= Date.now()) {
    if (failedPinAttempts.size >= 10_000) {
      const now = Date.now();
      for (const [client, entry] of failedPinAttempts) if (entry.resetAt <= now) failedPinAttempts.delete(client);
      if (failedPinAttempts.size >= 10_000) return true;
    }
    failedPinAttempts.set(address, { count: 0, resetAt: Date.now() + pinFailureWindowMs });
    return false;
  }
  return current.count >= pinFailureLimit;
}

function recordFailedPin(address: string) {
  const current = failedPinAttempts.get(address);
  if (current) current.count += 1;
  else failedPinAttempts.set(address, { count: 1, resetAt: Date.now() + pinFailureWindowMs });
}

function clearFailedPin(address: string) {
  failedPinAttempts.delete(address);
}

router.get("/attendance/admin/kiosk", requireSession, async (request: AuthenticatedRequest, response: Response) => {
  try {
    const pool = await getDataStore();
    const { rows } = await pool.query<{ value: unknown }>(
      "SELECT value FROM app_data WHERE owner = $1 AND key = $2 LIMIT 1",
      [request.appUser!.id, kioskConfigKey],
    );
    const token = isRecord(rows[0]?.value) && typeof rows[0].value.token === "string" ? rows[0].value.token : null;
    response.json({ token });
  } catch (error) {
    logger.error({ err: error }, "Attendance kiosk configuration read failed");
    response.status(503).json({ error: "تعذر تحميل إعدادات رابط تسجيل الحضور." });
  }
});

router.post("/attendance/admin/kiosk", requireSession, async (request: AuthenticatedRequest, response: Response) => {
  const token = randomBytes(32).toString("hex");
  const key = `attendance:kiosk:${token}`;
  try {
    const pool = await getDataStore();
    const client = await pool.connect();
    try {
      await client.query("BEGIN");
      await client.query(
        "INSERT INTO app_data (owner, key, value, updated_at) VALUES ($1, $2, '{}'::jsonb, NOW()) ON CONFLICT (owner, key) DO NOTHING",
        [request.appUser!.id, kioskConfigKey],
      );
      const currentResult = await client.query<{ value: unknown }>(
        "SELECT value FROM app_data WHERE owner = $1 AND key = $2 FOR UPDATE",
        [request.appUser!.id, kioskConfigKey],
      );
      const oldToken = isRecord(currentResult.rows[0]?.value) && typeof currentResult.rows[0].value.token === "string"
        ? currentResult.rows[0].value.token
        : null;
      await client.query(
        "INSERT INTO app_data (owner, key, value, updated_at) VALUES ($1, $2, $3::jsonb, NOW())",
        [request.appUser!.id, key, JSON.stringify({ enabled: true })],
      );
      await client.query(
        "UPDATE app_data SET value = $1::jsonb, updated_at = NOW() WHERE owner = $2 AND key = $3",
        [JSON.stringify({ token }), request.appUser!.id, kioskConfigKey],
      );
      if (oldToken) {
        await client.query(
          "DELETE FROM app_data WHERE owner = $1 AND key = $2",
          [request.appUser!.id, `attendance:kiosk:${oldToken}`],
        );
      }
      await client.query("COMMIT");
    } catch (error) {
      await client.query("ROLLBACK").catch(rollbackError => logger.error({ err: rollbackError }, "Attendance kiosk transaction rollback failed"));
      throw error;
    } finally {
      client.release();
    }
    response.json({ token });
  } catch (error) {
    logger.error({ err: error }, "Attendance kiosk configuration update failed");
    response.status(503).json({ error: "تعذر إنشاء رابط تسجيل الحضور." });
  }
});

router.get("/attendance/admin/staff", requireSession, async (request: AuthenticatedRequest, response: Response) => {
  try {
    const pool = await getDataStore();
    const { rows } = await pool.query<{ value: unknown }>(
      "SELECT value FROM app_data WHERE owner = $1 AND key = $2 LIMIT 1",
      [request.appUser!.id, staffKey],
    );
    const staff = isRecord(rows[0]?.value) && Array.isArray(rows[0].value.employees)
      ? rows[0].value.employees.filter(isRecord)
      : [];
    response.json({
      employees: staff.map(employee => ({
        employeeId: employee.employeeId,
        employeeNumber: employee.employeeNumber,
        pinConfigured: typeof employee.pinHash === "string" && employee.pinHash.length > 0,
      })),
    });
  } catch (error) {
    logger.error({ err: error }, "Attendance staff status read failed");
    response.status(503).json({ error: "تعذر تحميل إعدادات أرقام الموظفين السرية." });
  }
});

router.put("/attendance/admin/staff", requireSession, async (request: AuthenticatedRequest, response: Response) => {
  const input = request.body?.employees;
  if (!Array.isArray(input) || input.length > 1000) {
    response.status(400).json({ error: "قائمة الموظفين غير صالحة." });
    return;
  }
  const employeeIds = new Set<string>();
  const employeeNumbers = new Set<string>();
  for (const employee of input) {
    if (!isRecord(employee)
      || typeof employee.employeeId !== "string" || !employee.employeeId.trim() || employee.employeeId.length > 128
      || typeof employee.employeeNumber !== "string" || !employee.employeeNumber.trim() || employee.employeeNumber.length > 64
      || typeof employee.name !== "string" || !employee.name.trim() || employee.name.length > 200
      || (employee.pin !== undefined && employee.pin !== "" && (typeof employee.pin !== "string" || !/^\d{6,12}$/.test(employee.pin)))
      || (employee.clearPin !== undefined && typeof employee.clearPin !== "boolean")) {
      response.status(400).json({ error: "تأكد من رقم واسم كل موظف وأن الرقم السري من 6 إلى 12 رقماً." });
      return;
    }
    const id = employee.employeeId.trim();
    const number = normalizeDigits(employee.employeeNumber as string);
    if (employeeIds.has(id) || employeeNumbers.has(number)) {
      response.status(400).json({ error: "توجد أرقام موظفين أو معرفات مكررة؛ صححها قبل الحفظ." });
      return;
    }
    employeeIds.add(id);
    employeeNumbers.add(number);
  }

  try {
    const pool = await getDataStore();
    const existingResult = await pool.query<{ value: unknown }>(
      "SELECT value FROM app_data WHERE owner = $1 AND key = $2 LIMIT 1",
      [request.appUser!.id, staffKey],
    );
    const existing = isRecord(existingResult.rows[0]?.value) && Array.isArray(existingResult.rows[0].value.employees)
      ? existingResult.rows[0].value.employees.filter(isRecord)
      : [];
    const existingById = new Map(existing.map(employee => [String(employee.employeeId), employee]));
    const employees: AttendanceStaff[] = [];
    for (const rawEmployee of input as Array<Record<string, unknown>>) {
      const employeeId = (rawEmployee.employeeId as string).trim();
      const previous = existingById.get(employeeId);
      let pinHash = rawEmployee.clearPin === true ? "" : typeof previous?.pinHash === "string" ? previous.pinHash : "";
      if (typeof rawEmployee.pin === "string" && rawEmployee.pin) pinHash = await hashPassword(rawEmployee.pin);
      employees.push({
        employeeId,
        employeeNumber: (rawEmployee.employeeNumber as string).trim(),
        name: (rawEmployee.name as string).trim(),
        pinHash,
      });
    }
    await pool.query(
      `INSERT INTO app_data (owner, key, value, updated_at) VALUES ($1, $2, $3::jsonb, NOW())
       ON CONFLICT (owner, key) DO UPDATE SET value = EXCLUDED.value, updated_at = NOW()`,
      [request.appUser!.id, staffKey, JSON.stringify({ employees })],
    );
    response.json({ saved: true, configuredCount: employees.filter(employee => employee.pinHash).length });
  } catch (error) {
    logger.error({ err: error }, "Attendance employee PIN update failed");
    response.status(503).json({ error: "تعذر حفظ أرقام الموظفين السرية." });
  }
});

router.get("/attendance/admin/records", requireSession, async (request: AuthenticatedRequest, response: Response) => {
  const from = request.query.from;
  const to = request.query.to;
  if (!isValidDate(from) || !isValidDate(to) || from > to) {
    response.status(400).json({ error: "حدد فترة زمنية صحيحة لتقرير الحضور." });
    return;
  }
  const fromDate = new Date(`${from}T00:00:00Z`);
  const toDate = new Date(`${to}T00:00:00Z`);
  if (!Number.isFinite(fromDate.getTime()) || !Number.isFinite(toDate.getTime())
    || fromDate.toISOString().slice(0, 10) !== from || toDate.toISOString().slice(0, 10) !== to) {
    response.status(400).json({ error: "أدخل تاريخين صالحين لتقرير الحضور." });
    return;
  }
  if ((toDate.getTime() - fromDate.getTime()) / 86_400_000 > 730) {
    response.status(400).json({ error: "الفترة القصوى لتقرير الحضور سنتان." });
    return;
  }
  const months: string[] = [];
  const cursor = new Date(Date.UTC(fromDate.getUTCFullYear(), fromDate.getUTCMonth(), 1));
  const last = new Date(Date.UTC(toDate.getUTCFullYear(), toDate.getUTCMonth(), 1));
  while (cursor <= last) {
    months.push(`${cursor.getUTCFullYear()}-${String(cursor.getUTCMonth() + 1).padStart(2, "0")}`);
    cursor.setUTCMonth(cursor.getUTCMonth() + 1);
  }
  try {
    const pool = await getDataStore();
    const keys = months.map(month => recordsKey(request.appUser!.id, month));
    const { rows } = await pool.query<{ value: unknown }>(
      "SELECT value FROM app_data WHERE owner = $1 AND key = ANY($2::text[])",
      [request.appUser!.id, keys],
    );
    const records = rows.flatMap(row => Array.isArray(row.value) ? row.value : [])
      .filter((record): record is AttendancePunch => isRecord(record) && typeof record.date === "string" && record.date >= from && record.date <= to) as AttendancePunch[];
    response.json({ records });
  } catch (error) {
    logger.error({ err: error }, "Attendance report read failed");
    response.status(503).json({ error: "تعذر تحميل تقرير الحضور." });
  }
});

const publicPunchBody = express.json({ limit: "8kb" });
router.post("/attendance/public/:token/punch", publicPunchBody, async (request: Request, response: Response) => {
  const address = request.ip || "unknown";
  const tokenParam = request.params.token;
  const token = Array.isArray(tokenParam) ? tokenParam[0] ?? "" : tokenParam;
  const { employeeNumber, pin, action } = request.body ?? {};
  if (!/^[a-f0-9]{64}$/.test(token)
    || typeof employeeNumber !== "string" || employeeNumber.trim().length < 1 || employeeNumber.trim().length > 64
    || typeof pin !== "string" || !/^\d{6,12}$/.test(pin)
    || (action !== "in" && action !== "out")) {
    response.status(400).json({ error: "أدخل رقم الموظف ورقمه السري المكوّن من 6 إلى 12 رقماً، واختر تسجيل الحضور أو الانصراف." });
    return;
  }
  const normalizedEmployeeNumber = normalizeDigits(employeeNumber);
  const normalizedPin = normalizeDigits(pin);
  const failureKey = `${address}:${normalizedEmployeeNumber.toLocaleLowerCase("en-US")}`;
  if (readFailedPinStatus(failureKey)) {
    response.status(429).json({ error: "محاولات كثيرة لهذا الموظف من هذا الاتصال. انتظر 15 دقيقة ثم أعد المحاولة." });
    return;
  }
  try {
    const pool = await getDataStore();
    const { rows: owners } = await pool.query<{ owner: string }>(
      "SELECT owner FROM app_data WHERE key = $1 AND value->>'enabled' = 'true' LIMIT 1",
      [`attendance:kiosk:${token}`],
    );
    const owner = owners[0]?.owner;
    if (!owner) {
      response.status(404).json({ error: "رابط تسجيل الحضور غير صالح أو تم إيقافه." });
      return;
    }
    const { rows: users } = await pool.query<{ username: string }>(
      "SELECT username FROM app_users WHERE id = $1 AND status = 'approved' LIMIT 1",
      [owner],
    );
    const username = users[0]?.username;
    const { rows: staffRows } = await pool.query<{ value: unknown }>(
      "SELECT value FROM app_data WHERE owner = $1 AND key = $2 LIMIT 1",
      [owner, staffKey],
    );
    const employees = isRecord(staffRows[0]?.value) && Array.isArray(staffRows[0].value.employees)
      ? staffRows[0].value.employees.filter(isRecord) as unknown as AttendanceStaff[]
      : [];
    const employee = employees.find(item => normalizeDigits(item.employeeNumber).toLocaleLowerCase("en-US") === normalizedEmployeeNumber.toLocaleLowerCase("en-US") && item.pinHash);
    if (!username || !employee || !(await verifyPassword(normalizedPin, employee.pinHash))) {
      recordFailedPin(failureKey);
      response.status(401).json({ error: "رقم الموظف أو الرقم السري غير صحيح." });
      return;
    }
    clearFailedPin(failureKey);

    const now = new Date();
    const local = localDateTime(now);
    const key = recordsKey(owner, local.date.slice(0, 7));
    const client = await pool.connect();
    let record: AttendancePunch;
    try {
      await client.query("BEGIN");
      await client.query(
        "INSERT INTO app_data (owner, key, value, updated_at) VALUES ($1, $2, '[]'::jsonb, NOW()) ON CONFLICT (owner, key) DO NOTHING",
        [owner, key],
      );
      const existingResult = await client.query<{ value: unknown }>(
        "SELECT value FROM app_data WHERE owner = $1 AND key = $2 FOR UPDATE",
        [owner, key],
      );
      const existing = Array.isArray(existingResult.rows[0]?.value)
        ? existingResult.rows[0].value.filter((item): item is AttendancePunch => isRecord(item) && typeof item.date === "string")
        : [];
      const dayPunches = existing.filter(item => item.employeeId === employee.employeeId && item.date === local.date);
      const lastPunch = dayPunches[dayPunches.length - 1];
      if ((action === "in" && dayPunches.length > 0) || (action === "out" && (!lastPunch || lastPunch.action !== "in"))) {
        await client.query("ROLLBACK");
        response.status(409).json({
          error: action === "in" ? "تم تسجيل حضور هذا الموظف اليوم مسبقاً." : "لا يوجد تسجيل حضور مفتوح لهذا الموظف اليوم.",
        });
        return;
      }
      record = {
        id: randomUUID(),
        employeeId: employee.employeeId,
        employeeNumber: employee.employeeNumber,
        employeeName: employee.name,
        action,
        timestamp: now.toISOString(),
        date: local.date,
        time: local.time,
      };
      await client.query(
        "UPDATE app_data SET value = value || $1::jsonb, updated_at = NOW() WHERE owner = $2 AND key = $3",
        [JSON.stringify([record]), owner, key],
      );
      await client.query("COMMIT");
    } catch (error) {
      await client.query("ROLLBACK");
      throw error;
    } finally {
      client.release();
    }
    response.json({ saved: true, employeeName: employee.name, action, date: local.date, time: local.time });
  } catch (error) {
    logger.error({ err: error }, "Attendance punch submission failed");
    response.status(503).json({ error: "تعذر تسجيل الحركة حالياً. أعد المحاولة بعد قليل." });
  }
});

export default router;
