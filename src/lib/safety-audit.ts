export type SafetyAuditKind = "blocked" | "safe";

export interface SafetyAuditEntry {
  id: string;
  kind: SafetyAuditKind;
  occurredAt: string;
  screen: string;
  action: string;
  detail?: string;
  reviewedAt?: string;
}

export const SAFETY_LOCK_KEY = "recon_safety_lock";
export const SAFETY_AUDIT_KEY = "recon_safety_audit";
export const SAFETY_PASSWORDS_KEY = "recon_safety_passwords";
export const SAFETY_AUDIT_LIMIT = 5000;

type SafetyPasswordPurpose = "activation" | "unlock";

interface SafetyPasswords {
  activation?: string;
  unlock?: string;
}

function accountPassword(): string {
  try {
    const raw = localStorage.getItem("recon_credentials");
    const credentials = raw ? JSON.parse(raw) as { password?: unknown } : null;
    return typeof credentials?.password === "string" ? credentials.password : "admin123";
  } catch {
    return "admin123";
  }
}

function readSafetyPasswords(): SafetyPasswords {
  try {
    const raw = localStorage.getItem(SAFETY_PASSWORDS_KEY);
    const passwords = raw ? JSON.parse(raw) as SafetyPasswords : {};
    return {
      activation: typeof passwords.activation === "string" ? passwords.activation : undefined,
      unlock: typeof passwords.unlock === "string" ? passwords.unlock : undefined,
    };
  } catch {
    return {};
  }
}

export function saveSafetyPassword(purpose: SafetyPasswordPurpose, password: string): void {
  if (!password) throw new Error("أدخل كلمة مرور غير فارغة.");

  try {
    const raw = localStorage.getItem(SAFETY_PASSWORDS_KEY);
    const passwords = raw ? JSON.parse(raw) as SafetyPasswords : {};
    localStorage.setItem(SAFETY_PASSWORDS_KEY, JSON.stringify({ ...passwords, [purpose]: password }));
  } catch (error) {
    console.error("تعذر حفظ كلمة مرور خط الأمان:", error);
    throw new Error("تعذر حفظ كلمة المرور في هذا المتصفح.");
  }
}

export function readSafetyAudit(): SafetyAuditEntry[] {
  try {
    const value = JSON.parse(localStorage.getItem(SAFETY_AUDIT_KEY) || "[]");
    if (!Array.isArray(value)) return [];
    return value.filter((entry): entry is SafetyAuditEntry =>
      entry && typeof entry.id === "string" &&
      (entry.kind === "blocked" || entry.kind === "safe") &&
      typeof entry.occurredAt === "string" &&
      typeof entry.screen === "string" &&
      typeof entry.action === "string"
    );
  } catch {
    return [];
  }
}

export function writeSafetyAudit(entries: SafetyAuditEntry[]): void {
  try {
    localStorage.setItem(SAFETY_AUDIT_KEY, JSON.stringify(entries));
  } catch (error) {
    console.error("تعذر حفظ سجل خط الأمان:", error);
  }
}

export function verifySafetyActivationPassword(password: string): boolean {
  return password === (readSafetyPasswords().activation ?? accountPassword());
}

export function verifySafetyUnlockPassword(password: string): boolean {
  return password === (readSafetyPasswords().unlock ?? accountPassword());
}
