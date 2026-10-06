"use client"

import { Trash2, Database, ShieldAlert, LockKeyhole, Check, Clock3 } from "lucide-react"
import { useState, type FormEvent } from "react"
import { SAFETY_AUDIT_KEY, SAFETY_LOCK_KEY, SAFETY_PASSWORDS_KEY, saveSafetyPassword, type SafetyAuditEntry } from "../lib/safety-audit"
import { passwordPolicyError } from "../lib/password-policy"

function PasswordChangeForm({
  id,
  title,
  description,
  onSave,
  requireLetterAndNumber = false,
}: {
  id: string
  title: string
  description: string
  onSave: (password: string) => void
  requireLetterAndNumber?: boolean
}) {
  const [password, setPassword] = useState("")
  const [confirmation, setConfirmation] = useState("")
  const [message, setMessage] = useState("")
  const [failed, setFailed] = useState(false)

  const handleSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (!password) {
      setFailed(true)
      setMessage("أدخل كلمة المرور الجديدة.")
      return
    }
    if (requireLetterAndNumber) {
      const passwordError = passwordPolicyError(password)
      if (passwordError) {
        setFailed(true)
        setMessage(passwordError)
        return
      }
    }
    if (password !== confirmation) {
      setFailed(true)
      setMessage("كلمتا المرور غير متطابقتين.")
      return
    }
    try {
      onSave(password)
      setPassword("")
      setConfirmation("")
      setFailed(false)
      setMessage("تم تغيير كلمة المرور.")
    } catch (error) {
      setFailed(true)
      setMessage(error instanceof Error ? error.message : "تعذر تغيير كلمة المرور.")
    }
  }

  return (
    <form onSubmit={handleSubmit} className="rounded-xl border border-border bg-background p-4">
      <h3 className="text-sm font-semibold">{title}</h3>
      <p className="mt-1 text-xs leading-5 text-muted-foreground">{description}</p>
      <div className="mt-3 grid gap-3 sm:grid-cols-2">
        <label htmlFor={`${id}-new`} className="space-y-1.5 text-xs font-medium">
          <span>كلمة المرور الجديدة</span>
          <input
            id={`${id}-new`}
            type="password"
            autoComplete="new-password"
            value={password}
            onChange={event => setPassword(event.target.value)}
            minLength={requireLetterAndNumber ? 8 : undefined}
            maxLength={requireLetterAndNumber ? 256 : undefined}
            className="w-full rounded-lg border border-border bg-card px-3 py-2.5 text-sm outline-none focus:border-primary focus:ring-1 focus:ring-primary"
          />
        </label>
        <label htmlFor={`${id}-confirm`} className="space-y-1.5 text-xs font-medium">
          <span>تأكيد كلمة المرور الجديدة</span>
          <input
            id={`${id}-confirm`}
            type="password"
            autoComplete="new-password"
            value={confirmation}
            onChange={event => setConfirmation(event.target.value)}
            className="w-full rounded-lg border border-border bg-card px-3 py-2.5 text-sm outline-none focus:border-primary focus:ring-1 focus:ring-primary"
          />
        </label>
      </div>
      <div className="mt-3 flex flex-wrap items-center gap-3">
        <button type="submit" className="rounded-lg bg-primary px-4 py-2 text-sm font-semibold text-primary-foreground hover:bg-primary/90">
          تغيير كلمة المرور
        </button>
        {message && <p role={failed ? "alert" : "status"} className={`text-xs ${failed ? "text-destructive" : "text-emerald-700"}`}>{message}</p>}
      </div>
    </form>
  )
}

export function SettingsView({
  safetyLocked,
  auditEntries,
  onActivateSafetyLock,
  onMarkAuditReviewed,
  onChangeLoginPassword,
  centralStorage = false,
}: {
  safetyLocked: boolean
  auditEntries: SafetyAuditEntry[]
  onActivateSafetyLock: () => void
  onMarkAuditReviewed: () => void
  onChangeLoginPassword: (password: string) => void
  centralStorage?: boolean
}) {
  const [cleared, setCleared] = useState(false)
  const pendingEntries = auditEntries.filter(entry => entry.kind === "blocked" && !entry.reviewedAt)
  const blockedEntries = auditEntries.filter(entry => entry.kind === "blocked")
  const safeEntries = auditEntries.filter(entry => entry.kind === "safe")

  const clearReconData = () => {
    const keep = new Set(["recon_auth", "recon_credentials", "recon_sidebar_collapsed", SAFETY_AUDIT_KEY, SAFETY_LOCK_KEY, SAFETY_PASSWORDS_KEY])
    try {
      Object.keys(localStorage)
        .filter((k) => !keep.has(k))
        .forEach((k) => localStorage.removeItem(k))
      indexedDB.deleteDatabase("recon_store")
    } catch {
      /* ignore */
    }
    setCleared(true)
    setTimeout(() => setCleared(false), 2500)
  }

  return (
    <div dir="rtl" className="mx-auto min-h-[100dvh] max-w-3xl p-4 sm:p-8">
      <header className="mb-6">
        <div className="mb-2 text-xs font-semibold text-primary">التحكم والخصوصية</div>
        <h1 className="text-3xl font-bold">الإعدادات</h1>
        <p className="mt-2 text-sm text-muted-foreground">
          {centralStorage ? "بيانات التسوية والمشاريع تحفظ في قاعدة PostgreSQL المركزية." : "إدارة البيانات المحفوظة في هذا المتصفح"}
        </p>
      </header>

      <div className="space-y-4">
        <section className="rounded-xl border border-rose-300 bg-rose-50 p-5">
          <h2 className="mb-2 flex items-center gap-2 text-sm font-semibold text-rose-900">
            <ShieldAlert className="h-4 w-4" /> خط الأمان
          </h2>
          <p className="mb-4 text-xs leading-6 text-rose-800">
            يوقف التفاعل مع كل الشاشات والمنصات. المحاولات أثناء القفل تُسجل للمراجعة ولا تُطبق على ملفاتك أو مطابقتك.
          </p>
          <button
            type="button"
            onClick={onActivateSafetyLock}
            disabled={safetyLocked}
            className="flex items-center gap-2 rounded-lg bg-rose-700 px-4 py-2.5 text-sm font-semibold text-white hover:bg-rose-800 disabled:cursor-not-allowed disabled:opacity-50"
          >
            <LockKeyhole className="h-4 w-4" /> {safetyLocked ? "خط الأمان مفعّل" : "تفعيل خط الأمان"}
          </button>
          <p className="mt-3 text-[11px] leading-5 text-rose-800/80">
            كلمة التفعيل وكلمة فك القفل قابلتان للتغيير بشكل مستقل عن كلمة تسجيل الدخول.
          </p>
        </section>

        <section className="space-y-4 rounded-xl border border-border bg-card p-5">
          <div>
            <h2 className="text-sm font-semibold">إدارة كلمات المرور</h2>
            <p className="mt-1 text-xs leading-5 text-muted-foreground">
              لكل وظيفة كلمة مستقلة. تغيير كلمة الدخول لا يغيّر كلمات الأمان التي تخصصها هنا.
            </p>
          </div>
          {centralStorage
            ? <p className="rounded-xl border border-border bg-background p-4 text-xs leading-6 text-muted-foreground">
              كلمة مرور حسابك محفوظة كتجزئة آمنة في قاعدة البيانات. تغيير كلمة المرور غير متاح من هذه الصفحة.
            </p>
            : <PasswordChangeForm
              id="login-password"
              title="كلمة مرور تسجيل الدخول"
              description="تُستخدم عند تسجيل الدخول إلى المنصة فقط. يجب أن تتكون من 8 محارف على الأقل وتشمل حرفاً ورقماً."
              onSave={onChangeLoginPassword}
              requireLetterAndNumber
            />}
          <PasswordChangeForm
            id="safety-activation-password"
            title="كلمة تفعيل خط الأمان"
            description="اكتب @ ثم هذه الكلمة في حقل نصي لتفعيل القفل؛ كتابة الكلمة وحدها لا تفعّله. تبقى كلمة الدخول الحالية فعالة بعد @ إلى أن تغيّر هذه الكلمة."
            onSave={password => saveSafetyPassword("activation", password)}
          />
          <PasswordChangeForm
            id="safety-unlock-password"
            title="كلمة فك قفل الأمان"
            description="أدخل # ثم هذه الكلمة في شاشة القفل لفكّه. تبقى كلمة الدخول الحالية فعالة بعد # إلى أن تغيّر هذه الكلمة."
            onSave={password => saveSafetyPassword("unlock", password)}
          />
        </section>

        <section className="rounded-xl border border-red-300 bg-white p-5">
          <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
            <div>
              <h2 className="flex items-center gap-2 text-sm font-semibold text-red-800">
                <ShieldAlert className="h-4 w-4" /> مراجعة الحركات غير الآمنة
                <span className="rounded-full bg-red-100 px-2 py-0.5 text-xs text-red-800">{pendingEntries.length}</span>
              </h2>
              <p className="mt-1 text-xs text-muted-foreground">المحاولات المحجوبة مع وقتها والشاشة والإجراء المقصود.</p>
            </div>
            <button
              type="button"
              onClick={onMarkAuditReviewed}
              disabled={!pendingEntries.length}
              className="flex items-center gap-1.5 rounded-lg border border-red-200 px-3 py-2 text-xs font-medium text-red-800 hover:bg-red-50 disabled:cursor-not-allowed disabled:opacity-40"
            >
              <Check className="h-3.5 w-3.5" /> تمت المراجعة
            </button>
          </div>
          <div className="max-h-80 space-y-2 overflow-y-auto">
            {blockedEntries.map(entry => (
              <article key={entry.id} className={`border-r-4 p-3 ${entry.reviewedAt ? "border-slate-300 bg-slate-50" : "border-red-600 bg-red-50"}`}>
                <div className="flex flex-wrap items-center gap-x-3 gap-y-1 text-[11px] text-slate-600">
                  <span className="flex items-center gap-1"><Clock3 className="h-3 w-3" />{new Date(entry.occurredAt).toLocaleString("ar")}</span>
                  <span>{entry.screen}</span>
                  {entry.reviewedAt && <span className="text-emerald-700">تمت المراجعة</span>}
                </div>
                <p className="mt-1 text-xs font-semibold text-red-900">{entry.action}</p>
                {entry.detail && <p className="mt-1 break-words text-[11px] text-slate-600">{entry.detail}</p>}
              </article>
            ))}
            {!blockedEntries.length && <p className="py-5 text-center text-xs text-muted-foreground">لا توجد محاولات مسجلة.</p>}
          </div>
        </section>

        <section className="rounded-xl border border-emerald-200 bg-white p-5">
          <h2 className="mb-3 flex items-center gap-2 text-sm font-semibold text-emerald-800">
            <Check className="h-4 w-4" /> سجل المطابقات الآمنة
          </h2>
          <div className="max-h-72 space-y-2 overflow-y-auto">
            {safeEntries.map(entry => (
              <article key={entry.id} className="border-r-2 border-emerald-500 bg-emerald-50/60 p-3">
                <div className="flex flex-wrap items-center gap-x-3 gap-y-1 text-[11px] text-slate-600">
                  <span className="flex items-center gap-1"><Clock3 className="h-3 w-3" />{new Date(entry.occurredAt).toLocaleString("ar")}</span>
                  <span>{entry.screen}</span>
                </div>
                <p className="mt-1 text-xs font-medium text-emerald-900">{entry.action}</p>
                {entry.detail && <p className="mt-1 break-words text-[11px] text-slate-600">{entry.detail}</p>}
              </article>
            ))}
            {!safeEntries.length && <p className="py-5 text-center text-xs text-muted-foreground">ستظهر هنا المطابقات التي تحفظها أثناء العمل الطبيعي.</p>}
          </div>
        </section>

          <section className="rounded-2xl border border-border bg-card p-6 shadow-sm">
          <h2 className="mb-3 flex items-center gap-2 text-sm font-semibold">
            <Database className="h-4 w-4" />
            البيانات المحفوظة
          </h2>
          <p className="mb-4 text-xs text-muted-foreground">
            {centralStorage
              ? "التخزين المركزي مفعّل. لا تُحذف بيانات PostgreSQL من هذا الخيار."
              : "تحذف عمليات التسوية والملفات المحفوظة محلياً في هذا المتصفح. لن يؤثر ذلك على بيانات الدخول."}
          </p>
          {!centralStorage && <button
            onClick={clearReconData}
            className="flex items-center gap-2 rounded-lg border border-destructive/30 px-4 py-2.5 text-sm font-medium text-destructive transition-colors hover:bg-destructive/10"
          >
            <Trash2 className="h-4 w-4" />
            {cleared ? "تم مسح البيانات" : "مسح بيانات التسوية"}
          </button>}
        </section>
      </div>
    </div>
  )
}
