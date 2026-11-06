"use client"

import { useCallback, useEffect, useRef, useState } from "react"
import { BellRing, ChevronLeft } from "lucide-react"
import { LoginScreen } from "./login-screen"
import { AppSidebar, type ViewId } from "./app-sidebar"
import { ProfileView } from "./profile-view"
import { SettingsView } from "./settings-view"
import { UserGuideView } from "./user-guide-view"
import { SafetyLockOverlay, type BlockedAction } from "./safety-lock-overlay"
import App from "./reconciliation-app"
import { AccountStatementProcessor } from "./account-statement-processor"
import { DocumentStudio } from "./document-studio"
import { UserFilesView } from "./user-files-view"
import { DataIngestionHub } from "./data-ingestion-hub"
import { WeeklyReportDashboard } from "./weekly-report-dashboard"
import { AdminUsersView } from "./admin-users-view"
import { CustomerBalances, type MissingCustomerPolicy } from "./customer-balances"
import { CustomerNotes } from "./customer-notes"
import { FuelManagement } from "./fuel-management"
import { EmployeePayroll } from "./employee-payroll"
import { TransfersManagement } from "./transfers-management"
import { CashboxFinance } from "./cashbox-finance"
import { SalesInvoice } from "./sales-invoice"
import { AttendanceManagement } from "./attendance-management"
import { useGlobalLoading } from "./global-loading"
import { readSafetyAudit, SAFETY_AUDIT_LIMIT, SAFETY_LOCK_KEY, type SafetyAuditEntry, verifySafetyActivationPassword, verifySafetyUnlockPassword, writeSafetyAudit } from "../lib/safety-audit"
import { centralLogout, centralSession, centralStorageEnabled, centralStorageGet, centralStorageSet, type CentralUser } from "../lib/central-storage"

const VIEW_LABELS: Record<ViewId, string> = {
  home: "الصفحة الرئيسية",
  fuel: "منصة المحروقات",
  payroll: "كشف رواتب الموظفين",
  attendance: "الحضور والانصراف",
  transfers: "منصة الحوالات",
  cashbox: "الصندوق المالي",
  salesInvoice: "منصة الفواتير",
  customerBalances: "أرصدة الزبائن",
  customerNotes: "ملاحظات وتذكيرات",
  documentStudio: "المستندات الذكية",
  recon: "أدوات التسوية",
  assistant: "مساعد المطابقة",
  manual: "المطابقة اليدوية",
  report: "تقرير الفروقات",
  weeklyReport: "التقرير الأسبوعي الشامل",
  splitTool: "تسوية حوالات وفواتير متعددة",
  batchCutoff: "تجميع فترة التعطل",
  clearing: "تسكير السندات",
  processor: "معالج الكشوفات",
  guide: "دليل المستخدم",
  files: "ملفاتي",
  ingestion: "مركز رفع البيانات",
  adminUsers: "إدارة المستخدمين",
  profile: "الملف الشخصي",
  settings: "الإعدادات",
}

function removeTypedSafetyCommandPrefix(target: EventTarget | null, typedPrefix: string) {
  if (!typedPrefix) return

  if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) {
    if (!target.value.endsWith(typedPrefix)) return

    const valueSetter = Object.getOwnPropertyDescriptor(
      Object.getPrototypeOf(target),
      "value",
    )?.set
    valueSetter?.call(target, target.value.slice(0, -typedPrefix.length))
    target.dispatchEvent(new Event("input", { bubbles: true }))
  }
}

export function AppShell() {
  const [ready, setReady] = useState(false)
  const [username, setUsername] = useState<string | null>(null)
  const [isAdmin, setIsAdmin] = useState(false)
  const [centralAuthError, setCentralAuthError] = useState("")
  const [centralStorageError, setCentralStorageError] = useState("")
  const [customerPolicyError, setCustomerPolicyError] = useState("")
  const [view, setView] = useState<ViewId>("recon")
  const [customerReminderCount, setCustomerReminderCount] = useState(0)
  const [memoReminderCount, setMemoReminderCount] = useState(0)
  const [customerMissingPolicy, setCustomerMissingPolicy] = useState<MissingCustomerPolicy>("keep_zero")
  const [customerPolicyReady, setCustomerPolicyReady] = useState(false)
  const [collapsed, setCollapsed] = useState(false)
  const [hovering, setHovering] = useState(false)
  const [safetyLocked, setSafetyLocked] = useState(false)
  const [auditEntries, setAuditEntries] = useState<SafetyAuditEntry[]>([])
  const auditEntriesRef = useRef<SafetyAuditEntry[]>([])
  const safetyPasswordSequence = useRef({ value: "", lastKeyAt: 0 })
  const customerPolicyStorageKey = `customer_balances_missing_policy:${(username ?? "").normalize("NFKC").trim().toLocaleLowerCase("en-US")}`

  useGlobalLoading(!ready, "جاري تجهيز التطبيق...")
  useGlobalLoading(
    ready && view === "settings" && !customerPolicyReady && !customerPolicyError,
    "جاري تحميل إعدادات الحساب...",
  )

  useEffect(() => {
    if (!username) {
      setCustomerPolicyReady(false)
      return
    }
    let active = true
    setCustomerPolicyReady(false)
    const loadPolicy = async () => {
      try {
        const stored = centralStorageEnabled
          ? await centralStorageGet<MissingCustomerPolicy>(customerPolicyStorageKey)
          : localStorage.getItem(customerPolicyStorageKey)
        if (!active) return
        setCustomerMissingPolicy(stored === "hide_keep_history" ? "hide_keep_history" : "keep_zero")
      } catch (cause) {
        if (active) {
          setCustomerMissingPolicy("keep_zero")
          setCustomerPolicyError(cause instanceof Error ? `تعذر تحميل إعدادات أرصدة الزبائن: ${cause.message}` : "تعذر تحميل إعدادات أرصدة الزبائن.")
        }
      } finally {
        if (active) setCustomerPolicyReady(true)
      }
    }
    void loadPolicy()
    return () => { active = false }
  }, [customerPolicyStorageKey, username])

  const changeCustomerMissingPolicy = async (policy: MissingCustomerPolicy) => {
    try {
      if (centralStorageEnabled) {
        await centralStorageSet(customerPolicyStorageKey, policy)
      } else {
        localStorage.setItem(customerPolicyStorageKey, policy)
      }
      setCustomerMissingPolicy(policy)
      setCustomerPolicyError("")
    } catch (cause) {
      setCustomerPolicyError(cause instanceof Error ? `تعذر حفظ إعدادات أرصدة الزبائن: ${cause.message}` : "تعذر حفظ إعدادات أرصدة الزبائن.")
    }
  }

  const appendAuditEntries = (entries: Array<Omit<SafetyAuditEntry, "id" | "occurredAt">>) => {
    if (!entries.length) return
    const stamped = entries.map(entry => ({
      ...entry,
      id: `audit-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 9)}`,
      occurredAt: new Date().toISOString(),
    }))
    const combined = [...stamped, ...auditEntriesRef.current]
    const pending = combined.filter(entry => entry.kind === "blocked" && !entry.reviewedAt)
    const recent = combined
      .filter(entry => entry.kind !== "blocked" || !!entry.reviewedAt)
      .slice(0, Math.max(0, SAFETY_AUDIT_LIMIT - pending.length))
    const next = [...pending, ...recent].sort((a, b) => b.occurredAt.localeCompare(a.occurredAt))
    auditEntriesRef.current = next
    setAuditEntries(next)
    writeSafetyAudit(next)
  }

  const recordBlockedAction = (attempt: BlockedAction) => {
    appendAuditEntries([{ ...attempt, kind: "blocked" }])
  }

  const recordSafeActivities = (activities: string[], screen: string) => {
    appendAuditEntries(activities.map(detail => ({
      kind: "safe" as const,
      screen,
      action: "مطابقة آمنة محفوظة",
      detail,
    })))
  }

  useEffect(() => {
    let active = true
    const initialize = async () => {
      if (centralStorageEnabled) {
        try {
          const user = await centralSession()
          if (active) {
            setUsername(user?.username ?? null)
            setIsAdmin(user?.role === "admin")
          }
        } catch (cause) {
          if (active) setCentralAuthError(cause instanceof Error ? cause.message : "تعذر الاتصال بخادم تسجيل الدخول.")
        }
      } else {
        try {
          const raw = localStorage.getItem("recon_auth")
          if (raw) {
            setUsername((JSON.parse(raw) as { username: string }).username)
            setIsAdmin(true)
          }
        } catch {
          /* ignore */
        }
      }
      try {
        setCollapsed(localStorage.getItem("recon_sidebar_collapsed") === "1")
      } catch {
        /* ignore */
      }
      const savedAudit = readSafetyAudit()
      auditEntriesRef.current = savedAudit
      setAuditEntries(savedAudit)
      try {
        setSafetyLocked(localStorage.getItem(SAFETY_LOCK_KEY) === "1")
      } catch {
        setSafetyLocked(false)
      }
      if (active) setReady(true)
    }
    void initialize()
    return () => { active = false }
  }, [])

  useEffect(() => {
    const showStorageError = (event: Event) => {
      const detail = (event as CustomEvent<string>).detail
      setCentralStorageError(detail || "تعذر حفظ البيانات في قاعدة البيانات المركزية.")
    }
    window.addEventListener("central-storage-error", showStorageError)
    const clearStorageError = () => setCentralStorageError("")
    window.addEventListener("central-storage-recovered", clearStorageError)
    return () => {
      window.removeEventListener("central-storage-error", showStorageError)
      window.removeEventListener("central-storage-recovered", clearStorageError)
    }
  }, [])

  useEffect(() => {
    if (!auditEntries.some(entry => entry.kind === "blocked" && !entry.reviewedAt)) return
    const remindBeforeClose = (event: BeforeUnloadEvent) => {
      event.preventDefault()
      event.returnValue = "توجد حركات غير آمنة لم تتم مراجعتها."
    }
    window.addEventListener("beforeunload", remindBeforeClose)
    return () => window.removeEventListener("beforeunload", remindBeforeClose)
  }, [auditEntries])

  const activateSafetyLock = useCallback(() => {
    try {
      localStorage.setItem(SAFETY_LOCK_KEY, "1")
    } catch {
      /* Keep the in-memory lock active if storage is unavailable. */
    }
    setSafetyLocked(true)
  }, [])

  useEffect(() => {
    if (!ready || !username || safetyLocked) {
      safetyPasswordSequence.current = { value: "", lastKeyAt: 0 }
      return
    }

    const handleKeyDown = (event: KeyboardEvent) => {
      const target = event.target
      if (target instanceof HTMLElement && target.closest("[data-safety-unlock]")) return
      if (event.ctrlKey || event.altKey || event.metaKey || event.key.length !== 1) {
        safetyPasswordSequence.current = { value: "", lastKeyAt: 0 }
        return
      }

      const now = Date.now()
      const previous = safetyPasswordSequence.current
      const value = now - previous.lastKeyAt > 15_000 ? event.key : `${previous.value}${event.key}`.slice(-128)
      safetyPasswordSequence.current = { value, lastKeyAt: now }

      for (let start = 0; start < value.length; start++) {
        const candidate = value.slice(start)
        if (!candidate.startsWith("@") || !verifySafetyActivationPassword(candidate.slice(1))) continue
        event.preventDefault()
        event.stopPropagation()
        event.stopImmediatePropagation()

        const typedPrefix = candidate.slice(0, -event.key.length)
        removeTypedSafetyCommandPrefix(target, typedPrefix)

        safetyPasswordSequence.current = { value: "", lastKeyAt: 0 }
        activateSafetyLock()
        return
      }
    }

    const handlePaste = (event: ClipboardEvent) => {
      if (event.target instanceof HTMLElement && event.target.closest("[data-safety-unlock]")) return
      const pastedText = event.clipboardData?.getData("text") ?? ""
      if (!pastedText.startsWith("@") || !verifySafetyActivationPassword(pastedText.slice(1))) return
      event.preventDefault()
      event.stopPropagation()
      event.stopImmediatePropagation()
      safetyPasswordSequence.current = { value: "", lastKeyAt: 0 }
      activateSafetyLock()
    }

    window.addEventListener("keydown", handleKeyDown, true)
    window.addEventListener("paste", handlePaste, true)
    return () => {
      window.removeEventListener("keydown", handleKeyDown, true)
      window.removeEventListener("paste", handlePaste, true)
    }
  }, [ready, username, safetyLocked, activateSafetyLock])

  const unlockSafety = (password: string) => {
    if (!password.startsWith("#") || !verifySafetyUnlockPassword(password.slice(1))) {
      recordBlockedAction({ screen: "خط الأمان", action: "محاولة فك القفل", detail: "كلمة المرور غير صحيحة." })
      return false
    }
    try {
      localStorage.removeItem(SAFETY_LOCK_KEY)
    } catch {
      /* The in-memory lock can still be removed for this session. */
    }
    setSafetyLocked(false)
    return true
  }

  const changeLoginPassword = (password: string) => {
    try {
      const raw = localStorage.getItem("recon_credentials")
      const stored = raw ? JSON.parse(raw) as { username?: unknown; password?: unknown } : null
      const credentials = {
        username: typeof stored?.username === "string" ? stored.username : username ?? "admin",
        password,
      }
      localStorage.setItem("recon_credentials", JSON.stringify(credentials))
    } catch (error) {
      console.error("تعذر تغيير كلمة مرور تسجيل الدخول:", error)
      throw new Error("تعذر حفظ كلمة المرور في هذا المتصفح.")
    }
  }

  const markAuditReviewed = () => {
    const reviewedAt = new Date().toISOString()
    const next = auditEntriesRef.current.map(entry =>
      entry.kind === "blocked" && !entry.reviewedAt ? { ...entry, reviewedAt } : entry
    )
    auditEntriesRef.current = next
    setAuditEntries(next)
    writeSafetyAudit(next)
  }

  const toggleCollapsed = () => {
    setCollapsed((prev) => {
      const next = !prev
      try {
        localStorage.setItem("recon_sidebar_collapsed", next ? "1" : "0")
      } catch {
        /* ignore */
      }
      if (next) setHovering(false)
      return next
    })
  }

  const handleLogout = () => {
    if (centralStorageEnabled) {
      void centralLogout().catch(error => console.error("تعذر إنهاء جلسة الدخول المركزية:", error))
    } else {
      localStorage.removeItem("recon_auth")
    }
    setUsername(null)
    setIsAdmin(false)
    setView("recon")
  }

  if (!ready) return null
  if (!username) return <LoginScreen onLogin={(user: CentralUser) => { setCentralAuthError(""); setUsername(user.username); setIsAdmin(user.role === "admin") }} initialError={centralAuthError} />
  const sidebarVisible = !collapsed || hovering
  const pendingAuditCount = auditEntries.filter(entry => entry.kind === "blocked" && !entry.reviewedAt).length
  const onSafeActivities = (activities: string[]) => recordSafeActivities(activities, VIEW_LABELS[view])

  return (
    <div dir="rtl" className="app-shell flex h-[100dvh] max-h-[100dvh] overflow-hidden bg-background">
      {centralStorageError && (
        <div role="alert" className="fixed bottom-4 left-4 z-[100] flex max-w-xl items-center gap-3 rounded-xl border border-red-300 bg-red-50 px-4 py-3 text-sm text-red-900 shadow-xl">
          <span>{centralStorageError} لم يتم اعتماد الحفظ؛ أعد المحاولة بعد إصلاح الاتصال.</span>
          <button type="button" onClick={() => setCentralStorageError("")} className="shrink-0 rounded-lg border border-red-200 px-2 py-1 text-xs font-semibold hover:bg-red-100">إغلاق</button>
        </div>
      )}
      {collapsed && (
        <>
          <div className="fixed inset-y-0 right-0 z-30 w-4" onMouseEnter={() => setHovering(true)} />
          {!hovering && (
            <button
              type="button"
              onClick={() => setHovering(true)}
              aria-label="إظهار القائمة"
              title="إظهار القائمة"
              className="fixed right-0 top-1/2 z-30 -translate-y-1/2 rounded-l-lg border border-sidebar-border bg-sidebar px-1 py-3 text-sidebar-foreground shadow-md transition-opacity hover:opacity-100"
            >
              <ChevronLeft className="h-4 w-4" />
            </button>
          )}
        </>
      )}
      <div
        onMouseEnter={() => collapsed && setHovering(true)}
        onMouseLeave={() => collapsed && setHovering(false)}
        className={collapsed
          ? `fixed inset-y-0 right-0 z-40 transition-transform duration-200 ${sidebarVisible ? "translate-x-0" : "translate-x-full"}`
          : "relative shrink-0"}
      >
        <AppSidebar
          active={view}
          onNavigate={setView}
          username={username}
          isAdmin={isAdmin}
          centralStorage={centralStorageEnabled}
          onLogout={handleLogout}
          collapsed={collapsed}
          onToggleCollapse={toggleCollapsed}
        />
      </div>
      <div className="recon-scope flex-1 overflow-y-auto">
        {view === "home" && <App key="home" initialPage="main" onSafeActivities={onSafeActivities} />}
        {view === "fuel" && <FuelManagement username={username ?? ""} />}
        {view === "payroll" && <EmployeePayroll username={username ?? ""} />}
        {view === "attendance" && centralStorageEnabled && <AttendanceManagement username={username ?? ""} onNavigateToPayroll={() => setView("payroll")} />}
        {view === "transfers" && <TransfersManagement username={username ?? ""} />}
        {view === "cashbox" && <CashboxFinance username={username ?? ""} />}
        {view === "salesInvoice" && <SalesInvoice username={username ?? ""} />}
        <CustomerBalances
          onDueRemindersChange={setCustomerReminderCount}
          remindersEnabled={!safetyLocked}
          visible={view === "customerBalances"}
          username={username ?? ""}
          missingCustomerPolicy={customerMissingPolicy}
          missingCustomerPolicyReady={customerPolicyReady}
        />
        <CustomerNotes
          username={username ?? ""}
          visible={view === "customerNotes"}
          onDueRemindersChange={setMemoReminderCount}
        />
        {view === "documentStudio" && <DocumentStudio />}
        {view === "files" && centralStorageEnabled && <UserFilesView />}
        {view === "ingestion" && centralStorageEnabled && <DataIngestionHub />}
        {view === "weeklyReport" && centralStorageEnabled && <WeeklyReportDashboard />}
        {view === "adminUsers" && isAdmin && centralStorageEnabled && <AdminUsersView />}
        {view === "recon" && <App key="recon" initialPage="recon2" onSafeActivities={onSafeActivities} />}
        {view === "assistant" && <App key="assistant" initialPage="assist" onSafeActivities={onSafeActivities} />}
        {view === "manual" && <App key="manual" initialPage="manual" onSafeActivities={onSafeActivities} />}
        {view === "report" && <App key="report" initialPage="report" onSafeActivities={onSafeActivities} />}
        {view === "splitTool" && <App key="splitTool" initialPage="splitTool" onSafeActivities={onSafeActivities} />}
        {view === "batchCutoff" && <App key="batchCutoff" initialPage="batchCutoff" onSafeActivities={onSafeActivities} />}
        {view === "clearing" && <App key="clearing" initialPage="clearing" onSafeActivities={onSafeActivities} />}
        {view === "processor" && <AccountStatementProcessor />}
        {view === "guide" && <UserGuideView onNavigate={setView} />}
        {view === "profile" && <ProfileView username={username} onUsernameChange={setUsername} />}
        {view === "settings" && <SettingsView
          safetyLocked={safetyLocked}
          centralStorage={centralStorageEnabled}
          auditEntries={auditEntries}
          onActivateSafetyLock={activateSafetyLock}
          onMarkAuditReviewed={markAuditReviewed}
          onChangeLoginPassword={changeLoginPassword}
          customerMissingPolicy={customerMissingPolicy}
          customerPolicyReady={customerPolicyReady}
          customerPolicyError={customerPolicyError}
          onCustomerMissingPolicyChange={policy => void changeCustomerMissingPolicy(policy)}
          username={username}
        />}
      </div>
      {pendingAuditCount > 0 && !safetyLocked && (
        <button
          type="button"
          onClick={() => setView("settings")}
          className="fixed left-4 top-4 z-50 flex items-center gap-2 rounded-lg border border-red-300 bg-red-50 px-3 py-2 text-xs font-semibold text-red-800 shadow-lg hover:bg-red-100"
        >
          <span className="h-2 w-2 rounded-full bg-red-600" />
          {pendingAuditCount} محاولة غير آمنة بانتظار المراجعة
        </button>
      )}
      {customerReminderCount > 0 && view !== "customerBalances" && !safetyLocked && (
        <button
          type="button"
          onClick={() => setView("customerBalances")}
          aria-label={`${customerReminderCount.toLocaleString("ar")} تذكير مستحق لأرصدة الزبائن`}
          title={`${customerReminderCount.toLocaleString("ar")} تذكير مستحق لأرصدة الزبائن`}
          className="group fixed right-4 top-4 z-50 flex h-11 w-11 items-center justify-center overflow-hidden whitespace-nowrap rounded-full border border-amber-300 bg-amber-50 text-amber-900 shadow-lg transition-[width,border-radius,background-color] duration-300 ease-out hover:w-80 hover:justify-start hover:rounded-lg hover:bg-amber-100 focus-visible:w-80 focus-visible:justify-start focus-visible:rounded-lg focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-amber-500"
        >
          <span className="relative flex h-11 w-11 shrink-0 items-center justify-center">
            <BellRing className="h-5 w-5" />
            <span className="absolute -left-0.5 -top-0.5 flex h-4 min-w-4 items-center justify-center rounded-full bg-amber-700 px-0.5 text-[9px] font-bold leading-none text-white">
              {customerReminderCount.toLocaleString("ar")}
            </span>
          </span>
          <span className="max-w-0 overflow-hidden text-xs font-semibold opacity-0 transition-[max-width,opacity] duration-300 group-hover:mr-1 group-hover:max-w-[280px] group-hover:opacity-100 group-focus-visible:mr-1 group-focus-visible:max-w-[280px] group-focus-visible:opacity-100">
            {customerReminderCount.toLocaleString("ar")} تذكير مستحق لأرصدة الزبائن
          </span>
        </button>
      )}
      {memoReminderCount > 0 && view !== "customerNotes" && !safetyLocked && (
        <button
          type="button"
          onClick={() => setView("customerNotes")}
          aria-label={`${memoReminderCount.toLocaleString("ar")} تذكير مستحق في الملاحظات`}
          title={`${memoReminderCount.toLocaleString("ar")} تذكير مستحق في الملاحظات`}
          className="group fixed right-4 top-16 z-50 flex h-11 w-11 items-center justify-center overflow-hidden whitespace-nowrap rounded-full border border-cyan-300 bg-cyan-50 text-cyan-900 shadow-lg transition-[width,border-radius,background-color] duration-300 ease-out hover:w-80 hover:justify-start hover:rounded-lg hover:bg-cyan-100 focus-visible:w-80 focus-visible:justify-start focus-visible:rounded-lg focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-cyan-500"
        >
          <span className="relative flex h-11 w-11 shrink-0 items-center justify-center">
            <BellRing className="h-5 w-5" />
            <span className="absolute -left-0.5 -top-0.5 flex h-4 min-w-4 items-center justify-center rounded-full bg-cyan-700 px-0.5 text-[9px] font-bold leading-none text-white">
              {memoReminderCount.toLocaleString("ar")}
            </span>
          </span>
          <span className="max-w-0 overflow-hidden text-xs font-semibold opacity-0 transition-[max-width,opacity] duration-300 group-hover:mr-1 group-hover:max-w-[280px] group-hover:opacity-100 group-focus-visible:mr-1 group-focus-visible:max-w-[280px] group-focus-visible:opacity-100">
            {memoReminderCount.toLocaleString("ar")} تذكير مستحق في الملاحظات
          </span>
        </button>
      )}
      {safetyLocked && <SafetyLockOverlay activeScreen={VIEW_LABELS[view]} onAttempt={recordBlockedAction} onUnlock={unlockSafety} />}
    </div>
  )
}
