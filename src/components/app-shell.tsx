"use client"

import { useCallback, useEffect, useRef, useState } from "react"
import { ChevronLeft } from "lucide-react"
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
import { AdminUsersView } from "./admin-users-view"
import { readSafetyAudit, SAFETY_AUDIT_LIMIT, SAFETY_LOCK_KEY, type SafetyAuditEntry, verifySafetyActivationPassword, verifySafetyUnlockPassword, writeSafetyAudit } from "../lib/safety-audit"
import { centralLogout, centralSession, centralStorageEnabled, type CentralUser } from "../lib/central-storage"

const VIEW_LABELS: Record<ViewId, string> = {
  home: "الصفحة الرئيسية",
  documentStudio: "المستندات الذكية",
  recon: "أدوات التسوية",
  assistant: "مساعد المطابقة",
  manual: "المطابقة اليدوية",
  report: "تقرير الفروقات",
  splitTool: "تجزئة الحوالات",
  batchCutoff: "تجميع فترة التعطل",
  clearing: "تسكير السندات",
  processor: "معالج الكشوفات",
  guide: "دليل المستخدم",
  files: "ملفاتي",
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
  const [view, setView] = useState<ViewId>("recon")
  const [collapsed, setCollapsed] = useState(false)
  const [hovering, setHovering] = useState(false)
  const [safetyLocked, setSafetyLocked] = useState(false)
  const [auditEntries, setAuditEntries] = useState<SafetyAuditEntry[]>([])
  const auditEntriesRef = useRef<SafetyAuditEntry[]>([])
  const safetyPasswordSequence = useRef({ value: "", lastKeyAt: 0 })

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
    return () => window.removeEventListener("central-storage-error", showStorageError)
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
        {view === "documentStudio" && <DocumentStudio />}
        {view === "files" && centralStorageEnabled && <UserFilesView />}
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
      {safetyLocked && <SafetyLockOverlay activeScreen={VIEW_LABELS[view]} onAttempt={recordBlockedAction} onUnlock={unlockSafety} />}
    </div>
  )
}
