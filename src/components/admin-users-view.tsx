"use client"

import { useCallback, useEffect, useState } from "react"
import { Check, RefreshCw, UserRoundX, Users } from "lucide-react"
import { listAdminUsers, updateAdminUser, type AdminUser } from "../lib/central-storage"
import { useGlobalLoading } from "./global-loading"

const statusLabels: Record<AdminUser["status"], string> = {
  pending: "بانتظار الموافقة",
  approved: "مفعّل",
  rejected: "مرفوض",
}

export function AdminUsersView() {
  const [users, setUsers] = useState<AdminUser[]>([])
  const [loading, setLoading] = useState(true)
  const [busyId, setBusyId] = useState<string | null>(null)
  const [error, setError] = useState("")
  const [message, setMessage] = useState("")

  useGlobalLoading(loading || busyId !== null, loading ? "جاري تحميل الحسابات..." : "جاري تحديث الحساب...")

  const refresh = useCallback(async () => {
    setLoading(true)
    setError("")
    try {
      setUsers(await listAdminUsers())
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "تعذر تحميل المستخدمين.")
    } finally {
      setLoading(false)
    }
  }, [])

  useEffect(() => { void refresh() }, [refresh])

  const update = async (user: AdminUser, action: "approve" | "reject") => {
    setBusyId(user.id)
    setError("")
    setMessage("")
    try {
      await updateAdminUser(user.id, action)
      setMessage(action === "approve" ? `تم اعتماد حساب ${user.username}.` : `تم رفض حساب ${user.username}.`)
      await refresh()
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "تعذر تحديث الحساب.")
    } finally {
      setBusyId(null)
    }
  }

  return (
    <section className="mx-auto max-w-5xl p-5 sm:p-8" dir="rtl">
      <header className="mb-6 flex flex-wrap items-center justify-between gap-4">
        <div>
          <h1 className="flex items-center gap-2 text-2xl font-bold"><Users className="h-6 w-6" /> إدارة المستخدمين</h1>
          <p className="mt-2 text-sm text-muted-foreground">راجع طلبات التسجيل وفعّل الحسابات التي تسمح لها باستخدام المنصة.</p>
        </div>
        <button type="button" onClick={() => void refresh()} disabled={loading} className="inline-flex items-center gap-2 rounded-xl border px-4 py-2.5 text-sm font-semibold hover:bg-muted">
          <RefreshCw className={`h-4 w-4 ${loading ? "animate-spin" : ""}`} /> تحديث
        </button>
      </header>
      {error && <p role="alert" className="mb-4 rounded-lg bg-destructive/10 px-4 py-3 text-sm text-destructive">{error}</p>}
      {message && <p role="status" className="mb-4 rounded-lg bg-emerald-50 px-4 py-3 text-sm text-emerald-800">{message}</p>}
      <div className="overflow-hidden rounded-2xl border bg-card">
        {loading ? <p className="p-8 text-center text-sm text-muted-foreground">جارٍ تحميل الحسابات...</p> : users.length === 0 ? (
          <p className="p-8 text-center text-sm text-muted-foreground">لا توجد حسابات مسجلة.</p>
        ) : (
          <ul className="divide-y">
            {users.map(user => (
              <li key={user.id} className="flex flex-wrap items-center gap-4 px-4 py-4">
                <div className="min-w-0 flex-1">
                  <div className="font-semibold">{user.username} <span className="mr-2 rounded-full bg-muted px-2 py-1 text-xs font-normal">{statusLabels[user.status]}</span></div>
                  <div className="mt-1 text-sm text-muted-foreground">{user.email || "حساب مسؤول"}</div>
                  <div className="mt-1 text-xs text-muted-foreground">تاريخ التسجيل: {new Date(user.createdAt).toLocaleString("ar")}</div>
                </div>
                {user.role === "user" && (
                  <div className="flex gap-2">
                    {user.status !== "approved" && (
                      <button type="button" onClick={() => void update(user, "approve")} disabled={busyId === user.id} className="inline-flex items-center gap-1.5 rounded-lg bg-emerald-600 px-3 py-2 text-xs font-semibold text-white hover:bg-emerald-700 disabled:opacity-50">
                        <Check className="h-4 w-4" /> اعتماد
                      </button>
                    )}
                    {user.status !== "rejected" && (
                      <button type="button" onClick={() => void update(user, "reject")} disabled={busyId === user.id} className="inline-flex items-center gap-1.5 rounded-lg border border-destructive/30 px-3 py-2 text-xs font-semibold text-destructive hover:bg-destructive/10 disabled:opacity-50">
                        <UserRoundX className="h-4 w-4" /> رفض
                      </button>
                    )}
                  </div>
                )}
              </li>
            ))}
          </ul>
        )}
      </div>
    </section>
  )
}
