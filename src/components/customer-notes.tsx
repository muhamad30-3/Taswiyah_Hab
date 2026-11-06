"use client"

import { useEffect, useMemo, useState, type FormEvent } from "react"
import { Bell, Check, Clock3, Plus, Trash2 } from "lucide-react"
import { centralStorageEnabled, centralStorageGet, centralStorageSet } from "../lib/central-storage"
import { useGlobalLoading } from "./global-loading"

type CustomerMemo = {
  id: string
  text: string
  createdAt: string
  reminderAt: string
  completedAt?: string
}

function createId(): string {
  return typeof globalThis.crypto?.randomUUID === "function"
    ? globalThis.crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`
}

function toDateTimeLocal(date: Date): string {
  const local = new Date(date.getTime() - date.getTimezoneOffset() * 60_000)
  return local.toISOString().slice(0, 16)
}

function initialReminderTime(): string {
  return toDateTimeLocal(new Date(Date.now() + 60 * 60_000))
}

export function CustomerNotes({
  username,
  visible = true,
  onDueRemindersChange,
}: {
  username: string
  visible?: boolean
  onDueRemindersChange?: (count: number) => void
}) {
  const normalizedUsername = username.normalize("NFKC").trim().toLocaleLowerCase("en-US")
  const storageKey = `customer_memos:${normalizedUsername}`
  const [memos, setMemos] = useState<CustomerMemo[]>([])
  const [text, setText] = useState("")
  const [reminderAt, setReminderAt] = useState(initialReminderTime)
  const [readyKey, setReadyKey] = useState<string | null>(null)
  const [hasChanges, setHasChanges] = useState(false)
  const [filter, setFilter] = useState<"open" | "all" | "completed">("open")
  const [error, setError] = useState("")
  const [now, setNow] = useState(() => new Date())

  useGlobalLoading(visible && readyKey !== storageKey && !error, "جاري تحميل الملاحظات...")

  useEffect(() => {
    let active = true
    setReadyKey(null)
    setMemos([])
    setHasChanges(false)
    const load = async () => {
      try {
        const stored = centralStorageEnabled
          ? await centralStorageGet<CustomerMemo[]>(storageKey) ?? []
          : JSON.parse(localStorage.getItem(storageKey) ?? "[]") as CustomerMemo[]
        if (!Array.isArray(stored)) throw new Error("قائمة الملاحظات المحفوظة غير صالحة.")
        if (active) {
          setMemos(stored)
          setReadyKey(storageKey)
        }
      } catch (cause) {
        if (active) setError(cause instanceof Error ? `تعذر تحميل الملاحظات: ${cause.message}` : "تعذر تحميل الملاحظات.")
      }
    }
    void load()
    return () => { active = false }
  }, [storageKey])

  useEffect(() => {
    if (readyKey !== storageKey || !hasChanges) return
    let active = true
    const timer = window.setTimeout(() => {
      const save = async () => {
        try {
          if (centralStorageEnabled) {
            await centralStorageSet(storageKey, memos)
          } else {
            localStorage.setItem(storageKey, JSON.stringify(memos))
          }
          if (active) {
            setHasChanges(false)
            setError("")
          }
        } catch (cause) {
          if (active) setError(cause instanceof Error ? `تعذر حفظ الملاحظات: ${cause.message}` : "تعذر حفظ الملاحظات.")
        }
      }
      void save()
    }, 250)
    return () => {
      active = false
      window.clearTimeout(timer)
    }
  }, [hasChanges, memos, readyKey, storageKey])

  useEffect(() => {
    const refresh = () => setNow(new Date())
    const timer = window.setInterval(refresh, 30_000)
    window.addEventListener("focus", refresh)
    return () => {
      window.clearInterval(timer)
      window.removeEventListener("focus", refresh)
    }
  }, [])

  const dueMemos = useMemo(() => memos.filter(memo =>
    !memo.completedAt && !!memo.reminderAt && new Date(memo.reminderAt).getTime() <= now.getTime(),
  ), [memos, now])
  const visibleMemos = useMemo(() => memos.filter(memo =>
    filter === "all" || (filter === "completed" ? !!memo.completedAt : !memo.completedAt),
  ), [filter, memos])

  useEffect(() => {
    onDueRemindersChange?.(dueMemos.length)
  }, [dueMemos.length, onDueRemindersChange])

  const addMemo = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (!text.trim()) return
    if (!reminderAt || !Number.isFinite(new Date(reminderAt).getTime())) {
      setError("حدد وقتاً صالحاً للتذكير.")
      return
    }
    setMemos(current => [{
      id: createId(),
      text: text.trim(),
      createdAt: new Date().toISOString(),
      reminderAt,
    }, ...current])
    setHasChanges(true)
    setText("")
    setReminderAt(initialReminderTime())
    setError("")
  }

  const completeMemo = (memoId: string) => {
    setMemos(current => current.map(memo => memo.id === memoId
      ? { ...memo, completedAt: new Date().toISOString() }
      : memo))
    setHasChanges(true)
  }

  const deleteMemo = (memoId: string) => {
    setMemos(current => current.filter(memo => memo.id !== memoId))
    setHasChanges(true)
  }

  if (!visible) return null

  return (
    <main dir="rtl" className="min-h-full bg-[#f4f7fb] px-4 py-7 text-slate-900 sm:px-7 lg:px-10">
      <div className="mx-auto max-w-5xl space-y-6">
        <header className="rounded-3xl bg-[#10243c] px-6 py-8 text-white shadow-xl sm:px-9">
          <div className="inline-flex items-center gap-2 rounded-full border border-white/15 bg-white/5 px-3 py-1.5 text-xs text-cyan-100"><Bell className="h-3.5 w-3.5" />ملاحظات وتذكيرات شخصية</div>
          <h1 className="mt-3 text-2xl font-extrabold sm:text-3xl">منصة الملاحظات</h1>
          <p className="mt-2 text-sm leading-6 text-slate-300">سجل ما تريد متابعته وحدد وقت التذكير. تبقى ملاحظاتك محفوظة لحسابك حتى مع تحديث أو حذف كشف الأرصدة.</p>
          <p className="mt-3 text-xs text-cyan-100">الوقت الحالي: {now.toLocaleString("ar", { dateStyle: "medium", timeStyle: "short" })}</p>
        </header>

        {error && <div role="alert" className="rounded-xl border border-rose-200 bg-rose-50 p-3 text-sm text-rose-800">{error}</div>}

        <form onSubmit={addMemo} className="space-y-4 rounded-3xl border border-slate-200 bg-white p-5 shadow-sm sm:p-7">
          <label className="block space-y-1.5 text-sm font-semibold"><span>الملاحظة أو المطلوب</span><textarea required rows={3} value={text} onChange={event => setText(event.target.value)} placeholder="مثال: الاتصال بالزبون أو إرسال كشف الحساب..." className="w-full resize-y rounded-xl border border-slate-200 px-3 py-2.5 text-sm font-normal outline-none focus:border-cyan-500 focus:ring-2 focus:ring-cyan-100" /></label>
          <div className="flex flex-wrap items-end gap-3">
            <label className="min-w-[230px] flex-1 space-y-1.5 text-xs font-semibold text-slate-700"><span>موعد «ذكرني»</span><input type="datetime-local" required value={reminderAt} onChange={event => setReminderAt(event.target.value)} className="w-full rounded-xl border border-slate-200 px-3 py-2.5 text-sm font-normal" /></label>
            <button type="button" onClick={() => setReminderAt(initialReminderTime())} className="inline-flex items-center gap-2 rounded-xl border border-cyan-200 px-3 py-2.5 text-xs font-semibold text-cyan-800 hover:bg-cyan-50"><Clock3 className="h-4 w-4" />ذكرني كمان ساعة</button>
            <button type="submit" disabled={readyKey !== storageKey} className="inline-flex items-center gap-2 rounded-xl bg-cyan-800 px-4 py-2.5 text-sm font-semibold text-white hover:bg-cyan-900 disabled:opacity-50"><Plus className="h-4 w-4" />حفظ الملاحظة</button>
          </div>
          <p className="text-[11px] leading-5 text-slate-500">تظهر التذكيرات هنا وفي إشعار المنصة ما دامت مفتوحة. لا يمكن تشغيل تذكير ويب بعد إغلاق المنصة بالكامل.</p>
        </form>

        {dueMemos.length > 0 && <section role="alert" className="space-y-2 rounded-2xl border border-amber-300 bg-amber-50 p-4 text-amber-950">
          <h2 className="flex items-center gap-2 text-sm font-bold"><Bell className="h-4 w-4" />لديك {dueMemos.length.toLocaleString("ar")} تذكير مستحق</h2>
          {dueMemos.map(memo => <p key={memo.id} className="rounded-lg bg-white/80 p-3 text-sm">{memo.text}</p>)}
        </section>}

        <section className="overflow-hidden rounded-3xl border border-slate-200 bg-white shadow-sm">
          <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 p-4">
            <h2 className="font-bold">سجل الملاحظات ({visibleMemos.length.toLocaleString("ar")})</h2>
            <select value={filter} onChange={event => setFilter(event.target.value as typeof filter)} className="rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs"><option value="open">غير المكتملة</option><option value="all">الكل</option><option value="completed">المكتملة</option></select>
          </div>
          {!visibleMemos.length && <p className="p-10 text-center text-sm text-slate-500">{readyKey === storageKey ? "لا توجد ملاحظات في هذا العرض." : "جارٍ تحميل ملاحظاتك..."}</p>}
          <ul className="divide-y divide-slate-100">{visibleMemos.map(memo => <li key={memo.id} className={`flex flex-wrap items-start justify-between gap-3 p-4 ${memo.completedAt ? "bg-slate-50" : ""}`}>
            <div className="min-w-0 flex-1"><p className={`whitespace-pre-wrap text-sm leading-6 ${memo.completedAt ? "text-slate-500 line-through" : "text-slate-800"}`}>{memo.text}</p><p className="mt-2 text-[11px] text-slate-500">أُنشئت: {new Date(memo.createdAt).toLocaleString("ar")}{memo.reminderAt ? ` · التذكير: ${new Date(memo.reminderAt).toLocaleString("ar")}` : ""}{memo.completedAt ? ` · اكتملت: ${new Date(memo.completedAt).toLocaleString("ar")}` : ""}</p></div>
            <div className="flex shrink-0 gap-2">{!memo.completedAt && <button type="button" onClick={() => completeMemo(memo.id)} className="inline-flex items-center gap-1 rounded-lg bg-emerald-700 px-3 py-2 text-xs font-semibold text-white hover:bg-emerald-800"><Check className="h-3.5 w-3.5" />تم</button>}<button type="button" onClick={() => deleteMemo(memo.id)} aria-label="حذف الملاحظة" className="rounded-lg border border-rose-200 p-2 text-rose-700 hover:bg-rose-50"><Trash2 className="h-4 w-4" /></button></div>
          </li>)}</ul>
        </section>
      </div>
    </main>
  )
}
