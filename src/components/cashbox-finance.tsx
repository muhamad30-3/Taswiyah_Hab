"use client"

import { useEffect, useMemo, useState, type FormEvent } from "react"
import { ArrowDownLeft, ArrowUpRight, Banknote, CalendarDays, Check, Pencil, Plus, Search, Trash2, WalletCards, X } from "lucide-react"
import { centralStorageEnabled, centralStorageGet, centralStorageSet, verifyUserPassword } from "../lib/central-storage"
import { useGlobalLoading } from "./global-loading"

type CashBox = {
  id: string
  number: string
  name: string
  openingBalance: number
  balanceAdjustment?: number
  createdAt: string
}

type CashEntry = {
  id: string
  boxId: string
  date: string
  person: string
  debit: number
  credit: number
  note: string
  createdAt: string
}

type CashboxData = { boxes: CashBox[]; entries: CashEntry[] }
type EntryKind = "debit" | "credit"
type ProtectedEntryAction = { kind: "edit" | "delete"; entry: CashEntry }

const EMPTY_DATA: CashboxData = { boxes: [], entries: [] }
const INPUT_CLASS = "w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm outline-none transition focus:border-cyan-600 focus:ring-2 focus:ring-cyan-100"

function createId() {
  return typeof globalThis.crypto?.randomUUID === "function"
    ? globalThis.crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`
}

function roundAmount(value: number) {
  return Math.round((value + Number.EPSILON) * 100) / 100
}

function localDateValue() {
  const now = new Date()
  return new Date(now.getTime() - now.getTimezoneOffset() * 60_000).toISOString().slice(0, 10)
}

function formatAmount(value: number) {
  return new Intl.NumberFormat("ar-PS", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value)
}

function boxBalance(box: CashBox, entries: CashEntry[]) {
  return roundAmount(box.openingBalance + (box.balanceAdjustment ?? 0) + entries.reduce((sum, entry) =>
    entry.boxId === box.id ? sum + entry.debit - entry.credit : sum, 0,
  ))
}

export function CashboxFinance({ username }: { username: string }) {
  const normalizedUsername = username.normalize("NFKC").trim().toLocaleLowerCase("en-US")
  const storageKey = `cashbox_finance:${normalizedUsername}`
  const [data, setData] = useState<CashboxData>(EMPTY_DATA)
  const [loadedKey, setLoadedKey] = useState<string | null>(null)
  const [error, setError] = useState("")
  const [notice, setNotice] = useState("")
  const [busy, setBusy] = useState(false)
  const [selectedBoxId, setSelectedBoxId] = useState("")
  const [boxPickerOpen, setBoxPickerOpen] = useState(false)
  const [boxSearch, setBoxSearch] = useState("")
  const [showNewBox, setShowNewBox] = useState(false)
  const [newBoxNumber, setNewBoxNumber] = useState("")
  const [newBoxName, setNewBoxName] = useState("")
  const [newBoxOpeningBalance, setNewBoxOpeningBalance] = useState("0")
  const [editingBoxId, setEditingBoxId] = useState("")
  const [editBoxNumber, setEditBoxNumber] = useState("")
  const [editBoxName, setEditBoxName] = useState("")
  const [editBoxBalance, setEditBoxBalance] = useState("")
  const [entryKind, setEntryKind] = useState<EntryKind>("debit")
  const [entryDate, setEntryDate] = useState(localDateValue)
  const [entryPerson, setEntryPerson] = useState("")
  const [entryAmount, setEntryAmount] = useState("")
  const [entryNote, setEntryNote] = useState("")
  const [protectedEntryAction, setProtectedEntryAction] = useState<ProtectedEntryAction | null>(null)
  const [entryPassword, setEntryPassword] = useState("")
  const [passwordError, setPasswordError] = useState("")
  const [verifyingPassword, setVerifyingPassword] = useState(false)
  const [editingEntry, setEditingEntry] = useState<CashEntry | null>(null)
  const [editEntryBoxId, setEditEntryBoxId] = useState("")
  const [editEntryKind, setEditEntryKind] = useState<EntryKind>("debit")
  const [editEntryDate, setEditEntryDate] = useState("")
  const [editEntryPerson, setEditEntryPerson] = useState("")
  const [editEntryAmount, setEditEntryAmount] = useState("")
  const [editEntryNote, setEditEntryNote] = useState("")
  const [historySearch, setHistorySearch] = useState("")
  const [historyBoxId, setHistoryBoxId] = useState("all")
  const [dateFrom, setDateFrom] = useState("")
  const [dateTo, setDateTo] = useState("")

  useGlobalLoading(loadedKey !== storageKey && !error, "جاري تحميل بيانات الصناديق المالية...")
  useGlobalLoading(busy, "جاري حفظ حركة الصندوق...")

  useEffect(() => {
    let active = true
    setLoadedKey(null)
    setData(EMPTY_DATA)
    setSelectedBoxId("")
    setError("")
    setNotice("")
    const load = async () => {
      try {
        const stored = centralStorageEnabled
          ? await centralStorageGet<CashboxData>(storageKey) ?? EMPTY_DATA
          : JSON.parse(localStorage.getItem(storageKey) ?? "null") as CashboxData | null ?? EMPTY_DATA
        if (!stored || !Array.isArray(stored.boxes) || !Array.isArray(stored.entries)) {
          throw new Error("بيانات الصندوق المالي المحفوظة غير صالحة.")
        }
        if (active) {
          setData(stored)
          setSelectedBoxId(stored.boxes[0]?.id ?? "")
          setLoadedKey(storageKey)
        }
      } catch (cause) {
        if (active) setError(cause instanceof Error ? `تعذر تحميل بيانات الصناديق: ${cause.message}` : "تعذر تحميل بيانات الصناديق.")
      }
    }
    void load()
    return () => { active = false }
  }, [storageKey])

  const commitData = async (next: CashboxData, successMessage: string) => {
    setBusy(true)
    setError("")
    setNotice("")
    try {
      if (centralStorageEnabled) await centralStorageSet(storageKey, next)
      else localStorage.setItem(storageKey, JSON.stringify(next))
      setData(next)
      setNotice(successMessage)
      return true
    } catch (cause) {
      setError(cause instanceof Error ? `تعذر الحفظ؛ لم يتم اعتماد التعديل: ${cause.message}` : "تعذر الحفظ؛ لم يتم اعتماد التعديل.")
      return false
    } finally {
      setBusy(false)
    }
  }

  const balances = useMemo(() => new Map(data.boxes.map(box => [box.id, boxBalance(box, data.entries)])), [data])
  const selectedBox = data.boxes.find(box => box.id === selectedBoxId)
  const totalBalance = [...balances.values()].reduce((sum, balance) => sum + balance, 0)
  const totalDebit = data.entries.reduce((sum, entry) => sum + entry.debit, 0)
  const totalCredit = data.entries.reduce((sum, entry) => sum + entry.credit, 0)
  const matchingBoxes = data.boxes.filter(box => {
    const query = boxSearch.trim().toLocaleLowerCase("ar")
    return !query || box.number.toLocaleLowerCase("ar").includes(query) || box.name.toLocaleLowerCase("ar").includes(query)
  })
  const visibleEntries = useMemo(() => data.entries
    .filter(entry => historyBoxId === "all" || entry.boxId === historyBoxId)
    .filter(entry => (!dateFrom || entry.date >= dateFrom) && (!dateTo || entry.date <= dateTo))
    .filter(entry => {
      const box = data.boxes.find(item => item.id === entry.boxId)
      const query = historySearch.trim().toLocaleLowerCase("ar")
      return !query || [entry.person, entry.note, box?.number, box?.name]
        .some(value => value?.toLocaleLowerCase("ar").includes(query))
    })
    .sort((a, b) => b.date.localeCompare(a.date) || b.createdAt.localeCompare(a.createdAt)),
  [data.boxes, data.entries, dateFrom, dateTo, historyBoxId, historySearch])

  const createBox = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    const number = newBoxNumber.trim()
    const name = newBoxName.trim()
    const openingBalance = Number(newBoxOpeningBalance)
    if (!number || !name || !Number.isFinite(openingBalance) || openingBalance < 0) {
      setError("أدخل رقم الصندوق واسمه ورصيد افتتاحي صحيحاً (صفر أو أكثر).")
      return
    }
    if (data.boxes.some(box => box.number.trim().toLocaleLowerCase("ar") === number.toLocaleLowerCase("ar"))) {
      setError("رقم الصندوق مستخدم بالفعل؛ اختر رقماً آخر.")
      return
    }
    const box: CashBox = { id: createId(), number, name, openingBalance: roundAmount(openingBalance), createdAt: new Date().toISOString() }
    if (await commitData({ ...data, boxes: [box, ...data.boxes] }, `تم إنشاء الصندوق ${number} · ${name}.`)) {
      setSelectedBoxId(box.id)
      setHistoryBoxId(box.id)
      setShowNewBox(false)
      setNewBoxNumber("")
      setNewBoxName("")
      setNewBoxOpeningBalance("0")
    }
  }

  const startEditingBox = (box: CashBox) => {
    setError("")
    setEditingBoxId(box.id)
    setEditBoxNumber(box.number)
    setEditBoxName(box.name)
    setEditBoxBalance(String(balances.get(box.id) ?? 0))
  }

  const updateBox = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    const box = data.boxes.find(item => item.id === editingBoxId)
    const number = editBoxNumber.trim()
    const name = editBoxName.trim()
    const balance = Number(editBoxBalance)
    if (!box || !number || !name || !Number.isFinite(balance) || balance < 0) {
      setError("أدخل رقم الصندوق واسمه ورصيداً صحيحاً (صفر أو أكثر).")
      return
    }
    if (data.boxes.some(item => item.id !== box.id && item.number.trim().toLocaleLowerCase("ar") === number.toLocaleLowerCase("ar"))) {
      setError("رقم الصندوق مستخدم بالفعل؛ اختر رقماً آخر.")
      return
    }
    const movementBalance = data.entries.reduce((sum, entry) =>
      entry.boxId === box.id ? sum + entry.debit - entry.credit : sum, 0,
    )
    const updatedBox: CashBox = {
      ...box,
      number,
      name,
      balanceAdjustment: roundAmount(balance - box.openingBalance - movementBalance),
    }
    const next = { ...data, boxes: data.boxes.map(item => item.id === box.id ? updatedBox : item) }
    if (await commitData(next, `تم تحديث بيانات الصندوق ${number} · ${name} ورصيده.`)) {
      setEditingBoxId("")
    }
  }

  const deleteBox = async (box: CashBox) => {
    const next = { ...data, boxes: data.boxes.filter(item => item.id !== box.id) }
    if (await commitData(next, `تم حذف الصندوق ${box.number} · ${box.name} مع الاحتفاظ بسجل حركاته.`)) {
      if (selectedBoxId === box.id) setSelectedBoxId(next.boxes[0]?.id ?? "")
    }
  }

  const addEntry = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    const person = entryPerson.trim()
    const amount = Number(entryAmount)
    if (!selectedBox || !person || !entryDate || !Number.isFinite(amount) || amount <= 0) {
      setError("اختر صندوقاً وأدخل اسم الطرف والتاريخ ومبلغاً أكبر من صفر.")
      return
    }
    const entry: CashEntry = {
      id: createId(),
      boxId: selectedBox.id,
      date: entryDate,
      person,
      debit: entryKind === "debit" ? roundAmount(amount) : 0,
      credit: entryKind === "credit" ? roundAmount(amount) : 0,
      note: entryNote.trim(),
      createdAt: new Date().toISOString(),
    }
    if (await commitData({ ...data, entries: [entry, ...data.entries] }, entryKind === "debit" ? "تم تسجيل القبض وزيادة رصيد الصندوق." : "تم تسجيل الصرف ونقص رصيد الصندوق.")) {
      setEntryPerson("")
      setEntryAmount("")
      setEntryNote("")
      setHistoryBoxId(selectedBox.id)
    }
  }

  const requestEntryAction = (kind: ProtectedEntryAction["kind"], entry: CashEntry) => {
    if (kind === "delete" && !window.confirm("هل أنت متأكد من حذف هذه الحركة؟ سيُعاد احتساب رصيد الصندوق.")) return
    setProtectedEntryAction({ kind, entry })
    setEntryPassword("")
    setPasswordError("")
  }

  const authorizeEntryAction = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (!protectedEntryAction || verifyingPassword) return
    setVerifyingPassword(true)
    setPasswordError("")
    try {
      await verifyUserPassword(entryPassword)
      const { kind, entry } = protectedEntryAction
      if (kind === "edit") {
        setEditingEntry(entry)
        setEditEntryBoxId(entry.boxId)
        setEditEntryKind(entry.debit > 0 ? "debit" : "credit")
        setEditEntryDate(entry.date)
        setEditEntryPerson(entry.person)
        setEditEntryAmount(String(entry.debit > 0 ? entry.debit : entry.credit))
        setEditEntryNote(entry.note)
      } else {
        const next = { ...data, entries: data.entries.filter(item => item.id !== entry.id) }
        if (await commitData(next, "تم حذف الحركة وإعادة احتساب رصيد الصندوق.")) {
          setProtectedEntryAction(null)
        }
      }
      setEntryPassword("")
      setProtectedEntryAction(null)
    } catch (cause) {
      setPasswordError(cause instanceof Error ? cause.message : "كلمة المرور غير صحيحة أو تعذر التحقق منها.")
      setEntryPassword("")
    } finally {
      setVerifyingPassword(false)
    }
  }

  const updateEntry = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (!editingEntry) return
    const person = editEntryPerson.trim()
    const amount = Number(editEntryAmount)
    if (!data.boxes.some(box => box.id === editEntryBoxId) || !person || !editEntryDate || !Number.isFinite(amount) || amount <= 0) {
      setError("اختر صندوقاً وأدخل اسم الطرف والتاريخ ومبلغاً أكبر من صفر.")
      return
    }
    const updatedEntry: CashEntry = {
      ...editingEntry,
      boxId: editEntryBoxId,
      date: editEntryDate,
      person,
      debit: editEntryKind === "debit" ? roundAmount(amount) : 0,
      credit: editEntryKind === "credit" ? roundAmount(amount) : 0,
      note: editEntryNote.trim(),
    }
    const next = { ...data, entries: data.entries.map(item => item.id === editingEntry.id ? updatedEntry : item) }
    if (await commitData(next, "تم تعديل الحركة وإعادة احتساب رصيد الصندوق.")) setEditingEntry(null)
  }

  if (loadedKey !== storageKey && !error) return null

  return (
    <main className="min-h-full bg-slate-50 p-4 sm:p-6 lg:p-8">
      <div className="mx-auto max-w-7xl space-y-6">
        <header className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <h1 className="flex items-center gap-2 text-2xl font-bold text-slate-900"><WalletCards className="h-7 w-7 text-cyan-700" />الصندوق المالي</h1>
            <p className="mt-1 text-sm text-slate-600">سجل القبض والصرف واعرف رصيد كل صندوق وحركاته اليومية.</p>
          </div>
          <button type="button" onClick={() => { setError(""); setShowNewBox(true) }}
            className="inline-flex items-center gap-2 rounded-xl bg-cyan-700 px-4 py-2.5 text-sm font-semibold text-white shadow-sm hover:bg-cyan-800">
            <Plus className="h-4 w-4" />إنشاء صندوق
          </button>
        </header>

        {error && <div role="alert" className="flex items-center justify-between gap-3 rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-800"><span>{error}</span><button type="button" onClick={() => setError("")} aria-label="إغلاق" className="rounded p-1 hover:bg-red-100"><X className="h-4 w-4" /></button></div>}
        {notice && <div role="status" className="flex items-center gap-2 rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-800"><Check className="h-4 w-4" />{notice}</div>}

        <section className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
          <article className="rounded-2xl border border-cyan-100 bg-white p-4 shadow-sm">
            <p className="text-xs font-medium text-slate-500">الرصيد الحالي في كل الصناديق</p>
            <p className="mt-2 font-mono text-2xl font-bold text-cyan-800">{formatAmount(totalBalance)} ₪</p>
          </article>
          <article className="rounded-2xl border border-emerald-100 bg-white p-4 shadow-sm">
            <p className="text-xs font-medium text-slate-500">إجمالي المقبوض (مدين)</p>
            <p className="mt-2 font-mono text-2xl font-bold text-emerald-700">{formatAmount(totalDebit)} ₪</p>
          </article>
          <article className="rounded-2xl border border-rose-100 bg-white p-4 shadow-sm">
            <p className="text-xs font-medium text-slate-500">إجمالي المصروف (دائن)</p>
            <p className="mt-2 font-mono text-2xl font-bold text-rose-700">{formatAmount(totalCredit)} ₪</p>
          </article>
          <article className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
            <p className="text-xs font-medium text-slate-500">عدد الصناديق</p>
            <p className="mt-2 text-2xl font-bold text-slate-800">{data.boxes.length.toLocaleString("ar")}</p>
          </article>
        </section>

        <section className="grid items-start gap-5 xl:grid-cols-[minmax(0,1.35fr)_minmax(320px,0.65fr)]">
          <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
            <div className="mb-4 flex items-center justify-between gap-2">
              <div>
                <h2 className="text-base font-bold text-slate-900">تسجيل حركة يومية</h2>
                <p className="mt-1 text-xs text-slate-500">مدين يعني قبضاً يزيد الرصيد، ودائن يعني صرفاً ينقصه.</p>
              </div>
              {selectedBox && <span className="rounded-full bg-cyan-50 px-3 py-1 text-xs font-semibold text-cyan-800">{selectedBox.number} · {selectedBox.name}</span>}
            </div>
            {!data.boxes.length ? (
              <div className="rounded-xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center">
                <p className="text-sm font-semibold text-slate-700">أنشئ صندوقاً أولاً لتسجيل الحركات.</p>
                <button type="button" onClick={() => setShowNewBox(true)} className="mt-3 rounded-lg bg-cyan-700 px-3 py-2 text-xs font-semibold text-white hover:bg-cyan-800">إنشاء أول صندوق</button>
              </div>
            ) : (
              <form onSubmit={event => void addEntry(event)} className="space-y-4">
                <div className="relative">
                  <label className="mb-1.5 block text-xs font-semibold text-slate-700">الصندوق</label>
                  <button type="button" onClick={() => { setBoxPickerOpen(open => !open); setBoxSearch("") }}
                    aria-expanded={boxPickerOpen}
                    className={`${INPUT_CLASS} flex items-center justify-between text-right`}>
                    <span>{selectedBox ? `${selectedBox.number} · ${selectedBox.name}` : "اختر الصندوق..."}</span>
                    <span className="text-xs text-slate-500">عرض الصناديق</span>
                  </button>
                  {boxPickerOpen && (
                    <div className="absolute inset-x-0 top-full z-20 mt-1 rounded-xl border border-slate-200 bg-white p-2 shadow-xl">
                      <div className="relative">
                        <Search className="absolute right-3 top-2.5 h-4 w-4 text-slate-400" />
                        <input autoFocus value={boxSearch} onChange={event => setBoxSearch(event.target.value)} placeholder="ابحث برقم الصندوق أو اسمه" className={`${INPUT_CLASS} pr-9`} />
                      </div>
                      <div className="mt-2 max-h-52 overflow-y-auto">
                        {matchingBoxes.map(box => (
                          <button key={box.id} type="button" onClick={() => { setSelectedBoxId(box.id); setBoxPickerOpen(false) }}
                            className={`flex w-full items-center justify-between rounded-lg px-3 py-2 text-right text-sm hover:bg-cyan-50 ${box.id === selectedBoxId ? "bg-cyan-50 text-cyan-900" : "text-slate-700"}`}>
                            <span className="font-semibold">{box.number} · {box.name}</span>
                            <span className="font-mono text-xs text-slate-500">{formatAmount(balances.get(box.id) ?? 0)} ₪</span>
                          </button>
                        ))}
                        {!matchingBoxes.length && <p className="px-3 py-4 text-center text-xs text-slate-500">لا توجد صناديق مطابقة للبحث.</p>}
                      </div>
                    </div>
                  )}
                  {selectedBox && <p className="mt-1.5 text-xs text-slate-500">الرصيد الحالي: <span className="font-mono font-semibold text-slate-800">{formatAmount(balances.get(selectedBox.id) ?? 0)} ₪</span></p>}
                </div>

                <div className="grid gap-3 sm:grid-cols-2">
                  <label className="text-xs font-semibold text-slate-700">نوع الحركة
                    <select value={entryKind} onChange={event => setEntryKind(event.target.value as EntryKind)} className={`${INPUT_CLASS} mt-1.5`}>
                      <option value="debit">مدين · قبضت منه (يزيد الصندوق)</option>
                      <option value="credit">دائن · دفعت له (ينقص الصندوق)</option>
                    </select>
                  </label>
                  <label className="text-xs font-semibold text-slate-700">التاريخ
                    <span className="relative mt-1.5 block"><CalendarDays className="pointer-events-none absolute right-3 top-3 h-4 w-4 text-slate-400" /><input required type="date" value={entryDate} onChange={event => setEntryDate(event.target.value)} className={`${INPUT_CLASS} pr-9`} /></span>
                  </label>
                </div>
                <div className="grid gap-3 sm:grid-cols-2">
                  <label className="text-xs font-semibold text-slate-700">اسم الطرف (منه / له)
                    <input required value={entryPerson} onChange={event => setEntryPerson(event.target.value)} placeholder="اسم الشخص أو الجهة" className={`${INPUT_CLASS} mt-1.5`} />
                  </label>
                  <label className="text-xs font-semibold text-slate-700">المبلغ (₪)
                    <input required type="number" min="0.01" step="0.01" value={entryAmount} onChange={event => setEntryAmount(event.target.value)} placeholder="0.00" className={`${INPUT_CLASS} mt-1.5 text-left font-mono`} dir="ltr" />
                  </label>
                </div>
                <label className="block text-xs font-semibold text-slate-700">ملاحظة
                  <textarea value={entryNote} onChange={event => setEntryNote(event.target.value)} rows={2} placeholder="ملاحظة اختيارية عن الحركة..." className={`${INPUT_CLASS} mt-1.5 resize-y`} />
                </label>
                <button disabled={busy || loadedKey !== storageKey} type="submit"
                  className={`inline-flex w-full items-center justify-center gap-2 rounded-xl px-4 py-2.5 text-sm font-semibold text-white transition disabled:cursor-not-allowed disabled:opacity-50 ${entryKind === "debit" ? "bg-emerald-700 hover:bg-emerald-800" : "bg-rose-700 hover:bg-rose-800"}`}>
                  {entryKind === "debit" ? <ArrowDownLeft className="h-4 w-4" /> : <ArrowUpRight className="h-4 w-4" />}
                  {entryKind === "debit" ? "تسجيل قبض · مدين" : "تسجيل صرف · دائن"}
                </button>
              </form>
            )}
          </div>

          <section className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
            <h2 className="mb-3 text-base font-bold text-slate-900">أرصدة الصناديق</h2>
            {data.boxes.length ? (
              <div className="space-y-2">
                {data.boxes.map(box => (
                  <article key={box.id} className={`rounded-xl border p-3 transition ${selectedBoxId === box.id ? "border-cyan-300 bg-cyan-50/70" : "border-slate-100 bg-slate-50 hover:border-slate-200"}`}>
                    <button type="button" onClick={() => { setSelectedBoxId(box.id); setHistoryBoxId(box.id) }}
                      className="flex w-full items-center justify-between gap-3 text-right">
                      <span className="min-w-0">
                        <span className="block truncate text-sm font-semibold text-slate-800">{box.name}</span>
                        <span className="mt-0.5 block text-xs text-slate-500">رقم الصندوق: {box.number}</span>
                      </span>
                      <span className="shrink-0 text-left">
                        <span className="block font-mono text-sm font-bold text-cyan-800">{formatAmount(balances.get(box.id) ?? 0)} ₪</span>
                        <span className="text-[10px] text-slate-500">الرصيد الحالي</span>
                      </span>
                    </button>
                    <div className="mt-2 flex justify-end gap-1 border-t border-slate-200/70 pt-2">
                      <button type="button" onClick={() => startEditingBox(box)} disabled={busy}
                        aria-label={`تعديل الصندوق ${box.name}`} title="تعديل الصندوق"
                        className="inline-flex items-center gap-1 rounded-lg px-2 py-1 text-xs font-medium text-cyan-800 hover:bg-cyan-100 disabled:opacity-50">
                        <Pencil className="h-3.5 w-3.5" />تعديل
                      </button>
                      <button type="button" onClick={() => void deleteBox(box)} disabled={busy}
                        aria-label={`حذف الصندوق ${box.name}`} title="حذف الصندوق"
                        className="inline-flex items-center gap-1 rounded-lg px-2 py-1 text-xs font-medium text-rose-700 hover:bg-rose-100 disabled:opacity-50">
                        <Trash2 className="h-3.5 w-3.5" />حذف
                      </button>
                    </div>
                  </article>
                ))}
              </div>
            ) : <p className="rounded-xl bg-slate-50 p-4 text-center text-xs text-slate-500">لم يتم إنشاء صناديق بعد.</p>}
          </section>
        </section>

        <section className="overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-sm">
          <div className="border-b border-slate-100 p-4 sm:p-5">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div><h2 className="text-base font-bold text-slate-900">سجل حركات الصندوق</h2><p className="mt-1 text-xs text-slate-500">المدين قبض يضاف للرصيد، والدائن صرف يخصم منه.</p></div>
              <span className="rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-600">{visibleEntries.length.toLocaleString("ar")} حركة</span>
            </div>
            <div className="mt-4 grid gap-2 sm:grid-cols-2 xl:grid-cols-4">
              <div className="relative sm:col-span-2 xl:col-span-1"><Search className="absolute right-3 top-2.5 h-4 w-4 text-slate-400" /><input value={historySearch} onChange={event => setHistorySearch(event.target.value)} placeholder="بحث باسم الطرف أو الملاحظة" className={`${INPUT_CLASS} pr-9`} /></div>
              <select value={historyBoxId} onChange={event => setHistoryBoxId(event.target.value)} className={INPUT_CLASS}><option value="all">كل الصناديق</option>{data.boxes.map(box => <option key={box.id} value={box.id}>{box.number} · {box.name}</option>)}</select>
              <label className="flex items-center gap-2 rounded-xl border border-slate-200 px-3 text-xs text-slate-500"><span>من</span><input type="date" value={dateFrom} onChange={event => setDateFrom(event.target.value)} className="min-w-0 flex-1 py-2 text-xs text-slate-700 outline-none" /></label>
              <label className="flex items-center gap-2 rounded-xl border border-slate-200 px-3 text-xs text-slate-500"><span>إلى</span><input type="date" value={dateTo} onChange={event => setDateTo(event.target.value)} className="min-w-0 flex-1 py-2 text-xs text-slate-700 outline-none" /></label>
            </div>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full min-w-[900px] text-sm">
              <thead><tr className="bg-slate-50 text-xs text-slate-600"><th className="px-4 py-3 text-right">التاريخ</th><th className="px-4 py-3 text-right">الصندوق</th><th className="px-4 py-3 text-right">اسم الطرف</th><th className="px-4 py-3 text-right">مدين · قبض</th><th className="px-4 py-3 text-right">دائن · صرف</th><th className="px-4 py-3 text-right">ملاحظة</th><th className="px-4 py-3 text-right">إجراءات</th></tr></thead>
              <tbody>
                {visibleEntries.map(entry => {
                  const box = data.boxes.find(item => item.id === entry.boxId)
                  return <tr key={entry.id} className="border-t border-slate-100 hover:bg-slate-50/70">
                    <td className="whitespace-nowrap px-4 py-3 text-xs text-slate-600">{entry.date}</td>
                    <td className="px-4 py-3 text-xs font-medium text-slate-700">{box ? `${box.number} · ${box.name}` : "صندوق محذوف"}</td>
                    <td className="px-4 py-3 font-medium text-slate-800">{entry.person}</td>
                    <td className="px-4 py-3 font-mono font-semibold text-emerald-700">{entry.debit ? `${formatAmount(entry.debit)} ₪` : "—"}</td>
                    <td className="px-4 py-3 font-mono font-semibold text-rose-700">{entry.credit ? `${formatAmount(entry.credit)} ₪` : "—"}</td>
                    <td className="max-w-xs px-4 py-3 text-xs text-slate-500">{entry.note || "—"}</td>
                    <td className="whitespace-nowrap px-4 py-3">
                      <div className="flex items-center gap-1">
                        <button type="button" onClick={() => requestEntryAction("edit", entry)} disabled={busy || verifyingPassword}
                          aria-label={`تعديل حركة ${entry.person}`} title="تعديل الحركة (يتطلب كلمة المرور)"
                          className="inline-flex items-center gap-1 rounded-lg px-2 py-1 text-xs font-medium text-cyan-800 hover:bg-cyan-50 disabled:opacity-50">
                          <Pencil className="h-3.5 w-3.5" />تعديل
                        </button>
                        <button type="button" onClick={() => requestEntryAction("delete", entry)} disabled={busy || verifyingPassword}
                          aria-label={`حذف حركة ${entry.person}`} title="حذف الحركة (يتطلب كلمة المرور)"
                          className="inline-flex items-center gap-1 rounded-lg px-2 py-1 text-xs font-medium text-rose-700 hover:bg-rose-50 disabled:opacity-50">
                          <Trash2 className="h-3.5 w-3.5" />حذف
                        </button>
                      </div>
                    </td>
                  </tr>
                })}
                {!visibleEntries.length && <tr><td colSpan={7} className="px-4 py-10 text-center text-xs text-slate-500">{data.entries.length ? "لا توجد حركات تطابق البحث والفلاتر." : "لا توجد حركات بعد؛ أضف أول قبض أو صرف من النموذج أعلاه."}</td></tr>}
              </tbody>
            </table>
          </div>
        </section>
      </div>

      {showNewBox && (
        <div className="fixed inset-0 z-[80] flex items-center justify-center bg-slate-950/40 p-4" role="presentation" onMouseDown={event => { if (event.target === event.currentTarget) setShowNewBox(false) }}>
          <section role="dialog" aria-modal="true" aria-labelledby="cashbox-dialog-title" className="w-full max-w-md rounded-2xl bg-white p-5 shadow-2xl">
            <div className="mb-4 flex items-center justify-between gap-3">
              <div><h2 id="cashbox-dialog-title" className="text-lg font-bold text-slate-900">إنشاء صندوق جديد</h2><p className="mt-1 text-xs text-slate-500">أدخل رقم الصندوق واسمه، وحدد الرصيد الافتتاحي إن وجد.</p></div>
              <button type="button" onClick={() => setShowNewBox(false)} aria-label="إغلاق" className="rounded-lg p-2 text-slate-500 hover:bg-slate-100"><X className="h-4 w-4" /></button>
            </div>
            <form onSubmit={event => void createBox(event)} className="space-y-3">
              <label className="block text-xs font-semibold text-slate-700">رقم الصندوق<input required value={newBoxNumber} onChange={event => setNewBoxNumber(event.target.value)} placeholder="مثال: 001" className={`${INPUT_CLASS} mt-1.5`} /></label>
              <label className="block text-xs font-semibold text-slate-700">اسم الصندوق<input required value={newBoxName} onChange={event => setNewBoxName(event.target.value)} placeholder="مثال: الصندوق الرئيسي" className={`${INPUT_CLASS} mt-1.5`} /></label>
              <label className="block text-xs font-semibold text-slate-700">الرصيد الافتتاحي (₪)<input required type="number" min="0" step="0.01" value={newBoxOpeningBalance} onChange={event => setNewBoxOpeningBalance(event.target.value)} className={`${INPUT_CLASS} mt-1.5 text-left font-mono`} dir="ltr" /></label>
              <div className="flex justify-end gap-2 pt-2">
                <button type="button" onClick={() => setShowNewBox(false)} className="rounded-xl border border-slate-200 px-4 py-2 text-sm text-slate-700 hover:bg-slate-50">إلغاء</button>
                <button type="submit" disabled={busy} className="inline-flex items-center gap-2 rounded-xl bg-cyan-700 px-4 py-2 text-sm font-semibold text-white hover:bg-cyan-800 disabled:opacity-50"><Plus className="h-4 w-4" />إنشاء الصندوق</button>
              </div>
            </form>
          </section>
        </div>
      )}
      {editingBoxId && (
        <div className="fixed inset-0 z-[80] flex items-center justify-center bg-slate-950/40 p-4" role="presentation" onMouseDown={event => { if (event.target === event.currentTarget) setEditingBoxId("") }}>
          <section role="dialog" aria-modal="true" aria-labelledby="cashbox-edit-dialog-title" className="w-full max-w-md rounded-2xl bg-white p-5 shadow-2xl">
            <div className="mb-4 flex items-center justify-between gap-3">
              <div><h2 id="cashbox-edit-dialog-title" className="text-lg font-bold text-slate-900">تعديل الصندوق</h2><p className="mt-1 text-xs text-slate-500">تعديل الرصيد لا يغيّر سجل الحركات السابقة.</p></div>
              <button type="button" onClick={() => setEditingBoxId("")} aria-label="إغلاق" className="rounded-lg p-2 text-slate-500 hover:bg-slate-100"><X className="h-4 w-4" /></button>
            </div>
            {error && <div role="alert" className="mb-3 rounded-xl border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-800">{error}</div>}
            <form onSubmit={event => void updateBox(event)} className="space-y-3">
              <label className="block text-xs font-semibold text-slate-700">رقم الصندوق<input required value={editBoxNumber} onChange={event => setEditBoxNumber(event.target.value)} className={`${INPUT_CLASS} mt-1.5`} /></label>
              <label className="block text-xs font-semibold text-slate-700">اسم الصندوق<input required value={editBoxName} onChange={event => setEditBoxName(event.target.value)} className={`${INPUT_CLASS} mt-1.5`} /></label>
              <label className="block text-xs font-semibold text-slate-700">الرصيد الحالي (₪)<input required type="number" min="0" step="0.01" value={editBoxBalance} onChange={event => setEditBoxBalance(event.target.value)} className={`${INPUT_CLASS} mt-1.5 text-left font-mono`} dir="ltr" /></label>
              <div className="flex justify-end gap-2 pt-2">
                <button type="button" onClick={() => setEditingBoxId("")} className="rounded-xl border border-slate-200 px-4 py-2 text-sm text-slate-700 hover:bg-slate-50">إلغاء</button>
                <button type="submit" disabled={busy} className="inline-flex items-center gap-2 rounded-xl bg-cyan-700 px-4 py-2 text-sm font-semibold text-white hover:bg-cyan-800 disabled:opacity-50"><Check className="h-4 w-4" />حفظ التعديلات</button>
              </div>
            </form>
          </section>
        </div>
      )}
      {protectedEntryAction && (
        <div className="fixed inset-0 z-[90] flex items-center justify-center bg-slate-950/50 p-4" role="presentation" onMouseDown={event => { if (event.target === event.currentTarget && !verifyingPassword) setProtectedEntryAction(null) }}>
          <section role="dialog" aria-modal="true" aria-labelledby="cashbox-entry-password-title" className="w-full max-w-md rounded-2xl bg-white p-5 shadow-2xl">
            <div className="mb-4 flex items-center justify-between gap-3">
              <div>
                <h2 id="cashbox-entry-password-title" className="text-lg font-bold text-slate-900">تأكيد كلمة المرور</h2>
                <p className="mt-1 text-xs text-slate-500">{protectedEntryAction.kind === "edit" ? "أدخل كلمة مرور المستخدم للمتابعة إلى تعديل الحركة." : "أدخل كلمة مرور المستخدم لتأكيد حذف الحركة."}</p>
              </div>
              <button type="button" onClick={() => setProtectedEntryAction(null)} disabled={verifyingPassword} aria-label="إغلاق" className="rounded-lg p-2 text-slate-500 hover:bg-slate-100 disabled:opacity-50"><X className="h-4 w-4" /></button>
            </div>
            <p className="mb-3 rounded-xl bg-slate-50 p-3 text-xs text-slate-700">{protectedEntryAction.entry.date} · {protectedEntryAction.entry.person} · {formatAmount(protectedEntryAction.entry.debit || protectedEntryAction.entry.credit)} ₪</p>
            {passwordError && <div role="alert" className="mb-3 rounded-xl border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-800">{passwordError}</div>}
            <form onSubmit={event => void authorizeEntryAction(event)} className="space-y-3">
              <label className="block text-xs font-semibold text-slate-700">كلمة مرور المستخدم
                <input autoFocus required type="password" autoComplete="current-password" value={entryPassword} onChange={event => { setEntryPassword(event.target.value); setPasswordError("") }} className={`${INPUT_CLASS} mt-1.5`} />
              </label>
              <div className="flex justify-end gap-2 pt-2">
                <button type="button" onClick={() => setProtectedEntryAction(null)} disabled={verifyingPassword} className="rounded-xl border border-slate-200 px-4 py-2 text-sm text-slate-700 hover:bg-slate-50 disabled:opacity-50">إلغاء</button>
                <button type="submit" disabled={verifyingPassword || !entryPassword} className={`rounded-xl px-4 py-2 text-sm font-semibold text-white disabled:opacity-50 ${protectedEntryAction.kind === "edit" ? "bg-cyan-700 hover:bg-cyan-800" : "bg-rose-700 hover:bg-rose-800"}`}>
                  {verifyingPassword ? "جارٍ التحقق..." : protectedEntryAction.kind === "edit" ? "تحقق وتعديل" : "تحقق وحذف"}
                </button>
              </div>
            </form>
          </section>
        </div>
      )}
      {editingEntry && (
        <div className="fixed inset-0 z-[80] flex items-center justify-center bg-slate-950/40 p-4" role="presentation" onMouseDown={event => { if (event.target === event.currentTarget) setEditingEntry(null) }}>
          <section role="dialog" aria-modal="true" aria-labelledby="cashbox-entry-edit-title" className="w-full max-w-lg rounded-2xl bg-white p-5 shadow-2xl">
            <div className="mb-4 flex items-center justify-between gap-3">
              <div><h2 id="cashbox-entry-edit-title" className="text-lg font-bold text-slate-900">تعديل حركة الصندوق</h2><p className="mt-1 text-xs text-slate-500">تم التحقق من كلمة المرور؛ حفظ التعديل سيعيد احتساب رصيد الصندوق.</p></div>
              <button type="button" onClick={() => setEditingEntry(null)} aria-label="إغلاق" className="rounded-lg p-2 text-slate-500 hover:bg-slate-100"><X className="h-4 w-4" /></button>
            </div>
            {error && <div role="alert" className="mb-3 rounded-xl border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-800">{error}</div>}
            <form onSubmit={event => void updateEntry(event)} className="space-y-3">
              <label className="block text-xs font-semibold text-slate-700">الصندوق
                <select required value={editEntryBoxId} onChange={event => setEditEntryBoxId(event.target.value)} className={`${INPUT_CLASS} mt-1.5`}>
                  {data.boxes.map(box => <option key={box.id} value={box.id}>{box.number} · {box.name}</option>)}
                </select>
              </label>
              <div className="grid gap-3 sm:grid-cols-2">
                <label className="block text-xs font-semibold text-slate-700">نوع الحركة
                  <select value={editEntryKind} onChange={event => setEditEntryKind(event.target.value as EntryKind)} className={`${INPUT_CLASS} mt-1.5`}>
                    <option value="debit">مدين · قبض</option>
                    <option value="credit">دائن · صرف</option>
                  </select>
                </label>
                <label className="block text-xs font-semibold text-slate-700">التاريخ
                  <input required type="date" value={editEntryDate} onChange={event => setEditEntryDate(event.target.value)} className={`${INPUT_CLASS} mt-1.5`} />
                </label>
              </div>
              <div className="grid gap-3 sm:grid-cols-2">
                <label className="block text-xs font-semibold text-slate-700">اسم الطرف
                  <input required value={editEntryPerson} onChange={event => setEditEntryPerson(event.target.value)} className={`${INPUT_CLASS} mt-1.5`} />
                </label>
                <label className="block text-xs font-semibold text-slate-700">المبلغ (₪)
                  <input required type="number" min="0.01" step="0.01" value={editEntryAmount} onChange={event => setEditEntryAmount(event.target.value)} className={`${INPUT_CLASS} mt-1.5 text-left font-mono`} dir="ltr" />
                </label>
              </div>
              <label className="block text-xs font-semibold text-slate-700">ملاحظة
                <textarea value={editEntryNote} onChange={event => setEditEntryNote(event.target.value)} rows={2} className={`${INPUT_CLASS} mt-1.5 resize-y`} />
              </label>
              <div className="flex justify-end gap-2 pt-2">
                <button type="button" onClick={() => setEditingEntry(null)} className="rounded-xl border border-slate-200 px-4 py-2 text-sm text-slate-700 hover:bg-slate-50">إلغاء</button>
                <button type="submit" disabled={busy} className="inline-flex items-center gap-2 rounded-xl bg-cyan-700 px-4 py-2 text-sm font-semibold text-white hover:bg-cyan-800 disabled:opacity-50"><Check className="h-4 w-4" />حفظ الحركة</button>
              </div>
            </form>
          </section>
        </div>
      )}
    </main>
  )
}
