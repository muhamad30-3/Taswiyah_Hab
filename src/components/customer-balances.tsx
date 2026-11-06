"use client"

import { useEffect, useMemo, useRef, useState, type ChangeEvent, type FormEvent } from "react"
import * as XLSX from "xlsx"
import * as XLSXStyle from "xlsx-js-style"
import { AlertTriangle, Bell, BellRing, Check, ChevronDown, Download, FileSpreadsheet, Filter, Pencil, Plus, Search, ShieldCheck, Trash2, Upload, Users, X } from "lucide-react"
import { centralStorageEnabled, centralStorageGet, centralStorageSet, verifyUserPassword } from "../lib/central-storage"
import { REPORT_LOGO_PRINT_STYLES, reportLogoMarkup } from "../lib/report-branding"
import { useGlobalLoading } from "./global-loading"

type CustomerAction = {
  id: string
  description: string
  dueDate: string
  reminderDate: string
  reminderDates?: Array<{ id: string; date: string; notifiedFor?: string }>
  notifiedFor?: string
  completedAt?: string
  status?: "pending" | "completed" | "failed"
  failureReason?: string
  completed?: boolean
}

type AgreementFrequency = "weekly" | "monthly" | "bimonthly" | "once" | "flexible"
type AgreementReminder = { id: string; date: string; notifiedFor?: string }
type AgreementClosure = {
  id: string
  closedAt: string
  outcome: Exclude<AgreementCheckOutcome, "ignored">
  paidAmount: number
  balance: number
  balanceDirection: CustomerBalance["balanceDirection"]
}
type AgreementStatus = "active" | "completed" | "refused"
type AgreementSuggestion = "paid" | "partial" | "unpaid" | "not_comparable"
type AgreementCheckOutcome = "paid" | "partial" | "unpaid" | "ignored"

type AgreementCheck = {
  id: string
  checkedAt: string
  previousBalance: number
  previousDirection: CustomerBalance["balanceDirection"]
  newBalance: number
  newDirection: CustomerBalance["balanceDirection"]
  inferredPaid: number
  expectedPayment: number
  suggestion: AgreementSuggestion
  outcome?: AgreementCheckOutcome
  actualPaid?: number
  waivedAmount?: number
  note?: string
}

type AgreementBalanceSnapshot = {
  capturedAt: string
  balance: number
  direction: CustomerBalance["balanceDirection"]
}

type CustomerAgreement = {
  id: string
  startedAt: string
  description: string
  initialBalance: number
  balanceDirection: CustomerBalance["balanceDirection"]
  installmentAmount: number | null
  frequency: AgreementFrequency
  nextPaymentDate: string
  dueDateFrom?: string
  dueDateTo?: string
  reminders?: AgreementReminder[]
  outstandingInstallment: number
  status: AgreementStatus
  closedAt?: string
  closureOutcome?: Exclude<AgreementCheckOutcome, "ignored">
  closurePaidAmount?: number
  closureBalance?: number
  closureBalanceDirection?: CustomerBalance["balanceDirection"]
  closureHistory?: AgreementClosure[]
  resumedAt?: string
  balanceSnapshots?: AgreementBalanceSnapshot[]
  checks: AgreementCheck[]
}

type CustomerNote = {
  agreement: string
  agreementStartedAt?: string
  agreementHistory?: Array<{
    id: string
    agreement: string
    startedAt: string
    closedAt: string
    actions: CustomerAction[]
  }>
  actions: CustomerAction[]
  agreements: CustomerAgreement[]
}

type CustomerBalance = {
  customerNumber: string
  name: string
  paymentClassification?: PaymentClassification
  isInLatestStatement?: boolean
  statementAppearances?: number
  lastSeenInStatementAt?: string
  lastStatementBalance?: number
  lastStatementDirection?: "receivable" | "payable" | "unspecified"
  balance: number
  balanceDirection: "receivable" | "payable" | "unspecified"
  firstStatementBalance?: number
  firstStatementDirection?: "receivable" | "payable" | "unspecified"
  firstStatementAt?: string
  balanceUpdatedAt?: string
  duration: number
  invoiceAmount: number
  invoiceDate: string
  receiptAmount: number
  receiptDate: string
  phone: string
  note: CustomerNote
}

type PaymentClassification = "poor" | "acceptable" | "good" | "excellent"
type AutomaticPaymentClassification = PaymentClassification | "not_applicable"
type PaymentCriteria = {
  acceptableMinPercent: number
  goodMinPercent: number
  excellentMinPercent: number
  acceptableCreditLimitPercent: number
  goodCreditLimitPercent: number
  excellentCreditLimitPercent: number
  recentReceiptDays: number
}

type PaymentAssessment = {
  classification: AutomaticPaymentClassification
  paymentPercent: number | null
  creditLimit: number | null
  invoiceWarning: boolean
  recentReceiptDowngrade: boolean
}

type CustomMessageTemplate = {
  id: string
  name: string
  content: string
}

type BulkMessageRecipient = {
  customerNumber: string
  name: string
  phone: string
  message: string
  balance: number
  balanceDirection: CustomerBalance["balanceDirection"]
}

type CustomerCardTemplate = "modern" | "classic"

const MESSAGE_FIELDS = [
  { token: "{{name}}", label: "اسم الزبون" },
  { token: "{{balance}}", label: "الرصيد" },
  { token: "{{balanceLabel}}", label: "مستحق لنا / لصالحك" },
  { token: "{{due}}", label: "مستحق" },
] as const

const DEFAULT_PAYMENT_CRITERIA: PaymentCriteria = {
  acceptableMinPercent: 40,
  goodMinPercent: 70,
  excellentMinPercent: 100,
  acceptableCreditLimitPercent: 55,
  goodCreditLimitPercent: 95,
  excellentCreditLimitPercent: 120,
  recentReceiptDays: 30,
}

const PAYMENT_CLASSIFICATIONS: Array<{
  value: PaymentClassification | null
  label: string
  color: string
  excelColor: string
}> = [
  { value: "poor", label: "سيئ الدفع", color: "bg-rose-500", excelColor: "FECACA" },
  { value: "acceptable", label: "مقبول", color: "bg-amber-500", excelColor: "FDE68A" },
  { value: "good", label: "جيد وسالك", color: "bg-emerald-500", excelColor: "A7F3D0" },
  { value: "excellent", label: "ممتاز", color: "bg-sky-500", excelColor: "BAE6FD" },
  { value: null, label: "غير مصنف", color: "bg-slate-300", excelColor: "E2E8F0" },
]

const AUTOMATIC_CLASSIFICATIONS: Array<{
  value: AutomaticPaymentClassification
  label: string
}> = [
  { value: "poor", label: "سيئ الدفع" },
  { value: "acceptable", label: "مقبول" },
  { value: "good", label: "جيد وسالك" },
  { value: "excellent", label: "ممتاز" },
  { value: "not_applicable", label: "لا يوجد مستحق للتقييم" },
]

type CustomerPayment = {
  customerNumber: string
  customerName: string
  agreement: string
  date: string
  amount: number
  type: string
}

type CustomerReportType =
  | "balances"
  | "agreements"
  | "notes"
  | "paid_total"
  | "payments"
  | "blocked_balances"
  | "missing_phone"
  | "payment_classification"

type BlockedCustomer = {
  customerNumber: string
  name: string
  phone: string
  reason: string
  blockedAt: string
}

const DEFAULT_STATEMENT_MESSAGE = `الزبون الكريم: {{name}}
تحية طيبة،
نأمل أن تكون بأفضل حال!
نذكرك برصيد {{balanceLabel}} بقيمة *{{balance}} شيكل* لدى *سوبر ماركت بيت الجملة*. يسعدنا جداً زيارتك لنا في الفرع لتسوية الحساب، ونكون سعداء بلقائك ودعمك المستمر. أهلاً بك في أي وقت!
*التحويل على رقم حساب 1083594، رقم جوال للتحويل على حساب بنك فلسطين 0594466633*
*رقم دفع لتاجر بال بي 0566105106*
اسم صاحب الحساب محمود زكي عثمان`

const DEFAULT_WEEKLY_INSTALLMENT_MESSAGE = `الزبون الكريم: {{name}}

تحية طيبة وبعد،

نشكرك جزيل الشكر على تعاونك الدائم والتزامك معنا في *سوبر ماركت بيت الجملة*. كما اتفقنا مسبقاً بشأن تسوية الحساب المتبقي، نود تذكيرك بالدفعة الأسبوعية المستحقة لهذا الأسبوع بقيمة *{{weeklyAmount}} شيكل*.

هذه الأقساط المنتظمة تساعدنا كثيراً في الوفاء بالتزاماتنا واستمرار توفير المنتجات التي تفضلها بأفضل شكل ممكن. نحن نقدر حرصك، ويسعدنا دائماً استقبالك في الفرع لتسليم الدفعة ولقائك العزيز علينا.

كل الشكر لتعاونك المستمر، وأهلاً بك في أي وقت!`

function AgreementDetails({ record, onUpdateNote, onUpdateAction, onAddAction, onCloseSummary, onReopenSummary, onOpenCustomer }: {
  record: CustomerBalance
  onUpdateNote: (update: (note: CustomerNote) => CustomerNote) => void
  onUpdateAction: (actionId: string, update: (action: CustomerAction) => CustomerAction) => void
  onAddAction: () => void
  onCloseSummary: () => void
  onReopenSummary: (summaryId: string) => void
  onOpenCustomer: () => void
}) {
  const [expanded, setExpanded] = useState(false)
  const actionCount = record.note.actions.length
  const completedCount = record.note.actions.filter(action => action.status === "completed" || action.completed).length
  const failedCount = record.note.actions.filter(action => action.status === "failed").length
  const remindersForAction = (action: CustomerAction) => action.reminderDates?.length
    ? action.reminderDates
    : action.reminderDate ? [{ id: action.id, date: action.reminderDate, notifiedFor: action.notifiedFor }] : []

  return (
    <div className="min-w-[120px]">
      <button
        type="button"
        aria-expanded={expanded}
        onClick={() => setExpanded(value => !value)}
        className="inline-flex items-center gap-2 rounded-lg border border-cyan-200 bg-white px-3 py-2 text-xs font-semibold text-cyan-800 transition-colors hover:bg-cyan-50"
      >
        تم الاتفاق على
        {actionCount > 0 && <span className="rounded-full bg-cyan-100 px-1.5 py-0.5 text-[10px]">{completedCount}/{actionCount}</span>}
        {failedCount > 0 && <span className="rounded-full bg-rose-100 px-1.5 py-0.5 text-[10px] text-rose-700">{failedCount} لم ينفذ</span>}
        <ChevronDown className={`h-3.5 w-3.5 transition-transform ${expanded ? "rotate-180" : ""}`} />
      </button>
        <button type="button" onClick={onOpenCustomer} className="mt-1 block text-[10px] font-semibold text-cyan-800 hover:underline">ملف اتفاقات الزبون</button>
      {expanded && <div className="mt-2 min-w-[320px] space-y-3 rounded-xl border border-slate-200 bg-white p-3 shadow-sm">
        <label className="block space-y-1 text-[11px] font-semibold text-slate-700"><span>ملخص الاتفاق الحالي</span><textarea value={record.note.agreement} onChange={event => onUpdateNote(note => ({ ...note, agreement: event.target.value, agreementStartedAt: note.agreementStartedAt ?? new Date().toISOString() }))} rows={2} placeholder="سجل ملخص الاتفاق الحالي..." className="w-full resize-y rounded-lg border border-slate-200 px-3 py-2 text-xs font-normal outline-none focus:border-cyan-500 focus:ring-2 focus:ring-cyan-100" /></label>
        {record.note.agreementStartedAt && <p className="text-[10px] text-slate-500">تاريخ الاتفاق: {new Date(record.note.agreementStartedAt).toLocaleDateString("ar")}</p>}
        <div className="space-y-2">{record.note.actions.map((action, actionIndex) => <div key={action.id} className={`space-y-2 rounded-lg border p-2 ${action.status === "completed" || action.completed ? "border-emerald-200 bg-emerald-50/60" : action.status === "failed" ? "border-rose-200 bg-rose-50/60" : "border-slate-200 bg-slate-50/70"}`}>
          <div className="flex items-center gap-2"><span className="shrink-0 text-[11px] font-bold text-slate-500">{(actionIndex + 1).toLocaleString("ar")}.</span><input value={action.description} onChange={event => onUpdateAction(action.id, current => ({ ...current, description: event.target.value }))} placeholder="مثال: تسديد المبلغ المتفق عليه" className="min-w-0 flex-1 rounded-md border border-slate-200 bg-white px-2 py-1.5 text-xs font-normal outline-none focus:border-cyan-500" /><button type="button" aria-label="حذف بند الاتفاق" onClick={() => onUpdateNote(note => ({ ...note, actions: note.actions.filter(item => item.id !== action.id) }))} className="rounded-md p-1 text-slate-400 hover:bg-rose-50 hover:text-rose-700"><X className="h-4 w-4" /></button></div>
          <label className="block space-y-1 text-[10px] font-semibold text-slate-600"><span>موعد التنفيذ</span><input type="date" value={action.dueDate} onChange={event => onUpdateAction(action.id, current => ({ ...current, dueDate: event.target.value }))} className="w-full rounded-md border border-slate-200 bg-white px-2 py-1.5 text-[11px] font-normal" /></label>
          <div className="space-y-1.5"><span className="text-[10px] font-semibold text-slate-600">ذكرني في هذه التواريخ</span>{remindersForAction(action).map(reminder => <div key={reminder.id} className="flex items-center gap-2"><input type="date" value={reminder.date} onChange={event => onUpdateAction(action.id, current => ({ ...current, reminderDates: remindersForAction(current).map(item => item.id === reminder.id ? { ...item, date: event.target.value, notifiedFor: item.notifiedFor === event.target.value ? item.notifiedFor : undefined } : item), reminderDate: remindersForAction(current).length === 1 ? event.target.value : current.reminderDate }))} className="min-w-0 flex-1 rounded-md border border-slate-200 bg-white px-2 py-1.5 text-[11px] font-normal" /><button type="button" aria-label="حذف موعد تذكير" onClick={() => onUpdateAction(action.id, current => ({ ...current, reminderDates: remindersForAction(current).filter(item => item.id !== reminder.id), reminderDate: "" }))} className="rounded-md p-1 text-slate-400 hover:bg-rose-50 hover:text-rose-700"><X className="h-4 w-4" /></button></div>)}<button type="button" onClick={() => onUpdateAction(action.id, current => ({ ...current, reminderDates: [...remindersForAction(current), { id: createActionId(), date: "" }] }))} className="inline-flex items-center gap-1 rounded-md px-2 py-1 text-[10px] font-semibold text-cyan-800 hover:bg-cyan-50"><Plus className="h-3 w-3" />ذكرني بتاريخ آخر</button></div>
          <div className="flex flex-wrap gap-2">
            <button type="button" aria-pressed={action.status === "completed" || action.completed} onClick={() => onUpdateAction(action.id, current => ({ ...current, status: "completed", completed: true, completedAt: current.completedAt ?? new Date().toISOString(), failureReason: "" }))} className={`inline-flex items-center gap-1 rounded-md px-2.5 py-1.5 text-[11px] font-semibold ${action.status === "completed" || action.completed ? "bg-emerald-700 text-white" : "border border-emerald-200 bg-white text-emerald-800 hover:bg-emerald-50"}`}><Check className="h-3.5 w-3.5" />تم</button>
            <button type="button" aria-pressed={action.status === "failed"} onClick={() => onUpdateAction(action.id, current => ({ ...current, status: "failed", completed: false }))} className={`rounded-md px-2.5 py-1.5 text-[11px] font-semibold ${action.status === "failed" ? "bg-rose-700 text-white" : "border border-rose-200 bg-white text-rose-800 hover:bg-rose-50"}`}>لم ينفذ</button>
            {action.status === "pending" && <span className="self-center text-[10px] text-slate-500">بانتظار التنفيذ</span>}
          </div>
          {action.status === "failed" && <label className="block space-y-1 text-[10px] font-semibold text-rose-800"><span>سبب عدم التنفيذ</span><textarea required value={action.failureReason ?? ""} onChange={event => onUpdateAction(action.id, current => ({ ...current, failureReason: event.target.value }))} rows={2} placeholder="اكتب سبب عدم تنفيذ الاتفاق..." className="w-full resize-y rounded-md border border-rose-200 bg-white px-2 py-1.5 text-xs font-normal outline-none focus:border-rose-400" /></label>}
          {(action.status === "completed" || action.completed) && <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-emerald-700"><Check className="h-3.5 w-3.5" />تم تنفيذ هذا البند{action.completedAt ? ` بتاريخ ${formatDate(action.completedAt.slice(0, 10))}` : ""}</span>}
        </div>)}</div>
        <button type="button" onClick={onAddAction} className="inline-flex items-center gap-1 rounded-lg border border-cyan-200 px-2.5 py-1.5 text-[11px] font-semibold text-cyan-800 hover:bg-cyan-50"><Plus className="h-3.5 w-3.5" />إضافة بند مرقّم</button>
        {(record.note.agreement.trim() || actionCount > 0) && <button type="button" onClick={onCloseSummary} className="inline-flex items-center gap-1 rounded-lg bg-emerald-700 px-3 py-2 text-[11px] font-semibold text-white hover:bg-emerald-800">تسكير الملخص وحفظه بالتاريخ</button>}
        {!!record.note.agreementHistory?.length && <section className="space-y-2 border-t border-slate-200 pt-3"><h4 className="text-xs font-bold text-slate-700">ملخصات الاتفاق السابقة</h4>{[...record.note.agreementHistory].reverse().map(summary => <article key={summary.id} className="rounded-lg border border-slate-200 bg-slate-50 p-2"><p className="whitespace-pre-wrap text-xs">{summary.agreement}</p><p className="mt-1 text-[10px] text-slate-500">اتفق بتاريخ: {summary.startedAt ? formatDate(summary.startedAt.slice(0, 10)) : "غير مسجل"} · أُغلق بتاريخ: {formatDate(summary.closedAt.slice(0, 10))}</p>{summary.actions.map(action => <p key={action.id} className="mt-1 text-[10px] text-slate-600">{action.description} · {action.status === "completed" || action.completed ? `تم التنفيذ${action.completedAt ? ` بتاريخ ${formatDate(action.completedAt.slice(0, 10))}` : ""}` : action.status === "failed" ? "لم ينفذ" : "قيد المتابعة"}</p>)}<button type="button" onClick={() => onReopenSummary(summary.id)} className="mt-2 rounded-md border border-cyan-200 bg-white px-2.5 py-1.5 text-[10px] font-semibold text-cyan-800 hover:bg-cyan-50">فتح هذا الملخص</button></article>)}</section>}
      </div>}
    </div>
  )
}

function CustomerAgreementProfile({
  record,
  isBlocked,
  blockedReason,
  onClose,
  onCreateAgreement,
  onUpdateAgreement,
  onSetCheckOutcome,
  onCloseAgreement,
  onDeleteAgreement,
}: {
  record: CustomerBalance
  isBlocked: boolean
  blockedReason: string
  onClose: () => void
  onCreateAgreement: (agreement: Omit<CustomerAgreement, "id" | "checks" | "outstandingInstallment" | "status">) => void
  onUpdateAgreement: (agreementId: string, update: (agreement: CustomerAgreement) => CustomerAgreement) => void
  onSetCheckOutcome: (agreementId: string, checkId: string, outcome: AgreementCheckOutcome, actualPaid: number, note: string) => void
  onCloseAgreement: (agreementId: string, status: "completed" | "refused", outcome?: Exclude<AgreementCheckOutcome, "ignored">, paidAmount?: number) => void
  onDeleteAgreement: (agreementId: string) => void
}) {
  const agreementFormRef = useRef<HTMLFormElement>(null)
  const [description, setDescription] = useState("")
  const [installmentAmountType, setInstallmentAmountType] = useState<"fixed" | "unknown">("fixed")
  const [installmentAmount, setInstallmentAmount] = useState("")
  const [frequency, setFrequency] = useState<AgreementFrequency>("weekly")
  const [firstDueDate, setFirstDueDate] = useState("")
  const [lastDueDate, setLastDueDate] = useState("")
  const [reminderDates, setReminderDates] = useState<string[]>([])
  const [editingAgreementId, setEditingAgreementId] = useState<string | null>(null)
  const [formError, setFormError] = useState("")
  const [decisionNotes, setDecisionNotes] = useState<Record<string, string>>({})
  const [closingAgreementId, setClosingAgreementId] = useState<string | null>(null)
  const [closureOutcome, setClosureOutcome] = useState<Exclude<AgreementCheckOutcome, "ignored">>("paid")
  const [closurePaidAmount, setClosurePaidAmount] = useState("")
  const [closureError, setClosureError] = useState("")

  const submitAgreement = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    const amount = installmentAmountType === "fixed" ? Number(installmentAmount) : null
    if (!description.trim()) {
      setFormError("اكتب تفاصيل الاتفاق قبل حفظه.")
      return
    }
    if (amount !== null && (!Number.isFinite(amount) || amount <= 0)) {
      setFormError("أدخل مبلغ دفعة أكبر من صفر.")
      return
    }
    if (frequency !== "flexible" && !firstDueDate) {
      setFormError("حدد موعد أول دفعة.")
      return
    }
    if (firstDueDate && lastDueDate && firstDueDate > lastDueDate) {
      setFormError("تاريخ بداية فترة الدفعة يجب أن يسبق تاريخ نهايتها.")
      return
    }
    if (isBlocked && !window.confirm(`تنبيه: هذا الزبون موجود في قائمة المحظورين من التواصل.${blockedReason ? `\n\nالملاحظة: ${blockedReason}` : ""}\n\nهل تريد المتابعة وتسجيل الاتفاق رغم ذلك؟`)) return
    setFormError("")
    if (editingAgreementId) {
      onUpdateAgreement(editingAgreementId, agreement => ({
        ...agreement,
        description: description.trim(),
        installmentAmount: amount,
        frequency,
        nextPaymentDate: firstDueDate,
        dueDateFrom: firstDueDate,
        dueDateTo: lastDueDate,
        reminders: reminderDates.filter(Boolean).map(date => ({ id: createActionId(), date })),
      }))
      setEditingAgreementId(null)
    } else {
      onCreateAgreement({
        startedAt: new Date().toISOString(),
        description: description.trim(),
        initialBalance: record.balance,
        balanceDirection: record.balanceDirection,
        installmentAmount: amount,
        frequency,
        nextPaymentDate: firstDueDate,
        dueDateFrom: firstDueDate,
        dueDateTo: lastDueDate,
        reminders: reminderDates.filter(Boolean).map(date => ({ id: createActionId(), date })),
      })
    }
    setDescription("")
    setInstallmentAmount("")
    setInstallmentAmountType("fixed")
    setFrequency("weekly")
    setFirstDueDate("")
    setLastDueDate("")
    setReminderDates([])
  }

  const agreements = [...record.note.agreements].sort((left, right) => right.startedAt.localeCompare(left.startedAt))
  const activeAgreements = agreements.filter(agreement => agreement.status === "active").length
  const closingAgreement = agreements.find(agreement => agreement.id === closingAgreementId)
  const startEditingAgreement = (agreement: CustomerAgreement) => {
    setEditingAgreementId(agreement.id)
    setDescription(agreement.description)
    setInstallmentAmountType(agreement.installmentAmount === null ? "unknown" : "fixed")
    setInstallmentAmount(agreement.installmentAmount === null ? "" : String(agreement.installmentAmount))
    setFrequency(agreement.frequency)
    setFirstDueDate(agreement.dueDateFrom ?? agreement.nextPaymentDate)
    setLastDueDate(agreement.dueDateTo ?? "")
    setReminderDates((agreement.reminders ?? []).map(reminder => reminder.date))
    setFormError("")
  }

  const submitAgreementClosure = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (!closingAgreement) return
    const paidAmount = closureOutcome === "unpaid" ? 0 : Number(closurePaidAmount)
    const expected = Math.max(closingAgreement.outstandingInstallment, closingAgreement.installmentAmount ?? 0)
    if (!Number.isFinite(paidAmount) || paidAmount < 0 || (closureOutcome !== "unpaid" && paidAmount <= 0)) {
      setClosureError("أدخل مبلغاً صحيحاً تم تسديده.")
      return
    }
    if (closingAgreement.installmentAmount !== null && closureOutcome === "partial" && paidAmount >= expected) {
      setClosureError("مبلغ الدفع الجزئي يجب أن يكون أقل من الدفعة المستحقة.")
      return
    }
    if (closingAgreement.installmentAmount !== null && closureOutcome === "paid" && paidAmount < expected) {
      setClosureError("المبلغ الكامل يجب أن يساوي الدفعة المستحقة أو يتجاوزها.")
      return
    }
    onCloseAgreement(closingAgreement.id, "completed", closureOutcome, paidAmount)
    setClosingAgreementId(null)
    setClosureError("")
  }

  const deleteAgreement = async (agreement: CustomerAgreement) => {
    if (!window.confirm(`حذف الاتفاق «${agreement.description}» نهائياً؟`)) return
    const password = window.prompt("أدخل كلمة مرور المستخدم لتأكيد حذف الاتفاق:")
    if (password === null) return
    try {
      await verifyUserPassword(password)
      onDeleteAgreement(agreement.id)
    } catch (cause) {
      window.alert(cause instanceof Error ? cause.message : "كلمة المرور غير صحيحة أو تعذر التحقق منها.")
    }
  }

  const agreementStatusLabel = (status: AgreementStatus) => status === "active" ? "قيد التنفيذ" : status === "completed" ? "مكتمل" : "مرفوض"
  const agreementFrequencyLabel = (value: AgreementFrequency) =>
    value === "weekly" ? "أسبوعياً" :
      value === "monthly" ? "شهرياً" :
        value === "bimonthly" ? "كل شهرين" :
          value === "flexible" ? "غير دوري / حسب التواصل" :
            "مرة واحدة"
  const agreementOutcomeLabel = (value?: AgreementCheckOutcome) => value === "paid" ? "دفع كامل" : value === "partial" ? "دفع جزئي" : value === "unpaid" ? "لم يسدد" : value === "ignored" ? "تم التجاهل" : "بانتظار المراجعة"

  const exportAgreementWorkbook = () => {
    const workbook = XLSX.utils.book_new()
    const agreementRows = agreements.map(agreement => {
      const latestSnapshot = agreement.balanceSnapshots?.at(-1)
      const latestCheck = agreement.checks.at(-1)
      return [
        record.customerNumber,
        record.name,
        agreement.description,
        agreementStatusLabel(agreement.status),
        new Date(agreement.startedAt).toLocaleString("ar"),
        agreement.initialBalance,
        agreement.balanceDirection,
        agreement.installmentAmount === null ? "غير معلومة" : agreement.installmentAmount,
        agreementFrequencyLabel(agreement.frequency),
        formatDate(agreement.nextPaymentDate),
        agreement.closureOutcome ? agreementOutcomeLabel(agreement.closureOutcome) : "",
        agreement.closurePaidAmount ?? "",
        agreement.closedAt ? new Date(agreement.closedAt).toLocaleString("ar") : "",
        agreement.closureBalance ?? latestSnapshot?.balance ?? latestCheck?.newBalance ?? agreement.initialBalance,
        agreement.closureBalanceDirection ?? latestSnapshot?.direction ?? latestCheck?.newDirection ?? agreement.balanceDirection,
        agreement.outstandingInstallment,
      ]
    })
    const checkRows = agreements.flatMap(agreement => agreement.checks.map(check => [
      agreement.description,
      new Date(check.checkedAt).toLocaleString("ar"),
      check.previousBalance,
      check.previousDirection,
      check.newBalance,
      check.newDirection,
      check.inferredPaid,
      agreement.installmentAmount === null ? "غير محدد" : check.expectedPayment,
      check.outcome ? agreementOutcomeLabel(check.outcome) : "بانتظار المراجعة",
      check.actualPaid ?? "",
      check.waivedAmount ?? "",
      check.note ?? "",
    ]))
    const actionRows = record.note.actions.map((action, index) => [
      index + 1,
      action.description,
      action.status === "completed" || action.completed ? "تم" : action.status === "failed" ? "لم ينفذ" : "قيد المتابعة",
      action.dueDate ? formatDate(action.dueDate) : "",
      action.reminderDate ? formatDate(action.reminderDate) : "",
      action.failureReason ?? "",
    ])

    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
      ["رقم الزبون", "اسم الزبون", "تفاصيل الاتفاق", "الحالة", "تاريخ الاتفاق", "رصيد البداية", "اتجاه الرصيد", "مبلغ الدفعة", "التكرار", "الدفعة القادمة", "نتيجة الإنهاء", "المبلغ المدفوع عند الإنهاء", "تاريخ الإكمال", "آخر رصيد مسجل", "اتجاه آخر رصيد", "المتبقي من الدفعة"],
      ...agreementRows,
    ]), "الاتفاقات")
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
      ["الاتفاق", "تاريخ التحديث", "الرصيد السابق", "اتجاهه", "الرصيد الجديد", "اتجاهه", "الانخفاض المستنتج", "المبلغ المتوقع", "النتيجة المعتمدة", "المبلغ المدفوع", "المتجاوز عنه", "ملاحظة"],
      ...checkRows,
    ]), "تحديثات الرصيد")
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
      ["ملخص اتفاق الزبون", record.note.agreement],
      [],
      ["رقم", "بند الاتفاق", "الحالة", "موعد التنفيذ", "موعد التذكير", "سبب عدم التنفيذ"],
      ...actionRows,
    ]), "ملاحظات الزبون")
    XLSX.writeFile(workbook, `اتفاقات-${record.customerNumber}.xlsx`)
  }

  const exportAgreementPdf = () => {
    const printWindow = window.open("", "_blank", "width=1100,height=800")
    if (!printWindow) {
      window.alert("تعذر فتح تقرير PDF. اسمح بالنوافذ المنبثقة ثم أعد المحاولة.")
      return
    }
    const escapeHtml = (value: unknown) => String(value ?? "").replace(/[&<>"']/g, character => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
    })[character] ?? character)
    const statusLabel = (status: AgreementStatus) => agreementStatusLabel(status)
    const agreementRows = agreements.map(agreement => {
      const latest = agreement.balanceSnapshots?.at(-1)?.balance ?? agreement.checks.at(-1)?.newBalance ?? agreement.initialBalance
      return `<tr><td>${escapeHtml(agreement.description)}</td><td>${escapeHtml(statusLabel(agreement.status))}</td><td>${escapeHtml(new Date(agreement.startedAt).toLocaleDateString("ar"))}</td><td>${escapeHtml(formatAmount(agreement.initialBalance))}</td><td>${escapeHtml(formatAmount(latest))}</td><td>${escapeHtml(agreement.installmentAmount === null ? "غير معلومة" : formatAmount(agreement.installmentAmount))} · ${escapeHtml(agreementFrequencyLabel(agreement.frequency))}</td><td>${escapeHtml(agreement.closureOutcome ? agreementOutcomeLabel(agreement.closureOutcome) : "—")}${agreement.closurePaidAmount !== undefined ? ` · ${escapeHtml(formatAmount(agreement.closurePaidAmount))}` : ""}</td><td>${escapeHtml(agreement.closedAt ? new Date(agreement.closedAt).toLocaleDateString("ar") : "—")}</td></tr>`
    }).join("")
    const checkRows = agreements.flatMap(agreement => agreement.checks.map(check =>
      `<tr><td>${escapeHtml(agreement.description)}</td><td>${escapeHtml(new Date(check.checkedAt).toLocaleString("ar"))}</td><td>${escapeHtml(formatAmount(check.previousBalance))} ← ${escapeHtml(formatAmount(check.newBalance))}</td><td>${escapeHtml(agreement.installmentAmount === null ? "غير محدد" : formatAmount(check.expectedPayment))}</td><td>${escapeHtml(check.outcome ? agreementOutcomeLabel(check.outcome) : "بانتظار المراجعة")}</td><td>${escapeHtml(check.actualPaid !== undefined ? formatAmount(check.actualPaid) : "—")}</td><td>${escapeHtml(check.note ?? "—")}</td></tr>`
    )).join("")
    const actions = record.note.actions.map((action, index) =>
      `<li>${escapeHtml(index + 1)}. ${escapeHtml(action.description)} · ${escapeHtml(action.status === "completed" || action.completed ? "تم" : action.status === "failed" ? `لم ينفذ: ${action.failureReason ?? ""}` : "قيد المتابعة")}${action.dueDate ? ` · موعد التنفيذ: ${escapeHtml(formatDate(action.dueDate))}` : ""}${action.reminderDate ? ` · التذكير: ${escapeHtml(formatDate(action.reminderDate))}` : ""}</li>`
    ).join("")
    const title = `اتفاقات-${record.customerNumber}`
    printWindow.document.open()
    printWindow.document.write(`<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>${REPORT_LOGO_PRINT_STYLES}
      @page{size:A4 landscape;margin:14mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;color:#172033;font-size:11px;line-height:1.6}h1{font-size:20px;margin:0 0 4px}h2{font-size:14px;margin:18px 0 6px;border-bottom:1px solid #ccd5df;padding-bottom:4px}.meta{color:#4b5563;margin-bottom:16px}.summary{white-space:pre-wrap;border:1px solid #dbe2ea;padding:8px;border-radius:4px}table{width:100%;border-collapse:collapse;table-layout:auto}th,td{border:1px solid #cbd5e1;padding:5px;text-align:right;vertical-align:top}th{background:#eef2f6;font-weight:bold}ul{padding-right:18px}@media print{body{print-color-adjust:exact;-webkit-print-color-adjust:exact}}
      </style></head><body>${reportLogoMarkup()}<h1>كشف اتفاقات وملاحظات الزبون</h1><div class="meta">${escapeHtml(record.name)} · رقم الزبون: ${escapeHtml(record.customerNumber)} · الرصيد الحالي: ${escapeHtml(formatAmount(record.balance))}</div><h2>ملخص الاتفاق</h2><div class="summary">${escapeHtml(record.note.agreement || "لا يوجد ملخص اتفاق")}</div><h2>سجل الاتفاقات (${agreements.length})</h2><table><thead><tr><th>الاتفاق</th><th>الحالة</th><th>تاريخ الاتفاق</th><th>رصيد البداية</th><th>آخر رصيد</th><th>الدفعة</th><th>نتيجة الإنهاء</th><th>تاريخ الإكمال</th></tr></thead><tbody>${agreementRows || `<tr><td colspan="8">لا توجد اتفاقات مسجلة.</td></tr>`}</tbody></table><h2>تحديثات الرصيد</h2><table><thead><tr><th>الاتفاق</th><th>التاريخ</th><th>الرصيد السابق ← الجديد</th><th>المتوقع</th><th>النتيجة</th><th>المدفوع</th><th>ملاحظة</th></tr></thead><tbody>${checkRows || `<tr><td colspan="7">لا توجد تحديثات مسجلة.</td></tr>`}</tbody></table><h2>بنود وملاحظات الزبون</h2>${actions ? `<ul>${actions}</ul>` : "<p>لا توجد بنود إضافية.</p>"}<script>window.addEventListener('load',()=>{window.focus();window.print()},{once:true})</script></body></html>`)
    printWindow.document.close()
  }

  return (
    <div role="dialog" aria-modal="true" aria-labelledby="customer-agreement-title" className="fixed inset-0 z-[80] overflow-y-auto bg-slate-950/55 p-3 sm:p-6">
      <section dir="rtl" className="mx-auto min-h-[calc(100dvh-1.5rem)] max-w-5xl rounded-3xl bg-[#f4f7fb] p-4 text-slate-900 shadow-2xl sm:min-h-[calc(100dvh-3rem)] sm:p-7">
        <header className="flex flex-wrap items-start justify-between gap-4 rounded-2xl bg-[#10243c] p-5 text-white sm:p-7">
          <div>
            <p className="text-xs text-cyan-100">ملف الزبون · رقم {record.customerNumber}</p>
            <h2 id="customer-agreement-title" className="mt-1 text-xl font-extrabold sm:text-2xl">{record.name}</h2>
            <div className="mt-3 flex flex-wrap gap-x-5 gap-y-1 text-xs text-slate-200">
              <span>الرصيد الحالي: <strong className={record.balanceDirection === "receivable" ? "text-rose-300" : record.balanceDirection === "payable" ? "text-emerald-300" : "text-white"}>{formatAmount(record.balance)}{record.balanceDirection === "receivable" ? " م" : record.balanceDirection === "payable" ? " د" : ""}</strong></span>
              <span>اتفاقات قيد التنفيذ: <strong>{activeAgreements.toLocaleString("ar")}</strong></span>
              <span>الجوال: <strong dir="ltr">{record.phone || "—"}</strong></span>
            </div>
          </div>
          <button type="button" onClick={onClose} className="rounded-xl border border-white/20 px-3 py-2 text-sm font-semibold text-white hover:bg-white/10">إغلاق الملف</button>
        </header>

        <form ref={agreementFormRef} onSubmit={submitAgreement} className="mt-5 space-y-4 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-6">
          <div>
            <h3 className="text-base font-bold">{editingAgreementId ? "تعديل الاتفاق" : "تسجيل اتفاق جديد"}</h3>
            <p className="mt-1 text-xs leading-5 text-slate-500">يحفظ النظام رصيد الزبون وتاريخ الاتفاق كنقطة بداية للمقارنة مع كشوف التحديث. أي نتيجة مقارنة تبقى اقتراحاً حتى تراجعها.</p>
          </div>
          <label className="block space-y-1.5 text-xs font-semibold text-slate-700"><span>تفاصيل الاتفاق</span><textarea value={description} onChange={event => setDescription(event.target.value)} rows={2} required placeholder="مثال: تسديد 50 شيكل كل أسبوع" className="w-full resize-y rounded-xl border border-slate-200 px-3 py-2.5 text-sm font-normal outline-none focus:border-cyan-500 focus:ring-2 focus:ring-cyan-100" /></label>
          <div className="grid gap-3 sm:grid-cols-3">
            <label className="space-y-1.5 text-xs font-semibold text-slate-700"><span>نوع الدفعة</span><select value={installmentAmountType} onChange={event => setInstallmentAmountType(event.target.value as "fixed" | "unknown")} className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm font-normal"><option value="fixed">مبلغ محدد</option><option value="unknown">دفعة غير معلومة</option></select></label>
            {installmentAmountType === "fixed" && <label className="space-y-1.5 text-xs font-semibold text-slate-700"><span>مبلغ الدفعة المتفق عليها</span><input type="number" min="0.01" step="0.01" value={installmentAmount} onChange={event => setInstallmentAmount(event.target.value)} required placeholder="50" className="w-full rounded-xl border border-slate-200 px-3 py-2.5 text-sm font-normal" /></label>}
            <label className="space-y-1.5 text-xs font-semibold text-slate-700"><span>تكرار الدفعة</span><select value={frequency} onChange={event => setFrequency(event.target.value as AgreementFrequency)} className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm font-normal"><option value="weekly">أسبوعياً</option><option value="monthly">شهرياً</option><option value="bimonthly">كل شهرين</option><option value="once">دفعة واحدة</option><option value="flexible">غير دوري / حسب التواصل</option></select></label>
            <label className="space-y-1.5 text-xs font-semibold text-slate-700"><span>{frequency === "flexible" ? "الفترة المتوقعة للدفعة (اختياري)" : "موعد أول دفعة من"}</span><input type="date" value={firstDueDate} onChange={event => setFirstDueDate(event.target.value)} required={frequency !== "flexible"} className="w-full rounded-xl border border-slate-200 px-3 py-2.5 text-sm font-normal" /></label>
            <label className="space-y-1.5 text-xs font-semibold text-slate-700"><span>موعد أول دفعة إلى (اختياري)</span><input type="date" min={firstDueDate || undefined} value={lastDueDate} onChange={event => setLastDueDate(event.target.value)} className="w-full rounded-xl border border-slate-200 px-3 py-2.5 text-sm font-normal" /></label>
          </div>
          <section className="space-y-2 rounded-xl border border-cyan-100 bg-cyan-50/40 p-3"><div className="flex items-center justify-between gap-2"><h4 className="text-xs font-bold text-slate-700">ذكرني بمواعيد متابعة هذا الاتفاق</h4><button type="button" onClick={() => setReminderDates(current => [...current, ""])} className="inline-flex items-center gap-1 rounded-lg border border-cyan-200 bg-white px-2 py-1.5 text-[10px] font-semibold text-cyan-800"><Plus className="h-3 w-3" />موعد تذكير</button></div>{reminderDates.map((date, index) => <div key={`${editingAgreementId ?? "new"}-${index}`} className="flex items-center gap-2"><span className="text-[10px] text-slate-500">تذكير {index + 1}</span><input type="date" value={date} onChange={event => setReminderDates(current => current.map((item, itemIndex) => itemIndex === index ? event.target.value : item))} className="min-w-0 flex-1 rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-xs" /><button type="button" aria-label="حذف التذكير" onClick={() => setReminderDates(current => current.filter((_, itemIndex) => itemIndex !== index))} className="rounded p-1 text-slate-400 hover:bg-rose-50 hover:text-rose-700"><X className="h-4 w-4" /></button></div>)}</section>
          <div className="flex flex-wrap items-center gap-3">
            <button type="submit" className="inline-flex items-center gap-2 rounded-xl bg-cyan-700 px-4 py-2.5 text-sm font-semibold text-white hover:bg-cyan-800"><Plus className="h-4 w-4" />{editingAgreementId ? "حفظ التعديل" : "حفظ الاتفاق"}</button>
            {editingAgreementId && <button type="button" onClick={() => { setEditingAgreementId(null); setDescription(""); setInstallmentAmount(""); setFrequency("weekly"); setFirstDueDate(""); setLastDueDate(""); setReminderDates([]); setInstallmentAmountType("fixed"); setFormError("") }} className="rounded-xl border border-slate-200 px-4 py-2.5 text-sm font-semibold text-slate-600">إلغاء التعديل</button>}
            <span className="text-xs text-slate-500">رصيد بداية الاتفاق: {formatAmount(record.balance)}{record.balanceDirection === "receivable" ? " م" : record.balanceDirection === "payable" ? " د" : ""}</span>
          </div>
          {formError && <p role="alert" className="text-xs font-semibold text-rose-700">{formError}</p>}
        </form>

        <section className="mt-5 space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2"><h3 className="text-base font-bold">سجل الاتفاقات</h3><div className="flex flex-wrap items-center gap-2"><span className="text-xs text-slate-500">{agreements.length.toLocaleString("ar")} اتفاق</span><button type="button" onClick={exportAgreementWorkbook} className="inline-flex items-center gap-1.5 rounded-lg border border-emerald-200 bg-white px-2.5 py-2 text-xs font-semibold text-emerald-800 hover:bg-emerald-50"><FileSpreadsheet className="h-4 w-4" />Excel</button><button type="button" onClick={exportAgreementPdf} className="inline-flex items-center gap-1.5 rounded-lg border border-rose-200 bg-white px-2.5 py-2 text-xs font-semibold text-rose-800 hover:bg-rose-50"><Download className="h-4 w-4" />PDF</button></div></div>
          {!agreements.length && <div className="rounded-2xl border-2 border-dashed border-slate-200 bg-white p-8 text-center text-sm text-slate-500">لا توجد اتفاقات مسجلة لهذا الزبون بعد.</div>}
          {agreements.map(agreement => {
            const statusLabel = agreement.status === "active" ? "قيد التنفيذ" : agreement.status === "completed" ? "مكتمل" : "مرفوض"
            const statusClass = agreement.status === "active" ? "bg-amber-100 text-amber-800" : agreement.status === "completed" ? "bg-emerald-100 text-emerald-800" : "bg-slate-200 text-slate-700"
            return <article key={agreement.id} className="space-y-4 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-6">
              <div className="flex flex-wrap items-start justify-between gap-3">
                <div className="min-w-0 flex-1"><div className="flex flex-wrap items-center gap-2"><h4 className="font-bold">{agreement.description}</h4><span className={`rounded-full px-2.5 py-1 text-[11px] font-semibold ${statusClass}`}>{statusLabel}</span></div><p className="mt-2 text-xs leading-5 text-slate-600">تاريخ الاتفاق: {new Date(agreement.startedAt).toLocaleString("ar")} · رصيد البداية: {formatAmount(agreement.initialBalance)} · الدفعة: {agreement.installmentAmount === null ? "غير معلومة" : formatAmount(agreement.installmentAmount)} {agreement.frequency === "weekly" ? "أسبوعياً" : agreement.frequency === "monthly" ? "شهرياً" : "مرة واحدة"} · الدفعة القادمة: {formatDate(agreement.nextPaymentDate)}{agreement.status === "completed" && agreement.closedAt ? ` · تاريخ الاكتمال: ${new Date(agreement.closedAt).toLocaleString("ar")}` : ""}</p>{agreement.status === "completed" && agreement.closureOutcome && <p className="mt-1 text-xs font-semibold text-slate-700">عند الإنهاء: {agreement.closureOutcome === "paid" ? "دفع كاملاً" : agreement.closureOutcome === "partial" ? "دفع جزئياً" : "لم يسدد"} · المدفوع: {formatAmount(agreement.closurePaidAmount ?? 0)}</p>}</div>
                <div className="flex flex-wrap gap-2"><button type="button" onClick={() => { startEditingAgreement(agreement); agreementFormRef.current?.scrollIntoView({ behavior: "smooth", block: "start" }) }} className="inline-flex items-center gap-1.5 rounded-lg border border-cyan-200 px-3 py-2 text-xs font-semibold text-cyan-800 hover:bg-cyan-50"><Pencil className="h-3.5 w-3.5" />تعديل</button>{agreement.status === "active" && <><button type="button" onClick={() => { setClosureOutcome("paid"); setClosurePaidAmount(String(agreement.outstandingInstallment || agreement.installmentAmount || "")); setClosureError(""); setClosingAgreementId(agreement.id) }} className="rounded-lg bg-emerald-700 px-3 py-2 text-xs font-semibold text-white hover:bg-emerald-800">إنهاء الاتفاق</button><button type="button" onClick={() => { if (window.confirm("رفض / إلغاء هذا الاتفاق؟")) onCloseAgreement(agreement.id, "refused") }} className="rounded-lg border border-slate-300 px-3 py-2 text-xs font-semibold text-slate-700 hover:bg-slate-50">رفض / إلغاء</button></>}<button type="button" onClick={() => void deleteAgreement(agreement)} aria-label="حذف الاتفاق" title="حذف الاتفاق" className="rounded-lg border border-rose-200 p-2 text-rose-700 hover:bg-rose-50"><Trash2 className="h-4 w-4" /></button></div>
              </div>
              {(() => {
                const latestBalance = agreement.balanceSnapshots?.at(-1)
                const latestBalanceValue = agreement.closureBalance ?? latestBalance?.balance ?? agreement.checks.at(-1)?.newBalance ?? agreement.initialBalance
                const latestBalanceDirection = agreement.closureBalanceDirection ?? latestBalance?.direction ?? agreement.checks.at(-1)?.newDirection ?? agreement.balanceDirection
                return <div className="grid gap-3 rounded-xl bg-slate-50 p-3 text-xs sm:grid-cols-4">
                <span>الرصيد وقت الاتفاق: <strong>{formatAmount(agreement.initialBalance)}</strong></span>
                <span>الرصيد المسجل أخيراً: <strong>{formatAmount(latestBalanceValue)}{latestBalanceDirection === "receivable" ? " م" : latestBalanceDirection === "payable" ? " د" : ""}</strong></span>
                <span>فرق الرصيد منذ البداية: <strong className={agreement.initialBalance > latestBalanceValue ? "text-emerald-700" : "text-slate-700"}>{formatAmount(Math.max(0, agreement.initialBalance - latestBalanceValue))}</strong></span>
                <span>المتبقي من الدفعة: <strong>{agreement.installmentAmount === null ? "غير محدد" : formatAmount(agreement.outstandingInstallment)}</strong></span>
              </div>
              })()}
              {!!agreement.checks.length && <div className="space-y-2"><h5 className="text-xs font-bold text-slate-700">مقارنة تحديثات الكشف</h5>{[...agreement.checks].reverse().map(check => {
                const recommendation = check.suggestion === "paid" ? "يبدو أن الدفعة المستحقة سُددت" : check.suggestion === "partial" ? "يبدو أن الدفعة سُددت جزئياً" : check.suggestion === "unpaid" ? "لم يظهر انخفاض في الرصيد للدفعة المستحقة" : "لا يمكن استنتاج التسديد من هذا التغيير"
                const outcomeText = check.outcome === "paid" ? "تم تأكيد الدفعة" : check.outcome === "partial" ? "تم تأكيد دفع جزئي" : check.outcome === "unpaid" ? "تم تأكيد عدم التسديد" : check.outcome === "ignored" ? "تم تجاهل الاقتراح" : "بانتظار مراجعتك"
                const paidAmount = check.actualPaid ?? check.inferredPaid
                return <div key={check.id} className="space-y-3 rounded-xl border border-slate-200 p-3">
                  <div className="flex flex-wrap items-start justify-between gap-2"><div><p className="text-xs font-semibold">{new Date(check.checkedAt).toLocaleString("ar")}</p><p className="mt-1 text-xs text-slate-600">الرصيد: {formatAmount(check.previousBalance)} ← {formatAmount(check.newBalance)} · انخفاض مستنتج: {formatAmount(check.inferredPaid)} · المتوقع: {agreement.installmentAmount === null ? "غير محدد" : formatAmount(check.expectedPayment)}{check.outcome ? ` · المبلغ المعتمد: ${formatAmount(check.actualPaid ?? 0)}` : ""}{check.waivedAmount !== undefined ? ` · المتجاوز عنه: ${formatAmount(check.waivedAmount)}` : ""}</p><p className={`mt-1 text-xs font-semibold ${check.suggestion === "paid" ? "text-emerald-700" : check.suggestion === "partial" ? "text-amber-700" : check.suggestion === "unpaid" ? "text-rose-700" : "text-slate-600"}`}>{recommendation} · {outcomeText}</p></div></div>
                  {agreement.status === "active" && !check.outcome && <div className="flex flex-wrap items-end gap-2">
                    <label className="space-y-1 text-[10px] font-semibold text-slate-600"><span>المبلغ الذي تؤكد تسديده</span><input type="number" min="0" step="0.01" value={check.actualPaid ?? check.inferredPaid} onChange={event => onUpdateAgreement(agreement.id, current => ({ ...current, checks: current.checks.map(item => item.id === check.id ? { ...item, actualPaid: Number(event.target.value) || 0 } : item) }))} className="block w-36 rounded-lg border border-slate-200 px-2 py-2 text-xs font-normal" /></label>
                    <label className="min-w-[180px] flex-1 space-y-1 text-[10px] font-semibold text-slate-600"><span>ملاحظة المراجعة (اختياري)</span><input value={decisionNotes[check.id] ?? check.note ?? ""} onChange={event => setDecisionNotes(current => ({ ...current, [check.id]: event.target.value }))} className="block w-full rounded-lg border border-slate-200 px-2 py-2 text-xs font-normal" /></label>
                    <button type="button" disabled={agreement.installmentAmount === null && paidAmount <= 0} onClick={() => onSetCheckOutcome(agreement.id, check.id, "paid", Math.max(check.expectedPayment, paidAmount), decisionNotes[check.id] ?? "")} className="rounded-lg bg-emerald-700 px-3 py-2 text-xs font-semibold text-white hover:bg-emerald-800 disabled:cursor-not-allowed disabled:opacity-50">تأكيد تم</button>
                    <button type="button" disabled={agreement.installmentAmount === null && paidAmount <= 0} onClick={() => onSetCheckOutcome(agreement.id, check.id, "partial", paidAmount, decisionNotes[check.id] ?? "")} className="rounded-lg border border-amber-300 px-3 py-2 text-xs font-semibold text-amber-800 hover:bg-amber-50 disabled:cursor-not-allowed disabled:opacity-50">تم جزئياً</button>
                    <button type="button" onClick={() => onSetCheckOutcome(agreement.id, check.id, "unpaid", 0, decisionNotes[check.id] ?? "")} className="rounded-lg border border-rose-200 px-3 py-2 text-xs font-semibold text-rose-700 hover:bg-rose-50">لم يسدد</button>
                    <button type="button" onClick={() => onSetCheckOutcome(agreement.id, check.id, "ignored", 0, decisionNotes[check.id] ?? "")} className="rounded-lg border border-slate-200 px-3 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-50">تجاهل الاقتراح</button>
                  </div>}
                  {agreement.installmentAmount !== null && agreement.status === "active" && agreement.checks.at(-1)?.id === check.id && check.outcome === "partial" && check.waivedAmount === undefined && <button type="button" onClick={() => onUpdateAgreement(agreement.id, current => ({
                    ...current,
                    outstandingInstallment: 0,
                    checks: current.checks.map(item => item.id === check.id
                      ? { ...item, waivedAmount: Math.max(0, check.expectedPayment - (check.actualPaid ?? 0)) }
                      : item),
                  }))} className="rounded-lg border border-cyan-300 bg-cyan-50 px-3 py-2 text-xs font-semibold text-cyan-900 hover:bg-cyan-100">قبول الدفعة الجزئية وتجاوز المتبقي ({formatAmount(Math.max(0, check.expectedPayment - (check.actualPaid ?? 0)))})</button>}
                  {check.note && <p className="text-xs text-slate-600">ملاحظة: {check.note}</p>}
                </div>
              })}</div>}
            </article>
          })}
        </section>
        {closingAgreement && <div role="dialog" aria-modal="true" aria-labelledby="agreement-close-title" className="fixed inset-0 z-[90] flex items-center justify-center bg-slate-950/55 p-4"><form onSubmit={submitAgreementClosure} className="w-full max-w-md space-y-4 rounded-2xl bg-white p-5 shadow-2xl"><div><h3 id="agreement-close-title" className="font-bold">تسجيل نتيجة إنهاء الاتفاق</h3><p className="mt-1 text-xs text-slate-500">{closingAgreement.description}</p></div><label className="block space-y-1.5 text-xs font-semibold text-slate-700"><span>حالة التسديد</span><select value={closureOutcome} onChange={event => { const outcome = event.target.value as Exclude<AgreementCheckOutcome, "ignored">; setClosureOutcome(outcome); if (outcome === "unpaid") setClosurePaidAmount("0"); else if (closurePaidAmount === "0") setClosurePaidAmount(String(closingAgreement.outstandingInstallment || closingAgreement.installmentAmount || "")) }} className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2.5"><option value="paid">دفع كامل</option><option value="partial">دفع جزئي</option><option value="unpaid">لم يدفع</option></select></label>{closureOutcome !== "unpaid" && <label className="block space-y-1.5 text-xs font-semibold text-slate-700"><span>كم دفع؟</span><input type="number" min="0.01" step="0.01" required value={closurePaidAmount} onChange={event => setClosurePaidAmount(event.target.value)} className="w-full rounded-lg border border-slate-200 px-3 py-2.5 font-normal" /></label>}<p className="text-xs text-slate-600">رصيد الزبون عند الإنهاء: <strong>{formatAmount(record.balance)}{record.balanceDirection === "receivable" ? " م" : record.balanceDirection === "payable" ? " د" : ""}</strong></p>{closureError && <p role="alert" className="text-xs font-semibold text-rose-700">{closureError}</p>}<div className="flex justify-end gap-2"><button type="button" onClick={() => setClosingAgreementId(null)} className="rounded-lg border border-slate-300 px-3 py-2 text-xs font-semibold">إلغاء</button><button type="submit" className="rounded-lg bg-emerald-700 px-3 py-2 text-xs font-semibold text-white">حفظ وإنهاء الاتفاق</button></div></form></div>}
        {(record.note.agreement || record.note.actions.length > 0) && <section className="mt-5 rounded-2xl border border-slate-200 bg-white p-4"><h3 className="font-bold">ملاحظات واتفاقات سابقة</h3>{record.note.agreement && <p className="mt-2 whitespace-pre-wrap text-sm">{record.note.agreement}</p>}{record.note.actions.map((action, index) => <p key={action.id} className="mt-2 text-xs">{(index + 1).toLocaleString("ar")}. {action.description} · {action.status === "completed" || action.completed ? "تم" : action.status === "failed" ? `لم ينفذ: ${action.failureReason ?? ""}` : "قيد المتابعة"}</p>)}</section>}
      </section>
    </div>
  )
}

type CustomerImportField = Exclude<
  keyof CustomerBalance,
  | "note"
  | "paymentClassification"
  | "isInLatestStatement"
  | "statementAppearances"
  | "lastSeenInStatementAt"
  | "lastStatementBalance"
  | "lastStatementDirection"
  | "balanceDirection"
  | "firstStatementBalance"
  | "firstStatementDirection"
  | "firstStatementAt"
  | "balanceUpdatedAt"
>

type CustomerDataset = {
  fileName: string
  importedAt: string
  records: CustomerBalance[]
}

export type MissingCustomerPolicy = "keep_zero" | "hide_keep_history"

type DueReminder = {
  customerNumber: string
  customerName: string
  action: CustomerAction
  reminderDate: string
  source: "action" | "agreement"
  agreementId?: string
}

type Filters = {
  query: string
  searchField: "all" | "customerNumber" | "phone"
  notes: "all" | "withNotes"
  agreementStatus: "all" | "active"
  noteOrder: "default" | "first" | "dueSoon"
  dueFrom: string
  dueTo: string
  balanceOperator: "min" | "max" | "equal"
  balance: string
  balanceFrom: string
  balanceTo: string
  durationOperator: "min" | "max" | "equal"
  duration: string
  durationFrom: string
  durationTo: string
  receiptFrom: string
  receiptTo: string
  invoiceFrom: string
  invoiceTo: string
}

const STORAGE_KEY = "customer_balances_dataset"
const LEGACY_STORAGE_KEY = STORAGE_KEY
const DISPLAY_COLUMNS: Array<{ key: CustomerImportField; label: string }> = [
  { key: "customerNumber", label: "رقم الزبون" },
  { key: "name", label: "اسم الزبون / البيان" },
  { key: "balance", label: "الرصيد" },
  { key: "duration", label: "المدة" },
  { key: "invoiceAmount", label: "قيمة آخر فاتورة" },
  { key: "invoiceDate", label: "تاريخ آخر فاتورة" },
  { key: "receiptAmount", label: "قيمة آخر سند قبض" },
  { key: "receiptDate", label: "تاريخ آخر سند قبض" },
  { key: "phone", label: "الجوال" },
]

const COLUMN_ALIASES: Record<CustomerImportField, string[]> = {
  customerNumber: ["رقم الصفحة", "رقم الزبون", "رقم العميل", "رقم الحساب", "رقم", "الرقم", "page number", "customer number", "account number", "number", "no"],
  name: ["البيان", "اسم الزبون", "اسم العميل", "اسم الزبون البيان", "اسم الحساب", "الاسم", "customer name", "client name", "name", "description"],
  balance: ["الرصيد", "رصيد الزبون", "رصيد العميل", "الرصيد الحالي", "الرصيد المستحق", "balance", "current balance", "amount due"],
  duration: ["المدة", "مدة", "مدة الرصيد", "مدة الدين", "أيام التأخير", "عدد الأيام", "العمر", "days", "duration", "age"],
  invoiceAmount: ["قيمة آخر فاتورة", "قيمة اخر فاتورة", "قيمة اخر فاتوره", "آخر فاتورة", "اخر فاتورة", "مبلغ آخر فاتورة", "مبلغ الفاتورة الأخيرة", "قيمة الفاتورة الأخيرة", "last invoice amount", "invoice amount", "last invoice"],
  invoiceDate: ["تاريخ آخر فاتورة", "تاريخ اخر فاتورة", "تاريخ الفاتورة الأخيرة", "تاريخ الفاتورة", "last invoice date", "invoice date"],
  receiptAmount: ["قيمة آخر سند قبض", "قيمة اخر سند قبض", "قيمة اخر سند القبض", "آخر سند قبض", "اخر سند قبض", "مبلغ آخر سند قبض", "مبلغ سند القبض الأخير", "قيمة آخر قبض", "last receipt amount", "receipt amount", "last receipt"],
  receiptDate: ["تاريخ آخر سند قبض", "تاريخ اخر سند قبض", "تاريخ سند القبض الأخير", "تاريخ آخر قبض", "تاريخ القبض", "last receipt date", "receipt date"],
  phone: ["الجوال", "رقم الجوال", "الهاتف", "رقم الهاتف", "الموبايل", "رقم الموبايل", "mobile", "mobile number", "phone", "phone number"],
}

const defaultFilters: Filters = {
  query: "",
  searchField: "all",
  notes: "all",
  agreementStatus: "all",
  noteOrder: "default",
  dueFrom: "",
  dueTo: "",
  balanceOperator: "min",
  balance: "",
  balanceFrom: "",
  balanceTo: "",
  durationOperator: "min",
  duration: "",
  durationFrom: "",
  durationTo: "",
  receiptFrom: "",
  receiptTo: "",
  invoiceFrom: "",
  invoiceTo: "",
}

function normalizeHeader(value: unknown): string {
  return String(value ?? "")
    .toLowerCase()
    .normalize("NFKC")
    .replace(/[\u064b-\u065f\u0670\u0640\u200e\u200f\u202a-\u202e\u2066-\u2069\ufeff]/g, "")
    .replace(/[أإآ]/g, "ا")
    .replace(/[ى]/g, "ي")
    .replace(/[_-]+/g, " ")
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim()
}

const normalizedAliases = Object.fromEntries(
  Object.entries(COLUMN_ALIASES).map(([key, aliases]) => [
    key,
    new Set(aliases.map(normalizeHeader)),
  ]),
) as Record<CustomerImportField, Set<string>>

function findHeader(
  row: unknown[],
  nextRow?: unknown[],
): { columns: Partial<Record<CustomerImportField, number>>; rowCount: number } | null {
  const matchesRequiredColumns = (columns: Partial<Record<CustomerImportField, number>>) =>
    columns.customerNumber !== undefined &&
    columns.name !== undefined &&
    columns.balance !== undefined &&
    Object.keys(columns).length >= 6
  const findColumns = (headerRow: unknown[]) => {
    const columns: Partial<Record<CustomerImportField, number>> = {}
    for (const [index, value] of headerRow.entries()) {
      const header = normalizeHeader(value)
      if (!header) continue
      for (const key of Object.keys(normalizedAliases) as CustomerImportField[]) {
        if (columns[key] === undefined && normalizedAliases[key].has(header)) columns[key] = index
      }
    }
    return columns
  }
  const singleRowColumns = findColumns(row)
  if (matchesRequiredColumns(singleRowColumns)) return { columns: singleRowColumns, rowCount: 1 }
  if (!nextRow) return null
  const combinedRow = row.map((value, index) =>
    [value, nextRow[index]].filter(item => String(item ?? "").trim()).join(" "),
  )
  const combinedColumns = findColumns(combinedRow)
  return matchesRequiredColumns(combinedColumns) ? { columns: combinedColumns, rowCount: 2 } : null
}

function arabicDigits(value: string): string {
  return value
    .replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
    .replace(/[۰-۹]/g, digit => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)))
}

function cellString(value: unknown): string {
  if (value == null) return ""
  return arabicDigits(String(value))
    .replace(/[\u200b\u200e\u200f\u202a-\u202e\u2066-\u2069\ufeff]/g, "")
    .replace(/\s+/g, " ")
    .trim()
}

function whatsappNumberDigits(phone: string): string {
  let digits = arabicDigits(phone).replace(/\D/g, "")
  if (digits.startsWith("00")) digits = digits.slice(2)
  if (digits.startsWith("05") && digits.length === 10) {
    digits = `970${digits.slice(1)}`
  } else if (digits.startsWith("5") && digits.length === 9) {
    digits = `970${digits}`
  }
  return digits
}

function phoneSearchDigits(phone: string): string {
  let digits = arabicDigits(phone).replace(/\D/g, "")
  if (digits.startsWith("00")) digits = digits.slice(2)
  if (digits.startsWith("970") || digits.startsWith("972")) digits = digits.slice(3)
  if (digits.startsWith("0")) digits = digits.slice(1)
  return digits
}

function whatsappUrl(phone: string, message?: string, palestinianCode: "970" | "972" = "972"): string | null {
  let digits = whatsappNumberDigits(phone)
  if (digits.startsWith("970") || digits.startsWith("972")) {
    digits = `${palestinianCode}${digits.slice(3)}`
  }
  return digits.length >= 8 && digits.length <= 15
    ? `https://wa.me/${digits}${message ? `?text=${encodeURIComponent(message)}` : ""}`
    : null
}

function whatsappFallbackUrl(phone: string, message?: string): string | null {
  return whatsappNumberDigits(phone).startsWith("970") || whatsappNumberDigits(phone).startsWith("972")
    ? whatsappUrl(phone, message, "970")
    : null
}

async function createCustomerDebtImage(
  record: Pick<CustomerBalance, "name" | "balance" | "balanceDirection">,
  template: CustomerCardTemplate = "modern",
  message = "",
): Promise<File> {
  const width = 1080
  const height = 1350
  const canvas = document.createElement("canvas")
  canvas.width = width
  canvas.height = height
  const context = canvas.getContext("2d")
  if (!context) throw new Error("تعذر تجهيز مساحة إنشاء الصورة.")

  if (template === "classic") {
    context.fillStyle = "#FAF7F2"
    context.fillRect(0, 0, width, height)
    context.fillStyle = "#38A169"
    context.fillRect(0, 0, 18, height)
    context.fillStyle = "#D92A2A"
    context.fillRect(width - 18, 0, 18, height)
    context.fillStyle = "#FAF7F2"
    context.beginPath()
    context.roundRect(48, 48, width - 96, height - 96, 38)
    context.strokeStyle = "#E5E0D8"
    context.lineWidth = 3
    context.stroke()
  } else {
    const background = context.createLinearGradient(0, 0, width, height)
    background.addColorStop(0, "#fffaf0")
    background.addColorStop(0.55, "#f8f1e2")
    background.addColorStop(1, "#f3ead8")
    context.fillStyle = background
    context.fillRect(0, 0, width, height)
    context.fillStyle = "#c53c35"
    context.fillRect(0, 0, 18, height)
    context.fillStyle = "#55a858"
    context.fillRect(width - 18, 0, 18, height)
    context.fillStyle = "rgba(126,64,42,0.08)"
    context.beginPath()
    context.roundRect(38, 39, width - 76, height - 78, 52)
    context.fill()
    context.fillStyle = "#fffefa"
    context.beginPath()
    context.roundRect(45, 45, width - 90, height - 90, 46)
    context.fill()
    context.strokeStyle = "#d8c9a7"
    context.lineWidth = 2
    context.stroke()
    context.strokeStyle = "#eee4d1"
    context.lineWidth = 1
    context.beginPath()
    context.roundRect(60, 60, width - 120, height - 120, 38)
    context.stroke()
    context.fillStyle = "#c43b35"
    context.beginPath()
    context.roundRect(width / 2 - 86, 40, 172, 12, 6)
    context.fill()
    context.fillStyle = "#57a65a"
    context.beginPath()
    context.roundRect(width / 2 - 30, 40, 60, 12, 6)
    context.fill()
  }

  const logo = new Image()
  logo.src = "/beit-aljumla-logo.png"
  await new Promise<void>((resolve, reject) => {
    logo.onload = () => resolve()
    logo.onerror = () => reject(new Error("تعذر تحميل شعار المنصة لإنشاء صورة الزبون."))
  })
  const logoWidth = 238
  const logoHeight = logoWidth * (logo.naturalHeight / logo.naturalWidth)
  const logoX = template === "classic" ? (width - logoWidth) / 2 : (width - logoWidth) / 2 - 35
  const logoY = template === "classic" ? 72 : 65
  context.drawImage(logo, logoX, logoY, logoWidth, logoHeight)
  if (template === "classic") {
    context.direction = "rtl"
    context.textAlign = "right"
    const customerName = record.name.trim() || "زبوننا العزيز"
    const right = width - 78
    const contentWidth = width - 156

    context.fillStyle = "#1A202C"
    context.font = "bold 32px Arial, sans-serif"
    const greeting = `الزبون الكريم: ${customerName}`
    let messageStartY = 355
    if (context.measureText(greeting).width <= contentWidth) {
      context.fillText(greeting, right, 295)
    } else {
      context.fillText("الزبون الكريم:", right, 280)
      const nameLines: string[] = []
      let nameLine = ""
      for (const word of customerName.split(/\s+/)) {
        const nextLine = nameLine ? `${nameLine} ${word}` : word
        if (nameLine && context.measureText(nextLine).width > contentWidth) {
          nameLines.push(nameLine)
          nameLine = word
        } else {
          nameLine = nextLine
        }
      }
      if (nameLine) nameLines.push(nameLine)
      nameLines.slice(0, 2).forEach((line, index) => context.fillText(line, right, 320 + index * 36, contentWidth))
      messageStartY = 405
    }

    const messageLines: string[] = []
    context.font = "24px Arial, sans-serif"
    const paragraphs = message
      .split(/\r?\n/)
      .filter(paragraph => !/^\s*(الزبون الكريم\s*[:：]|(?:بيانات الدفع|اسم صاحب الحساب|رقم حساب|حساب بنك فلسطين|جوال بنك فلسطين|بنك فلسطين|بال بي|PalPay|التحويل على رقم حساب|رقم دفع))/i.test(paragraph.trim()))
    for (const paragraph of paragraphs) {
      if (!paragraph.trim()) {
        messageLines.push("")
        continue
      }
      let line = ""
      for (const word of paragraph.replace(/\*([^*]+)\*/g, "$1").split(/\s+/)) {
        const nextLine = line ? `${line} ${word}` : word
        if (line && context.measureText(nextLine).width > contentWidth) {
          messageLines.push(line)
          line = word
        } else {
          line = nextLine
        }
      }
      if (line) messageLines.push(line)
    }
    const visibleMessageLines = messageLines.slice(0, Math.min(8, Math.floor((660 - messageStartY) / 37)))
    if (messageLines.length > visibleMessageLines.length && visibleMessageLines.length) {
      visibleMessageLines[visibleMessageLines.length - 1] = `${visibleMessageLines.at(-1)?.trimEnd()} …`
    }
    context.fillStyle = "#4A5568"
    visibleMessageLines.forEach((line, index) => {
      if (line) context.fillText(line, right, messageStartY + index * 37, contentWidth)
    })

    const balanceLabel = record.balanceDirection === "receivable"
      ? "الرصيد المستحق لنا"
      : record.balanceDirection === "payable"
        ? "الرصيد المستحق لك"
        : "الرصيد المسجل"
    context.fillStyle = "#FFF5F5"
    context.beginPath()
    context.roundRect(92, 690, width - 184, 175, 26)
    context.fill()
    context.strokeStyle = "#FEB2B2"
    context.lineWidth = 3
    context.stroke()
    context.textAlign = "center"
    context.fillStyle = "#9B2C2C"
    context.font = "bold 28px Arial, sans-serif"
    context.fillText(balanceLabel, width / 2, 742)
    context.font = "bold 58px Arial, sans-serif"
    context.fillText(`${formatAmount(record.balance)} شيكل`, width / 2, 812, width - 240)

    context.fillStyle = "#4A5568"
    context.textAlign = "right"
    context.font = "24px Arial, sans-serif"
    context.fillText("يسعدنا جداً زيارتك لنا في الفرع لترتيب وتسوية الحساب،", right, 920, contentWidth)
    context.fillText("ونكون دائماً في أشد السعادة بلقائك. دمت بخير وعافية!", right, 956, contentWidth)

    context.fillStyle = "#FFFFFF"
    context.beginPath()
    context.roundRect(78, 990, width - 156, 270, 22)
    context.fill()
    context.strokeStyle = "#CBD5E0"
    context.lineWidth = 2
    context.stroke()
    context.textAlign = "center"
    context.fillStyle = "#9B2C2C"
    context.font = "bold 27px Arial, sans-serif"
    context.fillText("بيانات الدفع والتحويل", width / 2, 1035)
    context.textAlign = "right"
    context.fillStyle = "#2D3748"
    context.font = "23px Arial, sans-serif"
    context.fillText("اسم صاحب الحساب: محمود زكي عثمان", right - 14, 1085, contentWidth - 20)
    context.fillText("التحويل على رقم حساب: 1083594", right - 14, 1129, contentWidth - 20)
    context.fillText("بنك فلسطين (جوال): 0594466633", right - 14, 1173, contentWidth - 20)
    context.fillText("تاجر بال بي (PalPay): 0566105106", right - 14, 1217, contentWidth - 20)

    context.textAlign = "center"
    context.fillStyle = "#2D3748"
    context.font = "bold 25px Arial, sans-serif"
    context.fillText("سوبر ماركت بيت الجملة", width / 2, 1305)
    const blob = await new Promise<Blob | null>(resolve => canvas.toBlob(resolve, "image/png"))
    if (!blob) throw new Error("تعذر إنشاء ملف الصورة.")
    const safeName = customerName.replace(/[^\p{L}\p{N}-]+/gu, "_").slice(0, 40) || "customer"
    return new File([blob], `balance-${safeName}.png`, { type: "image/png" })
  }

  const taswiyahLogo = new Image()
  taswiyahLogo.src = "/taswiyah-logo.png"
  await new Promise<void>((resolve, reject) => {
    taswiyahLogo.onload = () => resolve()
    taswiyahLogo.onerror = () => reject(new Error("تعذر تحميل شعار تسوية لإنشاء صورة الزبون."))
  })
  const taswiyahWidth = 100
  const taswiyahHeight = taswiyahWidth * (taswiyahLogo.naturalHeight / taswiyahLogo.naturalWidth)
  context.drawImage(taswiyahLogo, width / 2 + 150, 100, taswiyahWidth, taswiyahHeight)

  context.direction = "rtl"
  context.textAlign = "center"
  if (template === "modern") {
    const drawBasket = (x: number, y: number, scale: number) => {
      context.save()
      context.translate(x, y)
      context.scale(scale, scale)
      context.strokeStyle = "#d8c9a7"
      context.lineWidth = 4
      context.lineCap = "round"
      context.lineJoin = "round"
      context.beginPath()
      context.moveTo(-22, -10)
      context.lineTo(-16, 12)
      context.lineTo(16, 12)
      context.lineTo(23, -10)
      context.closePath()
      context.moveTo(-28, -18)
      context.lineTo(-20, -18)
      context.moveTo(-10, -2)
      context.lineTo(-8, 7)
      context.moveTo(0, -2)
      context.lineTo(0, 7)
      context.moveTo(10, -2)
      context.lineTo(8, 7)
      context.stroke()
      context.restore()
    }
    const drawGift = (x: number, y: number, scale: number) => {
      context.save()
      context.translate(x, y)
      context.scale(scale, scale)
      context.strokeStyle = "#d8c9a7"
      context.lineWidth = 4
      context.lineCap = "round"
      context.lineJoin = "round"
      context.strokeRect(-19, -5, 38, 25)
      context.beginPath()
      context.moveTo(-23, -5)
      context.lineTo(23, -5)
      context.lineTo(23, 2)
      context.lineTo(-23, 2)
      context.closePath()
      context.moveTo(0, -5)
      context.lineTo(0, 20)
      context.ellipse(-8, -11, 8, 5, -0.5, Math.PI, 0)
      context.ellipse(8, -11, 8, 5, 0.5, Math.PI, 0)
      context.stroke()
      context.restore()
    }
    drawBasket(105, 270, 1.25)
    drawGift(967, 275, 1.15)
    drawGift(108, 1110, 1.1)
    drawBasket(968, 1100, 1.2)
    drawBasket(965, 475, 0.8)
  }
  context.fillStyle = "#172033"
  context.font = "bold 46px Arial, sans-serif"
  context.fillText("الزبون الكريم", width / 2, 330)
  context.font = "bold 48px Arial, sans-serif"
  const customerName = record.name.trim() || "زبوننا العزيز"
  const maxNameWidth = 880
  const nameWords = customerName.split(/\s+/)
  const nameLines: string[] = []
  let currentLine = ""
  for (const word of nameWords) {
    const nextLine = currentLine ? `${currentLine} ${word}` : word
    if (currentLine && context.measureText(nextLine).width > maxNameWidth) {
      nameLines.push(currentLine)
      currentLine = word
    } else {
      currentLine = nextLine
    }
  }
  if (currentLine) nameLines.push(currentLine)
  nameLines.slice(0, 2).forEach((line, index) => context.fillText(line, width / 2, 398 + index * 58, maxNameWidth))

  const balanceLabel = record.balanceDirection === "receivable"
    ? "الرصيد المستحق لنا"
    : record.balanceDirection === "payable"
      ? "الرصيد المستحق لك"
      : "الرصيد المسجل"
  context.fillStyle = "#5b5142"
  context.font = "bold 28px Arial, sans-serif"
  const messageLines: string[] = []
  for (const paragraph of message.split(/\r?\n/)) {
    if (!paragraph.trim()) {
      messageLines.push("")
      continue
    }
    let line = ""
    for (const word of paragraph.split(/\s+/)) {
      const nextLine = line ? `${line} ${word}` : word
      if (line && context.measureText(nextLine).width > 880) {
        messageLines.push(line)
        line = word
      } else {
        line = nextLine
      }
    }
    if (line) messageLines.push(line)
  }
  const visibleMessageLines = messageLines.slice(0, 8)
  if (messageLines.length > visibleMessageLines.length && visibleMessageLines.length) {
    visibleMessageLines[visibleMessageLines.length - 1] = `${visibleMessageLines.at(-1)?.replace(/\s+$/, "")} …`
  }
  visibleMessageLines.forEach((line, index) => {
    if (line) context.fillText(line, width / 2, 510 + index * 42, 900)
  })
  const balanceCardY = Math.max(730, 510 + visibleMessageLines.length * 42 + 44)
  const balanceCardHeight = 190
  context.fillStyle = "#fffaf0"
  context.beginPath()
  context.roundRect(130, balanceCardY, width - 260, balanceCardHeight, 34)
  context.fill()
  context.strokeStyle = "#ccb99b"
  context.lineWidth = 4
  context.stroke()
  context.fillStyle = "#7e302e"
  context.font = "bold 34px Arial, sans-serif"
  context.fillText(balanceLabel, width / 2, balanceCardY + 58)
  context.fillStyle = "#a92f2c"
  context.font = "bold 62px Arial, sans-serif"
  context.fillText(`${formatAmount(record.balance)} شيكل`, width / 2, balanceCardY + 145, width - 300)

  context.fillStyle = "#7e302e"
  context.font = "bold 30px Arial, sans-serif"
  context.fillText("نشكرك على ثقتك · سوبر ماركت بيت الجملة", width / 2, Math.max(1030, balanceCardY + 285))
  context.fillStyle = "#756c5e"
  context.font = "25px Arial, sans-serif"
  context.fillText("أهلاً بك دائماً، ونتمنى لك يوماً سعيداً", width / 2, Math.max(1075, balanceCardY + 330))

  const blob = await new Promise<Blob | null>(resolve => canvas.toBlob(resolve, "image/png"))
  if (!blob) throw new Error("تعذر إنشاء ملف الصورة.")
  const safeName = customerName.replace(/[^\p{L}\p{N}-]+/gu, "_").slice(0, 40) || "customer"
  return new File([blob], `balance-${safeName}.png`, { type: "image/png" })
}

function statementMessage(record: CustomerBalance, template: string): string {
  const balanceLabel = record.balanceDirection === "receivable"
    ? "مستحق لنا"
    : record.balanceDirection === "payable"
      ? "لصالحك"
      : "مسجل"
  return template
    .replaceAll("{{name}}", record.name)
    .replaceAll("{{customerNumber}}", record.customerNumber)
    .replaceAll("{{balance}}", formatAmount(record.balance))
    .replaceAll("{{balanceLabel}}", balanceLabel)
    .replaceAll("{{due}}", balanceLabel)
    .replaceAll("{{duration}}", record.duration.toLocaleString("ar"))
    .replaceAll("{{phone}}", record.phone)
    .replaceAll("{{invoiceAmount}}", formatAmount(record.invoiceAmount))
    .replaceAll("{{invoiceDate}}", formatDate(record.invoiceDate))
    .replaceAll("{{receiptAmount}}", formatAmount(record.receiptAmount))
    .replaceAll("{{receiptDate}}", formatDate(record.receiptDate))
}

function activeWeeklyAgreement(record: CustomerBalance): CustomerAgreement | undefined {
  return record.note.agreements.find(agreement =>
    agreement.status === "active" && agreement.frequency === "weekly",
  )
}

function weeklyInstallmentAmount(record: CustomerBalance): number | undefined {
  const agreement = activeWeeklyAgreement(record)
  const amount = agreement?.installmentAmount
  if (typeof amount === "number" && Number.isFinite(amount) && amount > 0) return amount
  const outstanding = agreement?.outstandingInstallment
  return typeof outstanding === "number" && Number.isFinite(outstanding) && outstanding > 0
    ? outstanding
    : undefined
}

function paymentClassificationLabel(classification?: PaymentClassification): string {
  return PAYMENT_CLASSIFICATIONS.find(item => item.value === (classification ?? null))?.label ?? "غير مصنف"
}

function normalizePaymentCriteria(value: unknown): PaymentCriteria {
  if (!value || typeof value !== "object") throw new Error("إعدادات معايير الدفع المحفوظة غير صالحة.")
  const criteria = value as Partial<PaymentCriteria>
  const normalized = { ...DEFAULT_PAYMENT_CRITERIA, ...criteria }
  if (!isPaymentCriteriaValid(normalized)) throw new Error("تحقق من أن نسب الدفع تصاعدية، وأن حدود الائتمان موجبة، وأيام حداثة السند بين 1 و3650.")
  return normalized
}

function isPaymentCriteriaValid(criteria: PaymentCriteria): boolean {
  return Object.values(criteria).every(number => typeof number === "number" && Number.isFinite(number)) &&
    criteria.acceptableMinPercent > 0 &&
    criteria.goodMinPercent > criteria.acceptableMinPercent &&
    criteria.excellentMinPercent > criteria.goodMinPercent &&
    criteria.excellentMinPercent <= 1000 &&
    criteria.acceptableMinPercent <= 1000 &&
    criteria.goodMinPercent <= 1000 &&
    criteria.acceptableCreditLimitPercent >= 0 &&
    criteria.goodCreditLimitPercent >= 0 &&
    criteria.excellentCreditLimitPercent >= 0 &&
    criteria.acceptableCreditLimitPercent <= 1000 &&
    criteria.goodCreditLimitPercent <= 1000 &&
    criteria.excellentCreditLimitPercent <= 1000 &&
    Number.isInteger(criteria.recentReceiptDays) &&
    criteria.recentReceiptDays >= 1 &&
    criteria.recentReceiptDays <= 3650
}

function assessCustomerPayment(
  record: CustomerBalance,
  criteria: PaymentCriteria,
  now = new Date(),
): PaymentAssessment {
  if (record.balanceDirection !== "receivable" || record.balance <= 0) {
    return {
      classification: "not_applicable",
      paymentPercent: null,
      creditLimit: null,
      invoiceWarning: false,
      recentReceiptDowngrade: false,
    }
  }

  const paymentPercent = Math.max(0, record.receiptAmount) / record.balance * 100
  let classification: PaymentClassification = paymentPercent < criteria.acceptableMinPercent
    ? "poor"
    : paymentPercent < criteria.goodMinPercent
      ? "acceptable"
      : paymentPercent < criteria.excellentMinPercent
        ? "good"
        : "excellent"
  const receiptTime = record.receiptDate ? new Date(`${record.receiptDate}T00:00:00`).getTime() : Number.NaN
  const elapsedDays = Number.isFinite(receiptTime)
    ? Math.floor((new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime() - receiptTime) / 86_400_000)
    : Number.POSITIVE_INFINITY
  const recentReceiptDowngrade = elapsedDays >= 0 && elapsedDays < criteria.recentReceiptDays
  if (recentReceiptDowngrade) {
    classification = classification === "excellent" ? "good"
      : classification === "good" ? "acceptable"
        : classification === "acceptable" ? "poor"
          : "poor"
  }

  const creditPercent = classification === "excellent"
    ? criteria.excellentCreditLimitPercent
    : classification === "good"
      ? criteria.goodCreditLimitPercent
      : classification === "acceptable"
        ? criteria.acceptableCreditLimitPercent
        : 0
  const creditLimit = record.balance * creditPercent / 100
  return {
    classification,
    paymentPercent,
    creditLimit,
    invoiceWarning: !!record.invoiceDate && record.invoiceAmount > creditLimit,
    recentReceiptDowngrade,
  }
}

function automaticClassificationLabel(classification: AutomaticPaymentClassification): string {
  return AUTOMATIC_CLASSIFICATIONS.find(item => item.value === classification)?.label ?? "لا يوجد مستحق للتقييم"
}

function automaticClassificationColor(classification: AutomaticPaymentClassification): string {
  if (classification === "not_applicable") return "bg-slate-300"
  return PAYMENT_CLASSIFICATIONS.find(item => item.value === classification)?.color ?? "bg-slate-300"
}

function classificationExcelColor(classification: AutomaticPaymentClassification | null): string {
  if (classification === "not_applicable") return "E2E8F0"
  return PAYMENT_CLASSIFICATIONS.find(item => item.value === classification)?.excelColor ?? "E2E8F0"
}

function paymentClassificationReportRows(
  records: CustomerBalance[],
  criteria: PaymentCriteria,
  now: Date,
) {
  return records.map(record => {
    const assessment = assessCustomerPayment(record, criteria, now)
    const alerts = [
      assessment.classification === "poor"
        ? "لا تمنح ديناً جديداً؛ اطلب السداد وراجع إغلاق الحساب"
        : "",
      assessment.recentReceiptDowngrade ? "خُفّض بسبب حداثة سند القبض" : "",
      assessment.invoiceWarning ? "آخر فاتورة أعلى من الحد المقترح؛ لا تتوفر معلومة تثبت سدادها" : "",
    ].filter(Boolean)
    return {
      record,
      assessment,
      manualClassification: paymentClassificationLabel(record.paymentClassification),
      note: alerts.join("؛ ") || "—",
    }
  })
}

function PaymentClassificationIndicators({ record, criteria, now }: {
  record: CustomerBalance
  criteria: PaymentCriteria
  now: Date
}) {
  const assessment = assessCustomerPayment(record, criteria, now)
  const manualClassification = PAYMENT_CLASSIFICATIONS.find(item => item.value === (record.paymentClassification ?? null))
  const warning = assessment.classification === "poor"
    ? "تصنيف النظام: سيئ الدفع. لا تمنح ديناً جديداً؛ اطلب السداد وراجع إمكانية إغلاق الحساب."
    : assessment.invoiceWarning
      ? "تنبيه: قيمة آخر فاتورة أعلى من حد الائتمان المقترح. الكشف لا يثبت إن كانت الفاتورة مدفوعة."
      : ""
  return <span className="inline-flex items-center gap-1.5">
    <span aria-label={`تصنيفك: ${manualClassification?.label ?? "غير مصنف"}`} title={`تصنيفك: ${manualClassification?.label ?? "غير مصنف"}`} className={`inline-block h-2.5 w-2.5 shrink-0 rounded-full ${manualClassification?.color ?? "bg-slate-300"}`} />
    <span aria-label={`تصنيف النظام: ${automaticClassificationLabel(assessment.classification)}`} title={`تصنيف النظام: ${automaticClassificationLabel(assessment.classification)}`} className={`inline-block h-2.5 w-2.5 shrink-0 rounded-full ${automaticClassificationColor(assessment.classification)}`} />
    {warning && <AlertTriangle aria-label={warning} className={`h-3.5 w-3.5 shrink-0 ${assessment.classification === "poor" ? "text-rose-600" : "text-amber-600"}`}><title>{warning}</title></AlertTriangle>}
  </span>
}

function customerMessage(record: CustomerBalance, template: string, installmentAmount: string): string {
  return statementMessage(record, template)
    .replaceAll("{{weeklyAmount}}", installmentAmount ? formatAmount(Number(installmentAmount)) : "[المبلغ الأسبوعي]")
}

function parseAmount(value: unknown): number {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0
  const raw = arabicDigits(cellString(value))
  const isNegative = /^\s*\(.*\)\s*$/.test(raw)
  const normalized = raw
    .replace(/[٬,]/g, "")
    .replace(/٫/g, ".")
    .replace(/[^\d.+-]/g, "")
  if (!normalized || normalized === "." || normalized === "-" || normalized === "+") return 0
  const amount = Number(normalized) * (isNegative ? -1 : 1)
  return Number.isFinite(amount) ? amount : 0
}

function parseBalance(value: unknown): Pick<CustomerBalance, "balance" | "balanceDirection"> {
  const raw = cellString(value)
    .normalize("NFKC")
    .replace(/[\u064b-\u065f\u0670\u0640]/g, "")
    .replace(/\s+/g, "")
    .toLowerCase()
  const hasReceivableMarker = /^(?:م|m)|(?:م|m)$/.test(raw)
  const hasPayableMarker = /^(?:د|d)|(?:د|d)$/.test(raw)
  const amount = parseAmount(value)
  const balanceDirection = hasReceivableMarker
    ? "receivable"
    : hasPayableMarker
      ? "payable"
      : "unspecified"
  const balance = Math.abs(amount)
  return {
    balance,
    balanceDirection: balanceDirection === "unspecified" && amount < 0
      ? "payable"
      : balanceDirection,
  }
}

function parseDate(value: unknown): string {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${value.getUTCFullYear()}-${String(value.getUTCMonth() + 1).padStart(2, "0")}-${String(value.getUTCDate()).padStart(2, "0")}`
  }
  if (typeof value === "number" && Number.isFinite(value) && value > 0) {
    const date = new Date(Date.UTC(1899, 11, 30) + value * 86_400_000)
    return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`
  }
  const text = arabicDigits(cellString(value))
  const iso = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/)
  if (iso) {
    const [, year, month, day] = iso
    return validDate(Number(year), Number(month), Number(day))
  }
  const dmy = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/)
  if (dmy) {
    const [, day, month, year] = dmy
    return validDate(Number(year), Number(month), Number(day))
  }
  const shortYear = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2})/)
  if (shortYear) {
    const [, day, month, year] = shortYear
    return validDate(2000 + Number(year), Number(month), Number(day))
  }
  return ""
}

function validDate(year: number, month: number, day: number): string {
  const date = new Date(Date.UTC(year, month - 1, day))
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return ""
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`
}

function readCell(row: unknown[], columns: Partial<Record<CustomerImportField, number>>, key: CustomerImportField): unknown {
  const index = columns[key]
  return index === undefined ? "" : row[index]
}

function parseWorkbook(data: ArrayBuffer): CustomerBalance[] {
  const workbook = XLSX.read(data, { type: "array", cellDates: true })
  const records: CustomerBalance[] = []
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet) continue
    const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: true, blankrows: false })
    const displayRows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: false, blankrows: false })
    let columns: Partial<Record<CustomerImportField, number>> | null = null
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex]
      const header = findHeader(row, rows[rowIndex + 1])
      if (header) {
        columns = header.columns
        rowIndex += header.rowCount - 1
        continue
      }
      if (!columns) continue
      const displayRow = displayRows[rowIndex] ?? row
      const name = cellString(readCell(row, columns, "name"))
      const customerNumber = normalizeCustomerNumber(readCell(displayRow, columns, "customerNumber"))
      const balance = parseBalance(readCell(displayRow, columns, "balance"))
      const normalizedName = normalizeHeader(name)
      const isPageMarker = row.some(value => /^(?:رقم\s*الصفحة|page\s*(?:no|number))\s*[:：#-]?\s*\d+$/i.test(cellString(value)))
      if (!name || !normalizedName || !customerNumber || isPageMarker || ["الاجمالي", "مجموع", "المجموع", "الرصيد السابق"].includes(normalizedName)) continue
      records.push({
        customerNumber,
        name,
        ...balance,
        duration: parseAmount(readCell(row, columns, "duration")),
        invoiceAmount: parseAmount(readCell(row, columns, "invoiceAmount")),
        invoiceDate: parseDate(readCell(row, columns, "invoiceDate")),
        receiptAmount: parseAmount(readCell(row, columns, "receiptAmount")),
        receiptDate: parseDate(readCell(row, columns, "receiptDate")),
        phone: cellString(readCell(displayRow, columns, "phone")),
        note: emptyCustomerNote(),
      })
    }
  }
  return records
}

function normalizeCustomerNumber(value: unknown): string {
  return arabicDigits(cellString(value))
    .normalize("NFKC")
    .replace(/\.0+$/, "")
    .trim()
}

function emptyCustomerNote(): CustomerNote {
  return { agreement: "", actions: [], agreements: [], agreementHistory: [] }
}

function normalizeAgreements(value: unknown): CustomerAgreement[] {
  if (!Array.isArray(value)) return []
  return value.filter((agreement): agreement is CustomerAgreement =>
    !!agreement &&
    typeof agreement.id === "string" &&
    typeof agreement.startedAt === "string" &&
    typeof agreement.description === "string" &&
    typeof agreement.initialBalance === "number" &&
    (typeof agreement.installmentAmount === "number" || agreement.installmentAmount === null) &&
    typeof agreement.nextPaymentDate === "string" &&
    (agreement.frequency === "weekly" || agreement.frequency === "monthly" || agreement.frequency === "bimonthly" || agreement.frequency === "once" || agreement.frequency === "flexible") &&
    (agreement.status === "active" || agreement.status === "completed" || agreement.status === "refused") &&
    Array.isArray(agreement.checks),
  ).map(agreement => ({
    ...agreement,
    dueDateFrom: typeof agreement.dueDateFrom === "string" ? agreement.dueDateFrom : agreement.nextPaymentDate,
    dueDateTo: typeof agreement.dueDateTo === "string" ? agreement.dueDateTo : "",
    reminders: Array.isArray(agreement.reminders)
      ? agreement.reminders.filter(reminder => !!reminder && typeof reminder.id === "string" && typeof reminder.date === "string")
      : [],
  }))
}

function normalizeCustomerNote(value: unknown): CustomerNote {
  if (!value || typeof value !== "object") return emptyCustomerNote()
  const note = value as Partial<CustomerNote>
  const normalizeActions = (actions: unknown): CustomerAction[] => Array.isArray(actions)
    ? actions.filter((action): action is CustomerAction =>
      !!action &&
      typeof action.id === "string" &&
      typeof action.description === "string" &&
      typeof action.dueDate === "string" &&
      typeof action.reminderDate === "string",
    ).map(action => ({
      ...action,
      status: action.status === "completed" || action.status === "failed" || action.status === "pending"
        ? action.status
        : action.completed ? "completed" : "pending",
      completed: action.completed === true,
      reminderDates: Array.isArray(action.reminderDates)
        ? action.reminderDates.filter(reminder => !!reminder && typeof reminder.id === "string" && typeof reminder.date === "string")
        : undefined,
    }))
    : []
  const actions = normalizeActions(note.actions)
  const agreement = typeof note.agreement === "string" ? note.agreement : ""
  const existingHistory = Array.isArray(note.agreementHistory)
    ? note.agreementHistory.filter(summary =>
      !!summary && typeof summary.id === "string" && typeof summary.agreement === "string" &&
      typeof summary.startedAt === "string" && typeof summary.closedAt === "string" && Array.isArray(summary.actions),
    ).map(summary => ({ ...summary, actions: normalizeActions(summary.actions) }))
    : []
  return {
    agreement,
    agreementStartedAt: typeof note.agreementStartedAt === "string"
      ? note.agreementStartedAt
      : agreement.trim() || actions.length ? new Date().toISOString() : undefined,
    agreementHistory: existingHistory,
    actions,
    agreements: normalizeAgreements(note.agreements),
  }
}

function normalizeStoredDataset(value: CustomerDataset | null): CustomerDataset | null {
  if (value === null) return null
  if (!Array.isArray(value.records)) throw new Error("قائمة أرصدة الزبائن المحفوظة غير صالحة. احتفظ بنسخة منها ثم أعد رفع الكشف.")
  const legacyRecords = value.records as Array<CustomerBalance & { pageNumber?: string }>
  return {
    ...value,
    records: legacyRecords
      .map(record => {
        const balance = Math.abs(Number(record.balance) || 0)
        const balanceDirection = record.balanceDirection === "receivable" || record.balanceDirection === "payable"
          ? record.balanceDirection
          : Number(record.balance) < 0 ? "payable" as const : "unspecified" as const
        return {
          ...record,
          customerNumber: normalizeCustomerNumber(record.customerNumber ?? record.pageNumber),
          paymentClassification: record.paymentClassification === "good" ||
            record.paymentClassification === "acceptable" ||
            record.paymentClassification === "poor" ||
            record.paymentClassification === "excellent"
            ? record.paymentClassification
            : undefined,
          isInLatestStatement: typeof record.isInLatestStatement === "boolean" ? record.isInLatestStatement : true,
          statementAppearances: Number.isInteger(record.statementAppearances) && Number(record.statementAppearances) > 0
            ? Number(record.statementAppearances)
            : 1,
          lastSeenInStatementAt: typeof record.lastSeenInStatementAt === "string"
            ? record.lastSeenInStatementAt
            : typeof record.balanceUpdatedAt === "string" ? record.balanceUpdatedAt : value.importedAt ?? "",
          lastStatementBalance: typeof record.lastStatementBalance === "number" && Number.isFinite(record.lastStatementBalance)
            ? Math.abs(record.lastStatementBalance)
            : balance,
          lastStatementDirection: record.lastStatementDirection === "receivable" ||
            record.lastStatementDirection === "payable" ||
            record.lastStatementDirection === "unspecified"
            ? record.lastStatementDirection
            : balanceDirection,
          balance,
          balanceDirection,
          firstStatementBalance: typeof record.firstStatementBalance === "number" && Number.isFinite(record.firstStatementBalance)
            ? Math.abs(record.firstStatementBalance)
            : balance,
          firstStatementDirection: record.firstStatementDirection === "receivable" ||
            record.firstStatementDirection === "payable" ||
            record.firstStatementDirection === "unspecified"
            ? record.firstStatementDirection
            : balanceDirection,
          firstStatementAt: typeof record.firstStatementAt === "string" ? record.firstStatementAt : value.importedAt ?? "",
          balanceUpdatedAt: typeof record.balanceUpdatedAt === "string" ? record.balanceUpdatedAt : value.importedAt ?? "",
          note: normalizeCustomerNote(record.note),
        }
      })
      .filter(record => record.customerNumber),
  }
}

function mergeCustomerRecords(
  existingRecords: CustomerBalance[],
  importedRecords: CustomerBalance[],
  checkedAt: string,
  missingCustomerPolicy: MissingCustomerPolicy,
): { records: CustomerBalance[]; added: number; updated: number } {
  const recordsByNumber = new Map<string, CustomerBalance>()
  for (const record of existingRecords) {
    const customerNumber = normalizeCustomerNumber(record.customerNumber)
    if (!customerNumber) continue
    recordsByNumber.set(customerNumber, {
      ...record,
      customerNumber,
      balance: Math.abs(Number(record.balance) || 0),
      balanceDirection: record.balanceDirection ?? "unspecified",
      isInLatestStatement: record.isInLatestStatement ?? true,
      statementAppearances: record.statementAppearances ?? 1,
      lastSeenInStatementAt: record.lastSeenInStatementAt ?? record.balanceUpdatedAt ?? checkedAt,
      lastStatementBalance: record.lastStatementBalance ?? Math.abs(Number(record.balance) || 0),
      lastStatementDirection: record.lastStatementDirection ?? record.balanceDirection ?? "unspecified",
      note: record.note ?? emptyCustomerNote(),
    })
  }
  let added = 0
  let updated = 0
  const importedNumbers = new Set(importedRecords.map(record => normalizeCustomerNumber(record.customerNumber)))
  for (const imported of importedRecords) {
    const customerNumber = normalizeCustomerNumber(imported.customerNumber)
    if (!customerNumber) continue
    const existing = recordsByNumber.get(customerNumber)
    if (existing) updated++
    else added++
    recordsByNumber.set(customerNumber, {
      ...existing,
      ...imported,
      customerNumber,
      isInLatestStatement: true,
      statementAppearances: (existing?.statementAppearances ?? 0) + 1,
      lastSeenInStatementAt: checkedAt,
      lastStatementBalance: imported.balance,
      lastStatementDirection: imported.balanceDirection,
      firstStatementBalance: existing?.firstStatementBalance ?? existing?.balance ?? imported.balance,
      firstStatementDirection: existing?.firstStatementDirection ?? existing?.balanceDirection ?? imported.balanceDirection,
      firstStatementAt: existing?.firstStatementAt ?? checkedAt,
      balanceUpdatedAt: checkedAt,
      note: existing
        ? {
          ...normalizeCustomerNote(existing.note),
          agreements: recordAgreementChecks(existing, imported, checkedAt),
        }
        : emptyCustomerNote(),
    })
  }
  if (missingCustomerPolicy === "keep_zero") {
    for (const [customerNumber, record] of recordsByNumber) {
      if (!importedNumbers.has(customerNumber)) {
        recordsByNumber.set(customerNumber, {
          ...record,
          isInLatestStatement: false,
          balance: 0,
          balanceDirection: "unspecified",
          duration: 0,
          balanceUpdatedAt: checkedAt,
        })
      }
    }
  } else {
    for (const [customerNumber, record] of recordsByNumber) {
      if (!importedNumbers.has(customerNumber)) {
        recordsByNumber.set(customerNumber, { ...record, isInLatestStatement: false })
      }
    }
  }
  return { records: [...recordsByNumber.values()], added, updated }
}

function localDateString(date = new Date()): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`
}

function createActionId(): string {
  return typeof globalThis.crypto?.randomUUID === "function"
    ? globalThis.crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`
}

function formatDate(value: string): string {
  if (!value) return "—"
  const [year, month, day] = value.split("-")
  return `${day}/${month}/${year}`
}

function formatAmount(value: number): string {
  return value.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
}

function dateIsInRange(value: string, from: string, to: string): boolean {
  if (!from && !to) return true
  if (!value) return false
  return (!from || value >= from) && (!to || value <= to)
}

function matchesThreshold(value: number, threshold: string, operator: Filters["balanceOperator"]): boolean {
  if (threshold === "") return true
  const target = Number(threshold)
  if (!Number.isFinite(target)) return true
  if (operator === "min") return value >= target
  if (operator === "max") return value <= target
  return value === target
}

function matchesNumericRange(value: number, from: string, to: string): boolean {
  const minimum = from === "" ? null : Number(from)
  const maximum = to === "" ? null : Number(to)
  if (minimum !== null && (!Number.isFinite(minimum) || value < minimum)) return false
  if (maximum !== null && (!Number.isFinite(maximum) || value > maximum)) return false
  return minimum === null || maximum === null || minimum <= maximum
}

function rangeLabel(label: string, from: string, to: string): string {
  return from || to ? `${label}: ${from || "بلا بداية"} — ${to || "بلا نهاية"}` : label
}

function actionDueTimestamp(action: CustomerAction): number | null {
  if (!action.dueDate) return null
  const [year, month, day] = action.dueDate.split("-").map(Number)
  if (!year || !month || !day) return null
  return new Date(year, month - 1, day, 23, 59, 59, 999).getTime()
}

function recordDueTimestamps(record: CustomerBalance): number[] {
  const actionDueDates = record.note.actions
    .filter(action => action.status !== "completed" && action.status !== "failed" && !action.completed)
    .map(actionDueTimestamp)
    .filter((timestamp): timestamp is number => timestamp !== null)
  const agreementDueDates = record.note.agreements
    .filter(agreement => agreement.status === "active")
    .flatMap(agreement => [
      ...[agreement.nextPaymentDate, agreement.dueDateTo ?? ""]
        .filter(Boolean)
        .map(date => new Date(`${date}T00:00:00`).getTime()),
      ...(agreement.reminders ?? []).map(reminder => new Date(`${reminder.date}T00:00:00`).getTime()),
    ])
    .filter(Number.isFinite)
  return [...actionDueDates, ...agreementDueDates]
}

function hasDueDateInRange(record: CustomerBalance, from: string, to: string): boolean {
  const actionMatches = record.note.actions.some(action => {
    if (!action.dueDate) return false
    const dueDate = action.dueDate.slice(0, 10)
    return (!from || dueDate >= from) && (!to || dueDate <= to)
  })
  const agreementMatches = record.note.agreements.some(agreement => {
    if (agreement.status !== "active") return false
    const dates = [agreement.nextPaymentDate, agreement.dueDateTo ?? "", ...(agreement.reminders ?? []).map(reminder => reminder.date)].filter(Boolean)
    return dates.some(date => (!from || date >= from) && (!to || date <= to))
  })
  return actionMatches || agreementMatches
}

function sortByNearestDue(records: CustomerBalance[], now: number): CustomerBalance[] {
  return records
    .map((record, index) => {
      const dueTimes = recordDueTimestamps(record)
      const overdueTimes = dueTimes.filter(timestamp => timestamp <= now)
      const upcomingTimes = dueTimes.filter(timestamp => timestamp > now)
      return {
        record,
        index,
        overdueAt: overdueTimes.length ? Math.max(...overdueTimes) : null,
        upcomingAt: upcomingTimes.length ? Math.min(...upcomingTimes) : null,
      }
    })
    .sort((left, right) => {
      const leftHasOverdue = left.overdueAt !== null
      const rightHasOverdue = right.overdueAt !== null
      if (leftHasOverdue !== rightHasOverdue) return leftHasOverdue ? -1 : 1
      if (leftHasOverdue && right.overdueAt !== null && left.overdueAt !== null) {
        return right.overdueAt - left.overdueAt || left.index - right.index
      }
      if (left.upcomingAt !== right.upcomingAt) {
        if (left.upcomingAt === null) return 1
        if (right.upcomingAt === null) return -1
        return left.upcomingAt - right.upcomingAt
      }
      return left.index - right.index
    })
    .map(({ record }) => record)
}

function hasCustomerNotes(record: CustomerBalance): boolean {
  return !!record.note.agreement.trim() ||
    record.note.actions.length > 0 ||
    !!record.note.agreementHistory?.length ||
    record.note.agreements.length > 0
}

function customerPayments(record: CustomerBalance): CustomerPayment[] {
  return record.note.agreements.flatMap(agreement => {
    const recordedChecks = agreement.checks
      .filter(check =>
        (check.outcome === "paid" || check.outcome === "partial") &&
        typeof check.actualPaid === "number" &&
        check.actualPaid > 0 &&
        !!check.checkedAt,
      )
      .map(check => ({
        customerNumber: record.customerNumber,
        customerName: record.name,
        agreement: agreement.description,
        date: check.checkedAt,
        amount: check.actualPaid!,
        type: "دفعة اتفاق",
      }))
    const closures = agreement.closureHistory?.length
      ? agreement.closureHistory
      : agreement.closedAt && agreement.closureOutcome
        ? [{
          id: `${agreement.id}-legacy-closure`,
          closedAt: agreement.closedAt,
          outcome: agreement.closureOutcome,
          paidAmount: agreement.closurePaidAmount ?? 0,
          balance: agreement.closureBalance ?? agreement.initialBalance,
          balanceDirection: agreement.closureBalanceDirection ?? agreement.balanceDirection,
        }]
        : []
    const closurePayments = closures
      .filter(closure =>
        (closure.outcome === "paid" || closure.outcome === "partial") &&
        closure.paidAmount > 0,
      )
      .filter(closure => !recordedChecks.some(payment =>
        payment.amount === closure.paidAmount &&
        payment.date.slice(0, 10) === closure.closedAt.slice(0, 10),
      ))
      .map(closure => ({
        customerNumber: record.customerNumber,
        customerName: record.name,
        agreement: agreement.description,
        date: closure.closedAt,
        amount: closure.paidAmount,
        type: "دفعة عند إنهاء الاتفاق",
      }))
    return [...recordedChecks, ...closurePayments]
  })
}

function agreementDueCount(agreement: CustomerAgreement, asOf: Date): number {
  if (agreement.frequency === "once") {
    if (agreement.checks.some(check => check.outcome === "paid" || check.outcome === "partial" || check.outcome === "unpaid")) return 0
    return new Date(`${agreement.nextPaymentDate}T00:00:00`).getTime() <= asOf.getTime() ? 1 : 0
  }
  if (agreement.frequency === "flexible") return 0
  const [year, month, day] = agreement.nextPaymentDate.split("-").map(Number)
  if (!year || !month || !day) return 0
  let count = 0
  let due = new Date(year, month - 1, day)
  const today = new Date(asOf.getFullYear(), asOf.getMonth(), asOf.getDate())
  while (due <= today && count < 1200) {
    count++
    if (agreement.frequency === "weekly") {
      due.setDate(due.getDate() + 7)
    } else {
      const nextMonth = due.getMonth() + (agreement.frequency === "bimonthly" ? 2 : 1)
      const nextYear = due.getFullYear() + Math.floor(nextMonth / 12)
      const monthIndex = nextMonth % 12
      due = new Date(nextYear, monthIndex, Math.min(day, new Date(nextYear, monthIndex + 1, 0).getDate()))
    }
  }
  return count
}

function advanceAgreementDueDate(dateString: string, frequency: AgreementFrequency, count: number): string {
  const [year, month, day] = dateString.split("-").map(Number)
  if (!year || !month || !day || frequency === "once" || frequency === "flexible") return dateString
  const date = new Date(year, month - 1, day)
  if (frequency === "weekly") {
    date.setDate(date.getDate() + count * 7)
  } else {
    const monthIndex = date.getMonth() + count * (frequency === "bimonthly" ? 2 : 1)
    const targetYear = date.getFullYear() + Math.floor(monthIndex / 12)
    const targetMonth = monthIndex % 12
    date.setFullYear(targetYear, targetMonth, 1)
    date.setDate(Math.min(day, new Date(targetYear, targetMonth + 1, 0).getDate()))
  }
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`
}

function recordAgreementChecks(previous: CustomerBalance, incoming: CustomerBalance, checkedAt: string): CustomerAgreement[] {
  return normalizeCustomerNote(previous.note).agreements.map(agreement => {
    const hasChanged = previous.balance !== incoming.balance || previous.balanceDirection !== incoming.balanceDirection
    const balanceSnapshots = hasChanged
      ? [...(agreement.balanceSnapshots ?? []), { capturedAt: checkedAt, balance: incoming.balance, direction: incoming.balanceDirection }]
      : agreement.balanceSnapshots ?? []
    if (agreement.status !== "active") return { ...agreement, balanceSnapshots }
    const asOf = new Date(checkedAt)
    const dueCount = agreementDueCount(agreement, asOf)
    const outstanding = Math.max(0, agreement.outstandingInstallment || 0)
    const expectedPayment = Math.min(previous.balance, outstanding + dueCount * (agreement.installmentAmount ?? 0))
    const comparable = previous.balanceDirection === incoming.balanceDirection &&
      previous.balanceDirection === agreement.balanceDirection &&
      previous.balanceDirection === "receivable"
    const inferredPaid = comparable ? Math.max(0, previous.balance - incoming.balance) : 0
    if (!hasChanged && expectedPayment <= 0) return agreement
    const latestCheck = agreement.checks.at(-1)
    if (
      latestCheck &&
      latestCheck.checkedAt.slice(0, 10) === checkedAt.slice(0, 10) &&
      latestCheck.previousBalance === previous.balance &&
      latestCheck.previousDirection === previous.balanceDirection &&
      latestCheck.newBalance === incoming.balance &&
      latestCheck.newDirection === incoming.balanceDirection &&
      latestCheck.expectedPayment === expectedPayment
    ) return agreement
    const suggestion: AgreementSuggestion = !comparable
      ? "not_comparable"
      : expectedPayment <= 0
        ? "not_comparable"
        : inferredPaid >= expectedPayment
          ? "paid"
          : inferredPaid > 0
            ? "partial"
            : "unpaid"
    const check: AgreementCheck = {
      id: createActionId(),
      checkedAt,
      previousBalance: previous.balance,
      previousDirection: previous.balanceDirection,
      newBalance: incoming.balance,
      newDirection: incoming.balanceDirection,
      inferredPaid,
      expectedPayment,
      suggestion,
    }
    return { ...agreement, checks: [...agreement.checks, check], balanceSnapshots }
  })
}

export function CustomerBalances({
  onDueRemindersChange,
  remindersEnabled = true,
  visible = true,
  username,
  missingCustomerPolicy = "keep_zero",
  missingCustomerPolicyReady = true,
}: {
  onDueRemindersChange?: (count: number) => void
  remindersEnabled?: boolean
  visible?: boolean
  username: string
  missingCustomerPolicy?: MissingCustomerPolicy
  missingCustomerPolicyReady?: boolean
}) {
  const normalizedUsername = username.normalize("NFKC").trim().toLocaleLowerCase("en-US")
  const storageKey = `${STORAGE_KEY}:${normalizedUsername}`
  const deletedDatasetKey = `${storageKey}:deleted`
  const blocklistStorageKey = `customer_balances_blocklist:${normalizedUsername}`
  const paymentCriteriaStorageKey = `${storageKey}:payment_criteria`
  const inputRef = useRef<HTMLInputElement>(null)
  const statementFileInputRef = useRef<HTMLInputElement>(null)
  const messageTemplateTextAreaRef = useRef<HTMLTextAreaElement>(null)
  const bulkImageMessageTextAreaRef = useRef<HTMLTextAreaElement>(null)
  const [dataset, setDataset] = useState<CustomerDataset | null>(null)
  const [openCustomerNumber, setOpenCustomerNumber] = useState<string | null>(null)
  const [customerShortcutNumber, setCustomerShortcutNumber] = useState<string | null>(null)
  const [statementCustomerNumber, setStatementCustomerNumber] = useState<string | null>(null)
  const [bulkAudienceOpen, setBulkAudienceOpen] = useState(false)
  const [bulkAudience, setBulkAudience] = useState<"all" | "exclude_blocked">("all")
  const [bulkSendMode, setBulkSendMode] = useState<"message" | "image">("message")
  const [bulkCardTemplate, setBulkCardTemplate] = useState<CustomerCardTemplate>("modern")
  const [bulkMessageType, setBulkMessageType] = useState("statement")
  const [bulkQueue, setBulkQueue] = useState<BulkMessageRecipient[] | null>(null)
  const [bulkQueueIndex, setBulkQueueIndex] = useState(0)
  const [bulkQueueError, setBulkQueueError] = useState("")
  const [bulkImageFile, setBulkImageFile] = useState<File | null>(null)
  const [bulkImagePreviewUrl, setBulkImagePreviewUrl] = useState<string | null>(null)
  const [bulkImageBusy, setBulkImageBusy] = useState(false)
  const [bulkImageError, setBulkImageError] = useState("")
  const [statementMode, setStatementMode] = useState<"both" | "file" | "message">("both")
  const [statementFile, setStatementFile] = useState<File | null>(null)
  const [statementShareError, setStatementShareError] = useState("")
  const [statementMessageTemplate, setStatementMessageTemplate] = useState(DEFAULT_STATEMENT_MESSAGE)
  const [weeklyInstallmentMessageTemplate, setWeeklyInstallmentMessageTemplate] = useState(DEFAULT_WEEKLY_INSTALLMENT_MESSAGE)
  const [paymentCriteria, setPaymentCriteria] = useState<PaymentCriteria>(DEFAULT_PAYMENT_CRITERIA)
  const [loadedPaymentCriteriaKey, setLoadedPaymentCriteriaKey] = useState<string | null>(null)
  const [hasPaymentCriteriaChanges, setHasPaymentCriteriaChanges] = useState(false)
  const [paymentCriteriaError, setPaymentCriteriaError] = useState("")
  const [statementMessageType, setStatementMessageType] = useState("statement")
  const [weeklyInstallmentAmountInput, setWeeklyInstallmentAmountInput] = useState("")
  const [customMessageTemplates, setCustomMessageTemplates] = useState<CustomMessageTemplate[]>([])
  const [messageTemplateEditorType, setMessageTemplateEditorType] = useState("statement")
  const [newMessageTemplateName, setNewMessageTemplateName] = useState("")
  const [showNewMessageTemplate, setShowNewMessageTemplate] = useState(false)
  const [loadedMessageTemplateKey, setLoadedMessageTemplateKey] = useState<string | null>(null)
  const [hasMessageTemplateChanges, setHasMessageTemplateChanges] = useState(false)
  const [messageEditorOpen, setMessageEditorOpen] = useState(false)
  const [blockedCustomers, setBlockedCustomers] = useState<BlockedCustomer[]>([])
  const [blockedPage, setBlockedPage] = useState(1)
  const [selectedCustomerNumbers, setSelectedCustomerNumbers] = useState<string[]>([])
  const [blockEditorNumbers, setBlockEditorNumbers] = useState<string[] | null>(null)
  const [blockReasons, setBlockReasons] = useState<Record<string, string>>({})
  const [showPaymentClassification, setShowPaymentClassification] = useState(true)
  const [showCustomerHistory, setShowCustomerHistory] = useState(false)
  const [customerHistoryPage, setCustomerHistoryPage] = useState(1)
  const [blockedContactNumber, setBlockedContactNumber] = useState<string | null>(null)
  const [blockedContactMessage, setBlockedContactMessage] = useState("")
  const [blockedContactAction, setBlockedContactAction] = useState<"whatsapp" | "statement" | "image" | null>(null)
  const [filters, setFilters] = useState<Filters>(defaultFilters)
  const [reportOpen, setReportOpen] = useState(false)
  const [reportType, setReportType] = useState<CustomerReportType>("balances")
  const [paymentDateFrom, setPaymentDateFrom] = useState("")
  const [paymentDateTo, setPaymentDateTo] = useState("")
  const [reportAmountFrom, setReportAmountFrom] = useState("")
  const [reportAmountTo, setReportAmountTo] = useState("")
  const [ready, setReady] = useState(false)
  const [loadedStorageKey, setLoadedStorageKey] = useState<string | null>(null)
  const [loadedBlocklistKey, setLoadedBlocklistKey] = useState<string | null>(null)
  const [hasChanges, setHasChanges] = useState(false)
  const [hasBlocklistChanges, setHasBlocklistChanges] = useState(false)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState("")
  const [notice, setNotice] = useState("")
  const [now, setNow] = useState(() => new Date())
  useGlobalLoading(visible && !ready && !error, "جاري تحميل أرصدة الزبائن...")
  useGlobalLoading(visible && busy, "جاري معالجة بيانات الزبائن...")
  const today = localDateString(now)
  const statementCustomer = dataset?.records.find(record => record.customerNumber === statementCustomerNumber)
  const selectedCustomMessageTemplate = customMessageTemplates.find(template => statementMessageType === `custom:${template.id}`)
  const selectedMessageTemplate = statementMessageType === "weekly_installment"
    ? weeklyInstallmentMessageTemplate
    : selectedCustomMessageTemplate?.content ?? statementMessageTemplate
  const selectedBulkCustomTemplate = customMessageTemplates.find(template => bulkMessageType === `custom:${template.id}`)
  const selectedBulkMessageTemplate = bulkMessageType === "weekly_installment"
    ? weeklyInstallmentMessageTemplate
    : selectedBulkCustomTemplate?.content ?? statementMessageTemplate
  const bulkAudienceRecords = useMemo(() => {
    const records = dataset?.records ?? []
    if (bulkAudience === "all") return records
    const blockedNumbers = new Set(blockedCustomers.map(customer => customer.customerNumber))
    return records.filter(record => !blockedNumbers.has(record.customerNumber))
  }, [dataset, blockedCustomers, bulkAudience])
  const bulkEligibleRecords = useMemo(() => bulkAudienceRecords.filter(record =>
    !!whatsappUrl(record.phone) &&
    (bulkSendMode === "image" || bulkMessageType !== "weekly_installment" || !!weeklyInstallmentAmount(record)),
  ), [bulkAudienceRecords, bulkMessageType, bulkSendMode])
  const bulkQueueRecipient = bulkQueue?.[bulkQueueIndex]
  const selectedCustomerMessage = statementCustomer
    ? customerMessage(statementCustomer, selectedMessageTemplate, weeklyInstallmentAmountInput)
    : ""
  const canSendSelectedMessage = selectedMessageTemplate.trim().length > 0 &&
    (statementMessageType !== "weekly_installment" ||
      (Number.isFinite(Number(weeklyInstallmentAmountInput)) && Number(weeklyInstallmentAmountInput) > 0))
  const paymentCriteriaValid = isPaymentCriteriaValid(paymentCriteria)
  const effectivePaymentCriteria = paymentCriteriaValid ? paymentCriteria : DEFAULT_PAYMENT_CRITERIA
  const [notificationPermission, setNotificationPermission] = useState<NotificationPermission | "unsupported">(
    () => typeof Notification === "undefined" ? "unsupported" : Notification.permission,
  )

  useEffect(() => {
    if (!bulkQueueRecipient || bulkSendMode !== "image") {
      setBulkImageFile(null)
      setBulkImageBusy(false)
      setBulkImageError("")
      return
    }
    let active = true
    setBulkImageFile(null)
    setBulkImageBusy(true)
    setBulkImageError("")
    void createCustomerDebtImage(bulkQueueRecipient, bulkCardTemplate, bulkQueueRecipient.message)
      .then(file => {
        if (active) setBulkImageFile(file)
      })
      .catch(error => {
        if (active) setBulkImageError(error instanceof Error ? error.message : "تعذر إنشاء صورة الزبون.")
      })
      .finally(() => {
        if (active) setBulkImageBusy(false)
      })
    return () => {
      active = false
    }
  }, [bulkCardTemplate, bulkQueueRecipient, bulkSendMode])

  useEffect(() => {
    if (!bulkImageFile) {
      setBulkImagePreviewUrl(null)
      return
    }
    const previewUrl = URL.createObjectURL(bulkImageFile)
    setBulkImagePreviewUrl(previewUrl)
    return () => URL.revokeObjectURL(previewUrl)
  }, [bulkImageFile])

  useEffect(() => {
    let active = true
    setReady(false)
    setHasChanges(false)
    setLoadedStorageKey(null)
    setDataset(null)
    setOpenCustomerNumber(null)
    setCustomerShortcutNumber(null)
    setStatementCustomerNumber(null)
    setStatementFile(null)
    setError("")
    setNotice("")
    setFilters(defaultFilters)
    const load = async () => {
      try {
        let stored: CustomerDataset | null = null
        if (centralStorageEnabled) {
          const wasDeleted = await centralStorageGet<boolean>(deletedDatasetKey) === true
          stored = wasDeleted ? null : await centralStorageGet<CustomerDataset>(storageKey) ?? null
          if (!wasDeleted && !stored && storageKey !== LEGACY_STORAGE_KEY) {
            const legacy = await centralStorageGet<CustomerDataset>(LEGACY_STORAGE_KEY)
            if (legacy) {
              stored = legacy
              await centralStorageSet(storageKey, legacy)
              await centralStorageSet(deletedDatasetKey, false)
            }
          }
        } else {
          const wasDeleted = localStorage.getItem(deletedDatasetKey) === "1"
          const raw = wasDeleted ? null : localStorage.getItem(storageKey)
          stored = raw ? JSON.parse(raw) as CustomerDataset : null
        }
        if (active) {
          setDataset(normalizeStoredDataset(stored))
          setLoadedStorageKey(storageKey)
        }
      } catch (cause) {
        if (active) setError(cause instanceof Error ? `تعذر تحميل البيانات المحفوظة: ${cause.message}` : "تعذر تحميل البيانات المحفوظة.")
      } finally {
        if (active) setReady(true)
      }
    }
    void load()
    return () => { active = false }
  }, [storageKey])

  useEffect(() => {
    if (!missingCustomerPolicyReady || loadedStorageKey !== storageKey || !dataset || missingCustomerPolicy !== "keep_zero") return
    const hasUnzeroedMissingCustomers = dataset.records.some(record =>
      record.isInLatestStatement === false &&
      (record.balance !== 0 || record.balanceDirection !== "unspecified" || record.duration !== 0),
    )
    if (!hasUnzeroedMissingCustomers) return
    setDataset(current => current ? {
      ...current,
      records: current.records.map(record => record.isInLatestStatement === false
        ? {
          ...record,
          balance: 0,
          balanceDirection: "unspecified",
          duration: 0,
          balanceUpdatedAt: current.importedAt,
        }
        : record),
    } : current)
    setHasChanges(true)
  }, [dataset, loadedStorageKey, missingCustomerPolicy, missingCustomerPolicyReady, storageKey])

  useEffect(() => {
    let active = true
    setLoadedBlocklistKey(null)
    setHasBlocklistChanges(false)
    setBlockedCustomers([])
    setSelectedCustomerNumbers([])
    setBlockEditorNumbers(null)
    const loadBlocklist = async () => {
      try {
        let stored: BlockedCustomer[] | null = null
        if (centralStorageEnabled) {
          stored = await centralStorageGet<BlockedCustomer[]>(blocklistStorageKey) ?? null
        } else {
          const raw = localStorage.getItem(blocklistStorageKey)
          stored = raw ? JSON.parse(raw) as BlockedCustomer[] : null
        }
        if (stored !== null && !Array.isArray(stored)) throw new Error("قائمة حظر التواصل المحفوظة غير صالحة.")
        if (active) {
          setBlockedCustomers(stored ?? [])
          setLoadedBlocklistKey(blocklistStorageKey)
        }
      } catch (cause) {
        if (active) setError(cause instanceof Error ? `تعذر تحميل قائمة حظر التواصل: ${cause.message}` : "تعذر تحميل قائمة حظر التواصل.")
      }
    }
    void loadBlocklist()
    return () => { active = false }
  }, [blocklistStorageKey])

  useEffect(() => {
    let active = true
    setLoadedMessageTemplateKey(null)
    setHasMessageTemplateChanges(false)
    const loadTemplate = async () => {
      try {
        let stored: string | null
        let storedWeekly: string | null
        let storedCustom: CustomMessageTemplate[] | null
        if (centralStorageEnabled) {
          stored = await centralStorageGet<string>(`${storageKey}:statement_message`) ?? null
          storedWeekly = await centralStorageGet<string>(`${storageKey}:weekly_installment_message`) ?? null
          storedCustom = await centralStorageGet<CustomMessageTemplate[]>(`${storageKey}:custom_messages`) ?? null
        } else {
          stored = localStorage.getItem(`${storageKey}:statement_message`)
          storedWeekly = localStorage.getItem(`${storageKey}:weekly_installment_message`)
          const rawCustom = localStorage.getItem(`${storageKey}:custom_messages`)
          storedCustom = rawCustom ? JSON.parse(rawCustom) as CustomMessageTemplate[] : null
        }
        if (storedCustom !== null && (!Array.isArray(storedCustom) || storedCustom.some(template =>
          !template || typeof template.id !== "string" || typeof template.name !== "string" || typeof template.content !== "string",
        ))) throw new Error("قوالب الرسائل الإضافية المحفوظة غير صالحة.")
        if (active) {
          setStatementMessageTemplate(stored ?? DEFAULT_STATEMENT_MESSAGE)
          setWeeklyInstallmentMessageTemplate(storedWeekly ?? DEFAULT_WEEKLY_INSTALLMENT_MESSAGE)
          setCustomMessageTemplates(storedCustom ?? [])
          setLoadedMessageTemplateKey(storageKey)
        }
      } catch (cause) {
        if (active) setError(cause instanceof Error ? `تعذر تحميل رسالة كشف الحساب: ${cause.message}` : "تعذر تحميل رسالة كشف الحساب.")
      }
    }
    void loadTemplate()
    return () => { active = false }
  }, [storageKey])

  useEffect(() => {
    let active = true
    setLoadedPaymentCriteriaKey(null)
    setHasPaymentCriteriaChanges(false)
    const loadPaymentCriteria = async () => {
      try {
        const stored = centralStorageEnabled
          ? await centralStorageGet<PaymentCriteria>(paymentCriteriaStorageKey) ?? null
          : (() => {
            const raw = localStorage.getItem(paymentCriteriaStorageKey)
            return raw ? JSON.parse(raw) as PaymentCriteria : null
          })()
        if (active) {
          setPaymentCriteria(stored ? normalizePaymentCriteria(stored) : DEFAULT_PAYMENT_CRITERIA)
          setPaymentCriteriaError("")
          setLoadedPaymentCriteriaKey(paymentCriteriaStorageKey)
        }
      } catch (cause) {
        if (active) {
          setPaymentCriteria(DEFAULT_PAYMENT_CRITERIA)
          setPaymentCriteriaError(cause instanceof Error
            ? `تعذر تحميل معايير الدفع: ${cause.message}`
            : "تعذر تحميل معايير الدفع.")
          setLoadedPaymentCriteriaKey(paymentCriteriaStorageKey)
        }
      }
    }
    void loadPaymentCriteria()
    return () => { active = false }
  }, [paymentCriteriaStorageKey])

  useEffect(() => {
    if (!ready || loadedStorageKey !== storageKey || !hasChanges) return
    let active = true
    const timer = window.setTimeout(() => {
      const save = async () => {
        try {
          if (centralStorageEnabled) {
            await centralStorageSet(storageKey, dataset)
            await centralStorageSet(deletedDatasetKey, dataset === null)
          } else if (dataset) {
            localStorage.setItem(storageKey, JSON.stringify(dataset))
            localStorage.removeItem(deletedDatasetKey)
          } else {
            localStorage.removeItem(storageKey)
            localStorage.setItem(deletedDatasetKey, "1")
          }
          if (active) {
            setHasChanges(false)
            setError("")
            setNotice("تم حفظ البيانات بنجاح.")
          }
        } catch (cause) {
          if (active) setError(cause instanceof Error ? `تعذر حفظ البيانات: ${cause.message}` : "تعذر حفظ البيانات.")
        }
      }
      void save()
    }, 250)
    return () => {
      active = false
      window.clearTimeout(timer)
    }
  }, [dataset, hasChanges, loadedStorageKey, ready, storageKey])

  useEffect(() => {
    if (loadedBlocklistKey !== blocklistStorageKey || !hasBlocklistChanges) return
    let active = true
    const timer = window.setTimeout(() => {
      const save = async () => {
        try {
          if (centralStorageEnabled) {
            await centralStorageSet(blocklistStorageKey, blockedCustomers)
          } else if (blockedCustomers.length) {
            localStorage.setItem(blocklistStorageKey, JSON.stringify(blockedCustomers))
          } else {
            localStorage.removeItem(blocklistStorageKey)
          }
          if (active) {
            setHasBlocklistChanges(false)
            setError("")
            setNotice("تم حفظ قائمة حظر التواصل.")
          }
        } catch (cause) {
          if (active) setError(cause instanceof Error ? `تعذر حفظ قائمة حظر التواصل: ${cause.message}` : "تعذر حفظ قائمة حظر التواصل.")
        }
      }
      void save()
    }, 250)
    return () => {
      active = false
      window.clearTimeout(timer)
    }
  }, [blockedCustomers, blocklistStorageKey, hasBlocklistChanges, loadedBlocklistKey])

  useEffect(() => {
    if (loadedMessageTemplateKey !== storageKey || !hasMessageTemplateChanges) return
    let active = true
    const timer = window.setTimeout(() => {
      const save = async () => {
        try {
          if (centralStorageEnabled) {
            await centralStorageSet(`${storageKey}:statement_message`, statementMessageTemplate)
            await centralStorageSet(`${storageKey}:weekly_installment_message`, weeklyInstallmentMessageTemplate)
            await centralStorageSet(`${storageKey}:custom_messages`, customMessageTemplates)
          } else {
            localStorage.setItem(`${storageKey}:statement_message`, statementMessageTemplate)
            localStorage.setItem(`${storageKey}:weekly_installment_message`, weeklyInstallmentMessageTemplate)
            localStorage.setItem(`${storageKey}:custom_messages`, JSON.stringify(customMessageTemplates))
          }
          if (active) {
            setHasMessageTemplateChanges(false)
            setError("")
            setNotice("تم حفظ قوالب الرسائل.")
          }
        } catch (cause) {
          if (active) setError(cause instanceof Error ? `تعذر حفظ قوالب الرسائل: ${cause.message}` : "تعذر حفظ قوالب الرسائل.")
        }
      }
      void save()
    }, 250)
    return () => {
      active = false
      window.clearTimeout(timer)
    }
  }, [customMessageTemplates, hasMessageTemplateChanges, loadedMessageTemplateKey, statementMessageTemplate, storageKey, weeklyInstallmentMessageTemplate])

  useEffect(() => {
    if (loadedPaymentCriteriaKey !== paymentCriteriaStorageKey || !hasPaymentCriteriaChanges || !paymentCriteriaValid) return
    let active = true
    const timer = window.setTimeout(() => {
      const save = async () => {
        try {
          if (centralStorageEnabled) {
            await centralStorageSet(paymentCriteriaStorageKey, paymentCriteria)
          } else {
            localStorage.setItem(paymentCriteriaStorageKey, JSON.stringify(paymentCriteria))
          }
          if (active) {
            setHasPaymentCriteriaChanges(false)
            setPaymentCriteriaError("")
            setNotice("تم حفظ معايير تقييم الدفع.")
          }
        } catch (cause) {
          if (active) setPaymentCriteriaError(cause instanceof Error
            ? `تعذر حفظ معايير الدفع: ${cause.message}`
            : "تعذر حفظ معايير الدفع.")
        }
      }
      void save()
    }, 250)
    return () => {
      active = false
      window.clearTimeout(timer)
    }
  }, [hasPaymentCriteriaChanges, loadedPaymentCriteriaKey, paymentCriteria, paymentCriteriaStorageKey, paymentCriteriaValid])

  useEffect(() => {
    const refreshClock = () => setNow(new Date())
    const timer = window.setInterval(refreshClock, 60_000)
    window.addEventListener("focus", refreshClock)
    return () => {
      window.clearInterval(timer)
      window.removeEventListener("focus", refreshClock)
    }
  }, [])

  const currentStatementRecords = useMemo(
    () => (dataset?.records ?? []).filter(record =>
      missingCustomerPolicy === "keep_zero" || record.isInLatestStatement !== false,
    ),
    [dataset, missingCustomerPolicy],
  )
  const filteredRecords = useMemo(() => {
    const query = normalizeHeader(filters.query)
    const records = currentStatementRecords.filter(record => {
      if (filters.notes === "withNotes" && !hasCustomerNotes(record)) return false
      if (filters.agreementStatus === "active" && !record.note.agreements.some(agreement => agreement.status === "active")) return false
      if ((filters.dueFrom || filters.dueTo) && !hasDueDateInRange(record, filters.dueFrom, filters.dueTo)) return false
      if (query) {
        if (filters.searchField === "customerNumber" &&
          !normalizeHeader(record.customerNumber).includes(query)) return false
        if (filters.searchField === "phone" &&
          (!phoneSearchDigits(filters.query) || !phoneSearchDigits(record.phone).includes(phoneSearchDigits(filters.query)))) return false
        if (filters.searchField === "all") {
        const searchableText = [
          record.name,
          record.customerNumber,
          record.phone,
          record.paymentClassification === "good" ? "جيد وسالك" :
            record.paymentClassification === "acceptable" ? "مقبول" :
              record.paymentClassification === "poor" ? "سيئ الدفع" :
                record.paymentClassification === "excellent" ? "ممتاز" : "",
          record.note.agreement,
          ...(record.note.agreementHistory ?? []).map(summary => summary.agreement),
          ...record.note.actions.flatMap(action => [action.description, action.failureReason ?? "", action.dueDate]),
          ...record.note.agreements.flatMap(agreement => [
            agreement.description,
            ...agreement.checks.flatMap(check => [check.note ?? "", check.outcome ?? "", check.suggestion]),
          ]),
        ].map(normalizeHeader).join(" ")
        if (!searchableText.includes(query)) return false
        }
      }
      if (!matchesThreshold(record.balance, filters.balance, filters.balanceOperator)) return false
      if (!matchesNumericRange(record.balance, filters.balanceFrom, filters.balanceTo)) return false
      if (!matchesThreshold(record.duration, filters.duration, filters.durationOperator)) return false
      if (!matchesNumericRange(record.duration, filters.durationFrom, filters.durationTo)) return false
      if (!dateIsInRange(record.receiptDate, filters.receiptFrom, filters.receiptTo)) return false
      if (!dateIsInRange(record.invoiceDate, filters.invoiceFrom, filters.invoiceTo)) return false
      return true
    })
    if (filters.noteOrder === "dueSoon") return sortByNearestDue(records, now.getTime())
    return filters.noteOrder === "first"
      ? records
        .map((record, index) => ({ record, index }))
        .sort((left, right) => Number(hasCustomerNotes(right.record)) - Number(hasCustomerNotes(left.record)) || left.index - right.index)
        .map(({ record }) => record)
      : records
  }, [currentStatementRecords, filters, now])

  const customerHistoryRecords = useMemo(
    () => [...(dataset?.records ?? [])].sort((left, right) => left.name.localeCompare(right.name, "ar")),
    [dataset],
  )
  const customerHistoryPageCount = Math.max(1, Math.ceil(customerHistoryRecords.length / 50))
  const visibleCustomerHistory = customerHistoryRecords.slice((customerHistoryPage - 1) * 50, customerHistoryPage * 50)

  const reportRecords = useMemo(
    () => currentStatementRecords.filter(hasCustomerNotes),
    [currentStatementRecords],
  )
  const blockedBalanceRows = useMemo(() => {
    const recordsByNumber = new Map((dataset?.records ?? []).map(record => [record.customerNumber, record]))
    return blockedCustomers.map(customer => ({
      customer,
      record: recordsByNumber.get(customer.customerNumber),
    }))
  }, [blockedCustomers, dataset])
  const blockedReceivableTotal = blockedBalanceRows.reduce(
    (total, row) => total + (row.record?.balanceDirection === "receivable" ? row.record.balance : 0),
    0,
  )
  const blockedPayableTotal = blockedBalanceRows.reduce(
    (total, row) => total + (row.record?.balanceDirection === "payable" ? row.record.balance : 0),
    0,
  )
  const blockedNetTotal = blockedReceivableTotal - blockedPayableTotal
  const blockedMissingBalanceCount = blockedBalanceRows.filter(row => !row.record).length
  const blockedPageCount = Math.max(1, Math.ceil(blockedBalanceRows.length / 3))
  const visibleBlockedRows = blockedBalanceRows.slice((blockedPage - 1) * 3, blockedPage * 3)
  const reportPayments = useMemo(
    () => reportRecords
      .flatMap(customerPayments)
      .filter(payment => {
        const date = payment.date.slice(0, 10)
        return (!paymentDateFrom || date >= paymentDateFrom) &&
          (!paymentDateTo || date <= paymentDateTo)
      }),
    [paymentDateFrom, paymentDateTo, reportRecords],
  )
  const paymentTotalsByCustomer = useMemo(() => {
    const totals = new Map<string, number>()
    for (const payment of reportPayments) {
      totals.set(payment.customerNumber, (totals.get(payment.customerNumber) ?? 0) + payment.amount)
    }
    return totals
  }, [reportPayments])
  const totalPayments = reportPayments.reduce((total, payment) => total + payment.amount, 0)
  const paymentDateRangeValid = !paymentDateFrom || !paymentDateTo || paymentDateFrom <= paymentDateTo
  const paymentPeriodLabel = paymentDateFrom || paymentDateTo
    ? `${paymentDateFrom || "من البداية"} إلى ${paymentDateTo || "دون حد أعلى"}`
    : "جميع التواريخ المسجلة"

  useEffect(() => {
    if (blockedPage > blockedPageCount) setBlockedPage(blockedPageCount)
  }, [blockedPage, blockedPageCount])

  const dueReminders = useMemo(() => currentStatementRecords.flatMap(record => [
    ...record.note.actions.flatMap(action => {
      if (action.status === "completed" || action.status === "failed" || action.completed) return []
      const reminders = action.reminderDates?.length
        ? action.reminderDates
        : action.reminderDate ? [{ id: action.id, date: action.reminderDate, notifiedFor: action.notifiedFor }] : []
      return reminders
        .filter(reminder => !!reminder.date && reminder.date <= today)
        .map(reminder => ({
          customerNumber: record.customerNumber,
          customerName: record.name,
          action: { ...action, reminderDate: reminder.date, notifiedFor: reminder.notifiedFor },
          reminderDate: reminder.date,
          source: "action" as const,
        }))
    }),
    ...record.note.agreements.flatMap(agreement => agreement.status !== "active" ? [] : (agreement.reminders ?? [])
      .filter(reminder => !!reminder.date && reminder.date <= today)
      .map(reminder => ({
        customerNumber: record.customerNumber,
        customerName: record.name,
        action: {
          id: reminder.id,
          description: `تذكير باتفاق: ${agreement.description}`,
          dueDate: "",
          reminderDate: reminder.date,
          notifiedFor: reminder.notifiedFor,
        },
        reminderDate: reminder.date,
        source: "agreement" as const,
        agreementId: agreement.id,
      }))),
  ]), [currentStatementRecords, today])

  useEffect(() => {
    onDueRemindersChange?.(dueReminders.length)
  }, [dueReminders.length, onDueRemindersChange])

  useEffect(() => {
    if (!remindersEnabled || notificationPermission !== "granted") return
    const pending = dueReminders.filter(reminder => reminder.action.notifiedFor !== reminder.reminderDate)
    if (!pending.length) return
    try {
      const names = pending.slice(0, 3).map(item => `${item.customerName} (${item.customerNumber})`).join("، ")
      const notification = new Notification("تذكيرات أرصدة الزبائن", {
        body: `${pending.length} تذكير مستحق: ${names}${pending.length > 3 ? "..." : ""}`,
        tag: "customer-balance-reminders",
      })
      notification.onclick = () => {
        window.focus()
        notification.close()
      }
    } catch (cause) {
      setError(cause instanceof Error ? `تعذر إظهار إشعار المتصفح: ${cause.message}` : "تعذر إظهار إشعار المتصفح.")
    }
    const notifiedDates = new Map<string, Set<string>>()
    for (const item of pending) {
      const dates = notifiedDates.get(item.customerNumber) ?? new Set<string>()
      dates.add(item.reminderDate)
      notifiedDates.set(item.customerNumber, dates)
    }
    setDataset(current => current ? {
      ...current,
      records: current.records.map(record => {
        const dates = notifiedDates.get(record.customerNumber)
        if (!dates) return record
        return {
          ...record,
          note: {
            ...record.note,
            actions: record.note.actions.map(action =>
              {
                const reminders = action.reminderDates?.length
                  ? action.reminderDates
                  : action.reminderDate ? [{ id: action.id, date: action.reminderDate, notifiedFor: action.notifiedFor }] : []
                const updatedReminders = reminders.map(reminder =>
                  dates.has(reminder.date) && reminder.date <= today
                    ? { ...reminder, notifiedFor: reminder.date }
                    : reminder,
                )
                const legacyReminder = updatedReminders.find(reminder => reminder.id === action.id)
                return {
                  ...action,
                  reminderDates: action.reminderDates?.length ? updatedReminders : action.reminderDates,
                  ...(legacyReminder ? { notifiedFor: legacyReminder.notifiedFor } : {}),
                }
              },
            ),
            agreements: record.note.agreements.map(agreement => ({
              ...agreement,
              reminders: agreement.reminders?.map(reminder =>
                dates.has(reminder.date) && reminder.date <= today
                  ? { ...reminder, notifiedFor: reminder.date }
                  : reminder,
              ),
            })),
          },
        }
      }),
    } : current)
    setHasChanges(true)
  }, [dueReminders, notificationPermission, remindersEnabled, today])

  const receivableBalance = useMemo(
    () => filteredRecords.reduce((total, record) => total + (record.balanceDirection === "receivable" ? record.balance : 0), 0),
    [filteredRecords],
  )
  const payableBalance = useMemo(
    () => filteredRecords.reduce((total, record) => total + (record.balanceDirection === "payable" ? record.balance : 0), 0),
    [filteredRecords],
  )
  const netBalance = receivableBalance - payableBalance

  const updateCustomerNote = (customerNumber: string, update: (note: CustomerNote) => CustomerNote) => {
    setDataset(current => current ? {
      ...current,
      records: current.records.map(record => record.customerNumber === customerNumber
        ? { ...record, note: update(normalizeCustomerNote(record.note)) }
        : record),
    } : current)
    setHasChanges(true)
    setNotice("")
    setError("")
  }

  const updatePaymentClassification = (
    customerNumber: string,
    paymentClassification: CustomerBalance["paymentClassification"],
  ) => {
    setDataset(current => current ? {
      ...current,
      records: current.records.map(record => record.customerNumber === customerNumber
        ? { ...record, paymentClassification }
        : record),
    } : current)
    setHasChanges(true)
    setNotice("")
    setError("")
  }

  const updatePaymentCriterion = <K extends keyof PaymentCriteria>(key: K, value: PaymentCriteria[K]) => {
    const next = { ...paymentCriteria, [key]: value }
    setPaymentCriteria(next)
    setHasPaymentCriteriaChanges(true)
    setPaymentCriteriaError(isPaymentCriteriaValid(next)
      ? ""
      : "تحقق من أن حدود الدفعات تصاعدية، وأن حدود الائتمان من 0 إلى 1000%، وأيام حداثة السند بين 1 و3650.")
  }

  const saveBlockedCustomers = () => {
    if (!blockEditorNumbers?.length) return
    const customers = new Map((dataset?.records ?? []).map(record => [record.customerNumber, record]))
    const existing = new Map(blockedCustomers.map(customer => [customer.customerNumber, customer]))
    const updated = blockEditorNumbers.map(customerNumber => {
      const customer = customers.get(customerNumber)
      const previous = existing.get(customerNumber)
      return {
        customerNumber,
        name: customer?.name ?? previous?.name ?? customerNumber,
        phone: customer?.phone ?? previous?.phone ?? "",
        reason: (blockReasons[customerNumber] ?? previous?.reason ?? "").trim(),
        blockedAt: previous?.blockedAt ?? new Date().toISOString(),
      }
    })
    setBlockedCustomers(current => [
      ...current.filter(customer => !blockEditorNumbers.includes(customer.customerNumber)),
      ...updated,
    ].sort((left, right) => left.name.localeCompare(right.name, "ar")))
    setHasBlocklistChanges(true)
    setSelectedCustomerNumbers(current => current.filter(number => !blockEditorNumbers.includes(number)))
    setBlockEditorNumbers(null)
    setBlockReasons({})
    setNotice("")
  }

  const updateBlockedReason = (customerNumber: string, reason: string) => {
    setBlockedCustomers(current => current.map(customer => customer.customerNumber === customerNumber
      ? { ...customer, reason }
      : customer))
    setHasBlocklistChanges(true)
  }

  const removeBlockedCustomer = (customerNumber: string) => {
    setBlockedCustomers(current => current.filter(customer => customer.customerNumber !== customerNumber))
    setHasBlocklistChanges(true)
  }

  const toggleCustomerSelection = (customerNumber: string) => {
    setSelectedCustomerNumbers(current => current.includes(customerNumber)
      ? current.filter(number => number !== customerNumber)
      : [...current, customerNumber])
  }

  const updateAction = (customerNumber: string, actionId: string, update: (action: CustomerAction) => CustomerAction) => {
    updateCustomerNote(customerNumber, note => ({
      ...note,
      actions: note.actions.map(action => action.id === actionId ? update(action) : action),
    }))
  }

  const addAction = (customerNumber: string) => {
    updateCustomerNote(customerNumber, note => ({
      ...note,
      agreementStartedAt: note.agreementStartedAt ?? new Date().toISOString(),
      actions: [...note.actions, {
        id: createActionId(),
        description: "",
        dueDate: "",
        reminderDate: "",
        reminderDates: [],
      }],
    }))
  }

  const closeCustomerSummary = (customerNumber: string) => {
    updateCustomerNote(customerNumber, note => {
      if (!note.agreement.trim() && !note.actions.length) return note
      const closedAt = new Date().toISOString()
      return {
        ...note,
        agreement: "",
        agreementStartedAt: undefined,
        actions: [],
        agreementHistory: [...(note.agreementHistory ?? []), {
          id: createActionId(),
          agreement: note.agreement.trim() || "اتفاق ببنود متابعة",
          startedAt: note.agreementStartedAt ?? closedAt,
          closedAt,
          actions: note.actions.map(action =>
            action.status === "completed" && !action.completedAt ? { ...action, completedAt: closedAt } : action,
          ),
        }],
      }
    })
  }

  const reopenCustomerSummary = (customerNumber: string, summaryId: string) => {
    updateCustomerNote(customerNumber, note => {
      const summary = note.agreementHistory?.find(item => item.id === summaryId)
      if (!summary) return note
      if ((note.agreement.trim() || note.actions.length) &&
        !window.confirm("سيتم حفظ الملخص الحالي في السجل ثم فتح الملخص السابق. هل تريد المتابعة؟")) return note
      const history = (note.agreementHistory ?? []).filter(item => item.id !== summaryId)
      const closedAt = new Date().toISOString()
      if (note.agreement.trim() || note.actions.length) {
        history.push({
          id: createActionId(),
          agreement: note.agreement.trim() || "اتفاق ببنود متابعة",
          startedAt: note.agreementStartedAt ?? closedAt,
          closedAt,
          actions: note.actions,
        })
      }
      return {
        ...note,
        agreement: summary.agreement,
        agreementStartedAt: summary.startedAt,
        actions: summary.actions,
        agreementHistory: history,
      }
    })
  }

  const updateAgreement = (customerNumber: string, agreementId: string, update: (agreement: CustomerAgreement) => CustomerAgreement) => {
    updateCustomerNote(customerNumber, note => ({
      ...note,
      agreements: note.agreements.map(agreement => agreement.id === agreementId ? update(agreement) : agreement),
    }))
  }

  const createAgreement = (customerNumber: string, agreement: Omit<CustomerAgreement, "id" | "checks" | "outstandingInstallment" | "status">) => {
    updateCustomerNote(customerNumber, note => ({
      ...note,
      agreements: [...note.agreements, {
        ...agreement,
        id: createActionId(),
        status: "active",
        outstandingInstallment: 0,
        checks: [],
      }],
    }))
  }

  const setAgreementCheckOutcome = (
    customerNumber: string,
    agreementId: string,
    checkId: string,
    outcome: AgreementCheckOutcome,
    actualPaid: number,
    note: string,
  ) => {
    updateAgreement(customerNumber, agreementId, agreement => {
      const check = agreement.checks.find(item => item.id === checkId)
      if (!check) return agreement
      const expected = check.expectedPayment || agreement.installmentAmount || 0
      const paid = Math.max(0, actualPaid)
      const dueCount = Math.max(1, agreementDueCount(agreement, new Date(check.checkedAt)))
      return {
        ...agreement,
        outstandingInstallment: outcome === "partial" ? Math.max(0, expected - paid) : outcome === "paid" ? 0 : outcome === "unpaid" ? expected : agreement.outstandingInstallment,
        nextPaymentDate: outcome === "paid" || outcome === "partial" || outcome === "unpaid"
          ? advanceAgreementDueDate(agreement.nextPaymentDate, agreement.frequency, dueCount)
          : agreement.nextPaymentDate,
        checks: agreement.checks.map(item => item.id === checkId
          ? { ...item, outcome, actualPaid: paid, note: note.trim() || undefined }
          : item),
      }
    })
  }

  const closeAgreement = (
    customerNumber: string,
    agreementId: string,
    status: "completed" | "refused",
    outcome?: Exclude<AgreementCheckOutcome, "ignored">,
    paidAmount = 0,
  ) => {
    updateAgreement(customerNumber, agreementId, agreement => ({
      ...agreement,
      status,
      closedAt: new Date().toISOString(),
      ...(status === "completed" ? {
        closureOutcome: outcome ?? "unpaid",
        closurePaidAmount: paidAmount,
        closureBalance: dataset?.records.find(record => record.customerNumber === customerNumber)?.balance ?? agreement.initialBalance,
        closureBalanceDirection: dataset?.records.find(record => record.customerNumber === customerNumber)?.balanceDirection ?? agreement.balanceDirection,
      } : {}),
    }))
  }

  const deleteAgreement = (customerNumber: string, agreementId: string) => {
    updateCustomerNote(customerNumber, note => ({
      ...note,
      agreements: note.agreements.filter(agreement => agreement.id !== agreementId),
    }))
  }

  const enableBrowserNotifications = async () => {
    if (typeof Notification === "undefined") {
      setNotificationPermission("unsupported")
      return
    }
    try {
      const permission = await Notification.requestPermission()
      setNotificationPermission(permission)
      setError(permission === "denied" ? "إشعارات المتصفح محظورة. غيّر إذن الإشعارات من إعدادات المتصفح لتفعيلها." : "")
    } catch (cause) {
      setError(cause instanceof Error ? `تعذر تفعيل إشعارات المتصفح: ${cause.message}` : "تعذر تفعيل إشعارات المتصفح.")
    }
  }

  const completeAction = (customerNumber: string, actionId: string) => {
    updateAction(customerNumber, actionId, action => ({ ...action, status: "completed", completed: true, failureReason: "" }))
  }

  const dismissAgreementReminder = (customerNumber: string, agreementId: string, reminderId: string) => {
    updateAgreement(customerNumber, agreementId, agreement => ({
      ...agreement,
      reminders: agreement.reminders?.map(reminder => reminder.id === reminderId
        ? { ...reminder, notifiedFor: reminder.date }
        : reminder),
    }))
  }

  const startStatementUpload = (record: CustomerBalance) => {
    const blocked = blockedCustomers.some(customer => customer.customerNumber === record.customerNumber)
    if (blocked) {
      setBlockedContactAction("statement")
      setBlockedContactMessage("")
      setBlockedContactNumber(record.customerNumber)
      return
    }
    setStatementCustomerNumber(record.customerNumber)
    setStatementMode("both")
    setStatementMessageType(activeWeeklyAgreement(record) ? "weekly_installment" : "statement")
    setWeeklyInstallmentAmountInput(String(weeklyInstallmentAmount(record) ?? ""))
    setStatementFile(null)
    setStatementShareError("")
  }

  const startCustomerImage = (record: CustomerBalance, allowBlocked = false) => {
    if (!allowBlocked && blockedCustomers.some(customer => customer.customerNumber === record.customerNumber)) {
      setBlockedContactAction("image")
      setBlockedContactMessage("")
      setBlockedContactNumber(record.customerNumber)
      return
    }
    setCustomerShortcutNumber(null)
    setBulkSendMode("image")
    setBulkQueue([{
      customerNumber: record.customerNumber,
      name: record.name,
      phone: record.phone,
      message: customerMessage(record, selectedBulkMessageTemplate, String(weeklyInstallmentAmount(record) ?? "")),
      balance: record.balance,
      balanceDirection: record.balanceDirection,
    }])
    setBulkQueueIndex(0)
  }

  const updateMessageTemplateContent = (templateId: string, content: string) => {
    if (templateId === "statement") {
      setStatementMessageTemplate(content)
    } else if (templateId === "weekly_installment") {
      setWeeklyInstallmentMessageTemplate(content)
    } else if (templateId.startsWith("custom:")) {
      const id = templateId.slice("custom:".length)
      setCustomMessageTemplates(current => current.map(template =>
        template.id === id ? { ...template, content } : template,
      ))
    }
    setHasMessageTemplateChanges(true)
  }

  const insertMessageField = (token: string) => {
    const textarea = messageTemplateTextAreaRef.current
    const templateId = messageTemplateEditorType
    const currentContent = templateId === "statement"
      ? statementMessageTemplate
      : templateId === "weekly_installment"
        ? weeklyInstallmentMessageTemplate
        : customMessageTemplates.find(template => templateId === `custom:${template.id}`)?.content ?? ""
    const start = textarea?.selectionStart ?? currentContent.length
    const end = textarea?.selectionEnd ?? start
    updateMessageTemplateContent(templateId, `${currentContent.slice(0, start)}${token}${currentContent.slice(end)}`)
    requestAnimationFrame(() => {
      textarea?.focus()
      textarea?.setSelectionRange(start + token.length, start + token.length)
    })
  }

  const updateBulkQueueTemplate = (templateId: string, content: string) => {
    setBulkQueue(current => {
      if (!current) return current
      return current.map(recipient => {
      const record = dataset?.records.find(item => item.customerNumber === recipient.customerNumber)
      return record
        ? { ...recipient, message: customerMessage(record, content, String(weeklyInstallmentAmount(record) ?? "")) }
        : recipient
      })
    })
    updateMessageTemplateContent(templateId, content)
  }

  const insertBulkImageMessageField = (token: string) => {
    const textarea = bulkImageMessageTextAreaRef.current
    const templateId = bulkMessageType
    const start = textarea?.selectionStart ?? selectedBulkMessageTemplate.length
    const end = textarea?.selectionEnd ?? start
    updateBulkQueueTemplate(templateId, `${selectedBulkMessageTemplate.slice(0, start)}${token}${selectedBulkMessageTemplate.slice(end)}`)
    requestAnimationFrame(() => {
      textarea?.focus()
      textarea?.setSelectionRange(start + token.length, start + token.length)
    })
  }

  const addCustomMessageTemplate = () => {
    const name = newMessageTemplateName.trim()
    if (!name) {
      setError("اكتب اسماً للقالب الجديد.")
      return
    }
    const id = typeof crypto !== "undefined" && "randomUUID" in crypto
      ? crypto.randomUUID()
      : `${Date.now()}-${Math.random().toString(36).slice(2)}`
    const template = { id, name, content: "" }
    setCustomMessageTemplates(current => [...current, template])
    setMessageTemplateEditorType(`custom:${id}`)
    setStatementMessageType(`custom:${id}`)
    setBulkMessageType(`custom:${id}`)
    setNewMessageTemplateName("")
    setShowNewMessageTemplate(false)
    setShowNewMessageTemplate(false)
    setHasMessageTemplateChanges(true)
    setError("")
  }

  const chooseStatementFile = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return
    if (!/\.(xlsx|xls|xlsm)$/i.test(file.name)) {
      setError("اختر ملف كشف حساب Excel بصيغة XLSX أو XLS أو XLSM.")
      event.target.value = ""
      return
    }
    setStatementFile(file)
    setStatementShareError("")
    setError("")
    event.target.value = ""
  }

  const shareStatement = async () => {
    if (!statementFile || !statementCustomer || statementMode === "message") return
    if (statementMessageType === "weekly_installment" && !canSendSelectedMessage) {
      setStatementShareError("أدخل مبلغ القسط الأسبوعي قبل إرسال الرسالة.")
      return
    }
    const message = statementMode === "both" ? selectedCustomerMessage : undefined
    try {
      if (!navigator.share || !navigator.canShare?.({ files: [statementFile] })) {
        setStatementShareError("المشاركة المباشرة للملفات غير مدعومة على هذا الجهاز. نزّل الملف ثم افتح واتساب لإرفاقه يدوياً.")
        return
      }
      await navigator.share({
        files: [statementFile],
        title: `كشف حساب ${statementCustomer.name}`,
        text: message,
      })
      setStatementFile(null)
      setStatementCustomerNumber(null)
      setStatementShareError("")
      setNotice("تم فتح قائمة المشاركة. اختر واتساب وأكمل إرسال الملف والرسالة.")
    } catch (cause) {
      if (cause instanceof Error && cause.name === "AbortError") return
      setStatementShareError(cause instanceof Error ? `تعذرت مشاركة الملف: ${cause.message}. يمكنك تنزيله وإرفاقه يدوياً في واتساب.` : "تعذرت مشاركة الملف. يمكنك تنزيله وإرفاقه يدوياً في واتساب.")
    }
  }

  const downloadStatementFile = () => {
    if (!statementFile) return
    const url = URL.createObjectURL(statementFile)
    const link = document.createElement("a")
    link.href = url
    link.download = statementFile.name
    document.body.appendChild(link)
    link.click()
    link.remove()
    window.setTimeout(() => URL.revokeObjectURL(url), 1000)
  }

  const shareSelectedStatementDirectly = statementMode !== "message" &&
    typeof navigator !== "undefined" &&
    !!navigator.share &&
    !!statementFile &&
    !!navigator.canShare?.({ files: [statementFile] })

  const importFile = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return
    setBusy(true)
    setError("")
    setNotice("")
    try {
      if (!/\.(xlsx|xls|xlsm|csv)$/i.test(file.name)) throw new Error("ارفع ملف Excel بصيغة XLSX أو XLS أو XLSM، أو ملف CSV.")
      const records = parseWorkbook(await file.arrayBuffer())
      if (!records.length) {
        throw new Error("لم يتم العثور على بيانات زبائن. تأكد من وجود ترويسة للأعمدة (رقم الزبون أو رقم الصفحة، البيان أو اسم الزبون، الرصيد، المدة، بيانات الفاتورة والسند، والجوال) داخل إحدى صفحات الملف.")
      }
      const uniqueRecords = [...new Map(records.map(record => [record.customerNumber, record])).values()]
      const importedAt = new Date().toISOString()
      const merged = mergeCustomerRecords(dataset?.records ?? [], uniqueRecords, importedAt, missingCustomerPolicy)
      const importedByNumber = new Map(uniqueRecords.map(record => [record.customerNumber, record]))
      const refreshedBlocklist = blockedCustomers.map(blocked => {
        const imported = importedByNumber.get(blocked.customerNumber)
        return imported ? { ...blocked, name: imported.name, phone: imported.phone } : blocked
      })
      if (refreshedBlocklist.some((blocked, index) =>
        blocked.name !== blockedCustomers[index]?.name || blocked.phone !== blockedCustomers[index]?.phone,
      )) {
        setBlockedCustomers(refreshedBlocklist)
        setHasBlocklistChanges(true)
      }
      setDataset({ fileName: file.name, importedAt, records: merged.records })
      setHasChanges(true)
      setFilters(defaultFilters)
      const missingCount = merged.records.filter(record => record.isInLatestStatement === false).length
      setNotice(`تم تحديث ${merged.updated.toLocaleString("ar")} زبوناً وإضافة ${merged.added.toLocaleString("ar")} جديداً. ${missingCustomerPolicy === "keep_zero"
        ? `تم إبقاء ${missingCount.toLocaleString("ar")} زبوناً غير ظاهر برصيد صفر.`
        : `تم الاحتفاظ بسجلات ${missingCount.toLocaleString("ar")} زبوناً غير ظاهر دون عرضهم في الأرصدة الحالية.`}`)
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "تعذرت معالجة ملف Excel.")
    } finally {
      setBusy(false)
      event.target.value = ""
    }
  }

  const clearDataset = async () => {
    if (!dataset || !window.confirm("سيؤدي حذف الكشف إلى حذف بيانات الأرصدة وملاحظات واتفاقات الزبائن وسجل المقارنات المرتبط به نهائياً. ستبقى قائمة حظر التواصل وملاحظات منصة التذكيرات محفوظة. هل تريد المتابعة؟")) return
    setDataset(null)
    setOpenCustomerNumber(null)
    setCustomerShortcutNumber(null)
    setStatementCustomerNumber(null)
    setStatementFile(null)
    setSelectedCustomerNumbers([])
    setHasChanges(false)
    setFilters(defaultFilters)
    setNotice("")
    setError("")
    try {
      if (centralStorageEnabled) {
        await centralStorageSet(deletedDatasetKey, true)
        await centralStorageSet(storageKey, null)
      } else {
        localStorage.setItem(deletedDatasetKey, "1")
        localStorage.removeItem(storageKey)
      }
      setNotice("تم حذف كشف الأرصدة وملاحظاته واتفاقاته نهائياً.")
    } catch (cause) {
      setError(cause instanceof Error ? `تعذر حفظ حذف الكشف: ${cause.message}` : "تعذر حفظ حذف الكشف.")
    }
  }

  const exportFiltered = () => {
    if (!filteredRecords.length) return
    const headers = [...DISPLAY_COLUMNS.map(column => column.label), "الاتفاق", "بنود الاتفاق والتذكير"]
    const rows = filteredRecords.map(record => [
      ...DISPLAY_COLUMNS.map(column => column.key === "balance"
        ? `${formatAmount(record.balance)}${record.balanceDirection === "receivable" ? "م" : record.balanceDirection === "payable" ? "د" : ""}`
        : record[column.key]),
      record.note.agreement,
      record.note.actions.map((action, index) => [
        `${index + 1}. ${action.description}`,
        action.dueDate ? `موعد التنفيذ: ${formatDate(action.dueDate)}` : "",
        action.reminderDate ? `التذكير: ${formatDate(action.reminderDate)}` : "",
        action.status === "completed" || action.completed ? "تم ✓" : action.status === "failed" ? `لم ينفذ${action.failureReason ? `: ${action.failureReason}` : ""}` : "بانتظار التنفيذ",
      ].filter(Boolean).join(" | ")).join("\n"),
    ])
    const sheet = XLSX.utils.aoa_to_sheet([headers, ...rows])
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, sheet, "أرصدة الزبائن")
    XLSX.writeFile(workbook, "أرصدة-الزبائن-المفلترة.xlsx")
  }

  const exportBlockedBalances = () => {
    if (!blockedBalanceRows.length) return
    const rows = blockedBalanceRows.map(({ customer, record }) => [
      customer.customerNumber,
      record?.name ?? customer.name,
      record?.phone ?? customer.phone,
      record ? record.balance : "غير متوفر",
      record ? record.balanceDirection === "receivable" ? "مستحق لنا" : record.balanceDirection === "payable" ? "مستحق للزبون" : "غير محدد" : "لا يوجد كشف",
      customer.reason,
    ])
    const worksheet = XLSX.utils.aoa_to_sheet([
      ["تقرير أرصدة المحظورين من التواصل"],
      ["تاريخ التقرير", new Date().toLocaleString("ar")],
      ["عدد المحظورين", blockedBalanceRows.length],
      ["عدد دون رصيد في الكشف الحالي", blockedMissingBalanceCount],
      [],
      ["رقم الزبون", "اسم الزبون", "الجوال", "الرصيد الحالي", "اتجاه الرصيد", "سبب الحظر"],
      ...rows,
      [],
      ["", "إجمالي مستحق لنا (م)", "", blockedReceivableTotal, "", ""],
      ["", "إجمالي مستحق للزبائن (د)", "", blockedPayableTotal, "", ""],
      ["", "الصافي (م − د)", "", blockedNetTotal, "", ""],
    ])
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, "أرصدة المحظورين")
    XLSX.writeFile(workbook, "تقرير-أرصدة-المحظورين.xlsx")
  }

  const exportBlockedBalancesPdf = () => {
    if (!blockedBalanceRows.length) return
    const printWindow = window.open("", "_blank", "width=1100,height=800")
    if (!printWindow) {
      window.alert("تعذر فتح تقرير PDF. اسمح بالنوافذ المنبثقة ثم أعد المحاولة.")
      return
    }
    const escapeHtml = (value: unknown) => String(value ?? "").replace(/[&<>"']/g, character => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
    })[character] ?? character)
    const rows = blockedBalanceRows.map(({ customer, record }) => `<tr>
      <td>${escapeHtml(customer.customerNumber)}</td>
      <td>${escapeHtml(record?.name ?? customer.name)}</td>
      <td dir="ltr">${escapeHtml((record?.phone ?? customer.phone) || "—")}</td>
      <td>${escapeHtml(record ? formatAmount(record.balance) : "غير متوفر")}</td>
      <td>${escapeHtml(record ? record.balanceDirection === "receivable" ? "مستحق لنا" : record.balanceDirection === "payable" ? "مستحق للزبون" : "غير محدد" : "لا يوجد كشف")}</td>
      <td>${escapeHtml(customer.reason || "—")}</td>
    </tr>`).join("")
    printWindow.document.open()
    printWindow.document.write(`<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>تقرير أرصدة المحظورين</title><style>${REPORT_LOGO_PRINT_STYLES}
      @page{size:A4 landscape;margin:14mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;color:#172033;font-size:11px;line-height:1.6}h1{font-size:20px;margin:0 0 6px}.meta{color:#4b5563;margin-bottom:14px}table{width:100%;border-collapse:collapse}th,td{border:1px solid #cbd5e1;padding:6px;text-align:right;vertical-align:top}th{background:#eef2f6}.totals{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-top:16px}.total{border:1px solid #cbd5e1;border-radius:6px;padding:10px;font-weight:bold}.total strong{display:block;margin-top:4px;font-size:16px}@media print{body{print-color-adjust:exact;-webkit-print-color-adjust:exact}}
      </style></head><body>${reportLogoMarkup()}<h1>تقرير أرصدة المحظورين من التواصل</h1><div class="meta">تاريخ التقرير: ${escapeHtml(new Date().toLocaleString("ar"))} · عدد المحظورين: ${blockedBalanceRows.length.toLocaleString("ar")} · دون رصيد في الكشف الحالي: ${blockedMissingBalanceCount.toLocaleString("ar")}</div><table><thead><tr><th>رقم الزبون</th><th>اسم الزبون</th><th>الجوال</th><th>الرصيد الحالي</th><th>اتجاه الرصيد</th><th>سبب الحظر</th></tr></thead><tbody>${rows}</tbody></table><div class="totals"><div class="total">إجمالي مستحق لنا (م)<strong>${escapeHtml(formatAmount(blockedReceivableTotal))}</strong></div><div class="total">إجمالي مستحق للزبائن (د)<strong>${escapeHtml(formatAmount(blockedPayableTotal))}</strong></div><div class="total">الصافي (م − د)<strong>${escapeHtml(formatAmount(blockedNetTotal))}</strong></div></div><script>window.addEventListener('load',()=>{window.focus();window.print()},{once:true})</script></body></html>`)
    printWindow.document.close()
  }

  const exportCustomerReport = () => {
    const records = reportRecords
    if (!records.length) return

    const formatReportDate = (value: string) => {
      if (!value) return ""
      if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return formatDate(value)
      const date = new Date(value)
      return Number.isNaN(date.getTime()) ? value : date.toLocaleString("ar")
    }
    const directionLabel = (direction: CustomerBalance["balanceDirection"]) =>
      direction === "receivable" ? "مستحق لنا" : direction === "payable" ? "مستحق للزبون" : "غير محدد"
    const agreementStatusLabel = (status: AgreementStatus) =>
      status === "active" ? "قيد التنفيذ" : status === "completed" ? "مكتمل" : "مرفوض"
    const frequencyLabel = (frequency: AgreementFrequency) =>
      frequency === "weekly" ? "أسبوعياً" : frequency === "monthly" ? "شهرياً" : "مرة واحدة"
    const outcomeLabel = (outcome?: AgreementCheckOutcome) =>
      outcome === "paid" ? "دفع كامل" : outcome === "partial" ? "دفع جزئي" : outcome === "unpaid" ? "لم يسدد" : outcome === "ignored" ? "تم التجاهل" : "بانتظار المراجعة"
    const actionStatusLabel = (action: CustomerAction) =>
      action.status === "completed" || action.completed ? "تم" : action.status === "failed" ? "لم ينفذ" : "قيد المتابعة"

    const summaryRows = records.map(record => [
      record.customerNumber,
      record.name,
      record.phone,
      record.firstStatementBalance ?? record.balance,
      directionLabel(record.firstStatementDirection ?? record.balanceDirection),
      formatReportDate(record.firstStatementAt ?? ""),
      record.balance,
      directionLabel(record.balanceDirection),
      formatReportDate(record.balanceUpdatedAt ?? dataset?.importedAt ?? ""),
      hasCustomerNotes(record) ? "نعم" : "لا",
      paymentTotalsByCustomer.get(record.customerNumber) ?? 0,
      record.note.agreement,
      record.note.actions.length,
      record.note.agreements.length,
      record.invoiceAmount,
      formatDate(record.invoiceDate),
      record.receiptAmount,
      formatDate(record.receiptDate),
    ])
    const agreementRows = records.flatMap(record => record.note.agreements.map(agreement => {
      const latestSnapshot = agreement.balanceSnapshots?.at(-1)
      const latestCheck = agreement.checks.at(-1)
      return [
        record.customerNumber,
        record.name,
        record.balance,
        directionLabel(record.balanceDirection),
        formatReportDate(record.balanceUpdatedAt ?? dataset?.importedAt ?? ""),
        agreement.description,
        agreementStatusLabel(agreement.status),
        formatReportDate(agreement.startedAt),
        agreement.initialBalance,
        directionLabel(agreement.balanceDirection),
        agreement.installmentAmount ?? "غير محدد",
        frequencyLabel(agreement.frequency),
        formatDate(agreement.nextPaymentDate),
        agreement.outstandingInstallment,
        latestSnapshot?.balance ?? latestCheck?.newBalance ?? agreement.initialBalance,
        directionLabel(latestSnapshot?.direction ?? latestCheck?.newDirection ?? agreement.balanceDirection),
        latestSnapshot ? formatReportDate(latestSnapshot.capturedAt) : latestCheck ? formatReportDate(latestCheck.checkedAt) : "",
        agreement.closureOutcome ? outcomeLabel(agreement.closureOutcome) : "",
        agreement.closurePaidAmount ?? "",
        formatReportDate(agreement.closedAt ?? ""),
        agreement.closureBalance ?? "",
      ]
    }))
    const noteRows = records.flatMap(record => {
      if (!hasCustomerNotes(record)) return []
      const rows: unknown[][] = []
      if (record.note.agreement.trim()) {
        rows.push([record.customerNumber, record.name, "ملخص الاتفاق/الملاحظة", record.note.agreement, "", "", ""])
      }
      for (const [index, action] of record.note.actions.entries()) {
        rows.push([
          record.customerNumber,
          record.name,
          `بند متابعة ${index + 1}`,
          action.description,
          actionStatusLabel(action),
          formatDate(action.dueDate),
          formatDate(action.reminderDate),
          action.failureReason ?? "",
        ])
      }
      return rows
    })
    const historyRows: unknown[][] = []
    for (const record of records) {
      historyRows.push([
        record.customerNumber,
        record.name,
        "رصيد أول كشف محفوظ",
        formatReportDate(record.firstStatementAt ?? ""),
        "",
        "",
        record.firstStatementBalance ?? record.balance,
        directionLabel(record.firstStatementDirection ?? record.balanceDirection),
        "",
      ])
      historyRows.push([
        record.customerNumber,
        record.name,
        "الرصيد الحالي",
        formatReportDate(record.balanceUpdatedAt ?? dataset?.importedAt ?? ""),
        "",
        "",
        record.balance,
        directionLabel(record.balanceDirection),
        "",
      ])
      for (const agreement of record.note.agreements) {
        for (const snapshot of agreement.balanceSnapshots ?? []) {
          historyRows.push([
            record.customerNumber,
            record.name,
            `لقطة رصيد الاتفاق: ${agreement.description}`,
            formatReportDate(snapshot.capturedAt),
            "",
            "",
            snapshot.balance,
            directionLabel(snapshot.direction),
            "",
          ])
        }
        for (const check of agreement.checks) {
          historyRows.push([
            record.customerNumber,
            record.name,
            `مراجعة الاتفاق: ${agreement.description}`,
            formatReportDate(check.checkedAt),
            check.previousBalance,
            directionLabel(check.previousDirection),
            check.newBalance,
            directionLabel(check.newDirection),
            [
              `المدفوع المستنتج: ${check.inferredPaid}`,
              `المتوقع: ${check.expectedPayment}`,
              `النتيجة: ${outcomeLabel(check.outcome)}`,
              check.note ?? "",
            ].filter(Boolean).join(" | "),
          ])
        }
      }
    }
    const paymentSummaryRows = reportRecords
      .map(record => [record.customerNumber, record.name, paymentTotalsByCustomer.get(record.customerNumber) ?? 0])
      .filter(([, , amount]) => typeof amount === "number" && amount > 0)
    const paymentDetailRows = reportPayments.map(payment => [
      payment.customerNumber,
      payment.customerName,
      formatReportDate(payment.date),
      payment.amount,
      payment.type,
      payment.agreement,
    ])

    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
      ["رقم الزبون", "اسم الزبون", "الجوال", "رصيد أول كشف محفوظ", "اتجاه الرصيد الأول", "تاريخ أول كشف محفوظ", "الرصيد الحالي", "اتجاه الرصيد الحالي", "تاريخ آخر تحديث للرصيد", "لديه ملاحظات/اتفاقات", `إجمالي المسدد (${paymentPeriodLabel})`, "ملخص الملاحظات", "عدد بنود المتابعة", "عدد الاتفاقات", "آخر فاتورة", "تاريخها", "آخر سند قبض", "تاريخه"],
      ...summaryRows,
    ]), "ملخص الأرصدة")
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
      ["رقم الزبون", "اسم الزبون", "الرصيد الحالي", "اتجاهه", "تاريخ تحديثه", "تفاصيل الاتفاق", "حالة الاتفاق", "تاريخ الاتفاق", "دين الزبون عند الاتفاق", "اتجاهه", "مبلغ الدفعة", "تكرارها", "الدفعة القادمة", "المتبقي من الدفعة", "آخر رصيد مسجل للاتفاق", "اتجاهه", "تاريخ آخر رصيد للاتفاق", "نتيجة الإنهاء", "المبلغ المدفوع عند الإنهاء", "تاريخ الإنهاء", "رصيد الإنهاء"],
      ...agreementRows,
    ]), "الاتفاقات")
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
      ["رقم الزبون", "اسم الزبون", "نوع الملاحظة", "التفاصيل", "الحالة", "موعد التنفيذ", "موعد التذكير", "سبب عدم التنفيذ"],
      ...noteRows,
    ]), "الملاحظات")
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
      ["رقم الزبون", "اسم الزبون", "نوع السجل", "التاريخ", "الرصيد السابق", "اتجاهه", "الرصيد اللاحق/المسجل", "اتجاهه", "تفاصيل الدفع والنتيجة"],
      ...historyRows,
    ]), "تاريخ الأرصدة")
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
      ["فترة التقرير", paymentPeriodLabel],
      ["رقم الزبون", "اسم الزبون", "إجمالي ما سدده خلال الفترة"],
      ...paymentSummaryRows,
      ["", "الإجمالي", totalPayments],
      [],
      ["رقم الزبون", "اسم الزبون", "تاريخ الدفعة", "المبلغ", "نوع الدفعة", "الاتفاق"],
      ...paymentDetailRows,
      ["", "الإجمالي", "", totalPayments, "", ""],
    ]), "إجمالي المسدد")
    XLSX.writeFile(workbook, "تقرير-زبائن-الملاحظات-والاتفاقات.xlsx")
  }

  const exportCustomerReportPdf = () => {
    const records = reportRecords
    if (!records.length) return

    const printWindow = window.open("", "_blank", "width=1200,height=850")
    if (!printWindow) {
      window.alert("تعذر فتح تقرير PDF. اسمح بالنوافذ المنبثقة ثم أعد المحاولة.")
      return
    }

    const escapeHtml = (value: unknown) => String(value ?? "").replace(/[&<>"']/g, character => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
    })[character] ?? character)
    const reportDate = (value: string) => {
      if (!value) return "—"
      if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return formatDate(value)
      const date = new Date(value)
      return Number.isNaN(date.getTime()) ? value : date.toLocaleString("ar")
    }
    const directionLabel = (direction: CustomerBalance["balanceDirection"]) =>
      direction === "receivable" ? "مستحق لنا" : direction === "payable" ? "مستحق للزبون" : "غير محدد"
    const agreementStatusLabel = (status: AgreementStatus) =>
      status === "active" ? "قيد التنفيذ" : status === "completed" ? "مكتمل" : "مرفوض"
    const frequencyLabel = (frequency: AgreementFrequency) =>
      frequency === "weekly" ? "أسبوعياً" : frequency === "monthly" ? "شهرياً" : "مرة واحدة"
    const outcomeLabel = (outcome?: AgreementCheckOutcome) =>
      outcome === "paid" ? "دفع كامل" : outcome === "partial" ? "دفع جزئي" : outcome === "unpaid" ? "لم يسدد" : outcome === "ignored" ? "تم التجاهل" : "بانتظار المراجعة"
    const actionStatusLabel = (action: CustomerAction) =>
      action.status === "completed" || action.completed ? "تم" : action.status === "failed" ? "لم ينفذ" : "قيد المتابعة"
    const renderTable = (headers: string[], rows: unknown[][], emptyMessage: string) =>
      `<table><thead><tr>${headers.map(header => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${rows.length
        ? rows.map(row => `<tr>${row.map(cell => `<td>${escapeHtml(cell === "" || cell === undefined || cell === null ? "—" : cell)}</td>`).join("")}</tr>`).join("")
        : `<tr><td colspan="${headers.length}" class="empty">${escapeHtml(emptyMessage)}</td></tr>`}</tbody></table>`

    const summaryRows = records.map(record => [
      record.customerNumber,
      record.name,
      record.phone || "—",
      `${formatAmount(record.firstStatementBalance ?? record.balance)} · ${directionLabel(record.firstStatementDirection ?? record.balanceDirection)}`,
      reportDate(record.firstStatementAt ?? ""),
      `${formatAmount(record.balance)} · ${directionLabel(record.balanceDirection)}`,
      reportDate(record.balanceUpdatedAt ?? dataset?.importedAt ?? ""),
      hasCustomerNotes(record) ? "نعم" : "لا",
      formatAmount(paymentTotalsByCustomer.get(record.customerNumber) ?? 0),
    ])
    const agreementRows = records.flatMap(record => record.note.agreements.map(agreement => {
      const latestSnapshot = agreement.balanceSnapshots?.at(-1)
      const latestCheck = agreement.checks.at(-1)
      return [
        record.customerNumber,
        record.name,
        agreement.description,
        agreementStatusLabel(agreement.status),
        reportDate(agreement.startedAt),
        `${formatAmount(agreement.initialBalance)} · ${directionLabel(agreement.balanceDirection)}`,
        agreement.installmentAmount === null ? "غير محدد" : formatAmount(agreement.installmentAmount),
        frequencyLabel(agreement.frequency),
        reportDate(agreement.nextPaymentDate),
        `${formatAmount(latestSnapshot?.balance ?? latestCheck?.newBalance ?? agreement.initialBalance)} · ${directionLabel(latestSnapshot?.direction ?? latestCheck?.newDirection ?? agreement.balanceDirection)}`,
        agreement.closureOutcome ? outcomeLabel(agreement.closureOutcome) : "—",
        agreement.closurePaidAmount === undefined ? "—" : formatAmount(agreement.closurePaidAmount),
        reportDate(agreement.closedAt ?? ""),
      ]
    }))
    const noteRows = records.flatMap(record => {
      const rows: unknown[][] = []
      if (record.note.agreement.trim()) {
        rows.push([record.customerNumber, record.name, "ملخص الاتفاق/الملاحظة", record.note.agreement, "—", "—", "—", "—"])
      }
      for (const [index, action] of record.note.actions.entries()) {
        rows.push([
          record.customerNumber,
          record.name,
          `بند متابعة ${index + 1}`,
          action.description,
          actionStatusLabel(action),
          reportDate(action.dueDate),
          reportDate(action.reminderDate),
          action.failureReason ?? "—",
        ])
      }
      return rows
    })
    const historyRows: unknown[][] = records.flatMap(record => {
      const rows: unknown[][] = [
        [
          record.customerNumber, record.name, "أول رصيد محفوظ", reportDate(record.firstStatementAt ?? ""),
          "—", "—", formatAmount(record.firstStatementBalance ?? record.balance),
          directionLabel(record.firstStatementDirection ?? record.balanceDirection), "—",
        ],
        [
          record.customerNumber, record.name, "الرصيد الحالي", reportDate(record.balanceUpdatedAt ?? dataset?.importedAt ?? ""),
          "—", "—", formatAmount(record.balance), directionLabel(record.balanceDirection), "—",
        ],
      ]
      for (const agreement of record.note.agreements) {
        for (const snapshot of agreement.balanceSnapshots ?? []) {
          rows.push([
            record.customerNumber, record.name, `لقطة رصيد: ${agreement.description}`, reportDate(snapshot.capturedAt),
            "—", "—", formatAmount(snapshot.balance), directionLabel(snapshot.direction), "—",
          ])
        }
        for (const check of agreement.checks) {
          rows.push([
            record.customerNumber, record.name, `مراجعة اتفاق: ${agreement.description}`, reportDate(check.checkedAt),
            formatAmount(check.previousBalance), directionLabel(check.previousDirection),
            formatAmount(check.newBalance), directionLabel(check.newDirection),
            `المتوقع: ${formatAmount(check.expectedPayment)} · المستنتج دفعه: ${formatAmount(check.inferredPaid)} · النتيجة: ${outcomeLabel(check.outcome)}${check.note ? ` · ${check.note}` : ""}`,
          ])
        }
      }
      return rows
    })
    const paymentSummaryRows: unknown[][] = reportRecords
      .map(record => [
        record.customerNumber,
        record.name,
        formatAmount(paymentTotalsByCustomer.get(record.customerNumber) ?? 0),
      ])
      .filter(row => paymentTotalsByCustomer.has(String(row[0])))
    const paymentDetailRows = reportPayments.map(payment => [
      payment.customerNumber,
      payment.customerName,
      reportDate(payment.date),
      formatAmount(payment.amount),
      payment.type,
      payment.agreement,
    ])
    const paymentSummaryWithTotal = [
      ...paymentSummaryRows,
      ["", "الإجمالي", formatAmount(totalPayments)],
    ]
    const paymentDetailsWithTotal = [
      ...paymentDetailRows,
      ["", "الإجمالي", "", formatAmount(totalPayments), "", ""],
    ]

    const sections = [
      `<section><h2>ملخص الأرصدة (${records.length.toLocaleString("ar")} زبون)</h2>${renderTable(
        ["رقم الزبون", "اسم الزبون", "الجوال", "رصيد أول كشف محفوظ", "تاريخ أول كشف", "الرصيد الحالي", "تاريخ آخر تحديث", "ملاحظات/اتفاقات", "المسدد خلال الفترة"],
        summaryRows,
        "لا توجد أرصدة مسجلة.",
      )}</section>`,
      `<section><h2>الاتفاقات (${agreementRows.length.toLocaleString("ar")} اتفاق)</h2>${renderTable(
        ["رقم الزبون", "اسم الزبون", "تفاصيل الاتفاق", "الحالة", "تاريخه", "دين البداية", "الدفعة", "التكرار", "موعدها القادم", "آخر رصيد مسجل", "نتيجة الإنهاء", "المبلغ المدفوع", "تاريخ الإنهاء"],
        agreementRows,
        "لا توجد اتفاقات مسجلة.",
      )}</section>`,
      `<section><h2>الملاحظات وبنود المتابعة</h2>${renderTable(
        ["رقم الزبون", "اسم الزبون", "النوع", "التفاصيل", "الحالة", "موعد التنفيذ", "موعد التذكير", "سبب عدم التنفيذ"],
        noteRows,
        "لا توجد ملاحظات أو بنود متابعة مسجلة.",
      )}</section>`,
      `<section><h2>إجمالي المسدد حسب الفترة</h2><p class="payment-total">الفترة: ${escapeHtml(paymentPeriodLabel)} · إجمالي المبالغ المسددة: ${escapeHtml(formatAmount(totalPayments))}</p>${renderTable(
        ["رقم الزبون", "اسم الزبون", "إجمالي ما سدده خلال الفترة"],
        paymentSummaryWithTotal,
        "لا توجد دفعات مسددة مسجلة ضمن الفترة المحددة.",
      )}<h3>تفاصيل الدفعات</h3>${renderTable(
        ["رقم الزبون", "اسم الزبون", "تاريخ الدفعة", "المبلغ", "نوع الدفعة", "الاتفاق"],
        paymentDetailsWithTotal,
        "لا توجد دفعات مسددة مسجلة ضمن الفترة المحددة.",
      )}</section>`,
      `<section><h2>تاريخ الأرصدة</h2>${renderTable(
        ["رقم الزبون", "اسم الزبون", "نوع السجل", "التاريخ", "الرصيد السابق", "اتجاهه", "الرصيد المسجل", "اتجاهه", "تفاصيل الدفع والنتيجة"],
        historyRows,
        "لا يوجد سجل أرصدة.",
      )}</section>`,
      `<section><h2>الإجمالي النهائي للمبالغ المسددة</h2><p class="payment-total">الفترة: ${escapeHtml(paymentPeriodLabel)} · إجمالي ما سدده جميع الزبائن: ${escapeHtml(formatAmount(totalPayments))}</p></section>`,
    ].join("")

    printWindow.document.open()
    printWindow.document.write(`<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>تقرير زبائن الملاحظات والاتفاقات</title><style>${REPORT_LOGO_PRINT_STYLES}
      @page{size:A4 landscape;margin:10mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;color:#172033;font-size:9px;line-height:1.5}h1{font-size:20px;margin:0 0 4px}h2{font-size:14px;margin:18px 0 7px;border-bottom:1px solid #ccd5df;padding-bottom:4px}h3{font-size:11px;margin:14px 0 5px}.meta{color:#4b5563;margin-bottom:14px}.payment-total{border:1px solid #a7f3d0;background:#ecfdf5;padding:9px;font-weight:bold;font-size:12px}section{break-before:page}section:first-of-type{break-before:auto}table{width:100%;border-collapse:collapse;table-layout:auto}th,td{border:1px solid #cbd5e1;padding:4px;text-align:right;vertical-align:top;overflow-wrap:anywhere}th{background:#eef2f6;font-weight:bold}thead{display:table-header-group}tr{break-inside:avoid}.empty{text-align:center;color:#64748b}@media print{body{print-color-adjust:exact;-webkit-print-color-adjust:exact}}
      </style></head><body>${reportLogoMarkup()}<h1>تقرير زبائن الملاحظات والاتفاقات</h1><div class="meta">الملف: ${escapeHtml(dataset?.fileName ?? "")} · عدد الزبائن: ${records.length.toLocaleString("ar")} · تاريخ التقرير: ${escapeHtml(new Date().toLocaleString("ar"))}</div>${sections}<script>window.addEventListener('load',()=>{window.focus();window.print()},{once:true})</script></body></html>`)
    printWindow.document.close()
  }

  const reportTypeLabels: Record<CustomerReportType, string> = {
    balances: "ملخص الأرصدة",
    agreements: "تقرير الاتفاقات",
    notes: "تقرير الملاحظات وبنود المتابعة",
    paid_total: "إجمالي المسدد حسب الزبون",
    payments: "تفاصيل الدفعات",
    blocked_balances: "أرصدة المحظورين",
    missing_phone: "الزبائن بدون أرقام جوال صالحة",
    payment_classification: "تصنيف الزبائن حسب الدفع",
  }
  const reportFiltersValid = reportType === "payment_classification" || reportType === "missing_phone" || (
    paymentDateRangeValid &&
    (!reportAmountFrom || !reportAmountTo || Number(reportAmountFrom) <= Number(reportAmountTo)) &&
    (!reportAmountFrom || (Number.isFinite(Number(reportAmountFrom)) && Number(reportAmountFrom) >= 0)) &&
    (!reportAmountTo || (Number.isFinite(Number(reportAmountTo)) && Number(reportAmountTo) >= 0))
  )

  const createSelectedReport = () => {
    type ReportRow = { values: Array<string | number>; date: string; amount?: number }
    let headers: string[]
    let sourceRows: ReportRow[]
    let amountColumn: number

    if (reportType === "payment_classification") {
      headers = ["رقم الزبون", "اسم الزبون", "الجوال", "تصنيف النظام", "تصنيفك اليدوي", "الرصيد المستحق", "آخر سند قبض", "نسبة السداد", "حد الائتمان المقترح", "آخر فاتورة", "التنبيهات"]
      sourceRows = paymentClassificationReportRows(dataset?.records ?? [], effectivePaymentCriteria, now).map(({ record, assessment, manualClassification, note }) => ({
        date: "",
        values: [
          record.customerNumber,
          record.name,
          record.phone || "—",
          automaticClassificationLabel(assessment.classification),
          manualClassification,
          `${formatAmount(record.balance)}${record.balanceDirection === "receivable" ? " م" : record.balanceDirection === "payable" ? " د" : ""}`,
          record.receiptDate ? `${formatAmount(record.receiptAmount)} · ${formatDate(record.receiptDate)}` : "—",
          assessment.paymentPercent === null ? "—" : `${assessment.paymentPercent.toFixed(1)}%`,
          assessment.creditLimit === null ? "—" : formatAmount(assessment.creditLimit),
          record.invoiceDate ? `${formatAmount(record.invoiceAmount)} · ${formatDate(record.invoiceDate)}` : "—",
          note,
        ],
      }))
      amountColumn = -1
    } else if (reportType === "missing_phone") {
      headers = ["رقم الزبون", "اسم الزبون", "رقم الهاتف المسجل", "الرصيد", "اتجاه الرصيد"]
      sourceRows = (dataset?.records ?? [])
        .filter(record => !whatsappUrl(record.phone))
        .map(record => ({
          date: "",
          values: [
            record.customerNumber,
            record.name,
            record.phone || "—",
            formatAmount(record.balance),
            record.balanceDirection === "receivable" ? "مستحق لنا" : record.balanceDirection === "payable" ? "مستحق للزبون" : "غير محدد",
          ],
        }))
      amountColumn = -1
    } else if (reportType === "balances") {
      headers = ["رقم الزبون", "اسم الزبون", "الجوال", "رصيد أول كشف", "تاريخ أول كشف", "الرصيد الحالي", "اتجاه الرصيد", "تاريخ آخر تحديث", "لديه ملاحظات/اتفاقات", "المبلغ"]
      sourceRows = reportRecords.map(record => ({
        date: record.balanceUpdatedAt ?? dataset?.importedAt ?? "",
        amount: record.balance,
        values: [
          record.customerNumber, record.name, record.phone || "—",
          record.firstStatementBalance ?? record.balance,
          record.firstStatementAt ? formatDate(record.firstStatementAt.slice(0, 10)) : "—",
          record.balance,
          record.balanceDirection === "receivable" ? "مستحق لنا" : record.balanceDirection === "payable" ? "مستحق للزبون" : "غير محدد",
          (record.balanceUpdatedAt ?? dataset?.importedAt ?? "").slice(0, 10) || "—",
          hasCustomerNotes(record) ? "نعم" : "لا",
          record.balance,
        ],
      }))
      amountColumn = 9
    } else if (reportType === "agreements") {
      headers = ["رقم الزبون", "اسم الزبون", "تفاصيل الاتفاق", "الحالة", "تاريخ الاتفاق", "رصيد البداية", "مبلغ الدفعة", "التكرار", "الدفعة القادمة", "المبلغ"]
      sourceRows = reportRecords.flatMap(record => record.note.agreements.map(agreement => ({
        date: agreement.startedAt.slice(0, 10),
        amount: agreement.initialBalance,
        values: [
          record.customerNumber, record.name, agreement.description,
          agreement.status === "active" ? "قيد التنفيذ" : agreement.status === "completed" ? "مكتمل" : "مرفوض",
          agreement.startedAt.slice(0, 10),
          agreement.initialBalance,
          agreement.installmentAmount ?? "غير محدد",
          agreement.frequency === "weekly" ? "أسبوعياً" : agreement.frequency === "monthly" ? "شهرياً" : "مرة واحدة",
          agreement.nextPaymentDate,
          agreement.initialBalance,
        ],
      })))
      amountColumn = 9
    } else if (reportType === "notes") {
      headers = ["رقم الزبون", "اسم الزبون", "نوع الملاحظة", "التفاصيل", "الحالة", "موعد التنفيذ", "موعد التذكير", "سبب عدم التنفيذ"]
      sourceRows = reportRecords.flatMap(record => {
        const rows: ReportRow[] = []
        if (record.note.agreement.trim()) {
          rows.push({
            date: "",
            values: [record.customerNumber, record.name, "ملخص الاتفاق/الملاحظة", record.note.agreement, "—", "—", "—", "—"],
          })
        }
        for (const [index, action] of record.note.actions.entries()) {
          rows.push({
            date: action.dueDate || action.reminderDate,
            values: [
              record.customerNumber, record.name, `بند متابعة ${index + 1}`, action.description,
              action.status === "completed" || action.completed ? "تم" : action.status === "failed" ? "لم ينفذ" : "قيد المتابعة",
              action.dueDate || "—", action.reminderDate || "—", action.failureReason ?? "—",
            ],
          })
        }
        return rows
      })
      amountColumn = -1
    } else if (reportType === "blocked_balances") {
      headers = ["رقم الزبون", "اسم الزبون", "الجوال", "الرصيد الحالي", "اتجاه الرصيد", "سبب الحظر", "تاريخ الحظر", "المبلغ"]
      sourceRows = blockedBalanceRows.map(({ customer, record }) => ({
        date: customer.blockedAt.slice(0, 10),
        amount: record?.balance,
        values: [
          customer.customerNumber, record?.name ?? customer.name, (record?.phone ?? customer.phone) || "—",
          record?.balance ?? "غير متوفر",
          record
            ? record.balanceDirection === "receivable" ? "مستحق لنا" : record.balanceDirection === "payable" ? "مستحق للزبون" : "غير محدد"
            : "لا يوجد كشف",
          customer.reason || "—", customer.blockedAt.slice(0, 10), record?.balance ?? "غير متوفر",
        ],
      }))
      amountColumn = 7
    } else {
      const totals = new Map<string, { record: CustomerBalance; amount: number }>()
      for (const payment of reportRecords.flatMap(customerPayments)) {
        const date = payment.date.slice(0, 10)
        if ((paymentDateFrom && date < paymentDateFrom) || (paymentDateTo && date > paymentDateTo)) continue
        const record = reportRecords.find(item => item.customerNumber === payment.customerNumber)
        if (!record) continue
        const current = totals.get(payment.customerNumber)
        totals.set(payment.customerNumber, { record, amount: (current?.amount ?? 0) + payment.amount })
      }
      if (reportType === "paid_total") {
        headers = ["رقم الزبون", "اسم الزبون", "إجمالي ما سدده", "المبلغ"]
        sourceRows = [...totals.values()].map(({ record, amount }) => ({
          date: paymentDateFrom || paymentDateTo || "",
          amount,
          values: [record.customerNumber, record.name, amount, amount],
        }))
        amountColumn = 3
      } else {
        headers = ["رقم الزبون", "اسم الزبون", "تاريخ الدفعة", "المبلغ", "نوع الدفعة", "الاتفاق"]
        sourceRows = reportRecords.flatMap(customerPayments).map(payment => ({
          date: payment.date.slice(0, 10),
          amount: payment.amount,
          values: [
            payment.customerNumber, payment.customerName, payment.date,
            payment.amount, payment.type, payment.agreement,
          ],
        }))
        amountColumn = 3
      }
    }

    const rows = sourceRows.filter(row => {
      if (reportType === "payment_classification" || reportType === "missing_phone") return true
      if (!paymentDateRangeValid) return false
      const rowDate = row.date.slice(0, 10)
      if ((paymentDateFrom || paymentDateTo) &&
        (!rowDate || (paymentDateFrom && rowDate < paymentDateFrom) || (paymentDateTo && rowDate > paymentDateTo))) return false
      if (reportType === "notes") return true
      if (reportAmountFrom && (row.amount === undefined || row.amount < Number(reportAmountFrom))) return false
      if (reportAmountTo && (row.amount === undefined || row.amount > Number(reportAmountTo))) return false
      return true
    })
    const receivableTotal = rows.reduce(
      (sum, row) => sum + (row.values[4] === "مستحق لنا" ? row.amount ?? 0 : 0),
      0,
    )
    const payableTotal = rows.reduce(
      (sum, row) => sum + (row.values[4] === "مستحق للزبون" ? row.amount ?? 0 : 0),
      0,
    )
    return {
      title: reportTypeLabels[reportType],
      headers,
      rows: rows.map(row => row.values),
      amountColumn,
      total: reportType === "blocked_balances"
        ? receivableTotal - payableTotal
        : rows.reduce((sum, row) => sum + (row.amount ?? 0), 0),
      totalLabel: reportType === "blocked_balances" ? "الصافي (م − د)" : "الإجمالي",
      extraTotals: reportType === "blocked_balances"
        ? [
          { label: "إجمالي مستحق لنا (م)", amount: receivableTotal },
          { label: "إجمالي مستحق للزبائن (د)", amount: payableTotal },
        ]
        : [],
    }
  }

  const exportSelectedReportExcel = () => {
    if (reportType === "payment_classification") {
      const records = dataset?.records ?? []
      const reportRows = paymentClassificationReportRows(records, effectivePaymentCriteria, now)
      const workbook = XLSXStyle.utils.book_new()
      for (const classification of PAYMENT_CLASSIFICATIONS) {
        const matchingRows = reportRows.filter(row =>
          (row.record.paymentClassification ?? null) === classification.value,
        )
        const worksheet = XLSXStyle.utils.aoa_to_sheet([
          [`الزبائن حسب تصنيفك اليدوي: ${classification.label}`],
          ["رقم الزبون", "اسم الزبون", "الجوال", "تصنيف النظام", "تصنيفك اليدوي", "الرصيد المستحق", "آخر سند قبض", "نسبة السداد", "حد الائتمان المقترح", "آخر فاتورة", "التنبيهات"],
          ...matchingRows.map(({ record, assessment, manualClassification, note }) => [
            record.customerNumber,
            record.name,
            record.phone || "—",
            automaticClassificationLabel(assessment.classification),
            manualClassification,
            `${formatAmount(record.balance)}${record.balanceDirection === "receivable" ? " م" : record.balanceDirection === "payable" ? " د" : ""}`,
            record.receiptDate ? `${formatAmount(record.receiptAmount)} · ${formatDate(record.receiptDate)}` : "—",
            assessment.paymentPercent === null ? "—" : `${assessment.paymentPercent.toFixed(1)}%`,
            assessment.creditLimit === null ? "—" : formatAmount(assessment.creditLimit),
            record.invoiceDate ? `${formatAmount(record.invoiceAmount)} · ${formatDate(record.invoiceDate)}` : "—",
            note,
          ]),
        ])
        for (let row = 3; row < matchingRows.length + 3; row++) {
          const systemCell = worksheet[`D${row}`]
          const manualCell = worksheet[`E${row}`]
          if (systemCell) systemCell.s = { fill: { fgColor: { rgb: classificationExcelColor(matchingRows[row - 3].assessment.classification) } } }
          if (manualCell) {
            manualCell.s = { fill: { fgColor: { rgb: classificationExcelColor(classification.value) } } }
          }
        }
        XLSXStyle.utils.book_append_sheet(workbook, worksheet, classification.label.slice(0, 31))
      }
      XLSXStyle.writeFile(workbook, "تقرير-تصنيف-الزبائن.xlsx")
      return
    }
    const report = createSelectedReport()
    const totalRows: Array<Array<string | number>> = report.amountColumn >= 0
      ? [
        ...report.extraTotals.map(total => {
          const row: Array<string | number> = report.headers.map(() => "")
          row[0] = total.label
          row[report.amountColumn] = total.amount
          return row
        }),
        (() => {
          const row: Array<string | number> = report.headers.map(() => "")
          row[0] = report.totalLabel
          row[report.amountColumn] = report.total
          return row
        })(),
      ]
      : []
    const worksheet = XLSX.utils.aoa_to_sheet([
      [report.title],
      [`الفترة: ${paymentPeriodLabel}`, `المبلغ: ${reportAmountFrom || "دون حد أدنى"} إلى ${reportAmountTo || "دون حد أعلى"}`],
      [],
      report.headers,
      ...report.rows,
      ...totalRows,
    ])
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, report.title.slice(0, 31))
    XLSX.writeFile(workbook, `${report.title.replace(/\s+/g, "-")}.xlsx`)
  }

  const exportSelectedReportPdf = () => {
    if (reportType === "payment_classification") {
      const printWindow = window.open("", "_blank", "width=1200,height=850")
      if (!printWindow) {
        window.alert("تعذر فتح تقرير PDF. اسمح بالنوافذ المنبثقة ثم أعد المحاولة.")
        return
      }
      const escapeHtml = (value: unknown) => String(value ?? "").replace(/[&<>"']/g, character => ({
        "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
      })[character] ?? character)
      const records = dataset?.records ?? []
      const reportRows = paymentClassificationReportRows(records, effectivePaymentCriteria, now)
      const sections = PAYMENT_CLASSIFICATIONS.map(classification => {
        const matchingRows = reportRows.filter(row =>
          (row.record.paymentClassification ?? null) === classification.value,
        )
        const rows = matchingRows.map(({ record, assessment, manualClassification, note }) => `<tr>
          <td>${escapeHtml(record.customerNumber)}</td>
          <td>${escapeHtml(record.name)}</td>
          <td dir="ltr">${escapeHtml(record.phone || "—")}</td>
          <td style="background:#${classificationExcelColor(assessment.classification)}">${escapeHtml(automaticClassificationLabel(assessment.classification))}</td>
          <td style="background:#${classificationExcelColor(record.paymentClassification ?? null)}">${escapeHtml(manualClassification)}</td>
          <td>${escapeHtml(`${formatAmount(record.balance)}${record.balanceDirection === "receivable" ? " م" : record.balanceDirection === "payable" ? " د" : ""}`)}</td>
          <td>${escapeHtml(record.receiptDate ? `${formatAmount(record.receiptAmount)} · ${formatDate(record.receiptDate)}` : "—")}</td>
          <td>${escapeHtml(assessment.paymentPercent === null ? "—" : `${assessment.paymentPercent.toFixed(1)}%`)}</td>
          <td>${escapeHtml(assessment.creditLimit === null ? "—" : formatAmount(assessment.creditLimit))}</td>
          <td>${escapeHtml(record.invoiceDate ? `${formatAmount(record.invoiceAmount)} · ${formatDate(record.invoiceDate)}` : "—")}</td>
          <td>${escapeHtml(note)}</td>
        </tr>`).join("")
        return `<section><h2>تصنيفك اليدوي: ${escapeHtml(classification.label)} <span>(${matchingRows.length.toLocaleString("ar")})</span></h2>
          <table><thead><tr><th>رقم الزبون</th><th>اسم الزبون</th><th>الجوال</th><th>تصنيف النظام</th><th>تصنيفك اليدوي</th><th>الرصيد المستحق</th><th>آخر سند قبض</th><th>نسبة السداد</th><th>حد الائتمان</th><th>آخر فاتورة</th><th>التنبيهات</th></tr></thead>
          <tbody>${rows || `<tr><td colspan="11" class="empty">لا يوجد زبائن في هذا التصنيف.</td></tr>`}</tbody></table></section>`
      }).join("")
      printWindow.document.open()
      printWindow.document.write(`<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>تقرير تصنيف الزبائن</title><style>${REPORT_LOGO_PRINT_STYLES}
        @page{size:A4 landscape;margin:8mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;color:#172033;font-size:8px;line-height:1.4}h1{font-size:20px;margin:0 0 8px}h2{font-size:14px;margin:0 0 7px;border-bottom:1px solid #cbd5e1;padding-bottom:4px}h2 span,.empty{color:#64748b}section{break-before:page}section:first-of-type{break-before:auto}table{width:100%;border-collapse:collapse;table-layout:fixed}th,td{border:1px solid #cbd5e1;padding:4px;text-align:right;vertical-align:top;overflow-wrap:anywhere}th{background:#eef2f6}@media print{body{print-color-adjust:exact;-webkit-print-color-adjust:exact}}
        </style></head><body>${reportLogoMarkup()}<h1>تقرير تصنيف الزبائن حسب الدفع</h1><p>تاريخ التقرير: ${escapeHtml(new Date().toLocaleString("ar"))} · إجمالي الزبائن: ${records.length.toLocaleString("ar")} · الجداول حسب تصنيفك اليدوي، ويظهر تصنيف النظام الآلي بجانبه بلونه · حداثة السند تخفض التصنيف درجة · الفاتورة تستخدم للتنبيه ولا يثبت الكشف سدادها.</p>${sections}<script>window.addEventListener('load',()=>{window.focus();window.print()},{once:true})</script></body></html>`)
      printWindow.document.close()
      return
    }
    const report = createSelectedReport()
    const printWindow = window.open("", "_blank", "width=1200,height=850")
    if (!printWindow) {
      window.alert("تعذر فتح تقرير PDF. اسمح بالنوافذ المنبثقة ثم أعد المحاولة.")
      return
    }
    const escapeHtml = (value: unknown) => String(value ?? "").replace(/[&<>"']/g, character => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
    })[character] ?? character)
    const rows = report.rows.map(row => `<tr>${row.map(cell => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("")
    const totalRows = report.amountColumn >= 0
      ? [
        ...report.extraTotals,
        { label: report.totalLabel, amount: report.total },
      ].map(total => `<tr class="total">${report.headers.map((_, index) => `<td>${index === 0 ? escapeHtml(total.label) : index === report.amountColumn ? escapeHtml(formatAmount(total.amount)) : ""}</td>`).join("")}</tr>`).join("")
      : ""
    printWindow.document.open()
    printWindow.document.write(`<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>${escapeHtml(report.title)}</title><style>${REPORT_LOGO_PRINT_STYLES}
      @page{size:A4 landscape;margin:12mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;color:#172033;font-size:10px;line-height:1.5}h1{font-size:20px;margin:0 0 8px}.meta{color:#4b5563;margin-bottom:14px}table{width:100%;border-collapse:collapse}th,td{border:1px solid #cbd5e1;padding:6px;text-align:right;vertical-align:top;overflow-wrap:anywhere}th{background:#eef2f6}.total{font-weight:bold;background:#ecfdf5}@media print{body{print-color-adjust:exact;-webkit-print-color-adjust:exact}}
      </style></head><body>${reportLogoMarkup()}<h1>${escapeHtml(report.title)}</h1><div class="meta">الفترة: ${escapeHtml(paymentPeriodLabel)} · المبلغ: ${escapeHtml(reportAmountFrom || "دون حد أدنى")} إلى ${escapeHtml(reportAmountTo || "دون حد أعلى")} · عدد السجلات: ${report.rows.length.toLocaleString("ar")}</div><table><thead><tr>${report.headers.map(header => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${rows || `<tr><td colspan="${report.headers.length}">لا توجد بيانات مطابقة للفلاتر.</td></tr>`}${totalRows}</tbody></table><script>window.addEventListener('load',()=>{window.focus();window.print()},{once:true})</script></body></html>`)
    printWindow.document.close()
  }

  const selectedReportPreview = reportOpen ? createSelectedReport() : null

  const updateFilter = <K extends keyof Filters>(key: K, value: Filters[K]) => {
    setFilters(current => ({ ...current, [key]: value }))
    setNotice("")
  }

  const resetFilters = () => {
    setFilters(defaultFilters)
    setNotice("")
  }

  const openCustomer = dataset?.records.find(record => record.customerNumber === openCustomerNumber)
  const shortcutCustomer = dataset?.records.find(record => record.customerNumber === customerShortcutNumber)
  const blockedContact = blockedCustomers.find(customer => customer.customerNumber === blockedContactNumber)
  const blockedContactRecord = dataset?.records.find(record => record.customerNumber === blockedContactNumber)
  const selectedRecords = (dataset?.records ?? []).filter(record => selectedCustomerNumbers.includes(record.customerNumber))
  const visibleCustomerNumbers = filteredRecords.slice(0, 1000).map(record => record.customerNumber)
  const allVisibleSelected = visibleCustomerNumbers.length > 0 && visibleCustomerNumbers.every(number => selectedCustomerNumbers.includes(number))

  if (!visible) return null

  return (
    <main dir="rtl" className="min-h-full bg-[#f4f7fb] px-4 py-7 text-slate-900 sm:px-7 lg:px-10">
      <div className="mx-auto max-w-7xl space-y-6">
        <header className="rounded-3xl bg-[#10243c] px-6 py-8 text-white shadow-xl sm:px-9">
          <div className="mb-3 inline-flex items-center gap-2 rounded-full border border-white/15 bg-white/5 px-3 py-1.5 text-xs text-cyan-100"><ShieldCheck className="h-3.5 w-3.5" />معالجة آمنة لبيانات الزبائن</div>
          <h1 className="text-2xl font-extrabold sm:text-3xl">أرصدة الزبائن</h1>
          <p className="mt-3 max-w-3xl text-sm leading-7 text-slate-300">ارفع كشف Excel، وسيتم جمع بيانات الزبائن تلقائياً من أوراق الملف وصفحاته حتى مع تكرار الترويسة، ثم ابحث وفلتر النتائج واحفظها في Excel.</p>
          <div className="mt-4 flex items-center gap-2 text-xs text-emerald-200"><ShieldCheck className="h-4 w-4" />{centralStorageEnabled ? "تُحفظ البيانات في مساحة حسابك" : "تُعالج البيانات وتُحفظ على هذا الجهاز"}</div>
        </header>

        {error && <div role="alert" className="flex items-start gap-2 rounded-xl border border-rose-200 bg-rose-50 p-4 text-sm text-rose-800"><AlertTriangle className="mt-0.5 h-4 w-4 shrink-0" /><span>{error}</span></div>}
        {notice && <div role="status" className="rounded-xl border border-emerald-200 bg-emerald-50 p-3 text-sm text-emerald-800">{notice}</div>}

        <section className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm sm:p-7">
          <div className="flex flex-wrap items-start justify-between gap-4">
            <div>
              <h2 className="text-base font-bold">ملف أرصدة الزبائن</h2>
              <p className="mt-1 text-xs leading-6 text-slate-500">يدعم XLSX وXLS وXLSM وCSV. يحدّث الكشف الجديد أرصدة الزبائن برقم الزبون ويضيف الجدد مع الحفاظ على الملاحظات، حتى إن لم يظهر بعض الزبائن في ملف التحديث.</p>
            </div>
            <div className="flex flex-wrap gap-2">
              <input ref={inputRef} type="file" accept=".xlsx,.xls,.xlsm,.csv" className="hidden" onChange={event => void importFile(event)} />
              <input ref={statementFileInputRef} type="file" accept=".xlsx,.xls,.xlsm" className="hidden" onChange={chooseStatementFile} />
              <button type="button" disabled={busy || !ready || loadedStorageKey !== storageKey || !missingCustomerPolicyReady} onClick={() => inputRef.current?.click()} className="inline-flex items-center gap-2 rounded-xl bg-[#10243c] px-4 py-2.5 text-sm font-semibold text-white hover:bg-[#193653] disabled:opacity-50">
                {busy ? <FileSpreadsheet className="h-4 w-4 animate-pulse" /> : <Upload className="h-4 w-4" />}{busy ? "جارٍ تحليل الملف..." : dataset ? "تحديث من ملف Excel" : "رفع ملف Excel"}
              </button>
              {dataset && <button type="button" onClick={clearDataset} className="inline-flex items-center gap-2 rounded-xl border border-rose-200 px-3 py-2.5 text-sm font-semibold text-rose-700 hover:bg-rose-50"><Trash2 className="h-4 w-4" />حذف البيانات</button>}
            </div>
          </div>
          {!centralStorageEnabled && <p className="mt-4 rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-xs leading-6 text-amber-900">البيانات معزولة حسب اسم المستخدم على هذا المتصفح فقط. لعرض كشف المستخدم نفسه من متصفح أو جهاز آخر، يجب تفعيل التخزين المركزي وربط المنصة بقاعدة البيانات.</p>}
          {dataset && <p className="mt-4 rounded-xl bg-slate-50 px-4 py-3 text-xs text-slate-600">الملف: <span className="font-semibold text-slate-800">{dataset.fileName}</span> · {dataset.records.length.toLocaleString("ar")} زبون محفوظ في السجل · آخر استيراد: {new Date(dataset.importedAt).toLocaleString("ar")}</p>}
          <div className="mt-4 rounded-2xl border border-emerald-200 bg-emerald-50/50 p-4">
            <div className="flex flex-wrap items-center justify-between gap-3"><div><h3 className="text-sm font-bold text-emerald-950">قوالب الرسائل على واتساب</h3><p className="mt-1 text-xs text-slate-600">رسالة كشف الحساب وتذكير القسط الأسبوعي والقوالب التي تضيفها متاحة لكل الزبائن، ويُقترح تذكير القسط تلقائياً عند وجود اتفاق أسبوعي.</p></div><button type="button" disabled={loadedMessageTemplateKey !== storageKey} onClick={() => setMessageEditorOpen(value => !value)} className="rounded-lg border border-emerald-300 bg-white px-3 py-2 text-xs font-semibold text-emerald-900">{messageEditorOpen ? "إخفاء الرسائل" : "تعديل الرسائل"}</button></div>
            {messageEditorOpen && <div className="mt-3 space-y-2">
              <label className="block space-y-1 text-xs font-semibold text-slate-700"><span>القالب المراد تعديله</span><select value={messageTemplateEditorType} onChange={event => setMessageTemplateEditorType(event.target.value)} className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm font-normal"><option value="statement">رسالة كشف الحساب</option><option value="weekly_installment">تذكير القسط الأسبوعي</option>{customMessageTemplates.map(template => <option key={template.id} value={`custom:${template.id}`}>{template.name}</option>)}</select></label>
              {messageTemplateEditorType.startsWith("custom:") && <label className="block space-y-1 text-xs font-semibold text-slate-700"><span>اسم الرسالة</span><input value={customMessageTemplates.find(template => messageTemplateEditorType === `custom:${template.id}`)?.name ?? ""} onChange={event => {
                const id = messageTemplateEditorType.slice("custom:".length)
                setCustomMessageTemplates(current => current.map(template => template.id === id ? { ...template, name: event.target.value } : template))
                setHasMessageTemplateChanges(true)
              }} className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm font-normal" /></label>}
              <label className="block space-y-1 text-xs font-semibold text-slate-700"><span>نص الرسالة</span><textarea ref={messageTemplateTextAreaRef} rows={9} value={messageTemplateEditorType === "statement"
                ? statementMessageTemplate
                : messageTemplateEditorType === "weekly_installment"
                  ? weeklyInstallmentMessageTemplate
                  : customMessageTemplates.find(template => messageTemplateEditorType === `custom:${template.id}`)?.content ?? ""}
                onChange={event => updateMessageTemplateContent(messageTemplateEditorType, event.target.value)}
                onDragOver={event => event.preventDefault()}
                onDrop={event => {
                  event.preventDefault()
                  const token = event.dataTransfer.getData("text/plain")
                  if (MESSAGE_FIELDS.some(field => field.token === token)) insertMessageField(token)
                }}
                className="w-full resize-y rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm font-normal" /></label>
              <div className="space-y-2">
                <p className="text-[11px] font-semibold text-slate-600">اسحب الحقل إلى موضع المؤشر في الرسالة أو اضغط لإضافته:</p>
                <div className="flex flex-wrap gap-2">{MESSAGE_FIELDS.map(field => <button key={field.token} type="button" draggable onDragStart={event => event.dataTransfer.setData("text/plain", field.token)} onClick={() => insertMessageField(field.token)} className="cursor-grab rounded-lg border border-cyan-200 bg-white px-3 py-2 text-xs font-semibold text-cyan-900 active:cursor-grabbing">{field.label} <code dir="ltr">{field.token}</code></button>)}</div>
              </div>
              <p className="text-[11px] leading-5 text-slate-600">حقول أخرى: <code dir="ltr">{"{{customerNumber}}"}</code> رقم الزبون، <code dir="ltr">{"{{duration}}"}</code> المدة، <code dir="ltr">{"{{phone}}"}</code> الجوال، <code dir="ltr">{"{{invoiceAmount}}"}</code> و<code dir="ltr">{"{{invoiceDate}}"}</code> آخر فاتورة، <code dir="ltr">{"{{receiptAmount}}"}</code> و<code dir="ltr">{"{{receiptDate}}"}</code> آخر سند قبض، و<code dir="ltr">{"{{weeklyAmount}}"}</code> مبلغ القسط الأسبوعي. تُحفظ القوالب لحسابك.</p>
              <div className="flex flex-wrap items-center gap-2">
                <button type="button" onClick={() => setShowNewMessageTemplate(value => !value)} className="inline-flex items-center gap-1 rounded-lg border border-emerald-300 bg-white px-3 py-2 text-xs font-semibold text-emerald-900"><Plus className="h-4 w-4" />إضافة رسالة</button>
                {showNewMessageTemplate && <div className="flex flex-1 flex-wrap gap-2"><input value={newMessageTemplateName} onChange={event => setNewMessageTemplateName(event.target.value)} onKeyDown={event => { if (event.key === "Enter") addCustomMessageTemplate() }} placeholder="اسم الرسالة الجديدة" className="min-w-48 flex-1 rounded-lg border border-slate-200 px-3 py-2 text-xs font-normal" /><button type="button" onClick={addCustomMessageTemplate} className="rounded-lg bg-emerald-700 px-3 py-2 text-xs font-semibold text-white">إنشاء</button></div>}
              </div>
              {!messageTemplateEditorType.startsWith("custom:") && <button type="button" onClick={() => {
                if (messageTemplateEditorType === "statement") setStatementMessageTemplate(DEFAULT_STATEMENT_MESSAGE)
                else setWeeklyInstallmentMessageTemplate(DEFAULT_WEEKLY_INSTALLMENT_MESSAGE)
                setHasMessageTemplateChanges(true)
              }} className="rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs font-semibold text-slate-600">استعادة الرسالة الافتراضية</button>}
            </div>}
          </div>
        </section>

        {dataset && <section className="space-y-4 rounded-3xl border border-indigo-200 bg-white p-5 shadow-sm sm:p-7">
          <div className="flex flex-wrap items-start justify-between gap-3">
            <div>
              <h2 className="text-base font-bold text-indigo-950">معايير التقييم الآلي وحدود الائتمان</h2>
              <p className="mt-1 max-w-3xl text-xs leading-5 text-slate-600">يقارن النظام آخر سند قبض بالرصيد المستحق لنا، ويخفض التصنيف درجة إذا كان السند أحدث من عدد الأيام المحدد. النسب قابلة للتعديل وتُضرب بالرصيد لتناسب المبالغ الأكبر. التصنيف الآلي لا يغيّر تصنيفك اليدوي.</p>
            </div>
            <span className={`rounded-full px-3 py-1.5 text-xs font-semibold ${hasPaymentCriteriaChanges && paymentCriteriaValid ? "bg-amber-100 text-amber-800" : paymentCriteriaValid ? "bg-emerald-100 text-emerald-800" : "bg-rose-100 text-rose-800"}`}>{!paymentCriteriaValid ? "راجع القيم" : hasPaymentCriteriaChanges ? "جارٍ حفظ المعايير" : "المعايير محفوظة"}</span>
          </div>
          <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
            {([
              ["acceptableMinPercent", "أقل نسبة للمقبول (%)", 1, 1000],
              ["goodMinPercent", "أقل نسبة للجيد (%)", 1, 1000],
              ["excellentMinPercent", "أقل نسبة للممتاز (%)", 1, 1000],
              ["acceptableCreditLimitPercent", "حد المقبول من الرصيد (%)", 0, 1000],
              ["goodCreditLimitPercent", "حد الجيد من الرصيد (%)", 0, 1000],
              ["excellentCreditLimitPercent", "حد الممتاز من الرصيد (%)", 0, 1000],
              ["recentReceiptDays", "أيام تخفيض السند الحديث", 1, 3650],
            ] as Array<[keyof PaymentCriteria, string, number, number]>).map(([key, label, min, max]) => <label key={key} className="space-y-1.5 text-xs font-semibold text-slate-700">
              <span>{label}</span>
              <input type="number" min={min} max={max} step="1" value={paymentCriteria[key]} disabled={loadedPaymentCriteriaKey !== paymentCriteriaStorageKey} onChange={event => updatePaymentCriterion(key, Number(event.target.value))} className="w-full rounded-xl border border-slate-200 px-3 py-2.5 text-sm font-normal disabled:bg-slate-100" />
            </label>)}
          </div>
          {paymentCriteriaError && <p role="alert" className="rounded-xl border border-rose-200 bg-rose-50 p-3 text-xs text-rose-800">{paymentCriteriaError}</p>}
          {!paymentCriteriaValid && <p role="alert" className="rounded-xl border border-amber-200 bg-amber-50 p-3 text-xs text-amber-800">لا تُحفظ القيم غير الصالحة، ويستمر العرض مؤقتاً على المعايير الافتراضية إلى أن تصلحها.</p>}
          <div className="flex flex-wrap gap-3 rounded-xl bg-indigo-50 p-3 text-xs leading-5 text-indigo-950">
            <span><span className="font-bold text-rose-700">سيئ:</span> أقل من {paymentCriteria.acceptableMinPercent}%؛ لا تمنح ديناً جديداً، واطلب السداد وراجع إغلاق الحساب.</span>
            <span><span className="font-bold text-amber-700">مقبول:</span> من {paymentCriteria.acceptableMinPercent}% إلى أقل من {paymentCriteria.goodMinPercent}%.</span>
            <span><span className="font-bold text-emerald-700">جيد:</span> من {paymentCriteria.goodMinPercent}% إلى أقل من {paymentCriteria.excellentMinPercent}%.</span>
            <span><span className="font-bold text-sky-700">ممتاز:</span> {paymentCriteria.excellentMinPercent}% فأكثر.</span>
            <span>مثال عند مستحق 100: حدود الائتمان المقترحة {paymentCriteria.acceptableCreditLimitPercent} للمقبول، {paymentCriteria.goodCreditLimitPercent} للجيد، و{paymentCriteria.excellentCreditLimitPercent} للممتاز.</span>
            <span>آخر فاتورة أعلى من الحد تُظهر تنبيهاً فقط؛ الكشف لا يثبت سداد الفاتورة ولا يغيّر التصنيف تلقائياً.</span>
          </div>
        </section>}

        {loadedBlocklistKey === blocklistStorageKey && <section className="space-y-4 rounded-3xl border border-rose-200 bg-rose-50/50 p-5 shadow-sm sm:p-7">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div><h2 className="text-base font-bold text-rose-900">المحظورون من التواصل</h2><p className="mt-1 text-xs leading-5 text-rose-800">تبقى هذه القائمة محفوظة بشكل مستقل عن كشوف الأرصدة، فلا يحذفها تحديث الملف أو حتى حذف الكشف.</p></div>
            <span className="rounded-full bg-rose-100 px-3 py-1.5 text-xs font-bold text-rose-800">{blockedCustomers.length.toLocaleString("ar")} زبون</span>
          </div>
          {blockedCustomers.length > 0 && <p className="text-xs font-semibold text-rose-900">إجمالي مستحق لنا: {formatAmount(blockedReceivableTotal)} · مستحق للزبائن: {formatAmount(blockedPayableTotal)} · الصافي: {formatAmount(blockedNetTotal)}{blockedMissingBalanceCount > 0 ? ` · ${blockedMissingBalanceCount.toLocaleString("ar")} بلا رصيد في الكشف الحالي` : ""}</p>}
          {!blockedCustomers.length && <p className="rounded-xl bg-white p-4 text-sm text-slate-500">لا يوجد زبائن في قائمة الحظر.</p>}
          {blockedCustomers.length > 0 && <div className="space-y-2">{visibleBlockedRows.map(({ customer, record }) => <div key={customer.customerNumber} className="grid gap-2 rounded-xl border border-rose-100 bg-white p-3 sm:grid-cols-[minmax(180px,1fr)_minmax(160px,1.5fr)_auto] sm:items-center">
            <div className="min-w-0"><p className="truncate text-sm font-semibold">{record?.name ?? customer.name}</p><p className="mt-0.5 text-[11px] text-slate-500">رقم الزبون: {customer.customerNumber}{(record?.phone ?? customer.phone) ? ` · الجوال: ${record?.phone ?? customer.phone}` : ""}</p><p className="mt-1 text-[11px] font-semibold text-rose-800">المبلغ المستحق: {record ? `${formatAmount(record.balance)}${record.balanceDirection === "receivable" ? " م" : record.balanceDirection === "payable" ? " د" : ""}` : "غير متوفر"}</p><p className="mt-0.5 text-[10px] text-slate-500">الرصيد: {record ? formatAmount(record.balance) : "—"} · المدة: {record ? `${record.duration.toLocaleString("ar")} يوم` : "—"}</p></div>
            <label className="space-y-1 text-[10px] font-semibold text-slate-600"><span>سبب الحظر (اختياري)</span><input value={customer.reason} onChange={event => updateBlockedReason(customer.customerNumber, event.target.value)} placeholder="اكتب ملاحظة إن وجدت..." className="w-full rounded-lg border border-slate-200 px-3 py-2 text-xs font-normal" /></label>
            <button type="button" onClick={() => removeBlockedCustomer(customer.customerNumber)} className="rounded-lg border border-rose-200 px-3 py-2 text-xs font-semibold text-rose-700 hover:bg-rose-50">رفع الحظر</button>
          </div>)}
          {blockedPageCount > 1 && <div className="flex items-center justify-center gap-3 pt-2"><button type="button" disabled={blockedPage <= 1} onClick={() => setBlockedPage(page => Math.max(1, page - 1))} className="rounded-lg border border-rose-200 bg-white px-3 py-1.5 text-xs font-semibold text-rose-800 disabled:opacity-40">السابق</button><span className="text-xs text-slate-600">صفحة {blockedPage.toLocaleString("ar")} من {blockedPageCount.toLocaleString("ar")} · عرض {visibleBlockedRows.length.toLocaleString("ar")} من {blockedCustomers.length.toLocaleString("ar")}</span><button type="button" disabled={blockedPage >= blockedPageCount} onClick={() => setBlockedPage(page => Math.min(blockedPageCount, page + 1))} className="rounded-lg border border-rose-200 bg-white px-3 py-1.5 text-xs font-semibold text-rose-800 disabled:opacity-40">التالي</button></div>}
          </div>}
        </section>}

        {dataset && <>
          <section className="overflow-hidden rounded-3xl border border-cyan-200 bg-white shadow-sm">
            <div className="flex flex-wrap items-center justify-between gap-3 p-5">
              <div>
                <h2 className="text-base font-bold text-cyan-950">سجل الزبائن</h2>
                <p className="mt-1 text-xs leading-5 text-slate-600">يحفظ كل زبون ظهر في كشف واحد على الأقل، مع مرات الظهور المسجلة وحالة ظهوره في آخر كشف ووجود ملاحظات أو اتفاقات. يبدأ العد من تفعيل السجل؛ للزبائن المحفوظين سابقاً يظهر الحد الأدنى المعروف.</p>
                <p className="mt-2 inline-flex flex-wrap items-center gap-2 text-[11px] text-slate-600"><span className="inline-block h-2.5 w-2.5 rounded-full bg-slate-400" />النقطة الأولى تصنيفك اليدوي <span className="inline-block h-2.5 w-2.5 rounded-full bg-cyan-500" />الثانية تصنيف النظام الآلي. السيئ: لا تمنح ديناً جديداً واطلب السداد؛ علامة التحذير توضح تجاوز الفاتورة للحد المقترح.</p>
              </div>
              <button type="button" aria-expanded={showCustomerHistory} onClick={() => setShowCustomerHistory(value => !value)} className="rounded-lg border border-cyan-200 px-3 py-2 text-xs font-semibold text-cyan-900 hover:bg-cyan-50">
                {showCustomerHistory ? "إخفاء السجل" : `عرض السجل (${dataset.records.length.toLocaleString("ar")})`}
              </button>
            </div>
            {showCustomerHistory && <div className="border-t border-cyan-100">
              <div className="max-h-[60vh] overflow-auto">
                <table className="w-full min-w-[850px] border-collapse text-right text-xs">
                  <thead className="sticky top-0 z-10 bg-cyan-50 text-cyan-950"><tr>
                    <th className="border-b border-cyan-100 px-3 py-3">رقم الزبون</th>
                    <th className="border-b border-cyan-100 px-3 py-3">الاسم</th>
                    <th className="border-b border-cyan-100 px-3 py-3">الرصيد المحفوظ</th>
                    <th className="border-b border-cyan-100 px-3 py-3">آخر رصيد بكشف فعلي</th>
                    <th className="border-b border-cyan-100 px-3 py-3">الظهور في آخر كشف</th>
                    <th className="border-b border-cyan-100 px-3 py-3">مرات الظهور المسجلة</th>
                    <th className="border-b border-cyan-100 px-3 py-3">آخر ظهور</th>
                    <th className="border-b border-cyan-100 px-3 py-3">الملاحظات والاتفاقات</th>
                  </tr></thead>
                  <tbody>
                    {visibleCustomerHistory.map(record => <tr key={record.customerNumber} className="hover:bg-cyan-50/40">
                      <td className="border-b border-slate-100 px-3 py-2.5">{record.customerNumber}</td>
                      <td className="border-b border-slate-100 px-3 py-2.5 font-semibold"><span aria-label={`تصنيف الدفع: ${paymentClassificationLabel(record.paymentClassification)}`} title={paymentClassificationLabel(record.paymentClassification)} className={`ml-2 inline-block h-2.5 w-2.5 rounded-full ${PAYMENT_CLASSIFICATIONS.find(item => item.value === (record.paymentClassification ?? null))?.color ?? "bg-slate-300"}`} />{record.name}</td>
                      <td className="border-b border-slate-100 px-3 py-2.5 font-semibold tabular-nums">{formatAmount(record.balance)}{record.balanceDirection === "receivable" ? " م" : record.balanceDirection === "payable" ? " د" : ""}</td>
                      <td className="border-b border-slate-100 px-3 py-2.5 tabular-nums">{formatAmount(record.lastStatementBalance ?? record.balance)}{(record.lastStatementDirection ?? record.balanceDirection) === "receivable" ? " م" : (record.lastStatementDirection ?? record.balanceDirection) === "payable" ? " د" : ""}</td>
                      <td className="border-b border-slate-100 px-3 py-2.5">{record.isInLatestStatement === false ? "غير ظاهر" : "ظاهر"}</td>
                      <td className="border-b border-slate-100 px-3 py-2.5">{(record.statementAppearances ?? 1).toLocaleString("ar")}</td>
                      <td className="border-b border-slate-100 px-3 py-2.5">{record.lastSeenInStatementAt ? new Date(record.lastSeenInStatementAt).toLocaleDateString("ar") : "—"}</td>
                      <td className="border-b border-slate-100 px-3 py-2.5">{hasCustomerNotes(record) ? "لديه سجل" : "لا يوجد سجل"}</td>
                    </tr>)}
                    {!visibleCustomerHistory.length && <tr><td colSpan={8} className="px-4 py-8 text-center text-slate-500">لا يوجد زبائن في السجل.</td></tr>}
                  </tbody>
                </table>
              </div>
              {customerHistoryPageCount > 1 && <div className="flex items-center justify-center gap-3 border-t border-cyan-100 p-3">
                <button type="button" disabled={customerHistoryPage <= 1} onClick={() => setCustomerHistoryPage(page => Math.max(1, page - 1))} className="rounded-lg border border-cyan-200 px-3 py-1.5 text-xs font-semibold text-cyan-900 disabled:opacity-40">السابق</button>
                <span className="text-xs text-slate-600">صفحة {customerHistoryPage.toLocaleString("ar")} من {customerHistoryPageCount.toLocaleString("ar")} · عرض {visibleCustomerHistory.length.toLocaleString("ar")} من {customerHistoryRecords.length.toLocaleString("ar")}</span>
                <button type="button" disabled={customerHistoryPage >= customerHistoryPageCount} onClick={() => setCustomerHistoryPage(page => Math.min(customerHistoryPageCount, page + 1))} className="rounded-lg border border-cyan-200 px-3 py-1.5 text-xs font-semibold text-cyan-900 disabled:opacity-40">التالي</button>
              </div>}
            </div>}
          </section>

          <section className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5">
            <div className="rounded-2xl border border-slate-200 bg-white p-4"><div className="flex items-center gap-2 text-xs text-slate-500"><Users className="h-4 w-4 text-cyan-700" />عدد النتائج</div><p className="mt-2 text-2xl font-extrabold text-[#10243c]">{filteredRecords.length.toLocaleString("ar")}<span className="mr-2 text-xs font-normal text-slate-500">من {currentStatementRecords.length.toLocaleString("ar")}</span></p></div>
            <div className={`rounded-2xl border p-4 ${netBalance >= 0 ? "border-cyan-200 bg-cyan-50" : "border-amber-200 bg-amber-50"}`}><div className="flex items-center gap-2 text-xs text-slate-600"><FileSpreadsheet className="h-4 w-4" />صافي المجموع (م − د)</div><p className={`mt-2 text-2xl font-extrabold ${netBalance >= 0 ? "text-cyan-800" : "text-amber-800"}`}>{formatAmount(netBalance)}</p></div>
            <div className="rounded-2xl border border-rose-200 bg-rose-50 p-4"><div className="flex items-center gap-2 text-xs text-rose-700"><FileSpreadsheet className="h-4 w-4" />م · مستحق لنا</div><p className="mt-2 text-2xl font-extrabold text-rose-700">{formatAmount(receivableBalance)}</p></div>
            <div className="rounded-2xl border border-emerald-200 bg-emerald-50 p-4"><div className="flex items-center gap-2 text-xs text-emerald-700"><FileSpreadsheet className="h-4 w-4" />د · مستحق للزبون</div><p className="mt-2 text-2xl font-extrabold text-emerald-700">{formatAmount(payableBalance)}</p></div>
            <div className="flex flex-col justify-between rounded-2xl border border-slate-200 bg-white p-4"><div><div className="flex items-center gap-2 text-xs text-slate-500"><Download className="h-4 w-4 text-cyan-700" />تقارير الزبائن</div><p className="mt-2 text-xs leading-5 text-slate-500">اختر نوع التقرير والفترة والمبالغ من واجهة واحدة.</p></div><button type="button" disabled={!dataset.records.length && !blockedCustomers.length} onClick={() => setReportOpen(true)} className="mt-4 rounded-xl bg-cyan-700 px-4 py-3 text-sm font-bold text-white hover:bg-cyan-800 disabled:opacity-40">التقارير</button></div>
          </section>

          <section className="space-y-4 rounded-3xl border border-slate-200 bg-white p-5 shadow-sm sm:p-7">
            <div className="flex flex-wrap items-center justify-between gap-3"><div><h2 className="flex items-center gap-2 text-base font-bold"><Filter className="h-4 w-4 text-cyan-700" />البحث والتصفية</h2><p className="mt-1 text-xs text-slate-500">يمكن البحث داخل الاتفاقات، وتصفية المواعيد أو ترتيبها من الأقرب إلى الوقت الحالي.</p><p className="mt-1 text-[11px] text-cyan-800">وقت المنصة الآن: {now.toLocaleString("ar", { dateStyle: "medium", timeStyle: "short" })}</p></div><button type="button" onClick={resetFilters} className="rounded-lg border border-slate-200 px-3 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-50">مسح الفلاتر</button></div>
            <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
              <label className="space-y-1.5 text-xs font-semibold text-slate-700"><span>بحث عن زبون</span><select aria-label="مجال البحث" value={filters.searchField} onChange={event => updateFilter("searchField", event.target.value as Filters["searchField"])} className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-xs font-normal"><option value="all">الاسم أو رقم الزبون أو الهاتف أو الاتفاق</option><option value="customerNumber">رقم الزبون فقط</option><option value="phone">رقم الهاتف فقط</option></select><div className="relative"><Search className="absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" /><input type="search" value={filters.query} onChange={event => updateFilter("query", event.target.value)} placeholder={filters.searchField === "phone" ? "أدخل رقم الهاتف..." : filters.searchField === "customerNumber" ? "أدخل رقم الزبون..." : "اكتب الاسم أو الرقم أو كلمة من الاتفاق..."} className="w-full rounded-xl border border-slate-200 py-2.5 pr-9 pl-3 text-sm font-normal outline-none focus:border-cyan-500 focus:ring-2 focus:ring-cyan-100" /></div></label>
              <label className="space-y-1.5 text-xs font-semibold text-slate-700"><span>عرض الزبائن</span><select value={filters.notes} onChange={event => updateFilter("notes", event.target.value as Filters["notes"])} className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-xs font-normal outline-none focus:border-cyan-500"><option value="all">كل الزبائن</option><option value="withNotes">الذين لديهم ملاحظة أو اتفاق فقط</option></select></label>
              <label className="space-y-1.5 text-xs font-semibold text-slate-700"><span>حالة الاتفاقات</span><select value={filters.agreementStatus} onChange={event => updateFilter("agreementStatus", event.target.value as Filters["agreementStatus"])} className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-xs font-normal outline-none focus:border-cyan-500"><option value="all">كل الحالات</option><option value="active">اتفاقات قيد التنفيذ فقط</option></select></label>
              <label className="space-y-1.5 text-xs font-semibold text-slate-700"><span>الترتيب</span><select value={filters.noteOrder} onChange={event => updateFilter("noteOrder", event.target.value as Filters["noteOrder"])} className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-xs font-normal outline-none focus:border-cyan-500"><option value="default">الترتيب الحالي</option><option value="first">أصحاب الملاحظات أولاً</option><option value="dueSoon">مواعيد التنفيذ الأقرب (المتأخرة أولاً)</option></select></label>
              <div className="space-y-1.5 text-xs font-semibold text-slate-700"><span>{rangeLabel("موعد تنفيذ الاتفاق", filters.dueFrom, filters.dueTo)}</span><div className="flex items-center gap-2"><input aria-label="من موعد تنفيذ الاتفاق" type="date" value={filters.dueFrom} onChange={event => updateFilter("dueFrom", event.target.value)} className="min-w-0 w-full rounded-xl border border-slate-200 px-2 py-2.5 text-xs font-normal" /><span className="text-slate-400">إلى</span><input aria-label="إلى موعد تنفيذ الاتفاق" type="date" value={filters.dueTo} onChange={event => updateFilter("dueTo", event.target.value)} className="min-w-0 w-full rounded-xl border border-slate-200 px-2 py-2.5 text-xs font-normal" /></div><p className="text-[10px] font-normal text-slate-500">يُعتبر الموعد اليومي مستحقاً حتى نهاية يومه.</p></div>
              <div className="space-y-2 text-xs font-semibold text-slate-700"><span>الرصيد</span><div className="flex gap-2"><select aria-label="طريقة تصفية الرصيد" value={filters.balanceOperator} onChange={event => updateFilter("balanceOperator", event.target.value as Filters["balanceOperator"])} className="min-w-0 flex-1 rounded-xl border border-slate-200 px-2 text-xs"><option value="min">المبلغ وما فوقه</option><option value="max">المبلغ وما دونه</option><option value="equal">يساوي المبلغ</option></select><input aria-label="الرصيد حسب المقارنة" type="number" value={filters.balance} onChange={event => updateFilter("balance", event.target.value)} placeholder="المبلغ" className="min-w-0 w-24 rounded-xl border border-slate-200 px-3 py-2.5 text-sm font-normal outline-none focus:border-cyan-500" /></div><div className="flex items-center gap-2"><input aria-label="الرصيد من" type="number" min="0" value={filters.balanceFrom} onChange={event => updateFilter("balanceFrom", event.target.value)} placeholder="من مبلغ" className="min-w-0 w-full rounded-xl border border-slate-200 px-2 py-2 text-xs font-normal" /><span className="text-slate-400">إلى</span><input aria-label="الرصيد إلى" type="number" min="0" value={filters.balanceTo} onChange={event => updateFilter("balanceTo", event.target.value)} placeholder="إلى مبلغ" className="min-w-0 w-full rounded-xl border border-slate-200 px-2 py-2 text-xs font-normal" /></div></div>
              <div className="space-y-2 text-xs font-semibold text-slate-700"><span>المدة (يوم)</span><div className="flex gap-2"><select aria-label="طريقة تصفية المدة" value={filters.durationOperator} onChange={event => updateFilter("durationOperator", event.target.value as Filters["durationOperator"])} className="min-w-0 flex-1 rounded-xl border border-slate-200 px-2 text-xs"><option value="min">المدة وما فوقها</option><option value="max">المدة وما دونها</option><option value="equal">تساوي المدة</option></select><input aria-label="المدة حسب المقارنة" type="number" min="0" value={filters.duration} onChange={event => updateFilter("duration", event.target.value)} placeholder="المدة" className="min-w-0 w-24 rounded-xl border border-slate-200 px-3 py-2.5 text-sm font-normal outline-none focus:border-cyan-500" /></div><div className="flex items-center gap-2"><input aria-label="المدة من" type="number" min="0" value={filters.durationFrom} onChange={event => updateFilter("durationFrom", event.target.value)} placeholder="من يوم" className="min-w-0 w-full rounded-xl border border-slate-200 px-2 py-2 text-xs font-normal" /><span className="text-slate-400">إلى</span><input aria-label="المدة إلى" type="number" min="0" value={filters.durationTo} onChange={event => updateFilter("durationTo", event.target.value)} placeholder="إلى يوم" className="min-w-0 w-full rounded-xl border border-slate-200 px-2 py-2 text-xs font-normal" /></div></div>
              <div className="space-y-1.5 text-xs font-semibold text-slate-700"><span>{rangeLabel("تاريخ آخر سند قبض", filters.receiptFrom, filters.receiptTo)}</span><div className="flex items-center gap-2"><input aria-label="من تاريخ آخر سند قبض" type="date" value={filters.receiptFrom} onChange={event => updateFilter("receiptFrom", event.target.value)} className="min-w-0 w-full rounded-xl border border-slate-200 px-2 py-2.5 text-xs font-normal" /><span className="text-slate-400">إلى</span><input aria-label="إلى تاريخ آخر سند قبض" type="date" value={filters.receiptTo} onChange={event => updateFilter("receiptTo", event.target.value)} className="min-w-0 w-full rounded-xl border border-slate-200 px-2 py-2.5 text-xs font-normal" /></div></div>
              <div className="space-y-1.5 text-xs font-semibold text-slate-700"><span>{rangeLabel("تاريخ آخر فاتورة", filters.invoiceFrom, filters.invoiceTo)}</span><div className="flex items-center gap-2"><input aria-label="من تاريخ آخر فاتورة" type="date" value={filters.invoiceFrom} onChange={event => updateFilter("invoiceFrom", event.target.value)} className="min-w-0 w-full rounded-xl border border-slate-200 px-2 py-2.5 text-xs font-normal" /><span className="text-slate-400">إلى</span><input aria-label="إلى تاريخ آخر فاتورة" type="date" value={filters.invoiceTo} onChange={event => updateFilter("invoiceTo", event.target.value)} className="min-w-0 w-full rounded-xl border border-slate-200 px-2 py-2.5 text-xs font-normal" /></div></div>
            </div>
          </section>

          <section className="overflow-hidden rounded-3xl border border-slate-200 bg-white shadow-sm">
            <div className="flex items-center justify-between border-b border-slate-100 px-5 py-4"><h2 className="text-sm font-bold">قائمة الزبائن</h2><span className="text-xs text-slate-500">{filteredRecords.length.toLocaleString("ar")} نتيجة</span></div>
            <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 bg-slate-50 px-4 py-3">
              <label className="inline-flex items-center gap-2 text-xs font-semibold text-slate-700"><input type="checkbox" checked={allVisibleSelected} onChange={event => setSelectedCustomerNumbers(current => event.target.checked
                ? [...new Set([...current, ...visibleCustomerNumbers])]
                : current.filter(number => !visibleCustomerNumbers.includes(number)))} className="h-4 w-4 rounded border-slate-300 text-rose-700 focus:ring-rose-500" />تحديد كل النتائج المعروضة ({visibleCustomerNumbers.length.toLocaleString("ar")})</label>
              <div className="flex flex-wrap items-center gap-3">
              <label className="inline-flex items-center gap-2 text-xs font-semibold text-slate-700"><input type="checkbox" checked={showPaymentClassification} onChange={event => setShowPaymentClassification(event.target.checked)} className="h-4 w-4 rounded border-slate-300 text-cyan-700 focus:ring-cyan-500" />إظهار تصنيف الزبائن</label>
              <button type="button" disabled={!selectedRecords.length || loadedBlocklistKey !== blocklistStorageKey} onClick={() => {
                setBlockReasons(Object.fromEntries(selectedRecords.map(record => [
                  record.customerNumber,
                  blockedCustomers.find(customer => customer.customerNumber === record.customerNumber)?.reason ?? "",
                ])))
                setBlockEditorNumbers(selectedRecords.map(record => record.customerNumber))
              }} className="rounded-lg bg-rose-700 px-3 py-2 text-xs font-semibold text-white hover:bg-rose-800 disabled:opacity-40">إضافة المحددين للحظر ({selectedCustomerNumbers.length.toLocaleString("ar")})</button>
              <button type="button" disabled={!dataset?.records.length || loadedBlocklistKey !== blocklistStorageKey || loadedMessageTemplateKey !== storageKey} onClick={() => {
                setBulkSendMode("message")
                setBulkAudienceOpen(true)
                setBulkQueueError("")
              }} className="rounded-lg bg-emerald-700 px-3 py-2 text-xs font-semibold text-white hover:bg-emerald-800 disabled:opacity-40">إرسال رسالة جماعية</button>
              <button type="button" disabled={!dataset?.records.length || loadedBlocklistKey !== blocklistStorageKey} onClick={() => {
                setBulkSendMode("image")
                setBulkAudienceOpen(true)
                setBulkQueueError("")
              }} className="rounded-lg border border-emerald-300 bg-white px-3 py-2 text-xs font-semibold text-emerald-800 hover:bg-emerald-50 disabled:opacity-40">إرسال صورة جماعية</button>
              </div>
            </div>
            {dueReminders.length > 0 && <div role="alert" className="mx-4 mt-4 space-y-2 rounded-xl border border-amber-300 bg-amber-50 p-4 text-amber-950">
              <div className="flex flex-wrap items-center justify-between gap-2"><h3 className="flex items-center gap-2 text-sm font-bold"><BellRing className="h-4 w-4" />لديك {dueReminders.length.toLocaleString("ar")} تذكير مستحق</h3>
                {notificationPermission !== "granted" && notificationPermission !== "unsupported" && <button type="button" onClick={() => void enableBrowserNotifications()} className="rounded-lg border border-amber-300 bg-white px-3 py-1.5 text-xs font-semibold hover:bg-amber-100"><Bell className="ml-1 inline h-3.5 w-3.5" />تفعيل إشعارات المتصفح</button>}
              </div>
              {notificationPermission === "unsupported" && <p className="text-xs">المتصفح لا يدعم الإشعارات؛ ستبقى التذكيرات ظاهرة هنا عند فتح المنصة.</p>}
              {notificationPermission === "denied" && <p className="text-xs">إشعارات المتصفح محظورة؛ ستبقى التذكيرات ظاهرة هنا عند فتح المنصة.</p>}
              <ul className="space-y-2">{dueReminders.slice(0, 20).map(({ customerNumber, customerName, action }) => <li key={`${customerNumber}-${action.id}`} className="flex flex-wrap items-center justify-between gap-2 rounded-lg bg-white/80 px-3 py-2 text-xs">
                <span><strong>{customerName} ({customerNumber})</strong> · {action.description || "تذكير"}</span><span className="flex items-center gap-2"><span>موعد التنبيه: {formatDate(action.reminderDate)}</span><button type="button" onClick={() => completeAction(customerNumber, action.id)} className="rounded-md bg-amber-700 px-2 py-1 font-semibold text-white hover:bg-amber-800">تم</button></span>
              </li>)}</ul>
              {dueReminders.length > 20 && <p className="text-xs">تظهر أول 20 تذكيراً مستحقاً.</p>}
              <p className="text-[11px]">يعمل تنبيه المنصة ما دامت مفتوحة؛ وإشعار المتصفح يظهر عند استحقاق الموعد إذا كان الإذن مفعلاً، لكنه لا يعمل عند إغلاق المنصة.</p>
            </div>}
            <div className="max-h-[58vh] overflow-auto">
              <table className="w-full min-w-[1220px] border-collapse text-right text-xs">
                <thead className="sticky top-0 z-10 bg-slate-50 text-slate-600"><tr><th className="border-b border-slate-200 px-3 py-3 font-bold">تحديد</th>{DISPLAY_COLUMNS.flatMap(column => [
                  <th key={column.key} className="whitespace-nowrap border-b border-slate-200 px-3 py-3 font-bold">{column.label}</th>,
                  ...(column.key === "customerNumber" && showPaymentClassification
                    ? [<th key="paymentClassification" className="whitespace-nowrap border-b border-slate-200 px-3 py-3 font-bold">تصنيف الزبون</th>]
                    : []),
                ])}<th className="w-[170px] border-b border-slate-200 px-3 py-3 font-bold">الاتفاق</th></tr></thead>
                <tbody>
                  {filteredRecords.slice(0, 1000).map((record) => <tr key={record.customerNumber} className="align-top hover:bg-cyan-50/40">
                    <td className="border-b border-slate-100 px-3 py-3"><input type="checkbox" aria-label={`تحديد ${record.name} للحظر`} checked={selectedCustomerNumbers.includes(record.customerNumber)} onChange={() => toggleCustomerSelection(record.customerNumber)} className="h-4 w-4 rounded border-slate-300 text-rose-700 focus:ring-rose-500" /></td><td className="border-b border-slate-100 px-3 py-3"><button type="button" onClick={() => setCustomerShortcutNumber(record.customerNumber)} className="font-semibold text-cyan-800 hover:underline">{record.customerNumber}</button>{blockedCustomers.some(customer => customer.customerNumber === record.customerNumber) && <span className="mr-2 rounded bg-rose-100 px-1.5 py-0.5 text-[10px] font-semibold text-rose-800">محظور</span>}</td>{showPaymentClassification && <td className="border-b border-slate-100 px-3 py-3"><select aria-label={`تصنيف ${record.name}`} value={record.paymentClassification ?? ""} onChange={event => {
                      const value = event.target.value
                      updatePaymentClassification(record.customerNumber, value === "good" || value === "acceptable" || value === "poor" || value === "excellent" ? value : undefined)
                    }} className="min-w-32 rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-xs focus:border-cyan-500 focus:outline-none"><option value="">بدون تصنيف</option><option value="poor">سيئ الدفع</option><option value="acceptable">مقبول</option><option value="good">جيد وسالك</option><option value="excellent">ممتاز</option></select></td>}<td className="border-b border-slate-100 px-3 py-3 font-semibold text-slate-800"><button type="button" onClick={() => setOpenCustomerNumber(record.customerNumber)} className="inline-flex items-center text-right hover:text-cyan-800 hover:underline"><PaymentClassificationIndicators record={record} criteria={effectivePaymentCriteria} now={now} />{record.name}</button></td><td className={`border-b border-slate-100 px-3 py-3 font-bold tabular-nums ${record.balanceDirection === "receivable" ? "text-rose-700" : record.balanceDirection === "payable" ? "text-emerald-700" : "text-slate-700"}`}><span>{formatAmount(record.balance)}{record.balanceDirection === "receivable" ? " م" : record.balanceDirection === "payable" ? " د" : ""}</span><span className="mt-1 block text-[10px] font-medium">{record.balanceDirection === "receivable" ? "مستحق لنا" : record.balanceDirection === "payable" ? "مستحق للزبون" : "غير محدد"}</span></td><td className="border-b border-slate-100 px-3 py-3 tabular-nums">{record.duration.toLocaleString("ar")}</td><td className="border-b border-slate-100 px-3 py-3 tabular-nums">{formatAmount(record.invoiceAmount)}</td><td className="border-b border-slate-100 px-3 py-3 tabular-nums">{formatDate(record.invoiceDate)}</td><td className="border-b border-slate-100 px-3 py-3 tabular-nums">{formatAmount(record.receiptAmount)}</td>                    <td className="border-b border-slate-100 px-3 py-3 tabular-nums">{formatDate(record.receiptDate)}</td><td dir="ltr" className="border-b border-slate-100 px-3 py-3 text-right tabular-nums">{whatsappUrl(record.phone) ? <>
                      <a href={whatsappUrl(record.phone)!} target="_blank" rel="noopener noreferrer" onClick={event => {
                      if (blockedCustomers.some(customer => customer.customerNumber === record.customerNumber)) {
                        event.preventDefault()
                        setBlockedContactMessage("")
                        setBlockedContactAction("whatsapp")
                        setBlockedContactNumber(record.customerNumber)
                      }
                    }} aria-label={`فتح محادثة واتساب مع ${record.name}`} title="فتح محادثة واتساب" className={`font-semibold hover:underline ${blockedCustomers.some(customer => customer.customerNumber === record.customerNumber) ? "text-rose-700" : "text-emerald-700"}`}>{record.phone}</a>
                      {whatsappFallbackUrl(record.phone) && <a href={whatsappFallbackUrl(record.phone)!} target="_blank" rel="noopener noreferrer" onClick={event => {
                        if (blockedCustomers.some(customer => customer.customerNumber === record.customerNumber)) {
                          event.preventDefault()
                          setBlockedContactMessage("")
                          setBlockedContactAction("whatsapp")
                          setBlockedContactNumber(record.customerNumber)
                        }
                      }} className="mr-2 text-[10px] font-semibold text-cyan-800 hover:underline" title="إعادة المحاولة باستخدام رمز الدولة +970 بدل +972">جرّب +970</a>}
                    </> : record.phone || "—"}</td>
                    <td className="border-b border-slate-100 px-3 py-3">
                      <AgreementDetails
                        record={record}
                        onUpdateNote={update => updateCustomerNote(record.customerNumber, update)}
                        onUpdateAction={(actionId, update) => updateAction(record.customerNumber, actionId, update)}
                        onAddAction={() => addAction(record.customerNumber)}
                        onCloseSummary={() => closeCustomerSummary(record.customerNumber)}
                        onReopenSummary={summaryId => reopenCustomerSummary(record.customerNumber, summaryId)}
                        onOpenCustomer={() => setOpenCustomerNumber(record.customerNumber)}
                      />
                    </td>
                  </tr>)}
                  {!filteredRecords.length && <tr><td colSpan={DISPLAY_COLUMNS.length + 2 + Number(showPaymentClassification)} className="px-4 py-12 text-center text-sm text-slate-500">لا توجد نتائج مطابقة. جرّب تغيير شروط البحث.</td></tr>}
                </tbody>
              </table>
            </div>
            {shortcutCustomer && <div role="dialog" aria-modal="true" aria-labelledby="customer-shortcut-title" className="fixed inset-0 z-[90] flex items-center justify-center bg-slate-950/55 p-4">
              <section dir="rtl" className="w-full max-w-sm rounded-2xl bg-white p-5 shadow-2xl">
                <h2 id="customer-shortcut-title" className="text-base font-bold">ماذا تريد أن تفتح؟</h2>
                <p className="mt-1 text-sm text-slate-600">{shortcutCustomer.name} · رقم {shortcutCustomer.customerNumber}</p>
                <div className="mt-5 grid gap-2">
                  <button type="button" onClick={() => {
                    setCustomerShortcutNumber(null)
                    setOpenCustomerNumber(shortcutCustomer.customerNumber)
                  }} className="rounded-xl bg-cyan-800 px-4 py-3 text-sm font-semibold text-white hover:bg-cyan-900">سجل الاتفاقيات</button>
                  {whatsappUrl(shortcutCustomer.phone)
                    ? <a href={whatsappUrl(shortcutCustomer.phone)!} target="_blank" rel="noopener noreferrer" onClick={event => {
                      if (blockedCustomers.some(customer => customer.customerNumber === shortcutCustomer.customerNumber)) {
                        event.preventDefault()
                        setCustomerShortcutNumber(null)
                        setBlockedContactMessage("")
                        setBlockedContactAction("whatsapp")
                        setBlockedContactNumber(shortcutCustomer.customerNumber)
                      } else {
                        setCustomerShortcutNumber(null)
                      }
                    }} className="rounded-xl bg-emerald-700 px-4 py-3 text-center text-sm font-semibold text-white hover:bg-emerald-800">محادثة واتساب <span dir="ltr">({shortcutCustomer.phone})</span></a>
                    : <p className="rounded-xl bg-amber-50 px-3 py-2 text-xs leading-5 text-amber-900">لا يوجد رقم جوال صالح في الكشف لفتح واتساب.</p>}
                  {whatsappFallbackUrl(shortcutCustomer.phone) && <a href={whatsappFallbackUrl(shortcutCustomer.phone)!} target="_blank" rel="noopener noreferrer" onClick={event => {
                    if (blockedCustomers.some(customer => customer.customerNumber === shortcutCustomer.customerNumber)) {
                      event.preventDefault()
                      setCustomerShortcutNumber(null)
                      setBlockedContactMessage("")
                      setBlockedContactAction("whatsapp")
                      setBlockedContactNumber(shortcutCustomer.customerNumber)
                    } else {
                      setCustomerShortcutNumber(null)
                    }
                  }} className="rounded-xl border border-cyan-300 px-4 py-2.5 text-center text-sm font-semibold text-cyan-800 hover:bg-cyan-50">إذا لم يعمل +972، جرّب +970</a>}
                  {whatsappUrl(shortcutCustomer.phone) && <button type="button" disabled={loadedMessageTemplateKey !== storageKey} onClick={() => {
                    setCustomerShortcutNumber(null)
                    startStatementUpload(shortcutCustomer)
                  }} className="rounded-xl border border-emerald-300 px-4 py-3 text-center text-sm font-semibold text-emerald-800 hover:bg-emerald-50 disabled:opacity-50">إرسال رسالة أو كشف حساب عبر واتساب <span className="block text-[10px] font-normal">اختر رسالة فقط أو ملفاً أو الاثنين معاً</span></button>}
                  {whatsappUrl(shortcutCustomer.phone) && <button type="button" onClick={() => startCustomerImage(shortcutCustomer)} className="rounded-xl border border-cyan-300 px-4 py-3 text-center text-sm font-semibold text-cyan-900 hover:bg-cyan-50">إرسال صورة رصيد الزبون <span className="block text-[10px] font-normal">إنشاء صورة باسمه ورصيده المستحق</span></button>}
                  <button type="button" onClick={() => setCustomerShortcutNumber(null)} className="rounded-xl border border-slate-200 px-4 py-2.5 text-sm font-semibold text-slate-600 hover:bg-slate-50">إلغاء</button>
                </div>
              </section>
            </div>}
            {openCustomer && <CustomerAgreementProfile
              record={openCustomer}
              isBlocked={blockedCustomers.some(customer => customer.customerNumber === openCustomer.customerNumber)}
              blockedReason={blockedCustomers.find(customer => customer.customerNumber === openCustomer.customerNumber)?.reason ?? ""}
              onClose={() => setOpenCustomerNumber(null)}
              onCreateAgreement={agreement => createAgreement(openCustomer.customerNumber, agreement)}
              onUpdateAgreement={(agreementId, update) => updateAgreement(openCustomer.customerNumber, agreementId, update)}
              onSetCheckOutcome={(agreementId, checkId, outcome, actualPaid, note) => setAgreementCheckOutcome(openCustomer.customerNumber, agreementId, checkId, outcome, actualPaid, note)}
              onCloseAgreement={(agreementId, status, outcome, paidAmount) => closeAgreement(openCustomer.customerNumber, agreementId, status, outcome, paidAmount)}
              onDeleteAgreement={agreementId => deleteAgreement(openCustomer.customerNumber, agreementId)}
            />}
            {filteredRecords.length > 1000 && <p className="border-t border-slate-100 px-4 py-3 text-xs text-slate-500">يتم عرض أول 1,000 نتيجة لتسهيل التصفح. صدّر إلى Excel لتنزيل جميع النتائج المطابقة.</p>}
          </section>
        </>}

        {bulkAudienceOpen && <div role="dialog" aria-modal="true" aria-labelledby="bulk-message-title" className="fixed inset-0 z-[100] flex items-center justify-center overflow-y-auto bg-slate-950/60 p-4">
          <section dir="rtl" className="my-auto max-h-[90vh] w-full max-w-xl space-y-4 overflow-y-auto rounded-2xl bg-white p-5 shadow-2xl">
            <div>
              <h2 id="bulk-message-title" className="text-lg font-bold">{bulkSendMode === "image" ? "إرسال صورة شخصية للزبائن" : "إرسال رسالة للزبائن"}</h2>
              <p className="mt-1 text-sm leading-6 text-slate-600">{bulkSendMode === "image"
                ? "سننشئ صورة مستقلة لكل زبون باسمه ورصيده، ثم تشاركها عبر واتساب أو تنزّلها وتُرفقها يدوياً."
                : "اختر المستلمين ونص الرسالة. ستُفتح محادثة كل زبون بالتتابع، ويجب الضغط على إرسال داخل واتساب ثم تأكيد الانتقال للزبون التالي."}</p>
            </div>
            <fieldset className="grid gap-2 sm:grid-cols-2">
              <legend className="mb-2 text-xs font-semibold text-slate-700">إلى من تريد الإرسال؟</legend>
              <label className={`flex cursor-pointer items-center gap-2 rounded-xl border p-3 text-sm font-semibold ${bulkAudience === "all" ? "border-emerald-400 bg-emerald-50 text-emerald-900" : "border-slate-200"}`}>
                <input type="radio" name="bulk-audience" checked={bulkAudience === "all"} onChange={() => setBulkAudience("all")} />
                إرسال للجميع
              </label>
              <label className={`flex cursor-pointer items-center gap-2 rounded-xl border p-3 text-sm font-semibold ${bulkAudience === "exclude_blocked" ? "border-emerald-400 bg-emerald-50 text-emerald-900" : "border-slate-200"}`}>
                <input type="radio" name="bulk-audience" checked={bulkAudience === "exclude_blocked"} onChange={() => setBulkAudience("exclude_blocked")} />
                إرسال للجميع ما عدا المحظورين
              </label>
            </fieldset>
            {bulkSendMode === "message" && <label className="block space-y-1 text-xs font-semibold text-slate-700">
              <span>نص الرسالة</span>
              <select value={bulkMessageType} onChange={event => setBulkMessageType(event.target.value)} className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm font-normal">
                <option value="statement">رسالة كشف الحساب</option>
                <option value="weekly_installment">تذكير القسط الأسبوعي</option>
                {customMessageTemplates.map(template => <option key={template.id} value={`custom:${template.id}`}>{template.name}</option>)}
              </select>
            </label>}
            {bulkSendMode === "image" && <>
              <fieldset className="grid gap-2 sm:grid-cols-2">
                <legend className="mb-2 text-xs font-semibold text-slate-700">تصميم البطاقة</legend>
                <label className={`flex cursor-pointer items-center gap-2 rounded-xl border p-3 text-sm font-semibold ${bulkCardTemplate === "modern" ? "border-emerald-400 bg-emerald-50 text-emerald-900" : "border-slate-200"}`}><input type="radio" name="bulk-card-template" checked={bulkCardTemplate === "modern"} onChange={() => setBulkCardTemplate("modern")} />التصميم الحديث</label>
                <label className={`flex cursor-pointer items-center gap-2 rounded-xl border p-3 text-sm font-semibold ${bulkCardTemplate === "classic" ? "border-emerald-400 bg-emerald-50 text-emerald-900" : "border-slate-200"}`}><input type="radio" name="bulk-card-template" checked={bulkCardTemplate === "classic"} onChange={() => setBulkCardTemplate("classic")} />التصميم الكلاسيكي مثل الصورة</label>
              </fieldset>
              <label className="block space-y-1 text-xs font-semibold text-slate-700">
                <span>الرسالة التي تظهر داخل البطاقة</span>
                <select value={bulkMessageType} onChange={event => {
                  setBulkMessageType(event.target.value)
                  setMessageTemplateEditorType(event.target.value)
                }} className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm font-normal">
                  <option value="statement">رسالة كشف الحساب</option>
                  <option value="weekly_installment">تذكير القسط الأسبوعي</option>
                  {customMessageTemplates.map(template => <option key={template.id} value={`custom:${template.id}`}>{template.name}</option>)}
                </select>
              </label>
              <div className="space-y-2">
                <textarea ref={bulkImageMessageTextAreaRef} rows={5} value={selectedBulkMessageTemplate}
                  onChange={event => updateMessageTemplateContent(bulkMessageType, event.target.value)}
                  onDragOver={event => event.preventDefault()}
                  onDrop={event => {
                    event.preventDefault()
                    const token = event.dataTransfer.getData("text/plain")
                    if (MESSAGE_FIELDS.some(field => field.token === token)) insertBulkImageMessageField(token)
                  }}
                  aria-label="نص الرسالة على البطاقة"
                  className="w-full resize-y rounded-xl border border-slate-200 px-3 py-2.5 text-sm font-normal" />
                <p className="text-[11px] font-semibold text-slate-600">اسحب الحقول إلى موضعها داخل الرسالة، أو اضغط لإضافتها:</p>
                <div className="flex flex-wrap gap-2">{MESSAGE_FIELDS.map(field => <button key={field.token} type="button" draggable onDragStart={event => event.dataTransfer.setData("text/plain", field.token)} onClick={() => insertBulkImageMessageField(field.token)} className="cursor-grab rounded-lg border border-cyan-200 bg-white px-3 py-2 text-xs font-semibold text-cyan-900 active:cursor-grabbing">{field.label} <code dir="ltr">{field.token}</code></button>)}</div>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <button type="button" onClick={() => setShowNewMessageTemplate(value => !value)} className="inline-flex items-center gap-1 rounded-lg border border-emerald-300 bg-white px-3 py-2 text-xs font-semibold text-emerald-900"><Plus className="h-4 w-4" />إضافة رسالة جديدة</button>
                {showNewMessageTemplate && <div className="flex flex-1 flex-wrap gap-2"><input value={newMessageTemplateName} onChange={event => setNewMessageTemplateName(event.target.value)} onKeyDown={event => { if (event.key === "Enter") addCustomMessageTemplate() }} placeholder="اسم الرسالة الجديدة" className="min-w-48 flex-1 rounded-lg border border-slate-200 px-3 py-2 text-xs font-normal" /><button type="button" onClick={addCustomMessageTemplate} className="rounded-lg bg-emerald-700 px-3 py-2 text-xs font-semibold text-white">إنشاء واستخدام</button></div>}
              </div>
            </>}
            <div className="rounded-xl border border-slate-200 bg-slate-50 p-3">
              <p className="text-xs font-semibold text-slate-700">سيتم تجهيز {bulkEligibleRecords.length.toLocaleString("ar")} {bulkSendMode === "image" ? "صورة" : "محادثة"} من أصل {bulkAudienceRecords.length.toLocaleString("ar")} زبون.</p>
              {bulkAudienceRecords.length > bulkEligibleRecords.length && <p className="mt-1 text-xs leading-5 text-amber-800">سيتم تخطي { (bulkAudienceRecords.length - bulkEligibleRecords.length).toLocaleString("ar") } زبوناً لعدم وجود رقم واتساب صالح{bulkSendMode === "message" && bulkMessageType === "weekly_installment" ? " أو مبلغ قسط أسبوعي مسجل" : ""}.</p>}
              {bulkSendMode === "message" && bulkEligibleRecords[0] && <p className="mt-2 max-h-28 overflow-y-auto whitespace-pre-wrap text-xs leading-5 text-slate-600">{customerMessage(bulkEligibleRecords[0], selectedBulkMessageTemplate, String(weeklyInstallmentAmount(bulkEligibleRecords[0]) ?? ""))}</p>}
              {bulkSendMode === "image" && bulkEligibleRecords[0] && <div className="mt-2 flex items-center gap-3 text-xs leading-5 text-slate-600"><span>الصورة تعرض: {bulkEligibleRecords[0].name} · {formatAmount(bulkEligibleRecords[0].balance)} شيكل</span></div>}
            </div>
            {bulkQueueError && <p role="alert" className="rounded-xl border border-rose-200 bg-rose-50 p-3 text-xs leading-5 text-rose-800">{bulkQueueError}</p>}
            <div className="flex flex-wrap justify-end gap-2">
              <button type="button" onClick={() => setBulkAudienceOpen(false)} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-semibold text-slate-600">إلغاء</button>
              <button type="button" disabled={!bulkEligibleRecords.length} onClick={() => {
                const recipients = bulkEligibleRecords.map(record => ({
                  customerNumber: record.customerNumber,
                  name: record.name,
                  phone: record.phone,
                  message: customerMessage(record, selectedBulkMessageTemplate, String(weeklyInstallmentAmount(record) ?? "")),
                  balance: record.balance,
                  balanceDirection: record.balanceDirection,
                }))
                if (!recipients.length) {
                  setBulkQueueError("لا يوجد زبائن بأرقام واتساب صالحة للإرسال حسب الاختيار.")
                  return
                }
                setBulkQueue(recipients)
                setBulkQueueIndex(0)
                setBulkAudienceOpen(false)
                setBulkQueueError("")
              }} className="rounded-lg bg-emerald-700 px-4 py-2 text-sm font-semibold text-white hover:bg-emerald-800 disabled:opacity-40">{bulkSendMode === "image" ? "ابدأ تجهيز الصور" : "ابدأ الإرسال"}</button>
            </div>
          </section>
        </div>}

        {bulkQueue && bulkQueueRecipient && <div role="dialog" aria-modal="true" aria-labelledby="bulk-queue-title" className="fixed inset-0 z-[110] flex items-center justify-center overflow-y-auto bg-slate-950/60 p-4">
          <section dir="rtl" className="my-auto max-h-[90vh] w-full max-w-xl space-y-4 overflow-y-auto rounded-2xl bg-white p-5 shadow-2xl">
            <div>
              <h2 id="bulk-queue-title" className="text-lg font-bold">{bulkSendMode === "image" ? "إرسال الصور" : "الإرسال الجماعي"} · { (bulkQueueIndex + 1).toLocaleString("ar") } من {bulkQueue.length.toLocaleString("ar")}</h2>
              <p className="mt-1 text-sm text-slate-600">{bulkQueueRecipient.name} · رقم الزبون {bulkQueueRecipient.customerNumber} · <span dir="ltr">{bulkQueueRecipient.phone}</span></p>
            </div>
            {bulkSendMode === "message" && <div className="rounded-xl border border-slate-200 bg-slate-50 p-3">
              <h3 className="mb-2 text-xs font-bold text-slate-700">معاينة الرسالة</h3>
              <p className="max-h-48 overflow-y-auto whitespace-pre-wrap text-sm leading-6 text-slate-800">{bulkQueueRecipient.message}</p>
            </div>}
            {bulkSendMode === "message"
              ? <>
                <p className="rounded-xl bg-amber-50 p-3 text-xs leading-5 text-amber-900">افتح المحادثة وأرسل الرسالة يدوياً من واتساب. لا تستطيع المنصة معرفة نجاح الإرسال تلقائياً.</p>
                <div className="flex flex-wrap gap-2">
                  <a href={whatsappUrl(bulkQueueRecipient.phone, bulkQueueRecipient.message, "972")!} target="_blank" rel="noopener noreferrer" className="rounded-lg bg-emerald-700 px-4 py-2 text-sm font-semibold text-white hover:bg-emerald-800">فتح واتساب عبر +972</a>
                  {whatsappFallbackUrl(bulkQueueRecipient.phone, bulkQueueRecipient.message) && <a href={whatsappFallbackUrl(bulkQueueRecipient.phone, bulkQueueRecipient.message)!} target="_blank" rel="noopener noreferrer" className="rounded-lg border border-cyan-300 px-4 py-2 text-sm font-semibold text-cyan-800 hover:bg-cyan-50">جرّب الرقم البديل +970</a>}
                </div>
              </>
              : <>
                <fieldset className="grid gap-2 sm:grid-cols-2">
                  <legend className="mb-2 text-xs font-semibold text-slate-700">تصميم البطاقة</legend>
                  <label className={`flex cursor-pointer items-center gap-2 rounded-xl border p-3 text-xs font-semibold ${bulkCardTemplate === "modern" ? "border-emerald-400 bg-emerald-50 text-emerald-900" : "border-slate-200"}`}><input type="radio" name="queue-card-template" checked={bulkCardTemplate === "modern"} onChange={() => setBulkCardTemplate("modern")} />التصميم الحديث</label>
                  <label className={`flex cursor-pointer items-center gap-2 rounded-xl border p-3 text-xs font-semibold ${bulkCardTemplate === "classic" ? "border-emerald-400 bg-emerald-50 text-emerald-900" : "border-slate-200"}`}><input type="radio" name="queue-card-template" checked={bulkCardTemplate === "classic"} onChange={() => setBulkCardTemplate("classic")} />التصميم مثل الصورة الثانية</label>
                </fieldset>
                <label className="block space-y-1 text-xs font-semibold text-slate-700">
                  <span>قالب النص المرفق داخل البطاقة</span>
                  <select value={bulkMessageType} onChange={event => {
                    const templateId = event.target.value
                    setBulkMessageType(templateId)
                    setMessageTemplateEditorType(templateId)
                    const content = templateId === "weekly_installment"
                      ? weeklyInstallmentMessageTemplate
                      : customMessageTemplates.find(template => templateId === `custom:${template.id}`)?.content ?? statementMessageTemplate
                    updateBulkQueueTemplate(templateId, content)
                  }} className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm font-normal">
                    <option value="statement">رسالة كشف الحساب</option>
                    <option value="weekly_installment">تذكير القسط الأسبوعي</option>
                    {customMessageTemplates.map(template => <option key={template.id} value={`custom:${template.id}`}>{template.name}</option>)}
                  </select>
                </label>
                <div className="space-y-2">
                  <textarea ref={bulkImageMessageTextAreaRef} rows={4} value={selectedBulkMessageTemplate}
                    onChange={event => updateBulkQueueTemplate(bulkMessageType, event.target.value)}
                    onDragOver={event => event.preventDefault()}
                    onDrop={event => {
                      event.preventDefault()
                      const token = event.dataTransfer.getData("text/plain")
                      if (MESSAGE_FIELDS.some(field => field.token === token)) insertBulkImageMessageField(token)
                    }}
                    aria-label="الرسالة المرفقة على صورة الزبون"
                    className="w-full resize-y rounded-xl border border-slate-200 px-3 py-2.5 text-sm font-normal" />
                  <p className="text-[11px] font-semibold text-slate-600">اسحب الحقول إلى موضعها داخل الرسالة، أو اضغط لإضافتها:</p>
                  <div className="flex flex-wrap gap-2">{MESSAGE_FIELDS.map(field => <button key={field.token} type="button" draggable onDragStart={event => event.dataTransfer.setData("text/plain", field.token)} onClick={() => insertBulkImageMessageField(field.token)} className="cursor-grab rounded-lg border border-cyan-200 bg-white px-3 py-2 text-xs font-semibold text-cyan-900 active:cursor-grabbing">{field.label}</button>)}</div>
                </div>
                <p className="rounded-xl bg-amber-50 p-3 text-xs leading-5 text-amber-900">هذه صورة مخصصة لهذا الزبون وتحتوي اسمه ورصيده ({formatAmount(bulkQueueRecipient.balance)} شيكل). شاركها مع واتساب أو نزّلها وأرفقها في المحادثة.</p>
                {bulkImageBusy && <p className="rounded-xl bg-slate-50 p-3 text-sm text-slate-600">جارٍ تجهيز الصورة...</p>}
                {bulkImageError && <p role="alert" className="rounded-xl border border-rose-200 bg-rose-50 p-3 text-xs leading-5 text-rose-800">{bulkImageError}</p>}
                {bulkImagePreviewUrl && <div className="flex justify-center rounded-xl border border-slate-200 bg-slate-50 p-3"><img src={bulkImagePreviewUrl} alt={`صورة رصيد ${bulkQueueRecipient.name}`} className="max-h-72 w-auto rounded-lg shadow-sm" /></div>}
                <div className="flex flex-wrap gap-2">
                  {bulkImageFile && bulkImagePreviewUrl && <a href={bulkImagePreviewUrl} download={bulkImageFile.name} className="rounded-lg border border-emerald-300 px-4 py-2 text-sm font-semibold text-emerald-800 hover:bg-emerald-50">تنزيل صورة الزبون</a>}
                  <a href={whatsappUrl(bulkQueueRecipient.phone)!} target="_blank" rel="noopener noreferrer" className="rounded-lg bg-emerald-700 px-4 py-2 text-sm font-semibold text-white hover:bg-emerald-800">فتح واتساب عبر +972</a>
                  {whatsappFallbackUrl(bulkQueueRecipient.phone) && <a href={whatsappFallbackUrl(bulkQueueRecipient.phone)!} target="_blank" rel="noopener noreferrer" className="rounded-lg border border-cyan-300 px-4 py-2 text-sm font-semibold text-cyan-800 hover:bg-cyan-50">جرّب الرقم البديل +970</a>}
                  {bulkImageFile && <button type="button" onClick={async () => {
                    if (!navigator.share || !navigator.canShare?.({ files: [bulkImageFile] })) {
                      setBulkImageError("مشاركة الصورة مباشرة غير مدعومة على هذا الجهاز. نزّل الصورة وافتح واتساب لإرفاقها.")
                      return
                    }
                    try {
                      await navigator.share({
                        files: [bulkImageFile],
                        title: `رصيد ${bulkQueueRecipient.name}`,
                        text: `صورة رصيد ${bulkQueueRecipient.name}`,
                      })
                    } catch (error) {
                      if (error instanceof DOMException && error.name === "AbortError") return
                      setBulkImageError(error instanceof Error ? `تعذرت مشاركة الصورة: ${error.message}` : "تعذرت مشاركة الصورة.")
                    }
                  }} className="rounded-lg bg-cyan-800 px-4 py-2 text-sm font-semibold text-white hover:bg-cyan-900">مشاركة الصورة</button>}
                </div>
              </>}
            <div className="flex flex-wrap justify-end gap-2">
              <button type="button" onClick={() => setBulkQueue(null)} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-semibold text-slate-600">إنهاء</button>
              <button type="button" onClick={() => {
                if (bulkQueueIndex + 1 >= bulkQueue.length) setBulkQueue(null)
                else setBulkQueueIndex(index => index + 1)
              }} className="rounded-lg bg-emerald-700 px-4 py-2 text-sm font-semibold text-white hover:bg-emerald-800">تم الإرسال، التالي</button>
              <button type="button" onClick={() => {
                if (bulkQueueIndex + 1 >= bulkQueue.length) setBulkQueue(null)
                else setBulkQueueIndex(index => index + 1)
              }} className="rounded-lg border border-amber-300 px-4 py-2 text-sm font-semibold text-amber-900 hover:bg-amber-50">تخطي هذا الزبون</button>
            </div>
          </section>
        </div>}

        {statementCustomer && statementCustomerNumber && <div role="dialog" aria-modal="true" aria-labelledby="statement-share-title" className="fixed inset-0 z-[100] flex items-center justify-center overflow-y-auto bg-slate-950/60 p-4">
          <section dir="rtl" className="my-auto max-h-[90vh] w-full max-w-2xl space-y-4 overflow-y-auto rounded-2xl bg-white p-5 shadow-2xl">
            <div><h2 id="statement-share-title" className="text-lg font-bold">إرسال إلى واتساب</h2><p className="mt-1 text-sm text-slate-600">{statementCustomer.name} · {statementCustomer.phone}</p></div>
            <fieldset className="grid gap-2 sm:grid-cols-3"><legend className="mb-2 text-xs font-semibold text-slate-700">ماذا تريد أن ترسل؟</legend>
              <label className={`flex cursor-pointer items-center gap-2 rounded-xl border p-3 text-xs font-semibold ${statementMode === "both" ? "border-emerald-400 bg-emerald-50 text-emerald-900" : "border-slate-200"}`}><input type="radio" name="statement-mode" checked={statementMode === "both"} onChange={() => setStatementMode("both")} />الرسالة والملف</label>
              <label className={`flex cursor-pointer items-center gap-2 rounded-xl border p-3 text-xs font-semibold ${statementMode === "file" ? "border-emerald-400 bg-emerald-50 text-emerald-900" : "border-slate-200"}`}><input type="radio" name="statement-mode" checked={statementMode === "file"} onChange={() => setStatementMode("file")} />الملف فقط</label>
              <label className={`flex cursor-pointer items-center gap-2 rounded-xl border p-3 text-xs font-semibold ${statementMode === "message" ? "border-emerald-400 bg-emerald-50 text-emerald-900" : "border-slate-200"}`}><input type="radio" name="statement-mode" checked={statementMode === "message"} onChange={() => setStatementMode("message")} />الرسالة فقط</label>
            </fieldset>
            {statementMode !== "file" && <>
              <fieldset className="grid gap-2 sm:grid-cols-2"><legend className="mb-2 text-xs font-semibold text-slate-700">اختر نص الرسالة</legend>
                {[{ id: "statement", label: "رسالة كشف الحساب" }, { id: "weekly_installment", label: "تذكير القسط الأسبوعي" }, ...customMessageTemplates.map(template => ({ id: `custom:${template.id}`, label: template.name }))].map(template => <button key={template.id} type="button" aria-pressed={statementMessageType === template.id} onClick={() => setStatementMessageType(template.id)} className={`rounded-xl border p-3 text-right text-xs font-semibold ${statementMessageType === template.id ? "border-emerald-400 bg-emerald-50 text-emerald-900" : "border-slate-200 text-slate-700 hover:bg-slate-50"}`}>{template.label}</button>)}
              </fieldset>
              {statementMessageType === "weekly_installment" && <label className="block space-y-1 text-xs font-semibold text-slate-700"><span>مبلغ القسط الأسبوعي</span><input type="number" min="0.01" step="0.01" value={weeklyInstallmentAmountInput} onChange={event => setWeeklyInstallmentAmountInput(event.target.value)} placeholder="أدخل مبلغ القسط الأسبوعي" className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm font-normal" />{!weeklyInstallmentAmountInput && <span className="font-normal text-amber-700">لا يوجد مبلغ قسط مسجل لهذا الزبون؛ أدخل المبلغ قبل الإرسال.</span>}</label>}
            </>}
            {statementMode !== "message" && <div className="flex flex-wrap items-center gap-3 rounded-xl border border-slate-200 p-3">
              <button type="button" onClick={() => {
                if (statementFileInputRef.current) {
                  statementFileInputRef.current.value = ""
                  statementFileInputRef.current.click()
                }
              }} className="rounded-lg border border-cyan-300 px-3 py-2 text-xs font-semibold text-cyan-800 hover:bg-cyan-50">{statementFile ? "اختيار ملف آخر" : "اختر ملف كشف Excel"}</button>
              <span className="text-xs text-slate-600">{statementFile?.name ?? "لم يتم اختيار ملف بعد"}</span>
            </div>}
            {statementMode !== "file" && <div className="rounded-xl border border-slate-200 bg-slate-50 p-3"><h3 className="mb-2 text-xs font-bold text-slate-700">معاينة الرسالة</h3><p className="max-h-48 overflow-y-auto whitespace-pre-wrap text-sm leading-6 text-slate-800">{selectedCustomerMessage}</p></div>}
            {statementShareError && <p role="alert" className="rounded-xl border border-rose-200 bg-rose-50 p-3 text-xs leading-5 text-rose-800">{statementShareError}</p>}
            {statementMode !== "message" && (shareSelectedStatementDirectly
              ? <div className="rounded-xl bg-emerald-50 p-3 text-xs leading-5 text-emerald-900">يمكن لجهازك مشاركة ملف Excel. اختر واتساب من قائمة المشاركة وأكمل الإرسال بنفسك.</div>
              : <div className="rounded-xl bg-amber-50 p-3 text-xs leading-5 text-amber-900">إذا لم يدعم جهازك مشاركة الملفات، نزّل الملف ثم افتح واتساب وأرفقه يدوياً.</div>)}
            <div className="flex flex-wrap justify-end gap-2">
              <button type="button" onClick={() => {
                setStatementFile(null)
                setStatementCustomerNumber(null)
              }} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-semibold text-slate-600">إلغاء</button>
              {statementMode !== "message" && statementFile && <button type="button" onClick={downloadStatementFile} className="rounded-lg border border-emerald-300 px-4 py-2 text-sm font-semibold text-emerald-800 hover:bg-emerald-50">تنزيل ملف Excel</button>}
              {whatsappUrl(statementCustomer.phone, statementMode === "file" ? undefined : selectedCustomerMessage) && canSendSelectedMessage && (statementMode === "message" || statementFile) && <a href={whatsappUrl(statementCustomer.phone, statementMode === "file" ? undefined : selectedCustomerMessage)!} target="_blank" rel="noopener noreferrer" className="rounded-lg border border-emerald-300 px-4 py-2 text-sm font-semibold text-emerald-800 hover:bg-emerald-50">{statementMode === "message" ? "فتح واتساب بالرسالة" : "فتح واتساب"}</a>}
              {whatsappFallbackUrl(statementCustomer.phone, statementMode === "file" ? undefined : selectedCustomerMessage) && canSendSelectedMessage && (statementMode === "message" || statementFile) && <a href={whatsappFallbackUrl(statementCustomer.phone, statementMode === "file" ? undefined : selectedCustomerMessage)!} target="_blank" rel="noopener noreferrer" className="rounded-lg border border-cyan-300 px-4 py-2 text-sm font-semibold text-cyan-800 hover:bg-cyan-50">إعادة المحاولة عبر +970</a>}
              {shareSelectedStatementDirectly && <button type="button" disabled={!statementFile || (statementMode === "both" && !canSendSelectedMessage)} onClick={() => void shareStatement()} className="rounded-lg bg-emerald-700 px-4 py-2 text-sm font-semibold text-white hover:bg-emerald-800 disabled:opacity-50">مشاركة {statementMode === "both" ? "الملف والرسالة" : "الملف"}</button>}
            </div>
          </section>
        </div>}

        {blockEditorNumbers && <div role="dialog" aria-modal="true" aria-labelledby="block-editor-title" className="fixed inset-0 z-[100] flex items-center justify-center overflow-y-auto bg-slate-950/60 p-4">
          <section dir="rtl" className="my-auto max-h-[90vh] w-full max-w-2xl space-y-4 overflow-y-auto rounded-2xl bg-white p-5 shadow-2xl">
            <div><h2 id="block-editor-title" className="text-lg font-bold text-rose-900">إضافة إلى قائمة حظر التواصل</h2><p className="mt-1 text-xs text-slate-600">يمكنك تسجيل سبب اختياري لكل زبون. ستبقى الأسماء في القائمة حتى ترفع الحظر يدوياً، حتى بعد تحديث ملف Excel.</p></div>
            <div className="space-y-3">{blockEditorNumbers.map(customerNumber => {
              const customer = dataset?.records.find(record => record.customerNumber === customerNumber)
              return <label key={customerNumber} className="block space-y-1.5 rounded-xl border border-slate-200 p-3 text-xs font-semibold text-slate-700">
                <span>{customer?.name ?? customerNumber} · رقم {customerNumber}</span>
                <textarea rows={2} value={blockReasons[customerNumber] ?? ""} onChange={event => setBlockReasons(current => ({ ...current, [customerNumber]: event.target.value }))} placeholder="سبب الحظر (اختياري)" className="w-full rounded-lg border border-slate-200 px-3 py-2 text-xs font-normal" />
              </label>
            })}</div>
            <div className="flex flex-wrap justify-end gap-2"><button type="button" onClick={() => setBlockEditorNumbers(null)} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-semibold text-slate-600">إلغاء</button><button type="button" onClick={saveBlockedCustomers} className="rounded-lg bg-rose-700 px-4 py-2 text-sm font-semibold text-white hover:bg-rose-800">حفظ {blockEditorNumbers.length.toLocaleString("ar")} زبوناً في قائمة الحظر</button></div>
          </section>
        </div>}

        {blockedContact && <div role="dialog" aria-modal="true" aria-labelledby="blocked-contact-title" className="fixed inset-0 z-[110] flex items-center justify-center bg-slate-950/60 p-4">
          <section dir="rtl" className="w-full max-w-md space-y-4 rounded-2xl border-t-4 border-rose-600 bg-white p-5 shadow-2xl">
            <div><h2 id="blocked-contact-title" className="text-lg font-bold text-rose-900">تنبيه: الزبون من قائمة المحظورين</h2><p className="mt-1 text-sm text-slate-700">{blockedContact.name} · رقم الزبون {blockedContact.customerNumber}</p>{blockedContact.phone && <p dir="ltr" className="mt-1 text-left text-sm text-slate-600">{blockedContact.phone}</p>}</div>
            {blockedContact.reason && <div className="rounded-xl bg-rose-50 p-3 text-sm leading-6 text-rose-900"><strong>الملاحظة:</strong> {blockedContact.reason}</div>}
            <p className="text-xs leading-5 text-slate-600">تأكد من سبب الحظر قبل المتابعة. لن تبدأ عملية التواصل إلا إذا اخترت المتابعة صراحةً.</p>
            <div className="flex flex-wrap justify-end gap-2"><button type="button" onClick={() => {
              setBlockedContactNumber(null)
              setBlockedContactMessage("")
              setBlockedContactAction(null)
            }} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-semibold text-slate-600">رجوع</button>{blockedContactAction === "image"
              ? blockedContactRecord
                ? <button type="button" onClick={() => {
                  startCustomerImage(blockedContactRecord, true)
                  setBlockedContactNumber(null)
                  setBlockedContactAction(null)
                }} className="rounded-lg bg-rose-700 px-4 py-2 text-sm font-semibold text-white hover:bg-rose-800">متابعة وإنشاء صورة الرصيد</button>
                : <p className="text-xs text-rose-700">لا يوجد كشف محفوظ لهذا الزبون لإنشاء صورة الرصيد.</p>
              : blockedContactAction === "statement"
              ? <button type="button" onClick={() => {
                setStatementCustomerNumber(blockedContact.customerNumber)
                setStatementMode("both")
                setStatementFile(null)
                setStatementShareError("")
                setBlockedContactNumber(null)
                setBlockedContactAction(null)
              }} className="rounded-lg bg-rose-700 px-4 py-2 text-sm font-semibold text-white hover:bg-rose-800">متابعة لاختيار رسالة أو ملف</button>
              : whatsappUrl(blockedContact.phone, blockedContactMessage) && <>
              <a href={whatsappUrl(blockedContact.phone, blockedContactMessage)!} target="_blank" rel="noopener noreferrer" onClick={() => {
              setBlockedContactNumber(null)
              setBlockedContactMessage("")
              setBlockedContactAction(null)
            }} className="rounded-lg bg-rose-700 px-4 py-2 text-sm font-semibold text-white hover:bg-rose-800">متابعة إلى واتساب</a>
              {whatsappFallbackUrl(blockedContact.phone, blockedContactMessage) && <a href={whatsappFallbackUrl(blockedContact.phone, blockedContactMessage)!} target="_blank" rel="noopener noreferrer" onClick={() => {
                setBlockedContactNumber(null)
                setBlockedContactMessage("")
                setBlockedContactAction(null)
              }} className="rounded-lg border border-rose-300 px-4 py-2 text-sm font-semibold text-rose-800 hover:bg-rose-50">المتابعة عبر +970</a>}
              </>}</div>
          </section>
        </div>}

        {!dataset && ready && <section className="rounded-3xl border-2 border-dashed border-slate-200 bg-white px-6 py-14 text-center">
          <div className="mx-auto flex h-14 w-14 items-center justify-center rounded-2xl bg-cyan-50 text-cyan-700"><Upload className="h-6 w-6" /></div>
          <h2 className="mt-4 text-base font-bold text-slate-800">ابدأ برفع ملف أرصدة الزبائن</h2>
          <p className="mx-auto mt-2 max-w-xl text-sm leading-7 text-slate-500">يعثر النظام على صف العناوين داخل كل ورقة، ويتجاهل العناوين والصفوف الفارغة، ويدمج سجلات الصفحات في قائمة واحدة.</p>
        </section>}
      </div>
      {reportOpen && <div role="dialog" aria-modal="true" aria-labelledby="customer-reports-title" className="fixed inset-0 z-[95] flex items-center justify-center overflow-y-auto bg-slate-950/55 p-4">
        <section className="my-auto w-full max-w-2xl space-y-5 rounded-3xl bg-white p-5 shadow-2xl sm:p-7">
          <header className="flex items-start justify-between gap-4">
            <div><h2 id="customer-reports-title" className="text-xl font-extrabold text-slate-900">تقارير الزبائن</h2><p className="mt-1 text-sm text-slate-500">اختر تقريراً وحدد نطاق التاريخ والمبلغ قبل التصدير.</p></div>
            <button type="button" onClick={() => setReportOpen(false)} aria-label="إغلاق التقارير" className="rounded-lg p-2 text-slate-500 hover:bg-slate-100"><X className="h-5 w-5" /></button>
          </header>
          <label className="block space-y-1.5 text-sm font-semibold text-slate-700"><span>نوع التقرير</span><select value={reportType} onChange={event => setReportType(event.target.value as CustomerReportType)} className="w-full rounded-xl border border-slate-200 bg-white px-3 py-3 text-sm">
            <option value="balances">ملخص الأرصدة</option>
            <option value="agreements">الاتفاقات</option>
            <option value="notes">الملاحظات وبنود المتابعة</option>
            <option value="paid_total">إجمالي المسدد حسب الزبون</option>
            <option value="payments">تفاصيل الدفعات</option>
            <option value="blocked_balances">أرصدة المحظورين من التواصل</option>
            <option value="missing_phone">الزبائن بدون أرقام جوال صالحة</option>
            <option value="payment_classification">تصنيف الزبائن حسب الدفع</option>
          </select></label>
          <div className="grid gap-4 sm:grid-cols-2">
            <fieldset className="space-y-2 rounded-2xl border border-slate-200 p-4" disabled={reportType === "payment_classification" || reportType === "missing_phone"}>
              <legend className="px-1 text-xs font-bold text-slate-700">الفترة الزمنية</legend>
              <label className="block space-y-1 text-xs font-semibold text-slate-600"><span>من تاريخ</span><input type="date" value={paymentDateFrom} onChange={event => setPaymentDateFrom(event.target.value)} className="w-full rounded-lg border border-slate-200 px-3 py-2.5 text-sm font-normal" /></label>
              <label className="block space-y-1 text-xs font-semibold text-slate-600"><span>إلى تاريخ</span><input type="date" value={paymentDateTo} onChange={event => setPaymentDateTo(event.target.value)} className="w-full rounded-lg border border-slate-200 px-3 py-2.5 text-sm font-normal" /></label>
            </fieldset>
            <fieldset className="space-y-2 rounded-2xl border border-slate-200 p-4" disabled={reportType === "notes" || reportType === "payment_classification" || reportType === "missing_phone"}>
              <legend className="px-1 text-xs font-bold text-slate-700">نطاق المبلغ</legend>
              <label className="block space-y-1 text-xs font-semibold text-slate-600"><span>من مبلغ</span><input type="number" min="0" step="0.01" value={reportAmountFrom} onChange={event => setReportAmountFrom(event.target.value)} placeholder="بلا حد أدنى" className="w-full rounded-lg border border-slate-200 px-3 py-2.5 text-sm font-normal disabled:bg-slate-100" /></label>
              <label className="block space-y-1 text-xs font-semibold text-slate-600"><span>إلى مبلغ</span><input type="number" min="0" step="0.01" value={reportAmountTo} onChange={event => setReportAmountTo(event.target.value)} placeholder="بلا حد أعلى" className="w-full rounded-lg border border-slate-200 px-3 py-2.5 text-sm font-normal disabled:bg-slate-100" /></label>
            </fieldset>
          </div>
          {reportType === "notes" && <p className="rounded-xl bg-amber-50 px-3 py-2 text-xs text-amber-800">تصفية المبلغ غير متاحة لتقرير الملاحظات لأنه لا يحتوي مبالغ مالية. تُفلتر البنود ذات المواعيد حسب الفترة المحددة.</p>}
          {reportType === "missing_phone" && <p className="rounded-xl bg-amber-50 px-3 py-2 text-xs text-amber-800">يعرض التقرير جميع الزبائن الذين لا يحتوي رقمهم على رقم جوال صالح لفتح واتساب، ويمكن تنزيل القائمة بصيغتي Excel وPDF.</p>}
          {reportType === "payment_classification" && <p className="rounded-xl bg-cyan-50 px-3 py-2 text-xs text-cyan-800">يشمل التقرير جميع الزبائن المحفوظين، ويضع كل تصنيف يدوي في جدول مستقل، بما في ذلك غير المصنفين؛ ويعرض بجانبه تصنيف النظام الآلي بلونه للمقارنة.</p>}
          {!paymentDateRangeValid && <p role="alert" className="text-xs font-semibold text-rose-700">تاريخ البداية يجب أن يسبق تاريخ النهاية.</p>}
          {paymentDateRangeValid && !reportFiltersValid && <p role="alert" className="text-xs font-semibold text-rose-700">تأكد أن مبلغ البداية لا يتجاوز مبلغ النهاية.</p>}
          {selectedReportPreview && <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl bg-slate-50 px-4 py-3 text-sm"><span className="font-semibold">{selectedReportPreview.title}</span><span>{selectedReportPreview.rows.length.toLocaleString("ar")} سجل{selectedReportPreview.amountColumn >= 0 ? ` · ${selectedReportPreview.totalLabel} ${formatAmount(selectedReportPreview.total)}` : ""}</span>{selectedReportPreview.extraTotals.map(total => <span key={total.label} className="text-xs text-slate-600">{total.label}: {formatAmount(total.amount)}</span>)}</div>}
          <div className="flex flex-wrap justify-end gap-2 border-t border-slate-100 pt-4">
            <button type="button" onClick={() => { setPaymentDateFrom(""); setPaymentDateTo(""); setReportAmountFrom(""); setReportAmountTo("") }} className="rounded-xl border border-slate-200 px-4 py-2.5 text-sm font-semibold text-slate-600 hover:bg-slate-50">مسح الحدود</button>
            <button type="button" onClick={() => setReportOpen(false)} className="rounded-xl border border-slate-200 px-4 py-2.5 text-sm font-semibold text-slate-600 hover:bg-slate-50">إغلاق</button>
            <button type="button" disabled={!reportFiltersValid || !selectedReportPreview?.rows.length} onClick={exportSelectedReportExcel} className="rounded-xl bg-cyan-700 px-4 py-2.5 text-sm font-semibold text-white hover:bg-cyan-800 disabled:opacity-40">تنزيل Excel</button>
            <button type="button" disabled={!reportFiltersValid || !selectedReportPreview?.rows.length} onClick={exportSelectedReportPdf} className="rounded-xl bg-rose-700 px-4 py-2.5 text-sm font-semibold text-white hover:bg-rose-800 disabled:opacity-40">تنزيل PDF</button>
          </div>
        </section>
      </div>}
    </main>
  )
}
