import { useEffect, useMemo, useState, type ReactNode } from "react";
import { ArrowLeft, Check, ChevronLeft, ChevronRight, Plus, Search, Star, Trash2, WalletCards, X } from "lucide-react";

export interface ToolBankRow {
  id: number;
  date: string;
  description: string;
  rawAmount: number;
  type: "مدفوع" | "مستلم";
  accountType: string;
  defaultRef?: string;
  _fileSessionId?: number;
}

export interface ToolCashierRow {
  id: number;
  name: string;
  rawName: string;
  amount: number;
  matchAmount: number;
  debit: number;
  credit: number;
  type: "مدفوع" | "مستلم";
  accountType: string;
  ref: string;
  defaultRef?: string;
  date: string;
  orig: Record<string, unknown>;
  _fileSessionId?: number;
}

export interface SplitAllocation {
  id: string;
  amount: number;
  kind: "company" | "external";
  bankId?: number;
  bankSessionId?: number;
  cashierId?: number;
  cashierSessionId?: number;
  note: string;
}

export interface SplitBankRef {
  id: number;
  sessionId: number;
}

export interface BankSplit {
  id: string;
  bankId: number;
  bankSessionId: number;
  sourceBanks?: SplitBankRef[];
  originalAmount: number;
  description: string;
  date: string;
  allocations: SplitAllocation[];
  savedAt: string;
}

export interface CutoffRowRef {
  id: number;
  sessionId: number;
}

export interface CutoffBatch {
  id: string;
  syntheticId: number;
  aggregateId: number;
  aggregateSessionId: number;
  sourceRows: CutoffRowRef[];
  dateFrom: string;
  dateTo: string;
  graceDays: number;
  total: number;
  note: string;
  savedAt: string;
}

export interface ClearingGroup {
  id: string;
  debitRows: CutoffRowRef[];
  creditRows: CutoffRowRef[];
  total: number;
  note: string;
  savedAt: string;
}

export function bankSplitSourceRefs(split: BankSplit): SplitBankRef[] {
  return split.sourceBanks?.length
    ? split.sourceBanks
    : [{ id: split.bankId, sessionId: split.bankSessionId }];
}

export function transactionKey(row: { id: number; _fileSessionId?: number }): string {
  return `${row._fileSessionId ?? 0}:${row.id}`;
}

export function syntheticTransactionId(key: string): number {
  let hash = 2166136261;
  for (let index = 0; index < key.length; index++) {
    hash ^= key.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return -1_000_000 - ((hash >>> 0) % 1_000_000_000);
}

function ReferenceButton({ row, onLocate }: {
  row: { id: number; _fileSessionId?: number; defaultRef?: string; rawAmount?: number };
  onLocate?: (row: { id: number; _fileSessionId?: number; defaultRef?: string; rawAmount?: number }) => void;
}) {
  if (!row.defaultRef || !onLocate) return null;
  return <button type="button" onClick={() => onLocate(row)} title="اعرض أماكن وجود هذه الحركة"
    className="rounded border border-indigo-200 bg-indigo-50 px-1.5 py-0.5 font-mono text-[10px] font-semibold text-indigo-700 hover:bg-indigo-100">
    {row.defaultRef}
  </button>;
}

function formatAmount(amount: number): string {
  return amount.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function normalizeDateDigits(value: string): string {
  return value
    .replace(/[٠-٩]/g, digit => String(digit.charCodeAt(0) - 0x660))
    .replace(/[۰-۹]/g, digit => String(digit.charCodeAt(0) - 0x6f0))
    .replace(/[\u200e\u200f\u202a-\u202e]/g, "")
    .trim();
}

export function toIsoDate(value: string): string | null {
  const normalized = normalizeDateDigits(value);
  const parts = normalized.match(/^(\d{1,4})[./-](\d{1,2})[./-](\d{1,4})/);
  if (parts) {
    const first = Number(parts[1]);
    const second = Number(parts[2]);
    const third = Number(parts[3]);
    const year = parts[1].length === 4 ? first : third;
    const month = second;
    const day = parts[1].length === 4 ? third : first;
    if (year >= 1000 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const date = new Date(Date.UTC(year, month - 1, day));
      if (date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day) {
        return `${year.toString().padStart(4, "0")}-${month.toString().padStart(2, "0")}-${day.toString().padStart(2, "0")}`;
      }
    }
  }
  const parsed = new Date(normalized);
  if (Number.isNaN(parsed.getTime())) return null;
  return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}-${String(parsed.getDate()).padStart(2, "0")}`;
}

function parseIsoDate(date: string): Date {
  const [year, month, day] = date.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function addDays(date: string, days: number): string {
  const next = parseIsoDate(date);
  next.setDate(next.getDate() + days);
  return `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, "0")}-${String(next.getDate()).padStart(2, "0")}`;
}

function todayIso(): string {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
}

function sourceKey(row: { id: number; _fileSessionId?: number }): string {
  return transactionKey(row);
}

function matchesSearch(fields: unknown[], search: string): boolean {
  const normalize = (value: unknown) => normalizeDateDigits(String(value ?? "")).toLowerCase().replace(/,/g, " ");
  return normalize(fields.join(" ")).includes(normalize(search.trim()));
}

function ToolFrame({ title, subtitle, onBack, children }: { title: string; subtitle: string; onBack: () => void; children: ReactNode }) {
  return (
    <main dir="rtl" className="min-h-full bg-background p-4 text-foreground sm:p-6">
      <div className="mx-auto max-w-6xl space-y-5">
        <header className="flex flex-wrap items-start justify-between gap-3 rounded-2xl border bg-card p-5">
          <div>
            <h1 className="text-xl font-bold">{title}</h1>
            <p className="mt-1 text-sm text-muted-foreground">{subtitle}</p>
          </div>
          <button type="button" onClick={onBack} className="inline-flex items-center gap-2 rounded-lg border px-3 py-2 text-sm hover:bg-muted">
            <ArrowLeft className="h-4 w-4" />العودة للتسوية
          </button>
        </header>
        {children}
      </div>
    </main>
  );
}

const panelClass = "space-y-4 rounded-2xl border bg-card p-4 sm:p-5";
const labelClass = "block space-y-1.5 text-xs font-medium text-muted-foreground";
const inputClass = "w-full rounded-lg border bg-background px-3 py-2 text-sm text-foreground";

export function SplitTool({
  bankRows,
  cashierRows,
  splits,
  unavailableBankKeys,
  unavailableCashierKeys,
  initialBankKeys,
  initialCashierKeys,
  accountTypeFilter,
  onSave,
  onDelete,
  onReview,
  onOpenReference,
  onBack,
}: {
  bankRows: ToolBankRow[];
  cashierRows: ToolCashierRow[];
  splits: BankSplit[];
  unavailableBankKeys: Set<string>;
  unavailableCashierKeys: Set<string>;
  initialBankKeys?: string[];
  initialCashierKeys?: string[];
  accountTypeFilter?: string;
  onSave: (split: BankSplit) => boolean;
  onDelete: (split: BankSplit) => void;
  onReview?: (cashier?: ToolCashierRow, bank?: ToolBankRow, detail?: string) => void;
  onOpenReference?: (row: { id: number; _fileSessionId?: number; defaultRef?: string; rawAmount?: number }) => void;
  onBack: () => void;
}) {
  const availableBanks = bankRows.filter(row =>
    !unavailableBankKeys.has(sourceKey(row)) &&
    !splits.some(split => bankSplitSourceRefs(split).some(reference =>
      reference.id === row.id && reference.sessionId === (row._fileSessionId ?? 0)
    ))
  );
  const availableCashiers = cashierRows.filter(row => !unavailableCashierKeys.has(sourceKey(row)));
  const [bankSearch, setBankSearch] = useState("");
  const [cashierSearch, setCashierSearch] = useState("");
  const [selectedBankKeys, setSelectedBankKeys] = useState<Set<string>>(() => new Set(initialBankKeys ?? []));
  const [bankPickerOpen, setBankPickerOpen] = useState(false);
  const [cashierPickerPartId, setCashierPickerPartId] = useState<string | null>(null);
  const [bankPage, setBankPage] = useState(1);
  const [cashierPage, setCashierPage] = useState(1);
  const pickerPageSize = 50;
  const presetBanks = availableBanks.filter(row => initialBankKeys?.includes(sourceKey(row)));
  const presetCashiers = availableCashiers.filter(row => initialCashierKeys?.includes(sourceKey(row)));
  const platformBanks = availableBanks.filter(row =>
    accountTypeFilter === undefined || row.accountType.trim() === accountTypeFilter
  );
  const platformCashiers = availableCashiers.filter(row =>
    accountTypeFilter === undefined || row.accountType.trim() === accountTypeFilter
  );
  const selectedBanks = availableBanks.filter(row => selectedBankKeys.has(sourceKey(row)));
  const bankOptions = platformBanks.filter(row =>
    matchesSearch([row.date, row.description, row.accountType, row.defaultRef, row.rawAmount, formatAmount(row.rawAmount)], bankSearch)
  );
  const cashierOptions = platformCashiers.filter(row =>
    (!selectedBanks.length || selectedBanks.some(bank => row.type === bank.type)) &&
    matchesSearch([row.date, row.name, row.rawName, row.ref, row.defaultRef, row.accountType, row.matchAmount, formatAmount(row.matchAmount)], cashierSearch)
  );
  const visibleBanks = bankOptions.slice((bankPage - 1) * pickerPageSize, bankPage * pickerPageSize);
  const visibleCashiers = cashierOptions.slice((cashierPage - 1) * pickerPageSize, cashierPage * pickerPageSize);
  const [parts, setParts] = useState<Array<{ id: string; bankKey: string; amount: string; kind: "company" | "external"; cashierKey: string; note: string }>>(() => {
    const makePart = (bankKey = "", cashierKey = "", amount = "") => ({
      id: `part-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      bankKey, amount, kind: "company" as const, cashierKey, note: "",
    });
    if (presetBanks.length === 1 && presetCashiers.length > 0) {
      return presetCashiers.map(cashier =>
        makePart(sourceKey(presetBanks[0]), sourceKey(cashier), String(cashier.matchAmount)),
      );
    }
    if (presetCashiers.length === 1 && presetBanks.length > 1) {
      return presetBanks.map(bank =>
        makePart(sourceKey(bank), sourceKey(presetCashiers[0]), String(bank.rawAmount)),
      );
    }
    return [makePart(
      presetBanks.length === 1 ? sourceKey(presetBanks[0]) : "",
      presetCashiers.length === 1 ? sourceKey(presetCashiers[0]) : "",
    )];
  });
  const [formError, setFormError] = useState("");
  const allocatedByBank = new Map<string, number>();
  parts.forEach(part => allocatedByBank.set(part.bankKey, (allocatedByBank.get(part.bankKey) ?? 0) + (Number(part.amount) || 0)));
  const selectionBankType = selectedBanks[0]?.type;
  const selectionAccountType = selectedBanks[0]?.accountType.trim();
  const bankSelectionCompatible = selectedBanks.every(bank =>
    bank.type === selectionBankType && bank.accountType.trim() === selectionAccountType
  );
  const allocatedByInvoice = new Map<string, { row: ToolCashierRow; total: number }>();
  parts.forEach(part => {
    if (part.kind !== "company") return;
    const row = availableCashiers.find(candidate => sourceKey(candidate) === part.cashierKey);
    if (!row) return;
    const key = sourceKey(row);
    const current = allocatedByInvoice.get(key) ?? { row, total: 0 };
    current.total += Number(part.amount) || 0;
    allocatedByInvoice.set(key, current);
  });

  const updatePart = (id: string, patch: Partial<(typeof parts)[number]>) => {
    setParts(current => current.map(part => part.id === id ? { ...part, ...patch } : part));
  };

  const toggleBank = (row: ToolBankRow) => {
    const key = sourceKey(row);
    const next = new Set(selectedBankKeys);
    if (next.has(key)) {
      next.delete(key);
      setParts(currentParts => currentParts.map(part => part.bankKey === key ? { ...part, bankKey: "" } : part));
    } else {
      next.add(key);
      if (!selectedBankKeys.size) {
        setParts(currentParts => currentParts.map(part => part.bankKey ? part : { ...part, bankKey: key }));
      }
    }
    setSelectedBankKeys(next);
    setFormError("");
  };

  const save = () => {
    setFormError("");
    if (!selectedBanks.length) return setFormError("اختر حوالة واحدة على الأقل.");
    if (selectedBanks.length > 1 && !bankSelectionCompatible) {
      return setFormError("اختر حوالات من الاتجاه والحساب نفسيهما في المجموعة.");
    }
    if (parts.length < 2) return setFormError("أضف حصتين على الأقل لربط الحوالات بالفواتير.");
    if (parts.some(part => !Number.isFinite(Number(part.amount)) || Number(part.amount) <= 0)) {
      return setFormError("أدخل مبلغاً موجباً لكل حصة.");
    }
    if (parts.some(part => !selectedBankKeys.has(part.bankKey))) {
      return setFormError("اختر الحوالة المقابلة لكل حصة.");
    }
    for (const bank of selectedBanks) {
      const allocated = allocatedByBank.get(sourceKey(bank)) ?? 0;
      if (Math.abs(allocated - bank.rawAmount) > 0.01) {
        return setFormError(`مجموع حصص حوالة «${bank.description}» (${formatAmount(allocated)}) يجب أن يساوي قيمتها (${formatAmount(bank.rawAmount)}).`);
      }
    }
    const allocations: SplitAllocation[] = [];
    const allocatedByCashier = new Map<string, { row: ToolCashierRow; total: number }>();
    for (const part of parts) {
      const amount = Number(part.amount);
      const bank = selectedBanks.find(row => sourceKey(row) === part.bankKey);
      if (!bank) return setFormError("إحدى الحوالات المختارة لم تعد متاحة.");
      if (part.kind === "external") {
        if (!part.note.trim()) return setFormError("اكتب سبباً واضحاً لكل حصة خارج شركتكم.");
        allocations.push({ id: part.id, amount, kind: "external", bankId: bank.id, bankSessionId: bank._fileSessionId ?? 0, note: part.note.trim() });
        continue;
      }
      const cashier = availableCashiers.find(row => sourceKey(row) === part.cashierKey);
      if (!cashier) return setFormError("اربط كل حصة تخص شركتكم بفاتورة أو سند متاح.");
      if (cashier.type !== bank.type) return setFormError("اتجاه الحركة في الفاتورة لا يطابق اتجاه الحوالة.");
      if (cashier.accountType.trim() !== bank.accountType.trim()) return setFormError("يجب ربط الفاتورة بالحوالة التابعة للمنصة نفسها.");
      const key = sourceKey(cashier);
      const allocated = allocatedByCashier.get(key) ?? { row: cashier, total: 0 };
      allocated.total += amount;
      allocatedByCashier.set(key, allocated);
      allocations.push({
        id: part.id, amount, kind: "company", bankId: bank.id, bankSessionId: bank._fileSessionId ?? 0, cashierId: cashier.id,
        cashierSessionId: cashier._fileSessionId ?? 0, note: part.note.trim(),
      });
    }
    for (const { row, total: allocated } of allocatedByCashier.values()) {
      if (Math.abs(row.matchAmount - allocated) > 0.01) {
        return setFormError(`مجموع الحصص المخصصة إلى «${row.name}» (${formatAmount(allocated)}) يجب أن يساوي مبلغ السند (${formatAmount(row.matchAmount)}).`);
      }
    }
    if (!allocations.some(allocation => allocation.kind === "company")) {
      return setFormError("اربط حصة واحدة على الأقل بفاتورة تخص شركتكم.");
    }
    const id = `split-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const primaryBank = selectedBanks[0];
    const didSave = onSave({
      id, bankId: primaryBank.id, bankSessionId: primaryBank._fileSessionId ?? 0,
      sourceBanks: selectedBanks.map(bank => ({ id: bank.id, sessionId: bank._fileSessionId ?? 0 })),
      originalAmount: selectedBanks.reduce((sum, bank) => sum + bank.rawAmount, 0),
      description: selectedBanks.length === 1 ? primaryBank.description : `${selectedBanks.length} حوالات`,
      date: primaryBank.date, allocations, savedAt: new Date().toISOString(),
    });
    if (!didSave) {
      setFormError("لم تحفظ التجزئة لأن إحدى الحركات لم تعد متاحة. راجع البيانات وحاول مجدداً.");
      return;
    }
    setSelectedBankKeys(new Set());
    setBankSearch("");
    setBankPage(1);
    setParts([{ id: `part-${Date.now()}`, bankKey: "", amount: "", kind: "company", cashierKey: "", note: "" }]);
  };

  return (
    <ToolFrame title="تسوية عدة حوالات مع عدة فواتير" subtitle={`${accountTypeFilter ? `منصة ${accountTypeFilter} · ` : ""}اختر حوالة واحدة أو أكثر، ثم وزّع مبالغها على فاتورة واحدة أو أكثر. يمكن ربط فاتورة بعدة حوالات أو حوالة بعدة فواتير؛ لا تحفظ المجموعة إلا بعد تغطية قيمة كل حوالة وكل فاتورة بالكامل.`} onBack={onBack}>
      {presetBanks.length > 0 && presetCashiers.length > 0 && (
        <section className="rounded-xl border border-emerald-200 bg-emerald-50 p-4">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <h2 className="font-semibold text-emerald-900">تعبئة تلقائية من منصة {accountTypeFilter || presetBanks[0].accountType}</h2>
            <span className="rounded-full bg-white px-2.5 py-1 text-xs font-medium text-emerald-800">
              {presetBanks.length} حوالة · {presetCashiers.length} فاتورة
            </span>
          </div>
          <div className="mt-3 space-y-2">
            {parts.map((part, index) => {
              const bank = selectedBanks.find(row => sourceKey(row) === part.bankKey);
              const cashier = availableCashiers.find(row => sourceKey(row) === part.cashierKey);
              if (!bank && !cashier) return null;
              return (
                <div key={part.id} className="grid gap-1 rounded-lg border border-emerald-100 bg-white px-3 py-2 text-sm sm:grid-cols-[1fr_auto_1fr_auto] sm:items-center">
                  <span><b>الحوالة:</b> {bank?.description || "لم تُحدد"} · {bank ? formatAmount(bank.rawAmount) : "—"}</span>
                  <span className="hidden text-emerald-700 sm:inline">←</span>
                  <span><b>فاتورة مقابل الحوالة {presetBanks.length === 1 ? (index === 0 ? "الأولى" : `رقم ${index + 1}`) : `رقم ${index + 1}`}:</b> {cashier?.name || "لم تُحدد"}</span>
                  <b className="font-mono text-emerald-700">{formatAmount(Number(part.amount) || 0)}</b>
                </div>
              );
            })}
          </div>
        </section>
      )}
      <section className="grid gap-4 lg:grid-cols-[minmax(0,1fr)_minmax(320px,0.8fr)]">
        <div className={panelClass}>
          <h2 className="font-semibold">الحوالات الداخلة في التسوية</h2>
          <label className={labelClass}>اختر حوالة واحدة أو أكثر
            <button type="button" onClick={() => { setBankPickerOpen(true); setBankSearch(""); setBankPage(1); }} className={`${inputClass} flex items-center justify-between gap-3 text-right hover:border-primary`}>
              <span className={selectedBanks.length ? "text-foreground" : "text-muted-foreground"}>
                {selectedBanks.length ? `${selectedBanks.length} حوالة محددة · ${formatAmount(selectedBanks.reduce((sum, bank) => sum + bank.rawAmount, 0))}` : "اختر الحوالات"}
              </span>
              <Search className="h-4 w-4 shrink-0" />
            </button>
          </label>
          {platformBanks.length > 0 && <p className="-mt-3 text-[11px] text-muted-foreground">الحركات المتاحة في المنصة: {platformBanks.length}. اضغط للاختيار والبحث بالاسم أو التاريخ أو المبلغ.</p>}
          {platformBanks.length === 0 && <p className="-mt-3 text-[11px] text-amber-700">لا توجد حوالات متاحة للتسوية في هذه المنصة.</p>}
          {selectedBanks.length > 0 && <div className="space-y-2">
            {selectedBanks.map(bank => {
              const allocated = allocatedByBank.get(sourceKey(bank)) ?? 0;
              const difference = bank.rawAmount - allocated;
              return <div key={sourceKey(bank)} data-default-reference={bank.defaultRef} className="flex flex-wrap items-center justify-between gap-2 rounded-xl border border-blue-200 bg-blue-50 p-3 text-sm text-blue-950">
                <div><div className="flex flex-wrap items-center gap-2 font-semibold">{bank.description} <ReferenceButton row={bank} onLocate={onOpenReference} /></div>
                  <div className="mt-1 text-xs">{bank.date || "بلا تاريخ"} · {bank.type} · {bank.accountType || "بلا حساب"} · الأصل {formatAmount(bank.rawAmount)} · الموزع {formatAmount(allocated)} · المتبقي {formatAmount(difference)}</div></div>
                <button type="button" onClick={() => toggleBank(bank)} className="rounded-lg border border-blue-300 bg-white px-2.5 py-1.5 text-xs hover:bg-blue-100">إزالة</button>
              </div>;
            })}
          </div>}
          {selectedBanks.length > 1 && !bankSelectionCompatible && <p role="alert" className="rounded-lg border border-amber-200 bg-amber-50 p-3 text-xs text-amber-800">يجب أن تكون الحوالات من الاتجاه والحساب نفسيهما.</p>}
          <div className="space-y-3">
            {parts.map((part, index) => (
              <article key={part.id} className="space-y-3 rounded-xl border p-3">
                <div className="flex items-center justify-between gap-2">
                  <h3 className="text-sm font-semibold">الحصة {index + 1}</h3>
                  {parts.length > 1 && <button type="button" onClick={() => setParts(current => current.filter(item => item.id !== part.id))} className="rounded-lg p-1.5 text-red-600 hover:bg-red-50" aria-label={`حذف الحصة ${index + 1}`}><Trash2 className="h-4 w-4" /></button>}
                </div>
                <div className="grid gap-3 sm:grid-cols-3">
                  <label className={labelClass}>الحوالة المصدر
                    <select className={inputClass} value={part.bankKey} onChange={event => updatePart(part.id, { bankKey: event.target.value })}>
                      <option value="">اختر حوالة</option>
                      {selectedBanks.map(bank => <option key={sourceKey(bank)} value={sourceKey(bank)}>{bank.date || "بلا تاريخ"} · {bank.description} · {formatAmount(bank.rawAmount)}</option>)}
                    </select>
                  </label>
                  <label className={labelClass}>المبلغ
                    <input className={inputClass} type="number" min="0.01" step="0.01" value={part.amount} onChange={event => updatePart(part.id, { amount: event.target.value })} />
                  </label>
                  <label className={labelClass}>تخصيص الحصة
                    <select className={inputClass} value={part.kind} onChange={event => updatePart(part.id, { kind: event.target.value as "company" | "external", cashierKey: "" })}>
                      <option value="company">تخص شركتنا — اربطها بفاتورة</option>
                      <option value="external">تخص جهة / شركة أخرى</option>
                    </select>
                  </label>
                </div>
                {part.kind === "company" ? (
                  <div className="space-y-2">
                    <label className={labelClass}>{presetCashiers.length > 1 ? `اسم الخانة — فاتورة مقابل الحوالة ${index === 0 ? "الأولى" : `رقم ${index + 1}`}` : "اسم الخانة — الفاتورة أو السند المقابل"} (يجوز ربط الفاتورة بعدة حوالات)
                    <button type="button" onClick={() => { setCashierPickerPartId(part.id); setCashierSearch(""); setCashierPage(1); }} className={`${inputClass} flex items-center justify-between gap-3 text-right hover:border-primary`}>
                      <span className={part.cashierKey ? "text-foreground" : "text-muted-foreground"}>
                        {(() => {
                          const row = availableCashiers.find(item => sourceKey(item) === part.cashierKey);
                          return row ? `${row.date || "بلا تاريخ"} · ${row.name} · ${formatAmount(row.matchAmount)}` : "اختر فاتورة أو سنداً";
                        })()}
                      </span>
                      <Search className="h-4 w-4 shrink-0" />
                    </button>
                    </label>
                  </div>
                ) : (
                  <label className={labelClass}>سبب / جهة تخصيص المبلغ (سيظهر في التقرير)
                    <textarea className={`${inputClass} min-h-20 resize-y`} value={part.note} onChange={event => updatePart(part.id, { note: event.target.value })} placeholder="مثال: حصة الشركة الثانية من الحوالة المشتركة" />
                  </label>
                )}
                {part.kind === "company" && <label className={labelClass}>ملاحظة اختيارية
                  <input className={inputClass} value={part.note} onChange={event => updatePart(part.id, { note: event.target.value })} placeholder="تظهر في سجل التجزئة" />
                </label>}
              </article>
            ))}
          </div>
          {!!allocatedByInvoice.size && <div className="space-y-2 rounded-xl border bg-muted/20 p-3">
            <h3 className="text-xs font-semibold">إجمالي التوزيع على كل فاتورة</h3>
            {[...allocatedByInvoice.values()].map(({ row, total }) => {
              const difference = row.matchAmount - total;
              return <div key={sourceKey(row)} className="flex flex-wrap justify-between gap-2 text-xs">
                <span>{row.name} · الفاتورة {formatAmount(row.matchAmount)}</span>
                <span className={Math.abs(difference) <= 0.01 ? "font-semibold text-green-700" : "font-semibold text-amber-700"}>
                  موزع {formatAmount(total)} · {Math.abs(difference) <= 0.01 ? "متطابق" : difference > 0 ? `متبقي ${formatAmount(difference)}` : `زيادة ${formatAmount(-difference)}`}
                </span>
              </div>;
            })}
          </div>}
          <div className="flex flex-wrap gap-2">
            <button type="button" onClick={() => setParts(current => [...current, { id: `part-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`, bankKey: selectedBanks[0] ? sourceKey(selectedBanks[0]) : "", amount: "", kind: "company", cashierKey: "", note: "" }])} className="inline-flex items-center gap-1.5 rounded-lg border border-blue-200 px-3 py-2 text-xs font-semibold text-blue-700 hover:bg-blue-50">
              <Plus className="h-4 w-4" />إضافة حصة
            </button>
            <button type="button" onClick={save} className="inline-flex items-center gap-1.5 rounded-lg bg-blue-600 px-4 py-2 text-xs font-semibold text-white hover:bg-blue-700">
              <Check className="h-4 w-4" />حفظ التسوية
            </button>
          </div>
          {formError && <p role="alert" className="rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-700">{formError}</p>}
          <p className="text-xs leading-6 text-muted-foreground">يجب توزيع كامل مبلغ كل حوالة، ومساواة إجمالي التوزيع على كل فاتورة لقيمتها، مع تطابق اتجاه وحساب الحوالات. يمكن ربط الفاتورة بعدة حوالات أو توزيع الحوالة على عدة فواتير. لا يتم إنشاء قيد في البرنامج المحاسبي.</p>
        </div>

        <div className={panelClass}>
          <div className="flex items-center gap-2"><WalletCards className="h-5 w-5 text-blue-600" /><h2 className="font-semibold">التجزئات المحفوظة</h2></div>
          {!splits.length && <p className="rounded-xl bg-muted/40 p-4 text-sm text-muted-foreground">لم تُجزّأ أي حوالة بعد.</p>}
          {splits.map(split => (
            <article key={split.id} data-default-reference={bankSplitSourceRefs(split).map(reference => bankRows.find(row => row.id === reference.id && (row._fileSessionId ?? 0) === reference.sessionId)?.defaultRef).filter(Boolean).join("|")}
              className="space-y-2 rounded-xl border p-3">
              <div className="flex items-start justify-between gap-2">
                <div><p className="font-semibold">{split.description}</p><p className="mt-1 flex flex-wrap items-center gap-2 text-xs text-muted-foreground">{bankSplitSourceRefs(split).length} حوالات · الإجمالي {formatAmount(split.originalAmount)}</p></div>
                <button type="button" onClick={() => onDelete(split)} className="rounded-lg p-1.5 text-red-600 hover:bg-red-50" title="إلغاء التجزئة"><Trash2 className="h-4 w-4" /></button>
              </div>
              <div className="space-y-1.5">
                {split.allocations.map((allocation, index) => {
                  const row = allocation.kind === "company" && allocation.cashierId != null
                    ? cashierRows.find(candidate => candidate.id === allocation.cashierId && (candidate._fileSessionId ?? 0) === (allocation.cashierSessionId ?? 0))
                    : undefined;
                  const sourceRef = { id: allocation.bankId ?? split.bankId, sessionId: allocation.bankSessionId ?? split.bankSessionId };
                  const sourceBank = bankRows.find(bank => bank.id === sourceRef.id && (bank._fileSessionId ?? 0) === sourceRef.sessionId);
                  return <div key={allocation.id} data-default-reference={row?.defaultRef} className="flex flex-wrap justify-between gap-2 rounded-lg bg-muted/40 px-3 py-2 text-xs">
                    <span className="flex flex-wrap items-center gap-2">{index + 1}. {sourceBank?.description || "حوالة"} → {allocation.kind === "company" ? row?.name || "فاتورة الشركة" : `حصة خارجية · ${allocation.note}`} {sourceBank && <ReferenceButton row={sourceBank} onLocate={onOpenReference} />} {row && <ReferenceButton row={row} onLocate={onOpenReference} />}</span>
                    <b className="font-mono">{formatAmount(allocation.amount)}</b>
                  </div>;
                })}
              </div>
            </article>
          ))}
        </div>
      </section>
      {(bankPickerOpen || cashierPickerPartId) && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/50 p-3 sm:p-6" onMouseDown={event => {
          if (event.target === event.currentTarget) { setBankPickerOpen(false); setCashierPickerPartId(null); }
        }}>
          <section role="dialog" aria-modal="true" aria-labelledby="split-picker-title" className="flex max-h-[90vh] w-full max-w-5xl flex-col overflow-hidden rounded-2xl border bg-card shadow-2xl">
            <header className="flex items-start justify-between gap-3 border-b p-4 sm:p-5">
              <div>
                <h2 id="split-picker-title" className="text-lg font-bold">{bankPickerOpen ? "اختيار الحوالات" : "اختيار الفاتورة أو السند"}</h2>
                <p className="mt-1 text-xs text-muted-foreground">{bankPickerOpen ? "يمكن تحديد أكثر من حوالة. حدد الحركات التي تريد تسويتها معاً ثم أغلق النافذة." : "ابحث بالاسم أو البيان أو التاريخ أو المبلغ، ثم اختر الحركة المطلوبة."}</p>
              </div>
              <button type="button" onClick={() => { setBankPickerOpen(false); setCashierPickerPartId(null); }} className="rounded-lg p-2 hover:bg-muted" aria-label="إغلاق"><X className="h-5 w-5" /></button>
            </header>
            <div className="border-b p-4">
              <div className="relative">
                <Search className="absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
                <input autoFocus className={`${inputClass} pr-10`} value={bankPickerOpen ? bankSearch : cashierSearch} onChange={event => {
                  if (bankPickerOpen) { setBankSearch(event.target.value); setBankPage(1); }
                  else { setCashierSearch(event.target.value); setCashierPage(1); }
                }} placeholder={bankPickerOpen ? "ابحث في الحوالات بالاسم أو التاريخ أو المبلغ..." : "ابحث في الفواتير والسندات بالاسم أو التاريخ أو المبلغ..."} />
              </div>
              <p className="mt-2 text-xs text-muted-foreground">{bankPickerOpen ? bankOptions.length : cashierOptions.length} حركة متاحة للعرض</p>
            </div>
            <div className="min-h-0 flex-1 overflow-auto p-3 sm:p-4">
              <div className="overflow-x-auto rounded-xl border">
                <table className="w-full min-w-[720px] text-sm">
                  <thead className="sticky top-0 bg-muted text-xs"><tr>
                    <th className="px-3 py-3 text-right">التاريخ</th><th className="px-3 py-3 text-right">البيان / اسم السند</th><th className="px-3 py-3 text-right">المرجع</th><th className="px-3 py-3 text-right">الحساب</th><th className="px-3 py-3 text-right">الاتجاه</th><th className="px-3 py-3 text-right">المبلغ</th><th className="px-3 py-3 text-center">مراجعة / اختيار</th>
                  </tr></thead>
                  <tbody>
                    {bankPickerOpen ? visibleBanks.map(row => <tr key={sourceKey(row)} data-default-reference={row.defaultRef} className="border-t hover:bg-muted/30">
                      <td className="px-3 py-2.5">{row.date || "—"}</td><td className="px-3 py-2.5 font-medium">{row.description || "حركة بلا بيان"} <ReferenceButton row={row} onLocate={onOpenReference} /></td><td className="px-3 py-2.5">—</td><td className="px-3 py-2.5 text-xs text-muted-foreground">{row.accountType || "—"}</td><td className="px-3 py-2.5">{row.type}</td><td className="px-3 py-2.5 font-mono">{formatAmount(row.rawAmount)}</td>
                      <td className="px-3 py-2.5 text-center"><div className="inline-flex items-center gap-1"><button type="button" onClick={() => onReview?.(undefined, row, "حوالة من أداة تسوية الحوالات والفواتير")} title="إضافة الحوالة للمراجعة" aria-label="مراجعة الحوالة" className="rounded-lg border p-1.5 text-amber-700 hover:bg-amber-50"><Star className="h-4 w-4" /></button><button type="button" onClick={() => toggleBank(row)} className={`rounded-lg px-3 py-1.5 text-xs font-semibold ${selectedBankKeys.has(sourceKey(row)) ? "bg-blue-100 text-blue-800" : "bg-primary text-primary-foreground hover:opacity-90"}`}>{selectedBankKeys.has(sourceKey(row)) ? "محدد ✓" : "تحديد"}</button></div></td>
                    </tr>) : visibleCashiers.map(row => <tr key={sourceKey(row)} data-default-reference={row.defaultRef} className="border-t hover:bg-muted/30">
                      <td className="px-3 py-2.5">{row.date || "—"}</td><td className="px-3 py-2.5 font-medium">{row.name || row.rawName || "سند بلا اسم"} <ReferenceButton row={row} onLocate={onOpenReference} /></td><td className="px-3 py-2.5">{row.ref || "—"}</td><td className="px-3 py-2.5 text-xs text-muted-foreground">{row.accountType || "—"}</td><td className="px-3 py-2.5">{row.type}</td><td className="px-3 py-2.5 font-mono">{formatAmount(row.matchAmount)}</td>
                      <td className="px-3 py-2.5 text-center"><div className="inline-flex items-center gap-1"><button type="button" onClick={() => {
                        const bank = selectedBanks[0];
                        onReview?.(row, bank, bank ? "فاتورة وحوالة مرشحتان للتسوية" : "فاتورة من أداة تسوية الحوالات والفواتير");
                      }} title="إضافة الفاتورة للمراجعة" aria-label="مراجعة الفاتورة" className="rounded-lg border p-1.5 text-amber-700 hover:bg-amber-50"><Star className="h-4 w-4" /></button><button type="button" onClick={() => {
                        if (cashierPickerPartId) updatePart(cashierPickerPartId, { cashierKey: sourceKey(row) });
                        setCashierPickerPartId(null);
                      }} className="rounded-lg bg-primary px-3 py-1.5 text-xs font-semibold text-primary-foreground hover:opacity-90">اختيار</button></div></td>
                    </tr>)}
                    {((bankPickerOpen && !visibleBanks.length) || (cashierPickerPartId && !visibleCashiers.length)) && <tr><td colSpan={7} className="py-12 text-center text-sm text-muted-foreground">لا توجد نتائج مطابقة لبحثك.</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
            {((bankPickerOpen && bankOptions.length > pickerPageSize) || (cashierPickerPartId && cashierOptions.length > pickerPageSize)) && <footer className="flex items-center justify-between border-t p-3 text-xs text-muted-foreground">
              <span>صفحة {bankPickerOpen ? bankPage : cashierPage} من {Math.ceil((bankPickerOpen ? bankOptions.length : cashierOptions.length) / pickerPageSize)}</span>
              <div className="flex gap-2">
                <button type="button" disabled={(bankPickerOpen ? bankPage : cashierPage) <= 1} onClick={() => bankPickerOpen ? setBankPage(page => Math.max(1, page - 1)) : setCashierPage(page => Math.max(1, page - 1))} className="inline-flex items-center gap-1 rounded-lg border px-3 py-1.5 disabled:opacity-40"><ChevronRight className="h-4 w-4" />السابق</button>
                <button type="button" disabled={(bankPickerOpen ? bankPage : cashierPage) >= Math.ceil((bankPickerOpen ? bankOptions.length : cashierOptions.length) / pickerPageSize)} onClick={() => bankPickerOpen ? setBankPage(page => Math.min(Math.ceil(bankOptions.length / pickerPageSize), page + 1)) : setCashierPage(page => Math.min(Math.ceil(cashierOptions.length / pickerPageSize), page + 1))} className="inline-flex items-center gap-1 rounded-lg border px-3 py-1.5 disabled:opacity-40">التالي<ChevronLeft className="h-4 w-4" /></button>
              </div>
            </footer>}
            {bankPickerOpen && <footer className="flex items-center justify-between gap-3 border-t p-3 text-xs">
              <span>الحوالات المحددة: {selectedBanks.length} · الإجمالي: {formatAmount(selectedBanks.reduce((sum, bank) => sum + bank.rawAmount, 0))}</span>
              <button type="button" onClick={() => setBankPickerOpen(false)} className="rounded-lg bg-primary px-4 py-2 font-semibold text-primary-foreground">اعتماد الاختيار</button>
            </footer>}
          </section>
        </div>
      )}
    </ToolFrame>
  );
}

export function ClearingTool({
  cashierRows,
  groups,
  unavailableCashierKeys,
  onSave,
  onDelete,
  onReview,
  onOpenReference,
  onBack,
}: {
  cashierRows: ToolCashierRow[];
  groups: ClearingGroup[];
  unavailableCashierKeys: Set<string>;
  onSave: (group: ClearingGroup) => boolean;
  onDelete: (group: ClearingGroup) => void;
  onReview?: (cashier?: ToolCashierRow, detail?: string) => void;
  onOpenReference?: (row: { id: number; _fileSessionId?: number; defaultRef?: string; rawAmount?: number }) => void;
  onBack: () => void;
}) {
  const [debitSearch, setDebitSearch] = useState("");
  const [creditSearch, setCreditSearch] = useState("");
  const [debitKeys, setDebitKeys] = useState<Set<string>>(new Set());
  const [creditKeys, setCreditKeys] = useState<Set<string>>(new Set());
  const [note, setNote] = useState("");
  const [formError, setFormError] = useState("");
  const availableRows = cashierRows.filter(row => !unavailableCashierKeys.has(sourceKey(row)));
  const debitRows = availableRows.filter(row => row.debit > 0 && row.credit <= 0);
  const creditRows = availableRows.filter(row => row.credit > 0 && row.debit <= 0);
  const filterRows = (rows: ToolCashierRow[], search: string) => rows
    .filter(row => !search.trim() || `${row.date} ${row.name} ${row.rawName} ${row.ref} ${row.debit} ${row.credit}`.toLowerCase().includes(search.trim().toLowerCase()))
    .slice(0, 200);
  const visibleDebits = filterRows(debitRows, debitSearch);
  const visibleCredits = filterRows(creditRows, creditSearch);
  const selectedDebits = debitRows.filter(row => debitKeys.has(sourceKey(row)));
  const selectedCredits = creditRows.filter(row => creditKeys.has(sourceKey(row)));
  const debitTotal = selectedDebits.reduce((sum, row) => sum + row.debit, 0);
  const creditTotal = selectedCredits.reduce((sum, row) => sum + row.credit, 0);
  const accountTypes = new Set([...selectedDebits, ...selectedCredits].map(row => row.accountType.trim()).filter(Boolean));
  const sameAccountType = accountTypes.size <= 1;
  const exact = selectedDebits.length > 0 && selectedCredits.length > 0 && sameAccountType && Math.abs(debitTotal - creditTotal) <= 0.01;

  const toggleRow = (key: string, checked: boolean, side: "debit" | "credit") => {
    const setter = side === "debit" ? setDebitKeys : setCreditKeys;
    setter(current => {
      const next = new Set(current);
      if (checked) next.add(key); else next.delete(key);
      return next;
    });
    setFormError("");
  };

  const save = () => {
    setFormError("");
    if (!selectedDebits.length || !selectedCredits.length) {
      setFormError("حدد حركة مدينة وحركة دائنة واحدة على الأقل.");
      return;
    }
    if (!exact) {
      setFormError(sameAccountType
        ? "لا يمكن التسكير إلا إذا تساوى مجموع الحركات المدينة والدائنة."
        : "لا يمكن جمع حركات من حسابات أو منصات مختلفة في مجموعة تسكير واحدة.");
      return;
    }
    const id = `clearing-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const saved = onSave({
      id,
      debitRows: selectedDebits.map(row => ({ id: row.id, sessionId: row._fileSessionId ?? 0 })),
      creditRows: selectedCredits.map(row => ({ id: row.id, sessionId: row._fileSessionId ?? 0 })),
      total: debitTotal,
      note: note.trim(),
      savedAt: new Date().toISOString(),
    });
    if (!saved) {
      setFormError("لم تحفظ المجموعة لأن إحدى الحركات لم تعد متاحة. حدّث الاختيارات ثم حاول مجدداً.");
      return;
    }
    setDebitKeys(new Set());
    setCreditKeys(new Set());
    setNote("");
  };

  const renderRows = (rows: ToolCashierRow[], selected: Set<string>, side: "debit" | "credit") => (
    <div className="max-h-[440px] overflow-auto rounded-xl border">
      <table className="w-full min-w-[520px] text-sm">
        <thead className="sticky top-0 bg-muted text-xs"><tr>
          <th className="w-10 px-3 py-2 text-center">اختيار</th>
          <th className="px-3 py-2 text-right">التاريخ</th>
          <th className="px-3 py-2 text-right">السند / البيان</th>
          <th className="px-3 py-2 text-right">الحساب</th>
          <th className="px-3 py-2 text-right">المبلغ</th>
          <th className="px-3 py-2 text-center">مراجعة</th>
        </tr></thead>
        <tbody>
          {rows.map(row => {
            const key = sourceKey(row);
            return <tr key={key} data-default-reference={row.defaultRef} className="border-t hover:bg-muted/30">
              <td className="px-3 py-2 text-center"><input type="checkbox" checked={selected.has(key)} onChange={event => toggleRow(key, event.target.checked, side)} /></td>
              <td className="px-3 py-2">{row.date || "—"}</td>
              <td className="px-3 py-2"><span className="font-medium">{row.name}</span>{row.ref && <span className="mr-2 text-xs text-muted-foreground">#{row.ref}</span>} <ReferenceButton row={row} onLocate={onOpenReference} /></td>
              <td className="px-3 py-2 text-xs text-muted-foreground">{row.accountType || "—"}</td>
              <td className="px-3 py-2 font-mono">{formatAmount(side === "debit" ? row.debit : row.credit)}</td>
              <td className="px-3 py-2 text-center"><button type="button" onClick={() => onReview?.(row, `حركة من شاشة تسكير السندات · ${side === "debit" ? "مدين" : "دائن"}`)} title="إضافة الحركة إلى المراجعة" aria-label={`مراجعة ${row.name}`} className="rounded-lg border p-1.5 text-amber-700 hover:bg-amber-50"><Star className="h-4 w-4" /></button></td>
            </tr>;
          })}
          {!rows.length && <tr><td colSpan={6} className="py-8 text-center text-xs text-muted-foreground">لا توجد حركات متاحة في هذه القائمة.</td></tr>}
        </tbody>
      </table>
    </div>
  );

  return (
    <ToolFrame title="شاشة تسكير السندات" subtitle="سكّر حركات الكشف المستورد يدوياً دون إنشاء قيود. اجمع سنداً أو أكثر في كل طرف، ولا تحفظ إلا عند تساوي مجموع المدين والدائن." onBack={onBack}>
      <section className="space-y-4">
        <div className="grid gap-4 xl:grid-cols-2">
          <div className={panelClass}>
            <div className="flex items-center justify-between gap-2">
              <div><h2 className="font-semibold text-green-700">الحركات المدينة</h2><p className="text-xs text-muted-foreground">المحدد: {selectedDebits.length} · {formatAmount(debitTotal)}</p></div>
              <label className="text-xs text-muted-foreground">بحث
                <input className={`${inputClass} mt-1 min-w-48`} value={debitSearch} onChange={event => setDebitSearch(event.target.value)} placeholder="الاسم أو التاريخ أو المبلغ" />
              </label>
            </div>
            {debitRows.length > 200 && <p className="text-xs text-muted-foreground">تُعرض أول 200 نتيجة؛ استخدم البحث لتحديد الحركات.</p>}
            {renderRows(visibleDebits, debitKeys, "debit")}
          </div>
          <div className={panelClass}>
            <div className="flex items-center justify-between gap-2">
              <div><h2 className="font-semibold text-red-700">الحركات الدائنة</h2><p className="text-xs text-muted-foreground">المحدد: {selectedCredits.length} · {formatAmount(creditTotal)}</p></div>
              <label className="text-xs text-muted-foreground">بحث
                <input className={`${inputClass} mt-1 min-w-48`} value={creditSearch} onChange={event => setCreditSearch(event.target.value)} placeholder="الاسم أو التاريخ أو المبلغ" />
              </label>
            </div>
            {creditRows.length > 200 && <p className="text-xs text-muted-foreground">تُعرض أول 200 نتيجة؛ استخدم البحث لتحديد الحركات.</p>}
            {renderRows(visibleCredits, creditKeys, "credit")}
          </div>
        </div>
        <div className={`${panelClass} space-y-3`}>
          <div className="flex flex-wrap items-center justify-between gap-3 rounded-xl border bg-muted/20 p-4">
            <span>المدين: <b className="font-mono text-green-700">{formatAmount(debitTotal)}</b></span>
            <span>الدائن: <b className="font-mono text-red-700">{formatAmount(creditTotal)}</b></span>
            <span className={`font-semibold ${exact ? "text-green-700" : "text-amber-700"}`}>{!sameAccountType ? "اختر حركات من الحساب نفسه" : exact ? "متساويان — جاهز للتسكير" : `الفرق: ${formatAmount(debitTotal - creditTotal)}`}</span>
          </div>
          <label className={labelClass}>ملاحظة تدقيقية اختيارية
            <input className={inputClass} value={note} onChange={event => setNote(event.target.value)} placeholder="سبب التسكير أو مرجع المراجعة" />
          </label>
          <div className="flex flex-wrap gap-2">
            <button type="button" onClick={save} disabled={!exact} className="inline-flex items-center gap-1.5 rounded-lg bg-teal-600 px-4 py-2.5 text-sm font-semibold text-white hover:bg-teal-700 disabled:cursor-not-allowed disabled:opacity-40"><Check className="h-4 w-4" />تسكير الحركات المحددة</button>
            <button type="button" onClick={() => { setDebitKeys(new Set()); setCreditKeys(new Set()); setFormError(""); }} className="inline-flex items-center gap-1.5 rounded-lg border px-4 py-2.5 text-sm hover:bg-muted"><X className="h-4 w-4" />مسح التحديد</button>
          </div>
          {formError && <p role="alert" className="rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-700">{formError}</p>}
          <p className="text-xs leading-5 text-muted-foreground">تختفي الحركات المسكّرة من قائمة المطابقة المفتوحة وتبقى تفاصيلها في سجل المجموعات والتقرير. لا يتغير ملف المصدر ولا يُنشأ أي قيد في النظام المحاسبي.</p>
        </div>
        <div className={panelClass}>
          <h2 className="font-semibold">سجل الحركات المسكّرة ({groups.length})</h2>
          {!groups.length && <p className="rounded-xl bg-muted/40 p-4 text-sm text-muted-foreground">لا توجد مجموعات مسكّرة.</p>}
          {groups.map(group => {
            const lookup = (refs: CutoffRowRef[]) => refs.map(ref => cashierRows.find(row => row.id === ref.id && (row._fileSessionId ?? 0) === ref.sessionId)).filter((row): row is ToolCashierRow => !!row);
            const debits = lookup(group.debitRows);
            const credits = lookup(group.creditRows);
            return <article key={group.id} className="space-y-2 rounded-xl border p-3">
              <div className="flex items-start justify-between gap-3">
                <div><p className="font-semibold">مجموعة تسكير · {formatAmount(group.total)}</p><p className="text-xs text-muted-foreground">{debits.length} مدين مقابل {credits.length} دائن · {new Date(group.savedAt).toLocaleString()}</p></div>
                <button type="button" onClick={() => onDelete(group)} className="rounded-lg p-1.5 text-red-600 hover:bg-red-50" title="إلغاء التسكير"><Trash2 className="h-4 w-4" /></button>
              </div>
              <div className="grid gap-2 md:grid-cols-2">
                {[{ label: "مدين", rows: debits }, { label: "دائن", rows: credits }].map(side => <div key={side.label} className="space-y-1">
                  <p className="text-xs font-semibold">{side.label}</p>
                  {side.rows.map(row => <div key={sourceKey(row)} data-default-reference={row.defaultRef} className="flex justify-between gap-2 rounded bg-muted/40 px-2 py-1.5 text-xs"><span>{row.date} · {row.name} <ReferenceButton row={row} onLocate={onOpenReference} /></span><b className="font-mono">{formatAmount(side.label === "مدين" ? row.debit : row.credit)}</b></div>)}
                </div>)}
              </div>
              {group.note && <p className="text-xs text-muted-foreground">{group.note}</p>}
            </article>;
          })}
        </div>
      </section>
    </ToolFrame>
  );
}

export function BatchCutoffTool({
  cashierRows,
  batches,
  unavailableCashierKeys,
  onSave,
  onDelete,
  onOpenMultiMatch,
  onReview,
  onOpenReference,
  onBack,
}: {
  cashierRows: ToolCashierRow[];
  batches: CutoffBatch[];
  unavailableCashierKeys: Set<string>;
  onSave: (batch: CutoffBatch) => boolean;
  onDelete: (batch: CutoffBatch) => void;
  onOpenMultiMatch: () => void;
  onReview?: (cashier?: ToolCashierRow, detail?: string) => void;
  onOpenReference?: (row: { id: number; _fileSessionId?: number; defaultRef?: string; rawAmount?: number }) => void;
  onBack: () => void;
}) {
  const [dateFrom, setDateFrom] = useState(todayIso);
  const [dateTo, setDateTo] = useState(todayIso);
  const [graceDays, setGraceDays] = useState(0);
  const [aggregateKey, setAggregateKey] = useState("");
  const [aggregateSearch, setAggregateSearch] = useState("");
  const [aggregatePickerOpen, setAggregatePickerOpen] = useState(false);
  const [aggregatePage, setAggregatePage] = useState(1);
  const [candidateSearch, setCandidateSearch] = useState("");
  const [candidatePage, setCandidatePage] = useState(1);
  const [selectedKeys, setSelectedKeys] = useState<Set<string>>(new Set());
  const [selectionTouched, setSelectionTouched] = useState(false);
  const [note, setNote] = useState("");
  const [formError, setFormError] = useState("");
  const aggregates = cashierRows.filter(row => !unavailableCashierKeys.has(sourceKey(row)));
  const aggregateOptions = aggregates.filter(row => matchesSearch(
    [row.date, row.name, row.rawName, row.ref, row.defaultRef, row.accountType, row.type, row.matchAmount, formatAmount(row.matchAmount)],
    aggregateSearch,
  ));
  const aggregatePageSize = 50;
  const aggregatePageCount = Math.max(1, Math.ceil(aggregateOptions.length / aggregatePageSize));
  const visibleAggregates = aggregateOptions.slice((aggregatePage - 1) * aggregatePageSize, aggregatePage * aggregatePageSize);
  const aggregate = aggregates.find(row => sourceKey(row) === aggregateKey) ?? null;
  const canSelectPeriod = !!dateFrom && !!dateTo && dateFrom <= dateTo;
  const lowerDate = canSelectPeriod ? addDays(dateFrom, -graceDays) : "";
  const upperDate = canSelectPeriod ? addDays(dateTo, graceDays) : "";
  const allCandidates = useMemo(() => cashierRows
    .filter(row => {
      if (!canSelectPeriod) return false;
      if (unavailableCashierKeys.has(sourceKey(row))) return false;
      if (!aggregate || row.type !== aggregate.type || sourceKey(row) === aggregateKey) return false;
      const date = toIsoDate(row.date);
      return !!date && date >= lowerDate && date <= upperDate;
    })
    .sort((a, b) => (toIsoDate(a.date) || "").localeCompare(toIsoDate(b.date) || "") || a.name.localeCompare(b.name)),
  [cashierRows, unavailableCashierKeys, aggregate, aggregateKey, lowerDate, upperDate, canSelectPeriod]);
  const candidates = allCandidates.filter(row => !candidateSearch.trim() || `${row.date} ${row.name} ${row.rawName} ${row.ref} ${row.defaultRef ?? ""} ${row.matchAmount}`.toLowerCase().includes(candidateSearch.trim().toLowerCase()));
  const selectedRows = allCandidates.filter(row => selectedKeys.has(sourceKey(row)));
  const pageSize = 100;
  const pageCount = Math.max(1, Math.ceil(candidates.length / pageSize));
  const visibleCandidates = candidates.slice((candidatePage - 1) * pageSize, candidatePage * pageSize);
  const selectedTotal = selectedRows.reduce((sum, row) => sum + row.matchAmount, 0);
  const difference = selectedTotal - (aggregate?.matchAmount ?? 0);
  const isExact = !!aggregate && selectedRows.length > 0 && Math.abs(difference) <= 0.01;
  useEffect(() => {
    if (!selectionTouched) {
      setSelectedKeys(new Set(allCandidates.map(sourceKey)));
    }
  }, [allCandidates, selectionTouched]);

  const save = () => {
    setFormError("");
    if (!aggregate) return setFormError("اختر سند الإجمالي الذي سجله المحاسب.");
    if (!canSelectPeriod) return setFormError("تحقق من تاريخي بداية ونهاية الفترة.");
    if (!selectedRows.length) return setFormError("حدد الفواتير الحقيقية التي تريد جمعها.");
    if (!isExact) return setFormError("لا يمكن إغلاق المجموعة ما دام مجموع الفواتير مختلفاً عن سند الإجمالي. وسّع نافذة التداخل أو راجع الحركات الناقصة قبل الحفظ.");
    const id = `cutoff-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const targetSession = aggregate._fileSessionId ?? 0;
    const didSave = onSave({
      id,
      syntheticId: syntheticTransactionId(id),
      aggregateId: aggregate.id,
      aggregateSessionId: targetSession,
      sourceRows: selectedRows.map(row => ({ id: row.id, sessionId: row._fileSessionId ?? 0 })),
      dateFrom, dateTo, graceDays, total: selectedTotal, note: note.trim(),
      savedAt: new Date().toISOString(),
    });
    if (!didSave) {
      setFormError("لم تحفظ المجموعة لأن إحدى الحركات لم تعد متاحة. حدّث الاختيارات وحاول مجدداً.");
      return;
    }
    setAggregateKey("");
    setSelectedKeys(new Set());
    setSelectionTouched(false);
    setNote("");
  };

  return (
    <ToolFrame title="أداة تجميع فترة التعطل (Batch Cut-off Tool)" subtitle="اجمع فواتير الفترة الفعلية مقابل سند الإجمالي، وافحص يوماً قبل أو بعد عند الحاجة. لا تُسوّى مجموعة ذات فرق ولا يُختلق قيد لإخفائه." onBack={onBack}>
      <section className="flex flex-wrap items-center justify-between gap-3 rounded-2xl border border-blue-200 bg-blue-50 p-4">
        <div>
          <h2 className="font-semibold text-blue-950">تريد تسوية حوالات البنك مع الفواتير؟</h2>
          <p className="mt-1 text-xs leading-5 text-blue-900">اختر حوالة أو عدة حوالات ثم وزعها على فاتورة أو عدة فواتير، أو اربط الفاتورة الواحدة بعدة حوالات.</p>
        </div>
        <button type="button" onClick={onOpenMultiMatch} className="inline-flex shrink-0 items-center gap-2 rounded-lg bg-blue-700 px-4 py-2.5 text-sm font-semibold text-white hover:bg-blue-800">
          <WalletCards className="h-4 w-4" />تسوية حوالات وفواتير متعددة
        </button>
      </section>
      <section className="grid gap-4 lg:grid-cols-[minmax(0,1fr)_minmax(300px,0.72fr)]">
        <div className={panelClass}>
          <h2 className="font-semibold">تكوين مجموعة فترة التعطل</h2>
          <div className="grid gap-3 sm:grid-cols-2">
            <label className={labelClass}>من تاريخ
              <input className={inputClass} type="date" value={dateFrom} onChange={event => { setDateFrom(event.target.value); setSelectedKeys(new Set()); setSelectionTouched(false); setCandidatePage(1); }} />
            </label>
            <label className={labelClass}>إلى تاريخ
              <input className={inputClass} type="date" value={dateTo} onChange={event => { setDateTo(event.target.value); setSelectedKeys(new Set()); setSelectionTouched(false); setCandidatePage(1); }} />
            </label>
          </div>
          <div className="grid gap-3 sm:grid-cols-2">
            <label className={labelClass}>نافذة التداخل حول الفترة
              <select className={inputClass} value={graceDays} onChange={event => { setGraceDays(Number(event.target.value)); setSelectedKeys(new Set()); setSelectionTouched(false); setCandidatePage(1); }}>
                <option value={0}>الفترة المحددة فقط</option>
                <option value={1}>إضافة يوم قبل ويوم بعد</option>
                <option value={2}>إضافة يومين قبل ويومين بعد</option>
                <option value={3}>إضافة 3 أيام قبل وبعد</option>
              </select>
            </label>
            <div className="space-y-2">
              <span className={labelClass}>سند الإجمالي المسجل</span>
              <button type="button" onClick={() => { setAggregateSearch(""); setAggregatePage(1); setAggregatePickerOpen(true); }} className={`${inputClass} flex items-center justify-between gap-2 text-right hover:bg-muted/40`}>
                <span className="truncate">{aggregate ? `${aggregate.date || "بلا تاريخ"} · ${aggregate.name} · ${formatAmount(aggregate.matchAmount)}` : "اختر سند الإجمالي"}</span>
                <Search className="h-4 w-4 shrink-0 text-muted-foreground" />
              </button>
            </div>
          </div>
          {aggregate && <div className="rounded-xl border border-blue-200 bg-blue-50 p-3 text-sm">
            <b>{aggregate.name}</b><div className="mt-1 text-xs">قيمة السند: {formatAmount(aggregate.matchAmount)} · {aggregate.date || "بلا تاريخ"}</div>
          </div>}
          <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl border bg-muted/20 p-3 text-sm">
            <span>{aggregate ? <>مجموع {selectedRows.length} من {allCandidates.length} فاتورة محددة: <b className="font-mono">{formatAmount(selectedTotal)}</b></> : "اختر سند الإجمالي لاحتساب فواتير الفترة."}</span>
            <span className={`font-semibold ${isExact ? "text-green-700" : difference < -0.01 ? "text-red-700" : difference > 0.01 ? "text-amber-700" : "text-muted-foreground"}`}>
              {aggregate ? isExact ? "مطابق تماماً" : `الفرق عن السند: ${difference > 0 ? "+" : ""}${formatAmount(difference)}` : "اختر سند الإجمالي"}
            </span>
          </div>
          {aggregate && !isExact && <p className="rounded-lg border border-amber-200 bg-amber-50 p-3 text-xs leading-5 text-amber-900">
            {difference < -0.01
              ? "المجموع أقل من السند: لا تُغلق المجموعة. افحص الفواتير الناقصة أو وسّع نافذة التداخل، وسجّل الفرق للمراجعة."
              : "المجموع أكبر من السند: لا تُغلق المجموعة. افحص إن كانت هناك فواتير زائدة أو سند تكميلي مطلوب من المحاسب."}
          </p>}
          <div className="flex flex-wrap items-end justify-between gap-2">
          <label className={`${labelClass} min-w-[240px] flex-1`}>ابحث داخل فواتير الفترة
            <input className={inputClass} value={candidateSearch} onChange={event => { setCandidateSearch(event.target.value); setCandidatePage(1); }} placeholder="اسم العميل أو رقم المرجع أو مبلغ الفاتورة" />
          </label>
            <div className="flex gap-2">
              <button type="button" onClick={() => { setSelectedKeys(new Set(allCandidates.map(sourceKey))); setSelectionTouched(true); }} className="rounded-lg border px-3 py-2 text-xs hover:bg-muted">تحديد كل فواتير الفترة</button>
              <button type="button" onClick={() => { setSelectedKeys(new Set()); setSelectionTouched(true); }} className="rounded-lg border px-3 py-2 text-xs hover:bg-muted">مسح التحديد</button>
            </div>
          </div>
          <div className="overflow-x-auto rounded-xl border">
            <table className="w-full min-w-[650px] text-sm">
              <thead className="bg-muted text-xs"><tr>
                <th className="w-10 px-3 py-2 text-center">تحديد</th><th className="px-3 py-2 text-right">التاريخ</th><th className="px-3 py-2 text-right">الفاتورة</th><th className="px-3 py-2 text-right">القيمة</th><th className="px-3 py-2 text-right">مصدر التاريخ</th><th className="px-3 py-2 text-center">مراجعة</th>
              </tr></thead>
              <tbody>
                {visibleCandidates.map(row => {
                  const key = sourceKey(row);
                  const date = toIsoDate(row.date);
                  const inChosenPeriod = !!date && date >= dateFrom && date <= dateTo;
                  return <tr key={key} data-default-reference={row.defaultRef} className="border-t hover:bg-muted/30">
                    <td className="px-3 py-2 text-center"><input type="checkbox" checked={selectedKeys.has(key)} onChange={event => {
                      setSelectionTouched(true);
                      setSelectedKeys(current => {
                        const next = new Set(current);
                        if (event.target.checked) next.add(key); else next.delete(key);
                        return next;
                      });
                    }} /></td>
                    <td className="px-3 py-2">{row.date || "—"}</td>
                    <td className="px-3 py-2"><span className="font-medium">{row.name}</span>{row.ref && <span className="mr-2 text-[10px] text-muted-foreground">#{row.ref}</span>} <ReferenceButton row={row} onLocate={onOpenReference} /></td>
                    <td className="px-3 py-2 font-mono">{formatAmount(row.matchAmount)}</td>
                    <td className="px-3 py-2 text-xs">{inChosenPeriod ? <span className="text-green-700">ضمن الفترة</span> : <span className="text-amber-700">ضمن نافذة التداخل</span>}</td>
                    <td className="px-3 py-2 text-center"><button type="button" onClick={() => onReview?.(row, aggregate ? `مجموعة مراجعة مع سند الإجمالي: ${aggregate.name}` : "فاتورة من أداة تجميع فترة التعطل")} title="إضافة الفاتورة إلى المراجعة" aria-label={`مراجعة ${row.name}`} className="rounded-lg border p-1.5 text-amber-700 hover:bg-amber-50"><Star className="h-4 w-4" /></button></td>
                  </tr>;
                })}
                {!candidates.length && <tr><td colSpan={6} className="py-8 text-center text-xs text-muted-foreground">
                  {!aggregate ? "اختر سند الإجمالي أولاً لعرض فواتير الفترة." : !canSelectPeriod ? "تحقق من تاريخ البداية والنهاية." : "لا توجد فواتير متاحة للفترة المحددة والاتجاه المختار."}
                </td></tr>}
              </tbody>
            </table>
          </div>
          <p className="text-xs leading-5 text-muted-foreground">تُحدّد جميع الحركات المتاحة في الفترة تلقائياً، بما فيها نافذة التداخل إن فُعّلت. راجع الأسماء والتواريخ، وأزل أي سند إجمالي أو حركة لا تمثل فاتورة فعلية قبل الإغلاق.</p>
          {candidates.length > pageSize && <div className="flex items-center justify-between text-xs text-muted-foreground">
            <span>عرض {(candidatePage - 1) * pageSize + 1}–{Math.min(candidatePage * pageSize, candidates.length)} من {candidates.length}</span>
            <div className="flex gap-2">
              <button type="button" disabled={candidatePage <= 1} onClick={() => setCandidatePage(page => Math.max(1, page - 1))} className="rounded-lg border px-3 py-1.5 disabled:opacity-40">السابق</button>
              <button type="button" disabled={candidatePage >= pageCount} onClick={() => setCandidatePage(page => Math.min(pageCount, page + 1))} className="rounded-lg border px-3 py-1.5 disabled:opacity-40">التالي</button>
            </div>
          </div>}
          <label className={labelClass}>ملاحظة تدقيقية اختيارية
            <textarea className={`${inputClass} min-h-20 resize-y`} value={note} onChange={event => setNote(event.target.value)} placeholder="سبب تكوين المجموعة أو توضيح نافذة التداخل" />
          </label>
          <button type="button" onClick={save} disabled={!isExact} className="inline-flex items-center gap-1.5 rounded-lg bg-teal-600 px-4 py-2.5 text-sm font-semibold text-white hover:bg-teal-700 disabled:cursor-not-allowed disabled:opacity-40">
            <Check className="h-4 w-4" />إغلاق المجموعة المطابقة
          </button>
          {!canSelectPeriod && <p role="alert" className="text-xs text-red-700">يجب أن يكون تاريخ البداية مساوياً لتاريخ النهاية أو أسبق منه.</p>}
          {formError && <p role="alert" className="rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-700">{formError}</p>}
          <p className="text-xs leading-6 text-muted-foreground">عند التطابق فقط، يستبدل النظام سند الإجمالي المحدد بتمثيل تجميعي للفواتير الحقيقية المختارة. تبقى تفاصيل الفواتير والفترة ونافذة التداخل محفوظة للتقرير والاسترجاع. لا تُنشأ قيود محاسبية.</p>
        </div>

        <div className={panelClass}>
          <h2 className="font-semibold">مجموعات الإقفال المحفوظة</h2>
          {!batches.length && <p className="rounded-xl bg-muted/40 p-4 text-sm text-muted-foreground">لا توجد مجموعات محفوظة.</p>}
          {batches.map(batch => {
            const aggregateRow = cashierRows.find(row => row.id === batch.aggregateId && (row._fileSessionId ?? 0) === batch.aggregateSessionId);
            const details = batch.sourceRows.map(ref => cashierRows.find(row => row.id === ref.id && (row._fileSessionId ?? 0) === ref.sessionId)).filter(Boolean);
            return <article key={batch.id} data-default-reference={[aggregateRow?.defaultRef, ...details.map(row => row?.defaultRef)].filter(Boolean).join("|")} className="space-y-2 rounded-xl border p-3">
              <div className="flex items-start justify-between gap-2">
                <div><p className="font-semibold">{aggregateRow?.name || "سند إجمالي"} {aggregateRow && <ReferenceButton row={aggregateRow} onLocate={onOpenReference} />}</p><p className="mt-1 text-xs text-muted-foreground">{batch.dateFrom} إلى {batch.dateTo} · {details.length} فاتورة · {formatAmount(batch.total)}</p></div>
                <button type="button" onClick={() => onDelete(batch)} className="rounded-lg p-1.5 text-red-600 hover:bg-red-50" title="إلغاء المجموعة"><Trash2 className="h-4 w-4" /></button>
              </div>
              <p className="text-xs text-muted-foreground">{batch.graceDays ? `نافذة تداخل ±${batch.graceDays} يوم` : "دون نافذة تداخل"}{batch.note ? ` · ${batch.note}` : ""}</p>
              <div className="max-h-36 space-y-1 overflow-auto">
                {details.map((row, index) => row && <div key={`${row._fileSessionId ?? 0}:${row.id}`} data-default-reference={row.defaultRef} className="flex justify-between gap-2 rounded bg-muted/40 px-2 py-1.5 text-xs"><span>{index + 1}. {row.date} · {row.name} <ReferenceButton row={row} onLocate={onOpenReference} /></span><b className="font-mono">{formatAmount(row.matchAmount)}</b></div>)}
              </div>
            </article>;
          })}
        </div>
      </section>
      {aggregatePickerOpen && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/50 p-3 sm:p-6" onMouseDown={event => {
          if (event.target === event.currentTarget) setAggregatePickerOpen(false);
        }}>
          <section role="dialog" aria-modal="true" aria-labelledby="cutoff-aggregate-picker-title" className="flex max-h-[90vh] w-full max-w-6xl flex-col overflow-hidden rounded-2xl border bg-card shadow-2xl">
            <header className="flex items-start justify-between gap-3 border-b p-4 sm:p-5">
              <div>
                <h2 id="cutoff-aggregate-picker-title" className="text-lg font-bold">اختيار الفاتورة أو السند الإجمالي</h2>
                <p className="mt-1 text-xs text-muted-foreground">ابحث بالاسم أو البيان أو التاريخ أو المرجع أو المبلغ، ثم اختر السند المطلوب.</p>
              </div>
              <button type="button" onClick={() => setAggregatePickerOpen(false)} className="rounded-lg p-2 hover:bg-muted" aria-label="إغلاق"><X className="h-5 w-5" /></button>
            </header>
            <div className="border-b p-4">
              <div className="relative">
                <Search className="absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
                <input autoFocus className={`${inputClass} pr-10`} value={aggregateSearch} onChange={event => { setAggregateSearch(event.target.value); setAggregatePage(1); }} placeholder="ابحث في الفواتير والسندات بالاسم أو التاريخ أو المبلغ..." />
              </div>
              <p className="mt-2 text-xs text-muted-foreground">{aggregateOptions.length} حركة متاحة للعرض</p>
            </div>
            <div className="min-h-0 flex-1 overflow-auto p-3 sm:p-4">
              <div className="overflow-x-auto rounded-xl border">
                <table className="w-full min-w-[780px] text-sm">
                  <thead className="sticky top-0 bg-muted text-xs"><tr>
                    <th className="px-3 py-3 text-right">التاريخ</th><th className="px-3 py-3 text-right">البيان / اسم السند</th><th className="px-3 py-3 text-right">المرجع</th><th className="px-3 py-3 text-right">الحساب</th><th className="px-3 py-3 text-right">الاتجاه</th><th className="px-3 py-3 text-right">المبلغ</th><th className="px-3 py-3 text-center">مراجعة / اختيار</th>
                  </tr></thead>
                  <tbody>
                    {visibleAggregates.map(row => <tr key={sourceKey(row)} data-default-reference={row.defaultRef} className="border-t hover:bg-muted/30">
                      <td className="px-3 py-2.5">{row.date || "—"}</td><td className="px-3 py-2.5 font-medium">{row.name || row.rawName || "سند بلا اسم"} <ReferenceButton row={row} onLocate={onOpenReference} /></td><td className="px-3 py-2.5">{row.ref || "—"}</td><td className="px-3 py-2.5 text-xs text-muted-foreground">{row.accountType || "—"}</td><td className="px-3 py-2.5">{row.type}</td><td className="px-3 py-2.5 font-mono">{formatAmount(row.matchAmount)}</td>
                      <td className="px-3 py-2.5 text-center"><div className="inline-flex items-center gap-1"><button type="button" onClick={() => onReview?.(row, "سند إجمالي من أداة تجميع فترة التعطل")} title="إضافة سند الإجمالي إلى المراجعة" aria-label={`مراجعة ${row.name}`} className="rounded-lg border p-1.5 text-amber-700 hover:bg-amber-50"><Star className="h-4 w-4" /></button><button type="button" onClick={() => {
                        setAggregateKey(sourceKey(row));
                        setSelectedKeys(new Set());
                        setSelectionTouched(false);
                        setCandidatePage(1);
                        setFormError("");
                        setAggregatePickerOpen(false);
                      }} className="rounded-lg bg-primary px-3 py-1.5 text-xs font-semibold text-primary-foreground hover:opacity-90">اختيار</button></div></td>
                    </tr>)}
                    {!visibleAggregates.length && <tr><td colSpan={7} className="py-8 text-center text-xs text-muted-foreground">لا توجد حركات متاحة تطابق البحث.</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
            <footer className="flex items-center justify-between border-t p-3 text-xs text-muted-foreground">
              <span>صفحة {aggregatePage} من {aggregatePageCount} · عرض {visibleAggregates.length} من {aggregateOptions.length}</span>
              <div className="flex gap-2">
                <button type="button" disabled={aggregatePage >= aggregatePageCount} onClick={() => setAggregatePage(page => Math.min(aggregatePageCount, page + 1))} className="inline-flex items-center gap-1 rounded-lg border px-3 py-1.5 disabled:opacity-40">التالي <ChevronLeft className="h-4 w-4" /></button>
                <button type="button" disabled={aggregatePage <= 1} onClick={() => setAggregatePage(page => Math.max(1, page - 1))} className="inline-flex items-center gap-1 rounded-lg border px-3 py-1.5 disabled:opacity-40"><ChevronRight className="h-4 w-4" /> السابق</button>
              </div>
            </footer>
          </section>
        </div>
      )}
    </ToolFrame>
  );
}
