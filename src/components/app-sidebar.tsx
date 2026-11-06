"use client"

import { User, Settings, Scale, LogOut, ChevronLeft, Sparkles, Link2, BarChart3, FileSpreadsheet, Split, CalendarRange, ListChecks, Home, BookOpen, ScanText, FolderOpen, UserCog, WalletCards, StickyNote, Upload, ChartNoAxesCombined, Fuel, Banknote, ArrowLeftRight, ReceiptText, CalendarCheck } from "lucide-react"

export type ViewId = "home" | "fuel" | "payroll" | "attendance" | "transfers" | "cashbox" | "salesInvoice" | "customerBalances" | "customerNotes" | "documentStudio" | "recon" | "assistant" | "manual" | "report" | "weeklyReport" | "splitTool" | "batchCutoff" | "clearing" | "processor" | "guide" | "files" | "ingestion" | "adminUsers" | "profile" | "settings"

export function AppSidebar({
  active,
  onNavigate,
  username,
  isAdmin,
  centralStorage,
  onLogout,
  collapsed,
  onToggleCollapse,
}: {
  active: ViewId
  onNavigate: (view: ViewId) => void
  username: string
  isAdmin: boolean
  centralStorage: boolean
  onLogout: () => void
  collapsed: boolean
  onToggleCollapse: () => void
}) {
  const navItems: { id: ViewId; label: string; icon: typeof User }[] = [
    { id: "home", label: "الصفحة الرئيسية", icon: Home },
    { id: "fuel", label: "منصة المحروقات", icon: Fuel },
    { id: "payroll", label: "كشف رواتب الموظفين", icon: Banknote },
    ...(centralStorage ? [{ id: "attendance" as const, label: "الحضور والانصراف", icon: CalendarCheck }] : []),
    { id: "transfers", label: "منصة الحوالات", icon: ArrowLeftRight },
    { id: "cashbox", label: "الصندوق المالي", icon: WalletCards },
    { id: "salesInvoice", label: "منصة الفواتير", icon: ReceiptText },
    { id: "customerBalances", label: "أرصدة الزبائن", icon: WalletCards },
    { id: "customerNotes", label: "ملاحظات وتذكيرات", icon: StickyNote },
    { id: "documentStudio", label: "المستندات الذكية", icon: ScanText },
    { id: "recon", label: "أدوات التسوية", icon: Scale },
    { id: "assistant", label: "مساعد المطابقة", icon: Sparkles },
    { id: "manual", label: "المطابقة اليدوية", icon: Link2 },
    { id: "report", label: "تقرير الفروقات", icon: BarChart3 },
    ...(centralStorage ? [{ id: "weeklyReport" as const, label: "التقرير الأسبوعي الشامل", icon: ChartNoAxesCombined }] : []),
    { id: "splitTool", label: "تسوية حوالات وفواتير متعددة", icon: Split },
    { id: "batchCutoff", label: "أداة تجميع فترة التعطل", icon: CalendarRange },
    { id: "clearing", label: "شاشة تسكير السندات", icon: ListChecks },
    { id: "processor", label: "معالج كشوفات الحسابات", icon: FileSpreadsheet },
    ...(centralStorage ? [{ id: "files" as const, label: "ملفاتي", icon: FolderOpen }] : []),
    ...(centralStorage ? [{ id: "ingestion" as const, label: "مركز رفع البيانات", icon: Upload }] : []),
    ...(centralStorage && isAdmin ? [{ id: "adminUsers" as const, label: "إدارة المستخدمين", icon: UserCog }] : []),
    { id: "guide", label: "دليل المستخدم", icon: BookOpen },
    { id: "profile", label: "الملف الشخصي", icon: User },
    { id: "settings", label: "الإعدادات", icon: Settings },
  ]

  return (
    <aside
      className={`flex h-screen w-60 shrink-0 flex-col border-l border-sidebar-border bg-sidebar text-sidebar-foreground ${
        collapsed ? "shadow-2xl" : ""
      }`}
    >
      <button
        type="button"
        onClick={onToggleCollapse}
        title="اضغط لإخفاء القائمة"
        className="flex items-center gap-2.5 border-b border-sidebar-border px-4 py-4 text-right transition-colors hover:bg-sidebar-accent"
      >
        <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-sidebar-primary text-sidebar-primary-foreground overflow-hidden">
          <img src="/taswiyah-logo.png" alt="Taswiyah Hub" className="h-full w-full object-cover mix-blend-multiply" />
        </div>
        <div className="min-w-0 flex-1">
            <div className="truncate text-sm font-bold">Taswiyah Hub</div>
            <div className="truncate text-[10px] text-sidebar-foreground/60">محطتك الأولى لتسوية الحسابات بدقة.</div>
        </div>
        <ChevronLeft className="h-4 w-4 shrink-0 text-muted-foreground" />
      </button>

      <nav className="flex-1 space-y-1 overflow-y-auto p-3">
        {navItems.map((item) => {
          const Icon = item.icon
          const isActive = active === item.id
          return (
            <button
              key={item.id}
              onClick={() => onNavigate(item.id)}
              className={`flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-sm font-medium transition-colors ${
                isActive
                  ? "bg-sidebar-primary text-sidebar-primary-foreground"
                  : "text-sidebar-foreground hover:bg-sidebar-accent hover:text-sidebar-primary"
              }`}
            >
              <Icon className="h-4 w-4 shrink-0" />
              <span>{item.label}</span>
            </button>
          )
        })}
      </nav>

      <div className="border-t border-sidebar-border p-3">
        <div className="mb-2 flex items-center gap-3 rounded-lg px-3 py-2">
          <div className="flex h-8 w-8 items-center justify-center rounded-full bg-sidebar-accent text-sidebar-accent-foreground">
            <User className="h-4 w-4" />
          </div>
          <div className="min-w-0">
            <div className="truncate text-sm font-medium">{username}</div>
            <div className="text-xs text-muted-foreground">مستخدم مسجّل</div>
          </div>
        </div>
        <button
          onClick={onLogout}
          className="flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-sm font-medium text-destructive transition-colors hover:bg-destructive/10"
        >
          <LogOut className="h-4 w-4 shrink-0" />
          <span>تسجيل الخروج</span>
        </button>
      </div>
    </aside>
  )
}
