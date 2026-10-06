import { useEffect, useState, type FormEvent, type MouseEvent } from "react";
import { LockKeyhole, ShieldAlert } from "lucide-react";

export interface BlockedAction {
  screen: string;
  action: string;
  detail?: string;
}

export function SafetyLockOverlay({
  activeScreen,
  onAttempt,
  onUnlock,
}: {
  activeScreen: string;
  onAttempt: (attempt: BlockedAction) => void;
  onUnlock: (password: string) => boolean;
}) {
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [lastAttempt, setLastAttempt] = useState("");

  const getScreen = () => {
    const heading = document.querySelector(".recon-scope h1, .recon-scope h2")?.textContent?.trim();
    return heading ? `${activeScreen} · ${heading}` : activeScreen;
  };

  const getActionLabel = (element: Element | null) => {
    const control = element?.closest<HTMLElement>("button, a, input, select, textarea, label, [role='button'], [role='menuitem']") ?? element;
    if (!control) return "نقرة على الشاشة";
    if (control instanceof HTMLInputElement || control instanceof HTMLSelectElement || control instanceof HTMLTextAreaElement) {
      const label = control.labels?.[0]?.textContent?.trim();
      return label || control.getAttribute("aria-label") || control.getAttribute("placeholder") || "حقل إدخال";
    }
    const visibleText = control instanceof HTMLElement ? control.innerText : control.textContent;
    return (control.getAttribute("aria-label") || control.getAttribute("title") || visibleText || "عنصر في الشاشة")
      .replace(/\s+/g, " ").trim().slice(0, 160);
  };

  const recordAttempt = (action: string, detail?: string) => {
    const screen = getScreen();
    setLastAttempt(`${action} · ${screen}`);
    onAttempt({ screen, action, detail });
  };

  const handleClick = (event: MouseEvent<HTMLDivElement>) => {
    const target = event.target as HTMLElement;
    if (target.closest("#safety-unlock-password, [data-safety-unlock] button, label[for='safety-unlock-password']")) return;
    const overlay = event.currentTarget;
    const previousPointerEvents = overlay.style.pointerEvents;
    overlay.style.pointerEvents = "none";
    const underlying = document.elementFromPoint(event.clientX, event.clientY);
    overlay.style.pointerEvents = previousPointerEvents;
    recordAttempt(getActionLabel(underlying), "لم يُنفذ الإجراء بسبب تفعيل خط الأمان.");
  };

  const handleSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!onUnlock(password)) {
      setError("كلمة المرور غير صحيحة. سُجلت محاولة فك القفل.");
      setPassword("");
      return;
    }
    setPassword("");
    setError("");
  };

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      const target = event.target;
      if (target instanceof HTMLElement && target.closest("[data-safety-unlock]")) {
        if (event.key === "Tab") {
          const controls = Array.from(document.querySelectorAll<HTMLElement>("[data-safety-unlock] input, [data-safety-unlock] button"));
          const first = controls[0];
          const last = controls[controls.length - 1];
          if ((!event.shiftKey && target === last) || (event.shiftKey && target === first)) {
            event.preventDefault();
            (event.shiftKey ? last : first)?.focus();
          }
        } else {
          event.stopPropagation();
          event.stopImmediatePropagation();
        }
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      if (["Shift", "Control", "Alt", "Meta"].includes(event.key)) return;
      recordAttempt("محاولة استخدام لوحة المفاتيح", "لم يُنفذ أي اختصار أو إدخال أثناء القفل.");
    };
    window.addEventListener("keydown", onKeyDown, true);
    return () => window.removeEventListener("keydown", onKeyDown, true);
  }, [activeScreen]);

  return (
    <div
      data-safety-overlay
      onClick={handleClick}
      onWheel={event => {
        if ((event.target as HTMLElement).closest("[data-safety-unlock]")) {
          event.stopPropagation();
          return;
        }
        event.preventDefault();
        recordAttempt("محاولة تمرير الشاشة", "لم يتغير موضع أي قائمة أثناء القفل.");
      }}
      onContextMenu={event => {
        if ((event.target as HTMLElement).closest("input, button, label")) return;
        event.preventDefault();
        recordAttempt("محاولة فتح قائمة سياق", "لم تُنفذ أي قائمة أو أمر أثناء القفل.");
      }}
      onDragOver={event => event.preventDefault()}
      onDrop={event => {
        event.preventDefault();
        recordAttempt("محاولة إضافة ملف", `عدد الملفات: ${event.dataTransfer.files.length}. لم تتم قراءة محتواها.`);
      }}
      className="fixed inset-0 z-[100] flex items-center justify-center bg-slate-950/45 p-4 backdrop-blur-[1px]"
      role="presentation"
    >
      <section
        data-safety-unlock
        role="alertdialog"
        aria-modal="true"
        aria-labelledby="safety-lock-title"
        className="max-h-[90dvh] w-full max-w-md overflow-y-auto rounded-xl border border-rose-300 bg-white p-6 text-right shadow-2xl"
      >
        <div className="mb-4 flex items-center gap-3 text-rose-700">
          <span className="flex h-11 w-11 items-center justify-center rounded-lg bg-rose-100">
            <ShieldAlert className="h-6 w-6" />
          </span>
          <div>
            <h1 id="safety-lock-title" className="text-lg font-bold">خط الأمان مفعّل</h1>
            <p className="text-xs text-rose-700">التطبيق في وضع القراءة فقط</p>
          </div>
        </div>
        <p className="text-sm leading-6 text-slate-700">
          لن تُحفظ أي مطابقة أو إضافة أو حذف أو تغيير إعدادات. محاولات التفاعل تُسجل للمراجعة، ولا تُطبق على بياناتك.
        </p>
        {lastAttempt && <p className="mt-3 rounded-md bg-rose-50 px-3 py-2 text-xs text-rose-800">آخر محاولة: {lastAttempt} · لم تُنفذ</p>}
        <form data-safety-unlock onSubmit={handleSubmit} className="mt-5 space-y-2">
          <label htmlFor="safety-unlock-password" className="text-xs font-medium text-slate-700">أدخل # ثم كلمة فك قفل الأمان</label>
          <input
            id="safety-unlock-password"
            type="password"
            autoComplete="current-password"
            autoFocus
            value={password}
            onChange={event => setPassword(event.target.value)}
            placeholder="#كلمة فك القفل"
            className="w-full rounded-lg border border-slate-300 px-3 py-2.5 text-sm outline-none focus:border-rose-500 focus:ring-2 focus:ring-rose-100"
          />
          {error && <p role="alert" className="text-xs text-rose-700">{error}</p>}
          <button type="submit" className="flex w-full items-center justify-center gap-2 rounded-lg bg-rose-700 px-4 py-2.5 text-sm font-semibold text-white hover:bg-rose-800">
            <LockKeyhole className="h-4 w-4" /> فك خط الأمان
          </button>
        </form>
      </section>
    </div>
  );
}
