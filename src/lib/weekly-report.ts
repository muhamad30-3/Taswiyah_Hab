import type { IngestionFileMetadata } from "./data-ingestion"

export type WeeklyReportSettings = {
  weeklyPayroll: number | null
  monthlyRent: number | null
  shrinkageCost: number | null
}

export type WeeklyReportFile = IngestionFileMetadata & {
  parsedData: NonNullable<IngestionFileMetadata["parsedData"]>
}

export type WeeklyReportSummary = {
  revenue: number | null
  cogs: number | null
  grossProfit: number | null
  payroll: number | null
  rent: number | null
  fuel: number
  utilities: number
  otherExpenses: number
  shrinkage: number | null
  totalOperatingExpenses: number | null
  netProfit: number | null
  profitMargin: number | null
  expectedApps: Array<number | null>
  actualApps: Array<number | null>
  appDifferences: Array<number | null>
  unmatchedWalletTotal: number
  newDebts: number | null
  collectedDebts: number | null
  vendorPayments: number | null
  fileCount: number
  parsedFileCount: number
  expenseFileCount: number
  warnings: string[]
}

function normalized(value: string) {
  return value.normalize("NFKC").toLocaleLowerCase("ar")
    .replace(/[\u064b-\u065f\u0670\u0640]/g, "")
    .replace(/[أإآ]/g, "ا")
}

function appIndex(label: string): number | null {
  const value = normalized(label)
  if (/(jawwal|جوال|جوال باي)/.test(value)) return 0
  if (/(bank|بنك|bank app|تطبيق البنك)/.test(value)) return 1
  if (/(wallet\s*3|app\s*3|application\s*3|تطبيق\s*3|محفظة\s*3|المحفظة الثالثة|wallet 3)/.test(value)) return 2
  return null
}

function isFuelFile(file: WeeklyReportFile) {
  return file.subCategory === "fuel_and_energy" || /(سولار|ديزل|وقود|محروقات|fuel|diesel)/i.test(`${file.customLabel} ${file.fileName}`)
}

function isUtilityFile(file: WeeklyReportFile) {
  return /(كهرباء|ماء|مياه|انترنت|إنترنت|اتصالات|مرافق|utilities|electricity|water|internet)/i.test(`${file.customLabel} ${file.fileName}`)
}

export function calculateWeeklyReport(
  files: WeeklyReportFile[],
  settings: WeeklyReportSettings,
): WeeklyReportSummary {
  const warnings: string[] = []
  const salesFiles = files.filter(file => file.categoryType === "SALES_AND_CATEGORIES")
  const walletFiles = files.filter(file => file.categoryType === "ELECTRONIC_WALLETS")
  const expenseFiles = files.filter(file => file.categoryType === "OPERATING_EXPENSES")
  const debtFiles = files.filter(file => file.categoryType === "DEBTS_AND_LEDGERS")
  const currencies = Array.from(new Set(files.map(file => file.parsedData.summary.currency).filter((value): value is string => !!value)))
  const salesWithTotal = salesFiles.filter(file => typeof file.parsedData.details.totalSales === "number")
  const salesWithCategories = salesFiles.filter(file => file.parsedData.categoriesBreakdown?.length)
  const posSalesFiles = salesWithTotal.filter(file => !file.parsedData.categoriesBreakdown?.length)

  const revenue = posSalesFiles.length
    ? posSalesFiles.reduce((sum, file) => sum + file.parsedData.details.totalSales!, 0)
    : salesWithCategories.length
      ? salesWithCategories.flatMap(file => file.parsedData.categoriesBreakdown ?? []).reduce((sum, item) => sum + item.sales, 0)
      : salesFiles.some(file => file.parsedData.amountDetected)
        ? salesFiles.filter(file => file.parsedData.amountDetected).reduce((sum, file) => sum + file.parsedData.summary.totalAmount, 0)
        : null

  let cogs: number | null = null
  const categoryCogs = salesWithCategories
    .filter(file => file.parsedData.details.cogsDetected !== false)
    .flatMap(file => file.parsedData.categoriesBreakdown ?? [])
  const explicitCogs = salesFiles.filter(file => typeof file.parsedData.details.cogs === "number"
    && file.parsedData.details.cogsDetected !== false)
  if (categoryCogs.length) {
    cogs = categoryCogs.reduce((sum, category) => sum + category.cogs, 0)
  } else if (explicitCogs.length) {
    cogs = explicitCogs.reduce((sum, file) => sum + file.parsedData.details.cogs!, 0)
  }
  const grossProfit = revenue !== null && cogs !== null ? revenue - cogs : null

  const expectedApps: Array<number | null> = [0, 1, 2].map(index => {
    const matches = salesFiles.filter(file => file.parsedData.details[`app${index + 1}Amount` as "app1Amount" | "app2Amount" | "app3Amount"] !== undefined)
    return matches.length
      ? matches.reduce((sum, file) => sum + file.parsedData.details[`app${index + 1}Amount` as "app1Amount" | "app2Amount" | "app3Amount"]!, 0)
      : null
  })
  const actualApps: Array<number | null> = [0, 1, 2].map(index => {
    const matches = walletFiles.filter(file => appIndex(`${file.customLabel} ${file.fileName}`) === index
      && typeof file.parsedData.details.actualStatementTotal === "number")
    return matches.length
      ? matches.reduce((sum, file) => sum + file.parsedData.details.actualStatementTotal!, 0)
      : null
  })
  const appDifferences = expectedApps.map((expected, index) =>
    expected !== null && actualApps[index] !== null ? actualApps[index]! - expected : null)
  const matchedWalletIds = new Set(walletFiles
    .filter(file => appIndex(`${file.customLabel} ${file.fileName}`) !== null)
    .map(file => file.fileId))
  const unmatchedWalletTotal = walletFiles
    .filter(file => !matchedWalletIds.has(file.fileId))
    .reduce((sum, file) => sum + (file.parsedData.details.actualStatementTotal ?? 0), 0)

  let fuel = 0
  let utilities = 0
  let otherExpenses = 0
  for (const file of expenseFiles) {
    const expenseAmount = file.parsedData.details.totalExpenseAmount
    if (expenseAmount === undefined) continue
    if (isFuelFile(file)) {
      const calculatedConsumption = file.parsedData.details.litersCount !== undefined
        && file.parsedData.details.costPerLiter !== undefined
        ? file.parsedData.details.litersCount * file.parsedData.details.costPerLiter
        : expenseAmount
      fuel += calculatedConsumption
    } else if (isUtilityFile(file)) {
      utilities += expenseAmount
    } else {
      otherExpenses += expenseAmount
    }
  }

  const payroll = settings.weeklyPayroll
  const rent = settings.monthlyRent === null ? null : settings.monthlyRent * 12 / 52
  const shrinkage = settings.shrinkageCost
  const allFixedCostsProvided = payroll !== null && rent !== null && shrinkage !== null
  const totalOperatingExpenses = allFixedCostsProvided
    ? payroll + rent + shrinkage + fuel + utilities + otherExpenses
    : null
  const netProfit = grossProfit !== null && totalOperatingExpenses !== null
    ? grossProfit - totalOperatingExpenses
    : null
  const profitMargin = netProfit !== null && revenue !== null && revenue !== 0 ? netProfit / revenue * 100 : null

  const salesFilesWithoutTotals = salesFiles.filter(file => !file.parsedData.amountDetected)
  if (!salesFiles.length) warnings.push("لا توجد ملفات مبيعات في هذا الأسبوع.")
  else if (salesFilesWithoutTotals.length) warnings.push("بعض ملفات المبيعات لا تحتوي إجمالياً مقروءاً، ولم تدخل في الإيرادات.")
  if (revenue === null) warnings.push("إجمالي المبيعات غير متاح حتى ترفع ملف POS يحوي عمود إجمالي.")
  if (cogs === null) warnings.push("تكلفة البضاعة غير متاحة؛ أضف ملف أرباح الأقسام الذي يحتوي على التكلفة.")
  if (payroll === null) warnings.push("أدخل الرواتب الأسبوعية لإكمال صافي الربح.")
  if (rent === null) warnings.push("أدخل الإيجار الشهري لحساب حصته الأسبوعية.")
  if (shrinkage === null) warnings.push("أدخل قيمة التوالف والعجز لهذا الأسبوع.")
  if (!expenseFiles.length) warnings.push("لا توجد ملفات مصاريف لهذا الأسبوع؛ لم تُحتسب المصاريف المتغيرة.")
  if (walletFiles.some(file => appIndex(`${file.customLabel} ${file.fileName}`) === null)) {
    warnings.push("بعض كشوف المحافظ لم تُربط تلقائياً بأحد التطبيقات الثلاثة، وتظهر ضمن غير المصنف.")
  }
  if (walletFiles.some(file => file.parsedData.extractionNote)) {
    warnings.push("تحقق يدوياً من كشوف المحافظ التي أظهر تحليلها ملاحظة حول اتجاه الحركة.")
  }
  if (currencies.length > 1) warnings.push(`توجد ملفات بعملات متعددة (${currencies.join("، ")}). لا يجري التقرير تحويل العملات؛ راجع الإجماليات قبل اعتمادها.`)
  if (salesFiles.length > 1) warnings.push("تأكد أن ملفات POS للفترة لا تتضمن فترات متداخلة حتى لا تتكرر المبيعات.")

  return {
    revenue,
    cogs,
    grossProfit,
    payroll,
    rent,
    fuel,
    utilities,
    otherExpenses,
    shrinkage,
    totalOperatingExpenses,
    netProfit,
    profitMargin,
    expectedApps,
    actualApps,
    appDifferences,
    unmatchedWalletTotal,
    newDebts: salesFiles.some(file => file.parsedData.details.newDebtsAmount !== undefined)
      ? salesFiles.reduce((sum, file) => sum + (file.parsedData.details.newDebtsAmount ?? 0), 0)
      : null,
    collectedDebts: debtFiles.some(file => file.parsedData.details.collectedDebtsAmount !== undefined)
      ? debtFiles.reduce((sum, file) => sum + (file.parsedData.details.collectedDebtsAmount ?? 0), 0)
      : null,
    vendorPayments: debtFiles.some(file => file.parsedData.details.vendorPaymentsMade !== undefined)
      ? debtFiles.reduce((sum, file) => sum + (file.parsedData.details.vendorPaymentsMade ?? 0), 0)
      : null,
    fileCount: files.length,
    parsedFileCount: files.length,
    expenseFileCount: expenseFiles.length,
    warnings,
  }
}
