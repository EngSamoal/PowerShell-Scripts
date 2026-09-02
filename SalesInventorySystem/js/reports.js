/**
 * reports.js — the "التقارير" screen. One shared report-type selector +
 * date range drives 11 report views (section 26 of the spec), each with
 * its own Excel export. "الأرباح" and "المخزون" are not separate menu
 * items (merged into المنتجات/التقارير per the approved simplification),
 * so "ملخص الأرباح" and "المخزون الحالي" live here as report types
 * instead of their own screens.
 *
 * Every report reuses the same data functions the other screens already
 * use (listAllInvoices, computeProfitSummary, productStockDisplay...) —
 * nothing here recomputes a number a different, possibly-inconsistent
 * way.
 */

const REPORT_TYPES = [
  { value: 'today-sales', label: 'مبيعات اليوم' },
  { value: 'monthly-sales', label: 'المبيعات الشهرية' },
  { value: 'retail-sales', label: 'مبيعات القطاعي' },
  { value: 'wholesale-sales', label: 'مبيعات الجملة' },
  { value: 'sales-by-product', label: 'المبيعات حسب المنتج' },
  { value: 'profit-by-product', label: 'الأرباح حسب المنتج' },
  { value: 'current-stock', label: 'المخزون الحالي' },
  { value: 'low-stock', label: 'المنتجات منخفضة المخزون' },
  { value: 'unpaid', label: 'المبالغ غير المدفوعة' },
  { value: 'expenses-report', label: 'المصروفات' },
  { value: 'profit-summary', label: 'ملخص الأرباح' },
];

/** Report types that describe current state, not a time period — their date range controls are hidden. */
const POINT_IN_TIME_REPORTS = new Set(['current-stock', 'low-stock']);

const reportsScreenState = { type: 'today-sales', dateFrom: '', dateTo: '' };

function defaultRangeForType(type) {
  const today = formatDateForStorage();
  const monthStart = `${today.slice(0, 7)}-01`;
  if (type === 'today-sales') return { from: today, to: today };
  if (type === 'unpaid') return { from: '', to: '' };
  if (POINT_IN_TIME_REPORTS.has(type)) return { from: '', to: '' };
  return { from: monthStart, to: today };
}

// ---- Shared helpers -----------------------------------------------------------

async function invoicesInDateRange(dateFrom, dateTo) {
  const invoices = await listAllInvoices();
  return invoices.filter((inv) => (
    inv.status === 'completed'
    && (!dateFrom || inv.date >= dateFrom)
    && (!dateTo || inv.date <= dateTo)
  ));
}

async function invoiceItemsForInvoices(invoices) {
  const invoiceById = new Map(invoices.map((inv) => [inv.id, inv]));
  const allItems = await dbGetAll('invoiceItems');
  return allItems
    .filter((item) => invoiceById.has(item.invoiceId))
    .map((item) => ({ ...item, invoiceDate: invoiceById.get(item.invoiceId).date, invoiceNumber: invoiceById.get(item.invoiceId).invoiceNumber }));
}

/** Builds the shared "list of invoices" report used by مبيعات اليوم / المبيعات الشهرية. */
async function buildInvoiceListReport(dateFrom, dateTo, settings) {
  const invoices = await invoicesInDateRange(dateFrom, dateTo);
  const [itemsByInvoice, customersById] = await Promise.all([groupAllInvoiceItemsByInvoiceId(), buildCustomerLookup()]);

  invoices.sort((a, b) => b.createdAt.localeCompare(a.createdAt));

  const totalValue = invoices.reduce((sum, inv) => sum + inv.grandTotalHalalas, 0);
  const summaryCards = [
    { label: 'عدد الفواتير', value: String(invoices.length) },
    { label: 'إجمالي المبيعات', value: formatCurrency(totalValue, settings.currencySymbol) },
  ];

  const columns = ['رقم الفاتورة', 'التاريخ', 'العميل', 'نوع البيع', 'قيمة المنتجات', 'الرسوم', 'الإجمالي', 'حالة الدفع'];
  const displayRows = [];
  const exportRows = [];
  for (const inv of invoices) {
    const items = itemsByInvoice.get(inv.id) || [];
    const customer = inv.customerId ? customersById.get(inv.customerId) : null;
    const saleType = invoiceSaleTypeSummary(items);
    const feesTotal = inv.feePercentHalalas + inv.feeFixedHalalas;
    displayRows.push([
      inv.invoiceNumber, inv.date, customer ? customer.name : '—', saleType,
      formatCurrency(inv.subtotalHalalas, settings.currencySymbol), formatCurrency(feesTotal, settings.currencySymbol),
      formatCurrency(inv.grandTotalHalalas, settings.currencySymbol), PAYMENT_STATUS_LABELS[inv.paymentStatus],
    ]);
    exportRows.push([
      inv.invoiceNumber, inv.date, customer ? customer.name : '', saleType,
      halalasToSar(inv.subtotalHalalas), halalasToSar(feesTotal), halalasToSar(inv.grandTotalHalalas), PAYMENT_STATUS_LABELS[inv.paymentStatus],
    ]);
  }

  return { summaryCards, columns, displayRows, exportHeaders: columns, exportRows };
}

/** Builds the line-item-level report used by مبيعات القطاعي / مبيعات الجملة (an invoice can mix both, so this filters at the line level, not the invoice level). */
async function buildSaleTypeLinesReport(saleType, dateFrom, dateTo, settings) {
  const invoices = await invoicesInDateRange(dateFrom, dateTo);
  const items = (await invoiceItemsForInvoices(invoices)).filter((i) => i.saleType === saleType);
  items.sort((a, b) => b.invoiceDate.localeCompare(a.invoiceDate));

  const totalValue = items.reduce((sum, i) => sum + i.lineTotalHalalas, 0);
  const summaryCards = [
    { label: 'عدد الأصناف المباعة', value: String(items.length) },
    { label: 'إجمالي القيمة', value: formatCurrency(totalValue, settings.currencySymbol) },
  ];

  const columns = ['التاريخ', 'رقم الفاتورة', 'المنتج', 'الكمية', 'السعر', 'الإجمالي'];
  const displayRows = items.map((i) => [
    i.invoiceDate, i.invoiceNumber, i.productName, i.quantity,
    formatCurrency(i.unitPriceHalalas, settings.currencySymbol), formatCurrency(i.lineTotalHalalas, settings.currencySymbol),
  ]);
  const exportRows = items.map((i) => [
    i.invoiceDate, i.invoiceNumber, i.productName, i.quantity,
    halalasToSar(i.unitPriceHalalas), halalasToSar(i.lineTotalHalalas),
  ]);

  return { summaryCards, columns, displayRows, exportHeaders: columns, exportRows };
}

async function buildSalesByProductReport(dateFrom, dateTo, settings) {
  const invoices = await invoicesInDateRange(dateFrom, dateTo);
  const items = await invoiceItemsForInvoices(invoices);

  const byProduct = new Map();
  for (const i of items) {
    const entry = byProduct.get(i.productId) || { name: i.productName, units: 0, value: 0 };
    entry.units += i.unitsPerLine;
    entry.value += i.lineTotalHalalas;
    byProduct.set(i.productId, entry);
  }
  const rows = Array.from(byProduct.values()).sort((a, b) => b.value - a.value);

  const totalValue = rows.reduce((sum, r) => sum + r.value, 0);
  const summaryCards = [
    { label: 'عدد المنتجات المباعة', value: String(rows.length) },
    { label: 'إجمالي المبيعات', value: formatCurrency(totalValue, settings.currencySymbol) },
  ];

  const columns = ['المنتج', 'الكمية المباعة (وحدة)', 'قيمة المبيعات'];
  const displayRows = rows.map((r) => [r.name, r.units, formatCurrency(r.value, settings.currencySymbol)]);
  const exportRows = rows.map((r) => [r.name, r.units, halalasToSar(r.value)]);

  return { summaryCards, columns, displayRows, exportHeaders: columns, exportRows };
}

async function buildProfitByProductReport(dateFrom, dateTo, settings) {
  const invoices = await invoicesInDateRange(dateFrom, dateTo);
  const items = await invoiceItemsForInvoices(invoices);

  const byProduct = new Map();
  for (const i of items) {
    const entry = byProduct.get(i.productId) || { name: i.productName, value: 0, cost: 0 };
    entry.value += i.lineTotalHalalas;
    entry.cost += i.lineCostHalalas || 0;
    byProduct.set(i.productId, entry);
  }
  const rows = Array.from(byProduct.values())
    .map((r) => ({ ...r, profit: r.value - r.cost }))
    .sort((a, b) => b.profit - a.profit);

  const totalProfit = rows.reduce((sum, r) => sum + r.profit, 0);
  const summaryCards = [
    { label: 'عدد المنتجات', value: String(rows.length) },
    { label: 'إجمالي الربح', value: formatCurrency(totalProfit, settings.currencySymbol) },
  ];

  const columns = ['المنتج', 'قيمة المبيعات', 'التكلفة', 'الربح'];
  const displayRows = rows.map((r) => [r.name, formatCurrency(r.value, settings.currencySymbol), formatCurrency(r.cost, settings.currencySymbol), formatCurrency(r.profit, settings.currencySymbol)]);
  const exportRows = rows.map((r) => [r.name, halalasToSar(r.value), halalasToSar(r.cost), halalasToSar(r.profit)]);

  return { summaryCards, columns, displayRows, exportHeaders: columns, exportRows };
}

async function buildStockReport(lowStockOnly, settings) {
  const products = await listAllProducts();
  const relevant = lowStockOnly ? products.filter((p) => p.status === 'active' && p.totalUnits <= p.minStockUnits) : products;
  relevant.sort((a, b) => a.name.localeCompare(b.name, 'ar'));

  const totalUnits = relevant.reduce((sum, p) => sum + p.totalUnits, 0);
  const summaryCards = [
    { label: 'عدد المنتجات', value: String(relevant.length) },
    { label: 'إجمالي الوحدات', value: String(totalUnits) },
  ];

  const columns = ['الكود', 'المنتج', 'التصنيف', 'المخزون', 'الحد الأدنى', 'الحالة'];
  const displayRows = relevant.map((p) => [
    p.sku, p.name, p.category, productStockDisplay(p), p.minStockUnits, p.status === 'active' ? 'فعال' : 'غير فعال',
  ]);
  const exportRows = relevant.map((p) => [p.sku, p.name, p.category, p.totalUnits, p.minStockUnits, p.status === 'active' ? 'فعال' : 'غير فعال']);

  return { summaryCards, columns, displayRows, exportHeaders: ['الكود', 'المنتج', 'التصنيف', 'المخزون (وحدة)', 'الحد الأدنى', 'الحالة'], exportRows };
}

async function buildUnpaidReport(dateFrom, dateTo, settings) {
  const invoices = await invoicesInDateRange(dateFrom, dateTo);
  const unpaid = invoices.filter((inv) => inv.remainingAmountHalalas > 0);
  const customersById = await buildCustomerLookup();
  unpaid.sort((a, b) => b.createdAt.localeCompare(a.createdAt));

  const totalRemaining = unpaid.reduce((sum, inv) => sum + inv.remainingAmountHalalas, 0);
  const summaryCards = [
    { label: 'عدد الفواتير غير المسددة بالكامل', value: String(unpaid.length) },
    { label: 'إجمالي المتبقي', value: formatCurrency(totalRemaining, settings.currencySymbol) },
  ];

  const columns = ['رقم الفاتورة', 'التاريخ', 'العميل', 'الإجمالي', 'المدفوع', 'المتبقي', 'حالة الدفع'];
  const displayRows = unpaid.map((inv) => {
    const customer = inv.customerId ? customersById.get(inv.customerId) : null;
    return [
      inv.invoiceNumber, inv.date, customer ? customer.name : '—',
      formatCurrency(inv.grandTotalHalalas, settings.currencySymbol), formatCurrency(inv.paidAmountHalalas, settings.currencySymbol),
      formatCurrency(inv.remainingAmountHalalas, settings.currencySymbol), PAYMENT_STATUS_LABELS[inv.paymentStatus],
    ];
  });
  const exportRows = unpaid.map((inv) => {
    const customer = inv.customerId ? customersById.get(inv.customerId) : null;
    return [
      inv.invoiceNumber, inv.date, customer ? customer.name : '',
      halalasToSar(inv.grandTotalHalalas), halalasToSar(inv.paidAmountHalalas), halalasToSar(inv.remainingAmountHalalas), PAYMENT_STATUS_LABELS[inv.paymentStatus],
    ];
  });

  return { summaryCards, columns, displayRows, exportHeaders: columns, exportRows };
}

async function buildExpensesReport(dateFrom, dateTo, settings) {
  const all = await listAllExpenses();
  const filtered = all.filter((e) => (!dateFrom || e.date >= dateFrom) && (!dateTo || e.date <= dateTo));
  filtered.sort((a, b) => b.date.localeCompare(a.date));

  const total = filtered.reduce((sum, e) => sum + e.amountHalalas, 0);
  const summaryCards = [
    { label: 'عدد المصروفات', value: String(filtered.length) },
    { label: 'الإجمالي', value: formatCurrency(total, settings.currencySymbol) },
  ];

  const columns = ['التاريخ', 'النوع', 'الوصف', 'المبلغ', 'ملاحظات'];
  const displayRows = filtered.map((e) => [e.date, EXPENSE_TYPE_LABELS[e.expenseType] || e.expenseType, e.description, formatCurrency(e.amountHalalas, settings.currencySymbol), e.notes]);
  const exportRows = filtered.map((e) => [e.date, EXPENSE_TYPE_LABELS[e.expenseType] || e.expenseType, e.description, halalasToSar(e.amountHalalas), e.notes]);

  return { summaryCards, columns, displayRows, exportHeaders: columns, exportRows };
}

async function buildProfitSummaryReport(dateFrom, dateTo, settings) {
  const s = await computeProfitSummary(dateFrom, dateTo);
  const fmt = (h) => formatCurrency(h, settings.currencySymbol);
  const summaryCards = [
    { label: 'قيمة بيع البضاعة', value: fmt(s.saleValueHalalas) },
    { label: 'تكلفة شراء البضاعة', value: fmt(s.costOfGoodsHalalas) },
    { label: 'ربح البضاعة', value: fmt(s.productProfitHalalas) },
    { label: 'رسوم النسبة المحصلة', value: fmt(s.percentFeeHalalas) },
    { label: 'الرسوم الثابتة المحصلة', value: fmt(s.fixedFeeHalalas) },
    { label: 'إجمالي المبلغ الذي دفعه العميل', value: fmt(s.paidByCustomerHalalas) },
    { label: 'إجمالي المصروفات', value: fmt(s.totalExpensesHalalas) },
    { label: 'الربح الصافي', value: fmt(s.netProfitHalalas) },
  ];
  const exportRows = summaryCards.map((c) => [c.label, c.value]);
  return { summaryCards, columns: [], displayRows: [], exportHeaders: ['البند', 'القيمة'], exportRows };
}

async function buildReportData(type, dateFrom, dateTo, settings) {
  switch (type) {
    case 'today-sales':
    case 'monthly-sales':
      return buildInvoiceListReport(dateFrom, dateTo, settings);
    case 'retail-sales':
      return buildSaleTypeLinesReport('retail', dateFrom, dateTo, settings);
    case 'wholesale-sales':
      return buildSaleTypeLinesReport('wholesale', dateFrom, dateTo, settings);
    case 'sales-by-product':
      return buildSalesByProductReport(dateFrom, dateTo, settings);
    case 'profit-by-product':
      return buildProfitByProductReport(dateFrom, dateTo, settings);
    case 'current-stock':
      return buildStockReport(false, settings);
    case 'low-stock':
      return buildStockReport(true, settings);
    case 'unpaid':
      return buildUnpaidReport(dateFrom, dateTo, settings);
    case 'expenses-report':
      return buildExpensesReport(dateFrom, dateTo, settings);
    case 'profit-summary':
      return buildProfitSummaryReport(dateFrom, dateTo, settings);
    default:
      return { summaryCards: [], columns: [], displayRows: [], exportHeaders: [], exportRows: [] };
  }
}

// ---- Rendering ------------------------------------------------------------------

async function renderReportsScreen() {
  const container = document.getElementById('screen-reports');
  const settings = await getSettings();

  const initialRange = defaultRangeForType(reportsScreenState.type);
  if (!reportsScreenState.dateFrom && !reportsScreenState.dateTo) {
    reportsScreenState.dateFrom = initialRange.from;
    reportsScreenState.dateTo = initialRange.to;
  }

  container.innerHTML = `
    <h2 class="screen-title">التقارير</h2>
    <div class="toolbar">
      <select id="report-type-select" class="toolbar-select">
        ${REPORT_TYPES.map((t) => `<option value="${t.value}" ${t.value === reportsScreenState.type ? 'selected' : ''}>${t.label}</option>`).join('')}
      </select>
      <label class="toolbar-inline" id="report-date-from-wrap"><span>من</span><input type="date" id="report-date-from"></label>
      <label class="toolbar-inline" id="report-date-to-wrap"><span>إلى</span><input type="date" id="report-date-to"></label>
      <div class="toolbar-spacer"></div>
      <button type="button" id="btn-export-report" class="btn btn-secondary">تصدير Excel</button>
    </div>
    <div class="dashboard-grid report-summary-grid" id="report-summary"></div>
    <div class="table-wrap">
      <table class="data-table">
        <thead id="report-table-head"></thead>
        <tbody id="report-table-body"></tbody>
      </table>
      <p id="report-empty-message" class="table-empty-message" hidden>لا توجد بيانات لهذا التقرير في الفترة المحددة.</p>
    </div>
  `;

  document.getElementById('report-date-from').value = reportsScreenState.dateFrom;
  document.getElementById('report-date-to').value = reportsScreenState.dateTo;
  toggleDateRangeVisibility();

  document.getElementById('report-type-select').addEventListener('change', (e) => {
    reportsScreenState.type = e.target.value;
    const range = defaultRangeForType(reportsScreenState.type);
    reportsScreenState.dateFrom = range.from;
    reportsScreenState.dateTo = range.to;
    document.getElementById('report-date-from').value = range.from;
    document.getElementById('report-date-to').value = range.to;
    toggleDateRangeVisibility();
    renderReportBody(settings);
  });
  document.getElementById('report-date-from').addEventListener('change', (e) => {
    reportsScreenState.dateFrom = e.target.value;
    renderReportBody(settings);
  });
  document.getElementById('report-date-to').addEventListener('change', (e) => {
    reportsScreenState.dateTo = e.target.value;
    renderReportBody(settings);
  });
  document.getElementById('btn-export-report').addEventListener('click', async () => {
    try {
      const data = await buildReportData(reportsScreenState.type, reportsScreenState.dateFrom, reportsScreenState.dateTo, settings);
      const typeLabel = REPORT_TYPES.find((t) => t.value === reportsScreenState.type).label;
      exportRowsToExcel(`${typeLabel}_${formatDateForStorage()}.xlsx`, typeLabel, data.exportHeaders, data.exportRows);
      UI.success('تم تصدير التقرير بنجاح.');
    } catch (err) {
      UI.error(friendlyError('تعذر تصدير التقرير. يرجى المحاولة مرة أخرى.', err));
    }
  });

  await renderReportBody(settings);
}

function toggleDateRangeVisibility() {
  const hide = POINT_IN_TIME_REPORTS.has(reportsScreenState.type);
  // .toolbar-inline sets its own `display`, which (being an author-stylesheet rule) always
  // outranks the UA default for [hidden] even when a plain .hidden=true toggle ties on
  // specificity — so this must set display directly rather than the `hidden` property.
  document.getElementById('report-date-from-wrap').style.display = hide ? 'none' : '';
  document.getElementById('report-date-to-wrap').style.display = hide ? 'none' : '';
}

async function renderReportBody(settings) {
  const summaryBox = document.getElementById('report-summary');
  const head = document.getElementById('report-table-head');
  const body = document.getElementById('report-table-body');
  const emptyMessage = document.getElementById('report-empty-message');
  if (!summaryBox) return;

  const data = await buildReportData(reportsScreenState.type, reportsScreenState.dateFrom, reportsScreenState.dateTo, settings);

  summaryBox.innerHTML = data.summaryCards.map((c) => `
    <div class="stat-card">
      <span class="stat-label">${escapeHtml(c.label)}</span>
      <span class="stat-value">${escapeHtml(c.value)}</span>
    </div>
  `).join('');

  if (data.columns.length === 0) {
    head.innerHTML = '';
    body.innerHTML = '';
    emptyMessage.hidden = true;
    return;
  }

  head.innerHTML = `<tr>${data.columns.map((c) => `<th>${escapeHtml(c)}</th>`).join('')}</tr>`;

  if (data.displayRows.length === 0) {
    body.innerHTML = '';
    emptyMessage.hidden = false;
    return;
  }
  emptyMessage.hidden = true;

  body.innerHTML = data.displayRows.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join('')}</tr>`).join('');
}
