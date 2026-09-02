/**
 * invoices.js — invoice numbering and the core sale-completion
 * transaction (stock check + deduction + invoice + invoiceItems),
 * shared by the "بيع جديد" screen now and the "الفواتير" screen in a
 * later phase.
 *
 * Every invoice snapshots the fee percentage/fixed fee from settings.js
 * at the moment of sale, so a later Settings change never alters an
 * already-saved invoice (see buildProduct-style snapshotting in
 * products.js for the same principle applied to stock).
 */

/** Thrown for an expected, user-facing failure (out of stock, inactive product...) inside the sale transaction. */
class InvoiceError extends Error {}

const PAYMENT_METHOD_LABELS = { cash: 'نقدي', bank_transfer: 'تحويل بنكي', card: 'بطاقة', other: 'أخرى' };
const PAYMENT_STATUS_LABELS = { paid: 'مدفوع', unpaid: 'غير مدفوع', partial: 'مدفوع جزئيًا' };

/**
 * Validates and completes a sale: checks every cart line's stock against
 * the live product record, deducts it, and writes the invoice + its line
 * items — all inside one IndexedDB transaction, so a sale either fully
 * succeeds with stock fully deducted, or nothing is written at all.
 *
 * cartItems: [{ productId, productName, sku, saleType, quantity, unitPriceHalalas, lineTotalHalalas, unitsPerLine }]
 * meta: { customerId, paymentMethod, paymentStatus, paidAmountHalalas, notes }
 */
async function createInvoice(cartItems, meta, settings) {
  if (!cartItems || cartItems.length === 0) {
    return { errors: ['السلة فارغة. أضف منتجًا واحدًا على الأقل قبل إتمام البيع.'] };
  }

  const subtotalHalalas = cartItems.reduce((sum, item) => sum + item.lineTotalHalalas, 0);
  const fees = calculateFees(subtotalHalalas, settings.feePercent, settings.feeFixedHalalas);

  let paidAmountHalalas;
  if (meta.paymentStatus === 'paid') paidAmountHalalas = fees.grandTotalHalalas;
  else if (meta.paymentStatus === 'unpaid') paidAmountHalalas = 0;
  else paidAmountHalalas = meta.paidAmountHalalas || 0;

  if (paidAmountHalalas < 0 || paidAmountHalalas > fees.grandTotalHalalas) {
    return { errors: ['المبلغ المدفوع غير صحيح.'] };
  }
  const remainingAmountHalalas = fees.grandTotalHalalas - paidAmountHalalas;

  const now = new Date();
  const year = now.getFullYear();

  try {
    const result = await dbTransaction(['products', 'invoices', 'invoiceItems', 'counters'], async (tx) => {
      const productsStore = tx.objectStore('products');

      // Aggregate requested units per product first — the same product can
      // appear twice (e.g. some sold retail, some wholesale, in one sale).
      const neededUnitsByProduct = new Map();
      for (const item of cartItems) {
        neededUnitsByProduct.set(item.productId, (neededUnitsByProduct.get(item.productId) || 0) + item.unitsPerLine);
      }

      // Pass 1: validate everything against the live stock. No writes yet,
      // so a failure here leaves the database completely untouched.
      const productRecords = new Map();
      for (const [productId, neededUnits] of neededUnitsByProduct) {
        const product = await requestToPromise(productsStore.get(productId));
        if (!product) throw new InvoiceError('أحد المنتجات في السلة لم يعد موجودًا. يرجى تحديث السلة.');
        if (product.status !== 'active') throw new InvoiceError(`المنتج "${product.name}" غير فعال ولا يمكن بيعه.`);
        if (product.totalUnits < neededUnits) {
          throw new InvoiceError(`الكمية المطلوبة من "${product.name}" أكبر من المخزون المتاح (المتاح: ${productStockDisplay(product)}).`);
        }
        productRecords.set(productId, product);
      }

      // Pass 2: everything validated — now it's safe to deduct stock.
      for (const [productId, neededUnits] of neededUnitsByProduct) {
        const product = productRecords.get(productId);
        product.totalUnits -= neededUnits;
        product.updatedAt = now.toISOString();
        productsStore.put(product);
      }

      const counterName = `invoice-${year}`;
      const counterStore = tx.objectStore('counters');
      const existingCounter = await requestToPromise(counterStore.get(counterName));
      const nextValue = (existingCounter ? existingCounter.value : 0) + 1;
      counterStore.put({ name: counterName, value: nextValue });
      const invoiceNumber = `${settings.invoicePrefix}-${year}-${String(nextValue).padStart(6, '0')}`;

      const invoiceRecord = {
        invoiceNumber,
        date: formatDateForStorage(now),
        createdAt: now.toISOString(),
        customerId: meta.customerId || null,
        subtotalHalalas,
        feePercentUsed: settings.feePercent,
        feePercentHalalas: fees.percentFeeHalalas,
        feeFixedHalalas: fees.fixedFeeHalalas,
        grandTotalHalalas: fees.grandTotalHalalas,
        paymentMethod: meta.paymentMethod,
        paymentStatus: meta.paymentStatus,
        paidAmountHalalas,
        remainingAmountHalalas,
        notes: meta.notes || '',
        status: 'completed',
      };
      const invoiceId = await requestToPromise(tx.objectStore('invoices').add(invoiceRecord));

      const invoiceItemsStore = tx.objectStore('invoiceItems');
      for (const item of cartItems) {
        invoiceItemsStore.add({
          invoiceId,
          productId: item.productId,
          productName: item.productName,
          productSku: item.sku,
          saleType: item.saleType,
          quantity: item.quantity,
          unitPriceHalalas: item.unitPriceHalalas,
          lineTotalHalalas: item.lineTotalHalalas,
          unitsPerLine: item.unitsPerLine,
        });
      }

      return { invoiceId, invoiceNumber, dateDisplay: formatDateTimeForDisplay(now) };
    });

    return {
      errors: [],
      invoiceId: result.invoiceId,
      invoiceNumber: result.invoiceNumber,
      dateDisplay: result.dateDisplay,
      subtotalHalalas,
      feePercentUsed: settings.feePercent,
      feePercentHalalas: fees.percentFeeHalalas,
      feeFixedHalalas: fees.fixedFeeHalalas,
      grandTotalHalalas: fees.grandTotalHalalas,
      paidAmountHalalas,
      remainingAmountHalalas,
    };
  } catch (err) {
    if (err instanceof InvoiceError) {
      return { errors: [err.message] };
    }
    throw err;
  }
}

/** Builds the printable receipt HTML for a just-created (or reopened) invoice. Reused by the Invoices screen in a later phase. */
function renderInvoiceReceiptHtml(invoice, items, settings) {
  return `
    <div class="receipt">
      <h2>${escapeHtml(settings.businessName || 'نظام المبيعات')}</h2>
      ${settings.phone ? `<p>${escapeHtml(settings.phone)}</p>` : ''}
      <p>فاتورة رقم: ${escapeHtml(invoice.invoiceNumber)}</p>
      <p>التاريخ: ${escapeHtml(invoice.dateDisplay)}</p>
      <table class="receipt-table">
        <thead><tr><th>المنتج</th><th>النوع</th><th>الكمية</th><th>السعر</th><th>الإجمالي</th></tr></thead>
        <tbody>
          ${items.map((item) => `
            <tr>
              <td>${escapeHtml(item.productName)}</td>
              <td>${item.saleType === 'wholesale' ? 'جملة' : 'قطاعي'}</td>
              <td>${item.quantity}</td>
              <td>${formatCurrency(item.unitPriceHalalas, settings.currencySymbol)}</td>
              <td>${formatCurrency(item.lineTotalHalalas, settings.currencySymbol)}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
      <div class="receipt-totals">
        <div><span>قيمة المنتجات</span><span>${formatCurrency(invoice.subtotalHalalas, settings.currencySymbol)}</span></div>
        <div><span>رسوم الموقع (${invoice.feePercentUsed}٪)</span><span>${formatCurrency(invoice.feePercentHalalas, settings.currencySymbol)}</span></div>
        <div><span>الرسم الثابت</span><span>${formatCurrency(invoice.feeFixedHalalas, settings.currencySymbol)}</span></div>
        <div class="receipt-grand-total"><span>الإجمالي النهائي</span><span>${formatCurrency(invoice.grandTotalHalalas, settings.currencySymbol)}</span></div>
      </div>
      <p>طريقة الدفع: ${PAYMENT_METHOD_LABELS[invoice.paymentMethod] || invoice.paymentMethod}</p>
      <p>حالة الدفع: ${PAYMENT_STATUS_LABELS[invoice.paymentStatus] || invoice.paymentStatus}</p>
      ${invoice.paymentStatus === 'partial' ? `<p>المدفوع: ${formatCurrency(invoice.paidAmountHalalas, settings.currencySymbol)} — المتبقي: ${formatCurrency(invoice.remainingAmountHalalas, settings.currencySymbol)}</p>` : ''}
      ${invoice.notes ? `<p>ملاحظات: ${escapeHtml(invoice.notes)}</p>` : ''}
      ${settings.extraInvoiceInfo ? `<p>${escapeHtml(settings.extraInvoiceInfo)}</p>` : ''}
    </div>
  `;
}

/** Prints one receipt via the hidden #print-area + print stylesheet (see style.css / index.html). */
function printReceipt(html) {
  document.getElementById('print-area').innerHTML = html;
  window.print();
}

// ---- Data access for the "الفواتير" screen -------------------------------------

async function listAllInvoices() {
  return dbGetAll('invoices');
}

async function getInvoiceById(id) {
  return dbGet('invoices', id);
}

async function getInvoiceItemsByInvoiceId(invoiceId) {
  return dbGetAllByIndex('invoiceItems', 'invoiceId', invoiceId);
}

/** One query for every invoice's line items, grouped by invoiceId — avoids an N+1 read when rendering the whole list. */
async function groupAllInvoiceItemsByInvoiceId() {
  const allItems = await dbGetAll('invoiceItems');
  const map = new Map();
  for (const item of allItems) {
    if (!map.has(item.invoiceId)) map.set(item.invoiceId, []);
    map.get(item.invoiceId).push(item);
  }
  return map;
}

/** "قطاعي" / "جملة" / "مختلط" — derived from the invoice's own line items, never stored, so it can never drift out of sync. */
function invoiceSaleTypeSummary(items) {
  const hasRetail = items.some((i) => i.saleType === 'retail');
  const hasWholesale = items.some((i) => i.saleType === 'wholesale');
  if (hasRetail && hasWholesale) return 'مختلط';
  if (hasWholesale) return 'جملة';
  return 'قطاعي';
}

/** Adapts a stored invoice record (date/createdAt) into the shape renderInvoiceReceiptHtml() expects (dateDisplay). */
function invoiceToReceiptData(invoice) {
  return { ...invoice, dateDisplay: formatDateTimeForDisplay(new Date(invoice.createdAt)) };
}

// ---- Rendering: "الفواتير" screen -----------------------------------------------

const invoicesScreenState = { search: '', dateFrom: '', dateTo: '', paymentStatusFilter: 'all' };

const PAYMENT_STATUS_BADGE_CLASS = { paid: 'badge-success', unpaid: 'badge-warning', partial: 'badge-muted' };

async function renderInvoicesScreen() {
  const container = document.getElementById('screen-invoices');

  container.innerHTML = `
    <h2 class="screen-title">الفواتير</h2>
    <div class="toolbar">
      <input type="search" id="invoice-search" class="toolbar-search" placeholder="ابحث برقم الفاتورة...">
      <label class="toolbar-inline"><span>من</span><input type="date" id="invoice-date-from"></label>
      <label class="toolbar-inline"><span>إلى</span><input type="date" id="invoice-date-to"></label>
      <select id="invoice-status-filter" class="toolbar-select">
        <option value="all">كل حالات الدفع</option>
        <option value="paid">مدفوع</option>
        <option value="unpaid">غير مدفوع</option>
        <option value="partial">مدفوع جزئيًا</option>
      </select>
    </div>
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>رقم الفاتورة</th><th>التاريخ</th><th>نوع البيع</th>
            <th>قيمة المنتجات</th><th>الرسوم</th><th>الإجمالي</th><th>حالة الدفع</th><th>إجراءات</th>
          </tr>
        </thead>
        <tbody id="invoices-table-body"></tbody>
      </table>
      <p id="invoices-empty-message" class="table-empty-message" hidden>لا توجد فواتير مطابقة.</p>
    </div>
  `;

  document.getElementById('invoice-search').value = invoicesScreenState.search;
  document.getElementById('invoice-date-from').value = invoicesScreenState.dateFrom;
  document.getElementById('invoice-date-to').value = invoicesScreenState.dateTo;
  document.getElementById('invoice-status-filter').value = invoicesScreenState.paymentStatusFilter;

  document.getElementById('invoice-search').addEventListener('input', (e) => {
    invoicesScreenState.search = e.target.value;
    renderInvoicesTableBody();
  });
  document.getElementById('invoice-date-from').addEventListener('change', (e) => {
    invoicesScreenState.dateFrom = e.target.value;
    renderInvoicesTableBody();
  });
  document.getElementById('invoice-date-to').addEventListener('change', (e) => {
    invoicesScreenState.dateTo = e.target.value;
    renderInvoicesTableBody();
  });
  document.getElementById('invoice-status-filter').addEventListener('change', (e) => {
    invoicesScreenState.paymentStatusFilter = e.target.value;
    renderInvoicesTableBody();
  });

  await renderInvoicesTableBody();
}

async function renderInvoicesTableBody() {
  const tbody = document.getElementById('invoices-table-body');
  const emptyMessage = document.getElementById('invoices-empty-message');
  if (!tbody) return;

  const settings = await getSettings();
  const [invoices, itemsByInvoice] = await Promise.all([listAllInvoices(), groupAllInvoiceItemsByInvoiceId()]);

  const search = invoicesScreenState.search.trim().toLowerCase();
  const filtered = invoices.filter((inv) => {
    if (search && !inv.invoiceNumber.toLowerCase().includes(search)) return false;
    if (invoicesScreenState.dateFrom && inv.date < invoicesScreenState.dateFrom) return false;
    if (invoicesScreenState.dateTo && inv.date > invoicesScreenState.dateTo) return false;
    if (invoicesScreenState.paymentStatusFilter !== 'all' && inv.paymentStatus !== invoicesScreenState.paymentStatusFilter) return false;
    return true;
  });

  filtered.sort((a, b) => b.createdAt.localeCompare(a.createdAt));

  if (filtered.length === 0) {
    tbody.innerHTML = '';
    emptyMessage.hidden = false;
    return;
  }
  emptyMessage.hidden = true;

  tbody.innerHTML = filtered.map((inv) => {
    const items = itemsByInvoice.get(inv.id) || [];
    return `
      <tr>
        <td>${escapeHtml(inv.invoiceNumber)}</td>
        <td>${escapeHtml(inv.date)}</td>
        <td>${invoiceSaleTypeSummary(items)}</td>
        <td>${formatCurrency(inv.subtotalHalalas, settings.currencySymbol)}</td>
        <td>${formatCurrency(inv.feePercentHalalas + inv.feeFixedHalalas, settings.currencySymbol)}</td>
        <td>${formatCurrency(inv.grandTotalHalalas, settings.currencySymbol)}</td>
        <td><span class="badge ${PAYMENT_STATUS_BADGE_CLASS[inv.paymentStatus]}">${PAYMENT_STATUS_LABELS[inv.paymentStatus]}</span></td>
        <td class="table-actions">
          <button type="button" class="link-btn" data-action="open" data-id="${inv.id}">فتح</button>
          <button type="button" class="link-btn" data-action="print" data-id="${inv.id}">طباعة</button>
        </td>
      </tr>
    `;
  }).join('');

  tbody.querySelectorAll('[data-action="open"]').forEach((btn) => {
    btn.addEventListener('click', () => openInvoiceDetailModal(Number(btn.dataset.id)));
  });
  tbody.querySelectorAll('[data-action="print"]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const invoice = invoices.find((i) => i.id === Number(btn.dataset.id));
      const items = itemsByInvoice.get(invoice.id) || [];
      printReceipt(renderInvoiceReceiptHtml(invoiceToReceiptData(invoice), items, settings));
    });
  });
}

// ---- Rendering: invoice detail modal (reopen / reprint) --------------------------

async function openInvoiceDetailModal(invoiceId) {
  const [invoice, items, settings] = await Promise.all([
    getInvoiceById(invoiceId),
    getInvoiceItemsByInvoiceId(invoiceId),
    getSettings(),
  ]);
  if (!invoice) {
    UI.error('تعذر العثور على الفاتورة.');
    return;
  }

  const receiptHtml = renderInvoiceReceiptHtml(invoiceToReceiptData(invoice), items, settings);

  UI.showModal(`
    <div class="modal modal-wide">
      <h3 class="modal-title">تفاصيل الفاتورة</h3>
      <div class="invoice-detail-preview">${receiptHtml}</div>
      <div class="modal-actions">
        <button type="button" class="btn btn-secondary" id="invoice-detail-close">إغلاق</button>
        <button type="button" class="btn btn-primary" id="invoice-detail-print">طباعة</button>
      </div>
    </div>
  `);

  document.getElementById('invoice-detail-close').addEventListener('click', () => UI.closeModal());
  document.getElementById('invoice-detail-print').addEventListener('click', () => printReceipt(receiptHtml));
}
