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
        // subtotal/fee/grandTotal are the CURRENT net figures — returnInvoiceItems()
        // reduces them as items are returned. The original* fields are a frozen
        // snapshot of the sale as it happened, and are never mutated again; they're
        // the fixed baseline proportional fee refunds are computed against, so a
        // second partial return on the same invoice still refunds correctly.
        subtotalHalalas,
        feePercentUsed: settings.feePercent,
        feePercentHalalas: fees.percentFeeHalalas,
        feeFixedHalalas: fees.fixedFeeHalalas,
        grandTotalHalalas: fees.grandTotalHalalas,
        originalSubtotalHalalas: subtotalHalalas,
        originalFeePercentHalalas: fees.percentFeeHalalas,
        originalFeeFixedHalalas: fees.fixedFeeHalalas,
        originalGrandTotalHalalas: fees.grandTotalHalalas,
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
        const unitCostHalalas = item.unitCostHalalas || 0;
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
          // Cost snapshot at sale time (per base unit, and for the whole line) — so a later
          // change to the product's purchase cost never alters a past invoice's profit.
          unitCostHalalas,
          lineCostHalalas: unitCostHalalas * item.unitsPerLine,
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

// ---- Void / Returns ---------------------------------------------------------------
//
// Voiding and returning are deliberately different operations:
//
// - void: the whole sale is cancelled outright (created by mistake, customer
//   walked away, etc). Stock is restored, the invoice is tagged 'voided', and
//   every report/profit/dashboard figure (which all filter `status !== 'voided'`)
//   drops it entirely — its subtotal/fees are left untouched so the Invoices
//   list can still show what it WAS for, just visibly marked cancelled.
//
// - return (full or partial): the sale genuinely happened; some or all of the
//   goods are coming back later. Because a partial return must keep counting
//   the *unreturned* remainder as real revenue, this can't just flip a status
//   filter — it reduces the touched invoiceItems' quantity/value/cost and the
//   invoice's own net subtotal/fees/grandTotal in place (mirroring how
//   products.totalUnits is the single mutable source of truth for stock, with
//   an audit trail alongside it). Every report already reads invoiceItems and
//   invoice totals directly, so once these are reduced correctly nothing else
//   needs to change to reflect a return. The `returns` store is that audit
//   trail: an immutable record of exactly what was returned, when, and for
//   how much — the original invoiceItems values before this function touches
//   them are only ever visible through it.

/** Restores an invoice's still-outstanding stock and marks it voided. Never allowed on an already-voided or fully-returned invoice. */
async function voidInvoice(invoiceId, reason) {
  try {
    await dbTransaction(['invoices', 'invoiceItems', 'products'], async (tx) => {
      const invoicesStore = tx.objectStore('invoices');
      const invoice = await requestToPromise(invoicesStore.get(invoiceId));
      if (!invoice) throw new InvoiceError('الفاتورة غير موجودة.');
      if (invoice.status === 'voided') throw new InvoiceError('هذه الفاتورة ملغاة بالفعل.');
      if (invoice.status === 'returned') throw new InvoiceError('لا يمكن إلغاء فاتورة تم إرجاعها بالكامل بالفعل.');

      const items = await requestToPromise(tx.objectStore('invoiceItems').index('invoiceId').getAll(invoiceId));
      const productsStore = tx.objectStore('products');
      for (const item of items) {
        if (item.unitsPerLine <= 0) continue; // already fully returned earlier — nothing left to restore
        const product = await requestToPromise(productsStore.get(item.productId));
        if (!product) continue; // product was deleted since the sale; nothing to restore stock to
        product.totalUnits += item.unitsPerLine;
        product.updatedAt = new Date().toISOString();
        productsStore.put(product);
      }

      invoice.status = 'voided';
      invoice.voidedAt = new Date().toISOString();
      invoice.voidReason = reason || '';
      invoicesStore.put(invoice);
    });
    return { errors: [] };
  } catch (err) {
    if (err instanceof InvoiceError) return { errors: [err.message] };
    throw err;
  }
}

/**
 * Returns some or all of an invoice's items. `returnLines` is
 * [{ invoiceItemId, quantityToReturn }] in the item's own unit (pieces for a
 * retail line, boxes for a wholesale line) — pass every current line at its
 * full remaining quantity for a full return.
 *
 * Fee refunds are proportional to originalSubtotalHalalas (the frozen
 * baseline from createInvoice), so a second partial return on the same
 * invoice still refunds correctly against what's genuinely left.
 */
async function returnInvoiceItems(invoiceId, returnLines, reason) {
  const positiveLines = (returnLines || []).filter((l) => l.quantityToReturn > 0);
  if (positiveLines.length === 0) {
    return { errors: ['يرجى تحديد كمية إرجاع واحدة على الأقل.'] };
  }

  try {
    const result = await dbTransaction(['invoices', 'invoiceItems', 'products', 'returns'], async (tx) => {
      const invoicesStore = tx.objectStore('invoices');
      const invoice = await requestToPromise(invoicesStore.get(invoiceId));
      if (!invoice) throw new InvoiceError('الفاتورة غير موجودة.');
      if (invoice.status === 'voided') throw new InvoiceError('لا يمكن إرجاع فاتورة ملغاة.');
      if (invoice.status === 'returned') throw new InvoiceError('تم إرجاع هذه الفاتورة بالكامل مسبقًا.');

      const invoiceItemsStore = tx.objectStore('invoiceItems');
      const productsStore = tx.objectStore('products');

      let subtotalReturnedHalalas = 0;
      const returnItemRecords = [];
      const productDeltas = new Map();

      // Pass 1: validate every line against its live invoiceItem and compute what
      // it means, but issue no writes yet — so a bad line anywhere in the request
      // leaves the database completely untouched, not partially applied.
      const plannedUpdates = [];
      for (const line of positiveLines) {
        const item = await requestToPromise(invoiceItemsStore.get(line.invoiceItemId));
        if (!item || item.invoiceId !== invoiceId) throw new InvoiceError('أحد عناصر الفاتورة غير موجود.');
        if (line.quantityToReturn > item.quantity) {
          throw new InvoiceError(`الكمية المطلوب إرجاعها من "${item.productName}" أكبر من الكمية القابلة للإرجاع حاليًا (${item.quantity}).`);
        }

        const unitsPerSingleQty = item.unitsPerLine / item.quantity; // exact: 1 for retail, unitsPerBox for wholesale
        const unitsToRestore = line.quantityToReturn * unitsPerSingleQty;
        const valueReturnedHalalas = line.quantityToReturn * item.unitPriceHalalas;
        const costReturnedHalalas = unitsToRestore * item.unitCostHalalas;
        plannedUpdates.push({ item, line, unitsToRestore, valueReturnedHalalas, costReturnedHalalas });
      }

      // Pass 2: everything validated — now it's safe to write.
      for (const { item, line, unitsToRestore, valueReturnedHalalas, costReturnedHalalas } of plannedUpdates) {
        item.quantity -= line.quantityToReturn;
        item.unitsPerLine -= unitsToRestore;
        item.lineTotalHalalas -= valueReturnedHalalas;
        item.lineCostHalalas -= costReturnedHalalas;
        invoiceItemsStore.put(item);

        productDeltas.set(item.productId, (productDeltas.get(item.productId) || 0) + unitsToRestore);
        subtotalReturnedHalalas += valueReturnedHalalas;
        returnItemRecords.push({
          invoiceItemId: item.id, productId: item.productId, productName: item.productName,
          saleType: item.saleType, quantityReturned: line.quantityToReturn, unitsReturned: unitsToRestore,
          valueReturnedHalalas, costReturnedHalalas,
        });
      }

      for (const [productId, units] of productDeltas) {
        const product = await requestToPromise(productsStore.get(productId));
        if (!product) continue;
        product.totalUnits += units;
        product.updatedAt = new Date().toISOString();
        productsStore.put(product);
      }

      const fraction = subtotalReturnedHalalas / invoice.originalSubtotalHalalas;
      const percentFeeRefundHalalas = roundHalfUp(invoice.originalFeePercentHalalas * fraction);
      const fixedFeeRefundHalalas = roundHalfUp(invoice.originalFeeFixedHalalas * fraction);
      const totalRefundedHalalas = subtotalReturnedHalalas + percentFeeRefundHalalas + fixedFeeRefundHalalas;

      invoice.subtotalHalalas -= subtotalReturnedHalalas;
      invoice.feePercentHalalas -= percentFeeRefundHalalas;
      invoice.feeFixedHalalas -= fixedFeeRefundHalalas;
      invoice.grandTotalHalalas = invoice.subtotalHalalas + invoice.feePercentHalalas + invoice.feeFixedHalalas;
      // Assumes a return also refunds cash already collected for that portion, so paid/remaining move together with the total.
      invoice.paidAmountHalalas = Math.max(0, invoice.paidAmountHalalas - totalRefundedHalalas);
      invoice.remainingAmountHalalas = Math.max(0, invoice.grandTotalHalalas - invoice.paidAmountHalalas);

      const remainingItems = await requestToPromise(invoiceItemsStore.index('invoiceId').getAll(invoiceId));
      const fullyReturned = remainingItems.every((i) => i.quantity <= 0);
      invoice.status = fullyReturned ? 'returned' : 'partially_returned';
      invoicesStore.put(invoice);

      const now = new Date();
      const returnRecord = {
        invoiceId, invoiceNumber: invoice.invoiceNumber,
        date: formatDateForStorage(now), createdAt: now.toISOString(),
        type: fullyReturned ? 'full' : 'partial',
        items: returnItemRecords,
        subtotalRefundedHalalas: subtotalReturnedHalalas,
        percentFeeRefundedHalalas: percentFeeRefundHalalas,
        fixedFeeRefundedHalalas: fixedFeeRefundHalalas,
        totalRefundedHalalas,
        reason: reason || '',
      };
      const returnId = await requestToPromise(tx.objectStore('returns').add(returnRecord));

      return { returnId, totalRefundedHalalas, status: invoice.status };
    });

    return { errors: [], ...result };
  } catch (err) {
    if (err instanceof InvoiceError) return { errors: [err.message] };
    throw err;
  }
}

async function getReturnsForInvoice(invoiceId) {
  return dbGetAllByIndex('returns', 'invoiceId', invoiceId);
}

/** Builds the printable receipt/invoice HTML for a just-created (or reopened) invoice. Shared by "بيع جديد" and the Invoices screen. */
function renderInvoiceReceiptHtml(invoice, items, settings, customer) {
  return `
    <div class="receipt">
      <h2>${escapeHtml(settings.businessName || 'نظام المبيعات')}</h2>
      ${settings.phone ? `<p>${escapeHtml(settings.phone)}</p>` : ''}
      <p>فاتورة رقم: ${escapeHtml(invoice.invoiceNumber)}</p>
      <p>التاريخ: ${escapeHtml(invoice.dateDisplay)}</p>
      ${customer ? `<p>العميل: ${escapeHtml(customer.name)}${customer.phone ? ' — ' + escapeHtml(customer.phone) : ''}</p>` : ''}
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

const INVOICE_STATUS_LABELS = {
  completed: 'مكتملة', voided: 'ملغاة', returned: 'مرتجعة بالكامل', partially_returned: 'مرتجعة جزئيًا',
};
const INVOICE_STATUS_BADGE_CLASS = {
  completed: 'badge-success', voided: 'badge-danger', returned: 'badge-muted', partially_returned: 'badge-info',
};

/** Void/return actions only make sense while there's still something left to void or return. */
function canVoidOrReturn(invoice) {
  return invoice.status === 'completed' || invoice.status === 'partially_returned';
}

async function renderInvoicesScreen() {
  const container = document.getElementById('screen-invoices');

  container.innerHTML = `
    <h2 class="screen-title">الفواتير</h2>
    <div class="toolbar">
      <input type="search" id="invoice-search" class="toolbar-search" placeholder="ابحث برقم الفاتورة أو اسم العميل...">
      <label class="toolbar-inline"><span>من</span><input type="date" id="invoice-date-from"></label>
      <label class="toolbar-inline"><span>إلى</span><input type="date" id="invoice-date-to"></label>
      <select id="invoice-status-filter" class="toolbar-select">
        <option value="all">كل حالات الدفع</option>
        <option value="paid">مدفوع</option>
        <option value="unpaid">غير مدفوع</option>
        <option value="partial">مدفوع جزئيًا</option>
      </select>
      <div class="toolbar-spacer"></div>
      <button type="button" id="btn-export-invoices" class="btn btn-secondary">تصدير Excel</button>
    </div>
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>رقم الفاتورة</th><th>التاريخ</th><th>العميل</th><th>نوع البيع</th>
            <th>قيمة المنتجات</th><th>الرسوم</th><th>الإجمالي</th><th>حالة الدفع</th><th>حالة الفاتورة</th><th>إجراءات</th>
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
  document.getElementById('btn-export-invoices').addEventListener('click', async () => {
    try {
      await exportVisibleInvoicesToExcel();
      UI.success('تم تصدير ملف الفواتير بنجاح.');
    } catch (err) {
      UI.error(friendlyError('تعذر تصدير الملف. يرجى المحاولة مرة أخرى.', err));
    }
  });

  await renderInvoicesTableBody();
}

/** Builds the id → customer map once per render, and a "بدون عميل" fallback for invoices with no customer attached. */
async function buildCustomerLookup() {
  const customers = await listAllCustomers();
  const byId = new Map(customers.map((c) => [c.id, c]));
  return byId;
}

function getFilteredInvoices(invoices, itemsByInvoice, customersById) {
  const search = invoicesScreenState.search.trim().toLowerCase();
  return invoices.filter((inv) => {
    if (search) {
      const customerName = inv.customerId ? (customersById.get(inv.customerId)?.name || '') : '';
      const matchesNumber = inv.invoiceNumber.toLowerCase().includes(search);
      const matchesCustomer = customerName.toLowerCase().includes(search);
      if (!matchesNumber && !matchesCustomer) return false;
    }
    if (invoicesScreenState.dateFrom && inv.date < invoicesScreenState.dateFrom) return false;
    if (invoicesScreenState.dateTo && inv.date > invoicesScreenState.dateTo) return false;
    if (invoicesScreenState.paymentStatusFilter !== 'all' && inv.paymentStatus !== invoicesScreenState.paymentStatusFilter) return false;
    return true;
  }).sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

async function renderInvoicesTableBody() {
  const tbody = document.getElementById('invoices-table-body');
  const emptyMessage = document.getElementById('invoices-empty-message');
  if (!tbody) return;

  const settings = await getSettings();
  const [invoices, itemsByInvoice, customersById] = await Promise.all([
    listAllInvoices(), groupAllInvoiceItemsByInvoiceId(), buildCustomerLookup(),
  ]);

  const filtered = getFilteredInvoices(invoices, itemsByInvoice, customersById);

  if (filtered.length === 0) {
    tbody.innerHTML = '';
    emptyMessage.hidden = false;
    return;
  }
  emptyMessage.hidden = true;

  tbody.innerHTML = filtered.map((inv) => {
    const items = itemsByInvoice.get(inv.id) || [];
    const customer = inv.customerId ? customersById.get(inv.customerId) : null;
    return `
      <tr>
        <td>${escapeHtml(inv.invoiceNumber)}</td>
        <td>${escapeHtml(inv.date)}</td>
        <td>${customer ? escapeHtml(customer.name) : '—'}</td>
        <td>${invoiceSaleTypeSummary(items)}</td>
        <td>${formatCurrency(inv.subtotalHalalas, settings.currencySymbol)}</td>
        <td>${formatCurrency(inv.feePercentHalalas + inv.feeFixedHalalas, settings.currencySymbol)}</td>
        <td>${formatCurrency(inv.grandTotalHalalas, settings.currencySymbol)}</td>
        <td><span class="badge ${PAYMENT_STATUS_BADGE_CLASS[inv.paymentStatus]}">${PAYMENT_STATUS_LABELS[inv.paymentStatus]}</span></td>
        <td><span class="badge ${INVOICE_STATUS_BADGE_CLASS[inv.status]}">${INVOICE_STATUS_LABELS[inv.status] || inv.status}</span></td>
        <td class="table-actions">
          <button type="button" class="link-btn" data-action="open" data-id="${inv.id}">فتح</button>
          <button type="button" class="link-btn" data-action="print" data-id="${inv.id}">طباعة</button>
          ${canVoidOrReturn(inv) ? `
            <button type="button" class="link-btn" data-action="return" data-id="${inv.id}">مرتجع</button>
            <button type="button" class="link-btn link-btn-danger" data-action="void" data-id="${inv.id}">إلغاء</button>
          ` : ''}
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
      const customer = invoice.customerId ? customersById.get(invoice.customerId) : null;
      printReceipt(renderInvoiceReceiptHtml(invoiceToReceiptData(invoice), items, settings, customer));
    });
  });
  tbody.querySelectorAll('[data-action="return"]').forEach((btn) => {
    btn.addEventListener('click', () => openReturnModal(Number(btn.dataset.id)));
  });
  tbody.querySelectorAll('[data-action="void"]').forEach((btn) => {
    btn.addEventListener('click', () => handleVoidInvoice(Number(btn.dataset.id)));
  });
}

async function handleVoidInvoice(invoiceId) {
  const invoice = await getInvoiceById(invoiceId);
  const confirmed = await UI.confirm({
    title: 'إلغاء الفاتورة',
    message: `سيتم إلغاء الفاتورة رقم "${escapeHtml(invoice.invoiceNumber)}" بالكامل وإعادة كامل كمياتها إلى المخزون. لا يمكن التراجع عن هذا الإجراء. هل أنت متأكد؟`,
    confirmLabel: 'إلغاء الفاتورة',
    danger: true,
  });
  if (!confirmed) return;

  const result = await voidInvoice(invoiceId);
  if (result.errors.length > 0) {
    UI.error(result.errors[0]);
    return;
  }
  UI.success('تم إلغاء الفاتورة وإعادة الكمية إلى المخزون.');
  renderInvoicesTableBody();
}

/** Exports whatever the current search/date/status filters are showing — "export what I'm looking at". */
async function exportVisibleInvoicesToExcel() {
  const settings = await getSettings();
  const [invoices, itemsByInvoice, customersById] = await Promise.all([
    listAllInvoices(), groupAllInvoiceItemsByInvoiceId(), buildCustomerLookup(),
  ]);
  const filtered = getFilteredInvoices(invoices, itemsByInvoice, customersById);

  const headers = ['رقم الفاتورة', 'التاريخ', 'العميل', 'نوع البيع', 'قيمة المنتجات', 'رسوم النسبة', 'الرسم الثابت', 'الإجمالي', 'طريقة الدفع', 'حالة الدفع', 'المدفوع', 'المتبقي', 'ملاحظات'];
  const rows = filtered.map((inv) => {
    const items = itemsByInvoice.get(inv.id) || [];
    const customer = inv.customerId ? customersById.get(inv.customerId) : null;
    return [
      inv.invoiceNumber, inv.date, customer ? customer.name : '',
      invoiceSaleTypeSummary(items),
      halalasToSar(inv.subtotalHalalas), halalasToSar(inv.feePercentHalalas), halalasToSar(inv.feeFixedHalalas),
      halalasToSar(inv.grandTotalHalalas), PAYMENT_METHOD_LABELS[inv.paymentMethod] || inv.paymentMethod,
      PAYMENT_STATUS_LABELS[inv.paymentStatus] || inv.paymentStatus,
      halalasToSar(inv.paidAmountHalalas), halalasToSar(inv.remainingAmountHalalas), inv.notes,
    ];
  });

  exportRowsToExcel(`الفواتير_${formatDateForStorage()}.xlsx`, 'الفواتير', headers, rows);
}

// ---- Rendering: invoice detail modal (reopen / reprint) --------------------------

async function openInvoiceDetailModal(invoiceId) {
  const [invoice, items, settings, returns] = await Promise.all([
    getInvoiceById(invoiceId),
    getInvoiceItemsByInvoiceId(invoiceId),
    getSettings(),
    getReturnsForInvoice(invoiceId),
  ]);
  if (!invoice) {
    UI.error('تعذر العثور على الفاتورة.');
    return;
  }

  const customer = invoice.customerId ? await getCustomerById(invoice.customerId) : null;
  const receiptHtml = renderInvoiceReceiptHtml(invoiceToReceiptData(invoice), items, settings, customer);

  const statusLine = `<span class="badge ${INVOICE_STATUS_BADGE_CLASS[invoice.status]}">${INVOICE_STATUS_LABELS[invoice.status] || invoice.status}</span>`;
  const returnsHistoryHtml = returns.length > 0 ? `
    <h4 class="panel-title">سجل المرتجعات</h4>
    <ul class="returns-history">
      ${returns.map((r) => `
        <li>
          ${escapeHtml(formatDateTimeForDisplay(new Date(r.createdAt)))} —
          ${r.type === 'full' ? 'مرتجع كامل' : 'مرتجع جزئي'} —
          استُرجع ${formatCurrency(r.totalRefundedHalalas, settings.currencySymbol)}
          (${r.items.map((i) => `${escapeHtml(i.productName)} × ${i.quantityReturned}`).join('، ')})
        </li>
      `).join('')}
    </ul>
  ` : '';
  const voidInfoHtml = invoice.status === 'voided'
    ? `<p class="field-hint">تم الإلغاء في: ${escapeHtml(formatDateTimeForDisplay(new Date(invoice.voidedAt)))}</p>`
    : '';

  UI.showModal(`
    <div class="modal modal-wide">
      <h3 class="modal-title">تفاصيل الفاتورة ${statusLine}</h3>
      ${voidInfoHtml}
      <div class="invoice-detail-preview">${receiptHtml}</div>
      ${returnsHistoryHtml}
      <div class="modal-actions">
        <button type="button" class="btn btn-secondary" id="invoice-detail-close">إغلاق</button>
        <button type="button" class="btn btn-primary" id="invoice-detail-print">طباعة</button>
        ${canVoidOrReturn(invoice) ? `
          <button type="button" class="btn btn-secondary" id="invoice-detail-return">مرتجع</button>
          <button type="button" class="btn btn-danger" id="invoice-detail-void">إلغاء الفاتورة</button>
        ` : ''}
      </div>
    </div>
  `);

  document.getElementById('invoice-detail-close').addEventListener('click', () => UI.closeModal());
  document.getElementById('invoice-detail-print').addEventListener('click', () => printReceipt(receiptHtml));
  const returnBtn = document.getElementById('invoice-detail-return');
  if (returnBtn) returnBtn.addEventListener('click', () => { UI.closeModal(); openReturnModal(invoiceId); });
  const voidBtn = document.getElementById('invoice-detail-void');
  if (voidBtn) voidBtn.addEventListener('click', () => { UI.closeModal(); handleVoidInvoice(invoiceId); });
}

// ---- Rendering: return modal ----------------------------------------------------

async function openReturnModal(invoiceId) {
  const [invoice, items, settings] = await Promise.all([
    getInvoiceById(invoiceId), getInvoiceItemsByInvoiceId(invoiceId), getSettings(),
  ]);
  const returnableItems = items.filter((i) => i.quantity > 0);

  if (returnableItems.length === 0) {
    UI.error('لا توجد كميات قابلة للإرجاع في هذه الفاتورة.');
    return;
  }

  UI.showModal(`
    <div class="modal modal-wide">
      <h3 class="modal-title">مرتجع: ${escapeHtml(invoice.invoiceNumber)}</h3>
      <p class="modal-message">أدخل الكمية المطلوب إرجاعها من كل منتج (اتركها 0 لعدم إرجاعه)، أو اضغط "إرجاع الكل" لمرتجع كامل.</p>
      <div id="return-form-errors" class="form-errors" hidden></div>
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr><th>المنتج</th><th>النوع</th><th>الكمية المباعة القابلة للإرجاع</th><th>كمية الإرجاع</th></tr></thead>
          <tbody>
            ${returnableItems.map((item) => `
              <tr>
                <td>${escapeHtml(item.productName)}</td>
                <td>${item.saleType === 'wholesale' ? 'جملة' : 'قطاعي'}</td>
                <td>${item.quantity}</td>
                <td>
                  <input type="number" class="return-qty-input" data-invoice-item-id="${item.id}" data-max="${item.quantity}" min="0" max="${item.quantity}" step="1" value="0" style="width:80px">
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
      <label class="form-field form-field-wide">
        <span>سبب الإرجاع (اختياري)</span>
        <textarea id="return-reason" rows="2"></textarea>
      </label>
      <div class="modal-actions">
        <button type="button" class="btn btn-secondary" id="return-cancel">إلغاء</button>
        <button type="button" class="btn btn-secondary" id="return-all">إرجاع الكل</button>
        <button type="button" class="btn btn-primary" id="return-confirm">تأكيد المرتجع</button>
      </div>
    </div>
  `);

  document.getElementById('return-cancel').addEventListener('click', () => UI.closeModal());
  document.getElementById('return-all').addEventListener('click', () => {
    document.querySelectorAll('.return-qty-input').forEach((input) => {
      input.value = input.dataset.max;
    });
  });

  document.getElementById('return-confirm').addEventListener('click', async (event) => {
    const submitBtn = event.currentTarget;
    if (submitBtn.disabled) return;
    submitBtn.disabled = true;

    const errorBox = document.getElementById('return-form-errors');
    const returnLines = Array.from(document.querySelectorAll('.return-qty-input'))
      .map((input) => ({ invoiceItemId: Number(input.dataset.invoiceItemId), quantityToReturn: Math.round(Number(input.value)) || 0 }));

    const reason = document.getElementById('return-reason').value;
    const result = await returnInvoiceItems(invoiceId, returnLines, reason);

    if (result.errors.length > 0) {
      errorBox.hidden = false;
      errorBox.innerHTML = `<ul>${result.errors.map((e) => `<li>${escapeHtml(e)}</li>`).join('')}</ul>`;
      submitBtn.disabled = false;
      return;
    }

    UI.closeModal();
    UI.success(`تم تسجيل المرتجع بنجاح. المبلغ المسترجع: ${formatCurrency(result.totalRefundedHalalas, settings.currencySymbol)}`);
    renderInvoicesTableBody();
  });
}
