/**
 * sales.js — the "بيع جديد" screen: pick sale type + product + quantity,
 * build a cart, review the live calculation, then complete the sale.
 *
 * Kept as simple as this screen is meant to be: two big sale-type
 * buttons, a filterable product list, one quantity field, and one large
 * "إتمام البيع" button that becomes unclickable the instant it's pressed
 * (isCompletingSale guard below) so a double-click can never create two
 * invoices for the same sale.
 */

let saleState = { cart: [], saleType: 'retail' };
let isCompletingSale = false;

async function renderNewSaleScreen() {
  saleState = { cart: [], saleType: 'retail' };
  isCompletingSale = false;

  const container = document.getElementById('screen-new-sale');
  const settings = await getSettings();
  const allProducts = await listAllProducts();
  const activeProducts = allProducts.filter((p) => p.status === 'active').sort((a, b) => a.name.localeCompare(b.name, 'ar'));

  if (activeProducts.length === 0) {
    container.innerHTML = `
      <h2 class="screen-title">بيع جديد</h2>
      <div class="placeholder-box">
        <p>لا توجد منتجات فعالة بعد. أضف منتجًا من شاشة "المنتجات" أولًا.</p>
      </div>
    `;
    return;
  }

  container.innerHTML = `
    <h2 class="screen-title">بيع جديد</h2>

    <div class="sale-layout">
      <div class="sale-add-panel">
        <h3 class="panel-title">إضافة منتج</h3>

        <div class="sale-type-toggle" id="sale-type-toggle">
          <button type="button" class="sale-type-btn sale-type-btn-active" data-type="retail">قطاعي</button>
          <button type="button" class="sale-type-btn" data-type="wholesale">جملة</button>
        </div>

        <label class="form-field">
          <span>ابحث عن منتج بالاسم أو الكود</span>
          <input type="text" id="sale-product-search" placeholder="اكتب هنا للبحث...">
        </label>

        <label class="form-field">
          <span>المنتج</span>
          <select id="sale-product-select"></select>
        </label>

        <div class="sale-line-info">
          <span>السعر: <strong id="sale-line-price">—</strong></span>
          <span>المتاح حاليًا: <strong id="sale-line-stock">—</strong></span>
        </div>

        <label class="form-field">
          <span>الكمية</span>
          <input type="number" id="sale-line-qty" min="1" step="1" value="1">
        </label>

        <div class="sale-line-info">
          <span>إجمالي هذا المنتج: <strong id="sale-line-total">—</strong></span>
        </div>

        <div id="sale-line-error" class="field-hint text-danger" hidden></div>

        <button type="button" id="btn-add-to-cart" class="btn btn-secondary btn-large">إضافة إلى السلة</button>
      </div>

      <div class="sale-cart-panel">
        <h3 class="panel-title">السلة</h3>
        <div class="table-wrap">
          <table class="data-table">
            <thead>
              <tr><th>المنتج</th><th>النوع</th><th>الكمية</th><th>السعر</th><th>الإجمالي</th><th></th></tr>
            </thead>
            <tbody id="sale-cart-body"></tbody>
          </table>
          <p id="sale-cart-empty" class="table-empty-message">السلة فارغة — أضف منتجًا من اليمين.</p>
        </div>

        <div class="sale-summary">
          <div class="sale-summary-row"><span>قيمة المنتجات</span><strong id="sum-subtotal">0.00 ${settings.currencySymbol}</strong></div>
          <div class="sale-summary-row"><span>رسوم الموقع (${settings.feePercent}٪)</span><strong id="sum-fee-percent">0.00 ${settings.currencySymbol}</strong></div>
          <div class="sale-summary-row"><span>الرسم الثابت</span><strong id="sum-fee-fixed">0.00 ${settings.currencySymbol}</strong></div>
          <div class="sale-summary-row sale-summary-total"><span>الإجمالي النهائي</span><strong id="sum-grand-total">0.00 ${settings.currencySymbol}</strong></div>
        </div>

        <fieldset class="form-section">
          <legend>الدفع</legend>
          <label class="form-field">
            <span>طريقة الدفع</span>
            <select id="sale-payment-method">
              <option value="cash">نقدي</option>
              <option value="bank_transfer">تحويل بنكي</option>
              <option value="card">بطاقة</option>
              <option value="other">أخرى</option>
            </select>
          </label>
          <label class="form-field">
            <span>حالة الدفع</span>
            <select id="sale-payment-status">
              <option value="paid">مدفوع</option>
              <option value="unpaid">غير مدفوع</option>
              <option value="partial">مدفوع جزئيًا</option>
            </select>
          </label>
          <div id="sale-partial-fields" class="form-field-wide" hidden>
            <label class="form-field">
              <span>المبلغ المدفوع (${settings.currencySymbol})</span>
              <input type="number" id="sale-paid-amount" min="0" step="0.01" value="0">
            </label>
            <p class="field-hint">المبلغ المتبقي: <strong id="sale-remaining-amount">0.00 ${settings.currencySymbol}</strong></p>
          </div>
          <label class="form-field form-field-wide">
            <span>ملاحظات (اختياري)</span>
            <textarea id="sale-notes" rows="2"></textarea>
          </label>
        </fieldset>

        <div id="sale-complete-error" class="form-errors" hidden></div>
        <button type="button" id="btn-complete-sale" class="btn btn-primary btn-large btn-complete-sale" disabled>إتمام البيع</button>
      </div>
    </div>
  `;

  wireAddProductPanel(activeProducts, settings);
  wirePaymentSection(settings);
  renderCartTable(settings);

  document.getElementById('btn-complete-sale').addEventListener('click', () => handleCompleteSale(settings));
}

// ---- Add-to-cart panel ---------------------------------------------------------

function wireAddProductPanel(activeProducts, settings) {
  const toggle = document.getElementById('sale-type-toggle');
  toggle.querySelectorAll('.sale-type-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      toggle.querySelectorAll('.sale-type-btn').forEach((b) => b.classList.remove('sale-type-btn-active'));
      btn.classList.add('sale-type-btn-active');
      saleState.saleType = btn.dataset.type;
      updateLinePreview(activeProducts, settings);
    });
  });

  const populateSelect = (filterText) => {
    const select = document.getElementById('sale-product-select');
    const search = (filterText || '').trim().toLowerCase();
    const filtered = search
      ? activeProducts.filter((p) => p.name.toLowerCase().includes(search) || p.sku.toLowerCase().includes(search))
      : activeProducts;
    select.innerHTML = filtered.map((p) => `<option value="${p.id}">${escapeHtml(p.sku)} — ${escapeHtml(p.name)}</option>`).join('')
      || '<option value="">لا توجد نتائج</option>';
    updateLinePreview(activeProducts, settings);
  };

  document.getElementById('sale-product-search').addEventListener('input', (e) => populateSelect(e.target.value));
  document.getElementById('sale-product-select').addEventListener('change', () => updateLinePreview(activeProducts, settings));
  document.getElementById('sale-line-qty').addEventListener('input', () => updateLinePreview(activeProducts, settings));

  populateSelect('');

  document.getElementById('btn-add-to-cart').addEventListener('click', () => addLineToCart(activeProducts, settings));
}

function getSelectedProduct(activeProducts) {
  const select = document.getElementById('sale-product-select');
  const id = Number(select.value);
  return activeProducts.find((p) => p.id === id) || null;
}

function priceForSaleType(product, saleType) {
  return saleType === 'wholesale' ? product.wholesaleBoxPriceHalalas : product.retailUnitPriceHalalas;
}

function alreadyReservedUnits(productId) {
  return saleState.cart
    .filter((line) => line.productId === productId)
    .reduce((sum, line) => sum + line.unitsPerLine, 0);
}

function updateLinePreview(activeProducts, settings) {
  const product = getSelectedProduct(activeProducts);
  const qty = Number(document.getElementById('sale-line-qty').value) || 0;
  const errorBox = document.getElementById('sale-line-error');
  errorBox.hidden = true;

  if (!product) {
    document.getElementById('sale-line-price').textContent = '—';
    document.getElementById('sale-line-stock').textContent = '—';
    document.getElementById('sale-line-total').textContent = '—';
    return;
  }

  const price = priceForSaleType(product, saleState.saleType);
  const reserved = alreadyReservedUnits(product.id);
  const remainingProduct = { ...product, totalUnits: product.totalUnits - reserved };

  document.getElementById('sale-line-price').textContent = price > 0 ? formatCurrency(price, settings.currencySymbol) : 'غير متاح بهذا النوع';
  document.getElementById('sale-line-stock').textContent = productStockDisplay(remainingProduct);

  const unitsPerLine = saleState.saleType === 'wholesale' ? qty * product.unitsPerBox : qty;
  document.getElementById('sale-line-total').textContent = formatCurrency(price * qty, settings.currencySymbol);

  if (price <= 0) {
    errorBox.hidden = false;
    errorBox.textContent = 'هذا المنتج غير متاح للبيع بهذه الطريقة.';
  } else if (unitsPerLine > remainingProduct.totalUnits) {
    errorBox.hidden = false;
    errorBox.textContent = `الكمية المطلوبة أكبر من المخزون المتاح (المتاح: ${productStockDisplay(remainingProduct)}).`;
  }
}

function addLineToCart(activeProducts, settings) {
  const product = getSelectedProduct(activeProducts);
  const qtyInput = document.getElementById('sale-line-qty');
  const qty = Math.round(Number(qtyInput.value));
  const errorBox = document.getElementById('sale-line-error');

  if (!product) {
    errorBox.hidden = false;
    errorBox.textContent = 'يرجى اختيار منتج.';
    return;
  }
  if (!Number.isInteger(qty) || qty <= 0) {
    errorBox.hidden = false;
    errorBox.textContent = 'يرجى إدخال كمية صحيحة أكبر من صفر.';
    return;
  }

  const price = priceForSaleType(product, saleState.saleType);
  if (price <= 0) {
    errorBox.hidden = false;
    errorBox.textContent = 'هذا المنتج غير متاح للبيع بهذه الطريقة.';
    return;
  }

  const unitsPerLine = saleState.saleType === 'wholesale' ? qty * product.unitsPerBox : qty;
  const reserved = alreadyReservedUnits(product.id);
  if (unitsPerLine + reserved > product.totalUnits) {
    const remaining = { ...product, totalUnits: product.totalUnits - reserved };
    errorBox.hidden = false;
    errorBox.textContent = `الكمية المطلوبة أكبر من المخزون المتاح (المتاح: ${productStockDisplay(remaining)}).`;
    return;
  }

  errorBox.hidden = true;
  saleState.cart.push({
    productId: product.id,
    productName: product.name,
    sku: product.sku,
    saleType: saleState.saleType,
    quantity: qty,
    unitPriceHalalas: price,
    lineTotalHalalas: price * qty,
    unitsPerLine,
  });

  qtyInput.value = 1;
  renderCartTable(settings);
  updateLinePreview(activeProducts, settings);
}

// ---- Cart + summary ------------------------------------------------------------

function renderCartTable(settings) {
  const tbody = document.getElementById('sale-cart-body');
  const emptyMessage = document.getElementById('sale-cart-empty');

  if (saleState.cart.length === 0) {
    tbody.innerHTML = '';
    emptyMessage.hidden = false;
  } else {
    emptyMessage.hidden = true;
    tbody.innerHTML = saleState.cart.map((line, index) => `
      <tr>
        <td>${escapeHtml(line.productName)}</td>
        <td>${line.saleType === 'wholesale' ? 'جملة' : 'قطاعي'}</td>
        <td>${line.quantity}</td>
        <td>${formatCurrency(line.unitPriceHalalas, settings.currencySymbol)}</td>
        <td>${formatCurrency(line.lineTotalHalalas, settings.currencySymbol)}</td>
        <td><button type="button" class="link-btn link-btn-danger" data-remove-index="${index}">حذف</button></td>
      </tr>
    `).join('');

    tbody.querySelectorAll('[data-remove-index]').forEach((btn) => {
      btn.addEventListener('click', () => {
        saleState.cart.splice(Number(btn.dataset.removeIndex), 1);
        renderCartTable(settings);
      });
    });
  }

  const subtotal = saleState.cart.reduce((sum, line) => sum + line.lineTotalHalalas, 0);
  const fees = calculateFees(subtotal, settings.feePercent, settings.feeFixedHalalas);

  document.getElementById('sum-subtotal').textContent = formatCurrency(subtotal, settings.currencySymbol);
  document.getElementById('sum-fee-percent').textContent = formatCurrency(fees.percentFeeHalalas, settings.currencySymbol);
  document.getElementById('sum-fee-fixed').textContent = formatCurrency(fees.fixedFeeHalalas, settings.currencySymbol);
  document.getElementById('sum-grand-total').textContent = formatCurrency(fees.grandTotalHalalas, settings.currencySymbol);

  document.getElementById('btn-complete-sale').disabled = saleState.cart.length === 0;
  updatePartialRemainingDisplay(settings);
}

// ---- Payment section ------------------------------------------------------------

function wirePaymentSection(settings) {
  document.getElementById('sale-payment-status').addEventListener('change', (e) => {
    document.getElementById('sale-partial-fields').hidden = e.target.value !== 'partial';
    updatePartialRemainingDisplay(settings);
  });
  document.getElementById('sale-paid-amount').addEventListener('input', () => updatePartialRemainingDisplay(settings));
}

function currentGrandTotalHalalas(settings) {
  const subtotal = saleState.cart.reduce((sum, line) => sum + line.lineTotalHalalas, 0);
  return calculateFees(subtotal, settings.feePercent, settings.feeFixedHalalas).grandTotalHalalas;
}

function updatePartialRemainingDisplay(settings) {
  const statusEl = document.getElementById('sale-payment-status');
  if (!statusEl || statusEl.value !== 'partial') return;
  const grandTotal = currentGrandTotalHalalas(settings);
  const paid = sarToHalalas(document.getElementById('sale-paid-amount').value || 0);
  const remaining = Math.max(0, grandTotal - paid);
  document.getElementById('sale-remaining-amount').textContent = formatCurrency(remaining, settings.currencySymbol);
}

// ---- Completing the sale ---------------------------------------------------------

async function handleCompleteSale(settings) {
  if (isCompletingSale) return;
  isCompletingSale = true;

  const btn = document.getElementById('btn-complete-sale');
  const errorBox = document.getElementById('sale-complete-error');
  errorBox.hidden = true;
  btn.disabled = true;
  const originalLabel = btn.textContent;
  btn.textContent = 'جارٍ الحفظ...';

  const releaseGuard = () => {
    isCompletingSale = false;
    btn.disabled = false;
    btn.textContent = originalLabel;
  };

  try {
    const paymentMethod = document.getElementById('sale-payment-method').value;
    const paymentStatus = document.getElementById('sale-payment-status').value;
    const notes = document.getElementById('sale-notes').value;
    const paidAmountHalalas = paymentStatus === 'partial' ? sarToHalalas(document.getElementById('sale-paid-amount').value || 0) : undefined;

    const meta = { customerId: null, paymentMethod, paymentStatus, paidAmountHalalas, notes };
    const result = await createInvoice(saleState.cart, meta, settings);

    if (result.errors.length > 0) {
      errorBox.hidden = false;
      errorBox.innerHTML = `<ul>${result.errors.map((e) => `<li>${escapeHtml(e)}</li>`).join('')}</ul>`;
      releaseGuard();
      return;
    }

    renderSaleSuccessView(result, saleState.cart, settings);
  } catch (err) {
    UI.error(friendlyError('تعذر إتمام عملية البيع. يرجى المحاولة مرة أخرى.', err));
    releaseGuard();
  }
}

function renderSaleSuccessView(invoiceResult, items, settings) {
  const container = document.getElementById('screen-new-sale');
  const receiptHtml = renderInvoiceReceiptHtml(invoiceResult, items, settings);

  container.innerHTML = `
    <div class="sale-success">
      <h2 class="screen-title">تم إتمام البيع بنجاح ✔</h2>
      <p class="screen-hint">رقم الفاتورة: <strong>${escapeHtml(invoiceResult.invoiceNumber)}</strong></p>
      <p class="screen-hint">الإجمالي النهائي: <strong>${formatCurrency(invoiceResult.grandTotalHalalas, settings.currencySymbol)}</strong></p>
      <div class="modal-actions">
        <button type="button" id="btn-print-receipt" class="btn btn-secondary">طباعة الفاتورة</button>
        <button type="button" id="btn-new-sale-again" class="btn btn-primary btn-large">بيع جديد</button>
      </div>
    </div>
  `;

  document.getElementById('btn-print-receipt').addEventListener('click', () => printReceipt(receiptHtml));
  document.getElementById('btn-new-sale-again').addEventListener('click', () => renderNewSaleScreen());
}
