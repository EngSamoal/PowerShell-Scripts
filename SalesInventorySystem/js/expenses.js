/**
 * expenses.js — the "المصروفات" screen, plus the profit summary this
 * phase is named for.
 *
 * Profit is never assumed to be the platform's percentage + fixed fee —
 * those are collected on behalf of the platform, not earned margin. Real
 * profit is (sale value − cost of goods sold), and net profit subtracts
 * this screen's own expenses from that. Cost of goods is read from each
 * invoiceItem's own unitCostHalalas/lineCostHalalas snapshot (see
 * invoices.js createInvoice()), which was fixed at the moment of sale —
 * so a later change to a product's purchase price never rewrites a past
 * invoice's profit, exactly like the fee snapshot already does for totals.
 */

const EXPENSE_TYPES = [
  { value: 'delivery', label: 'التوصيل' },
  { value: 'packaging', label: 'التغليف' },
  { value: 'advertising', label: 'الإعلان' },
  { value: 'transport', label: 'النقل' },
  { value: 'other', label: 'مصروفات أخرى' },
];
const EXPENSE_TYPE_LABELS = Object.fromEntries(EXPENSE_TYPES.map((t) => [t.value, t.label]));

const expensesScreenState = {
  dateFrom: '',
  dateTo: '',
  typeFilter: 'all',
};

// ---- Data access / mutations ----------------------------------------------

async function listAllExpenses() {
  return dbGetAll('expenses');
}

async function getExpenseById(id) {
  return dbGet('expenses', id);
}

function buildExpense(fields) {
  const errors = [];
  const date = String(fields.date || '').trim();
  const description = String(fields.description || '').trim();
  const amountHalalas = sarToHalalas(fields.amount || 0);
  const expenseType = EXPENSE_TYPE_LABELS[fields.expenseType] ? fields.expenseType : 'other';

  if (!date) errors.push('يرجى إدخال تاريخ المصروف.');
  if (!description) errors.push('يرجى إدخال وصف المصروف.');
  if (Number.isNaN(amountHalalas) || amountHalalas <= 0) errors.push('يرجى إدخال مبلغ أكبر من صفر.');

  if (errors.length > 0) return { errors, expense: null };

  return {
    errors: [],
    expense: { date, expenseType, description, amountHalalas, notes: String(fields.notes || '').trim() },
  };
}

async function addExpense(fields) {
  const { errors, expense } = buildExpense(fields);
  if (errors.length > 0) return { errors };
  const now = new Date().toISOString();
  const id = await dbPut('expenses', { ...expense, createdAt: now, updatedAt: now });
  return { errors: [], id };
}

async function updateExpense(id, fields) {
  const existing = await getExpenseById(id);
  if (!existing) return { errors: ['تعذر العثور على المصروف.'] };
  const { errors, expense } = buildExpense(fields);
  if (errors.length > 0) return { errors };
  await dbPut('expenses', { ...existing, ...expense, updatedAt: new Date().toISOString() });
  return { errors: [] };
}

async function deleteExpense(id) {
  await dbDelete('expenses', id);
  return { errors: [] };
}

// ---- Profit summary ---------------------------------------------------------

/**
 * Computes the section-14 profit breakdown for [dateFrom, dateTo] (inclusive,
 * either end optional). Fees are reported separately and never folded into
 * productProfitHalalas — they are money collected on the platform's behalf,
 * not margin earned on the goods.
 */
async function computeProfitSummary(dateFrom, dateTo) {
  const [invoices, allItems, expenses] = await Promise.all([
    listAllInvoices(), dbGetAll('invoiceItems'), listAllExpenses(),
  ]);

  const invoicesInRange = invoices.filter((inv) => (
    inv.status === 'completed'
    && (!dateFrom || inv.date >= dateFrom)
    && (!dateTo || inv.date <= dateTo)
  ));
  const invoiceIds = new Set(invoicesInRange.map((inv) => inv.id));
  const itemsInRange = allItems.filter((item) => invoiceIds.has(item.invoiceId));

  const saleValueHalalas = invoicesInRange.reduce((sum, inv) => sum + inv.subtotalHalalas, 0);
  const costOfGoodsHalalas = itemsInRange.reduce((sum, item) => sum + (item.lineCostHalalas || 0), 0);
  const productProfitHalalas = saleValueHalalas - costOfGoodsHalalas;
  const percentFeeHalalas = invoicesInRange.reduce((sum, inv) => sum + inv.feePercentHalalas, 0);
  const fixedFeeHalalas = invoicesInRange.reduce((sum, inv) => sum + inv.feeFixedHalalas, 0);
  const paidByCustomerHalalas = invoicesInRange.reduce((sum, inv) => sum + inv.paidAmountHalalas, 0);

  const expensesInRange = expenses.filter((e) => (!dateFrom || e.date >= dateFrom) && (!dateTo || e.date <= dateTo));
  const totalExpensesHalalas = expensesInRange.reduce((sum, e) => sum + e.amountHalalas, 0);
  const netProfitHalalas = productProfitHalalas - totalExpensesHalalas;

  return {
    invoiceCount: invoicesInRange.length,
    saleValueHalalas,
    costOfGoodsHalalas,
    productProfitHalalas,
    percentFeeHalalas,
    fixedFeeHalalas,
    paidByCustomerHalalas,
    totalExpensesHalalas,
    netProfitHalalas,
  };
}

function currentMonthStart() {
  return `${formatDateForStorage().slice(0, 7)}-01`;
}

// ---- Rendering: "المصروفات" screen ------------------------------------------

async function renderExpensesScreen() {
  const container = document.getElementById('screen-expenses');
  const settings = await getSettings();

  if (!expensesScreenState.dateFrom) expensesScreenState.dateFrom = currentMonthStart();
  if (!expensesScreenState.dateTo) expensesScreenState.dateTo = formatDateForStorage();

  container.innerHTML = `
    <h2 class="screen-title">المصروفات</h2>

    <div class="toolbar">
      <label class="toolbar-inline"><span>من</span><input type="date" id="expense-date-from"></label>
      <label class="toolbar-inline"><span>إلى</span><input type="date" id="expense-date-to"></label>
      <select id="expense-type-filter" class="toolbar-select">
        <option value="all">كل الأنواع</option>
        ${EXPENSE_TYPES.map((t) => `<option value="${t.value}">${t.label}</option>`).join('')}
      </select>
      <div class="toolbar-spacer"></div>
      <button type="button" id="btn-add-expense" class="btn btn-primary">+ إضافة مصروف</button>
    </div>

    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr><th>التاريخ</th><th>النوع</th><th>الوصف</th><th>المبلغ</th><th>ملاحظات</th><th>إجراءات</th></tr>
        </thead>
        <tbody id="expenses-table-body"></tbody>
      </table>
      <p id="expenses-empty-message" class="table-empty-message" hidden>لا توجد مصروفات مطابقة.</p>
    </div>

    <h3 class="panel-title profit-summary-title">ملخص الأرباح</h3>
    <p class="screen-hint">للفترة المحددة أعلاه (من/إلى) — الرسوم المحصلة من العميل منفصلة دائمًا عن ربح البضاعة، ولا تُحتسب ربحًا.</p>
    <div class="profit-summary" id="profit-summary"></div>
  `;

  document.getElementById('expense-date-from').value = expensesScreenState.dateFrom;
  document.getElementById('expense-date-to').value = expensesScreenState.dateTo;
  document.getElementById('expense-type-filter').value = expensesScreenState.typeFilter;

  document.getElementById('expense-date-from').addEventListener('change', (e) => {
    expensesScreenState.dateFrom = e.target.value;
    renderExpensesTableBody(settings);
    renderProfitSummary(settings);
  });
  document.getElementById('expense-date-to').addEventListener('change', (e) => {
    expensesScreenState.dateTo = e.target.value;
    renderExpensesTableBody(settings);
    renderProfitSummary(settings);
  });
  document.getElementById('expense-type-filter').addEventListener('change', (e) => {
    expensesScreenState.typeFilter = e.target.value;
    renderExpensesTableBody(settings);
  });
  document.getElementById('btn-add-expense').addEventListener('click', () => openExpenseFormModal(null, settings));

  await renderExpensesTableBody(settings);
  await renderProfitSummary(settings);
}

async function renderExpensesTableBody(settings) {
  const tbody = document.getElementById('expenses-table-body');
  const emptyMessage = document.getElementById('expenses-empty-message');
  if (!tbody) return;

  const all = await listAllExpenses();
  const filtered = all.filter((e) => {
    if (expensesScreenState.dateFrom && e.date < expensesScreenState.dateFrom) return false;
    if (expensesScreenState.dateTo && e.date > expensesScreenState.dateTo) return false;
    if (expensesScreenState.typeFilter !== 'all' && e.expenseType !== expensesScreenState.typeFilter) return false;
    return true;
  });
  filtered.sort((a, b) => b.date.localeCompare(a.date));

  if (filtered.length === 0) {
    tbody.innerHTML = '';
    emptyMessage.hidden = false;
    return;
  }
  emptyMessage.hidden = true;

  tbody.innerHTML = filtered.map((e) => `
    <tr>
      <td>${escapeHtml(e.date)}</td>
      <td>${escapeHtml(EXPENSE_TYPE_LABELS[e.expenseType] || e.expenseType)}</td>
      <td>${escapeHtml(e.description)}</td>
      <td>${formatCurrency(e.amountHalalas, settings.currencySymbol)}</td>
      <td>${escapeHtml(e.notes)}</td>
      <td class="table-actions">
        <button type="button" class="link-btn" data-action="edit" data-id="${e.id}">تعديل</button>
        <button type="button" class="link-btn link-btn-danger" data-action="delete" data-id="${e.id}">حذف</button>
      </td>
    </tr>
  `).join('');

  tbody.querySelectorAll('[data-action="edit"]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const expense = await getExpenseById(Number(btn.dataset.id));
      openExpenseFormModal(expense, settings);
    });
  });
  tbody.querySelectorAll('[data-action="delete"]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const id = Number(btn.dataset.id);
      const confirmed = await UI.confirm({
        title: 'حذف المصروف',
        message: 'هل أنت متأكد من حذف هذا المصروف؟',
        confirmLabel: 'حذف',
        danger: true,
      });
      if (!confirmed) return;
      await deleteExpense(id);
      UI.success('تم حذف المصروف.');
      renderExpensesTableBody(settings);
      renderProfitSummary(settings);
    });
  });
}

async function renderProfitSummary(settings) {
  const box = document.getElementById('profit-summary');
  if (!box) return;

  const summary = await computeProfitSummary(expensesScreenState.dateFrom, expensesScreenState.dateTo);
  const fmt = (h) => formatCurrency(h, settings.currencySymbol);

  box.innerHTML = `
    <div class="profit-summary-row"><span>قيمة بيع البضاعة</span><strong>${fmt(summary.saleValueHalalas)}</strong></div>
    <div class="profit-summary-row"><span>تكلفة شراء البضاعة</span><strong>${fmt(summary.costOfGoodsHalalas)}</strong></div>
    <div class="profit-summary-row profit-summary-highlight"><span>ربح البضاعة</span><strong>${fmt(summary.productProfitHalalas)}</strong></div>
    <div class="profit-summary-row"><span>رسوم النسبة المحصلة</span><strong>${fmt(summary.percentFeeHalalas)}</strong></div>
    <div class="profit-summary-row"><span>الرسوم الثابتة المحصلة</span><strong>${fmt(summary.fixedFeeHalalas)}</strong></div>
    <div class="profit-summary-row"><span>إجمالي المبلغ الذي دفعه العميل</span><strong>${fmt(summary.paidByCustomerHalalas)}</strong></div>
    <div class="profit-summary-row"><span>إجمالي المصروفات</span><strong>${fmt(summary.totalExpensesHalalas)}</strong></div>
    <div class="profit-summary-row profit-summary-total"><span>الربح الصافي (بعد المصروفات)</span><strong>${fmt(summary.netProfitHalalas)}</strong></div>
  `;
}

// ---- Rendering: add/edit expense modal ---------------------------------------

function openExpenseFormModal(existingExpense, settings) {
  const isEdit = Boolean(existingExpense);
  const e = existingExpense || { date: formatDateForStorage(), expenseType: 'other', description: '', amount: 0, notes: '' };
  const amountSar = isEdit ? halalasToSar(existingExpense.amountHalalas) : 0;

  UI.showModal(`
    <div class="modal">
      <h3 class="modal-title">${isEdit ? 'تعديل مصروف' : 'إضافة مصروف'}</h3>
      <div id="expense-form-errors" class="form-errors" hidden></div>
      <form id="expense-form" class="form-grid">
        <fieldset class="form-section">
          <label class="form-field">
            <span>التاريخ</span>
            <input type="date" name="date" value="${escapeHtml(e.date)}">
          </label>
          <label class="form-field">
            <span>نوع المصروف</span>
            <select name="expenseType">
              ${EXPENSE_TYPES.map((t) => `<option value="${t.value}" ${e.expenseType === t.value ? 'selected' : ''}>${t.label}</option>`).join('')}
            </select>
          </label>
          <label class="form-field form-field-wide">
            <span>الوصف</span>
            <input type="text" name="description" value="${escapeHtml(e.description)}">
          </label>
          <label class="form-field">
            <span>المبلغ (${escapeHtml(settings.currencySymbol)})</span>
            <input type="number" step="0.01" min="0" name="amount" value="${amountSar}">
          </label>
          <label class="form-field form-field-wide">
            <span>ملاحظات</span>
            <textarea name="notes" rows="2">${escapeHtml(e.notes)}</textarea>
          </label>
        </fieldset>
        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" id="expense-form-cancel">إلغاء</button>
          <button type="submit" class="btn btn-primary">${isEdit ? 'حفظ التعديلات' : 'إضافة المصروف'}</button>
        </div>
      </form>
    </div>
  `);

  document.getElementById('expense-form-cancel').addEventListener('click', () => UI.closeModal());

  document.getElementById('expense-form').addEventListener('submit', async (event) => {
    event.preventDefault();
    const submitBtn = event.target.querySelector('button[type="submit"]');
    if (submitBtn.disabled) return; // already submitting — ignore a rapid double-click
    submitBtn.disabled = true;

    const formData = new FormData(event.target);
    const fields = Object.fromEntries(formData.entries());

    try {
      const result = isEdit ? await updateExpense(existingExpense.id, fields) : await addExpense(fields);
      if (result.errors.length > 0) {
        const errorBox = document.getElementById('expense-form-errors');
        errorBox.hidden = false;
        errorBox.innerHTML = `<ul>${result.errors.map((err) => `<li>${escapeHtml(err)}</li>`).join('')}</ul>`;
        submitBtn.disabled = false;
        return;
      }
      UI.closeModal();
      UI.success(isEdit ? 'تم حفظ تعديلات المصروف.' : 'تم إضافة المصروف بنجاح.');
      renderExpensesTableBody(settings);
      renderProfitSummary(settings);
    } catch (err) {
      UI.error(friendlyError('تعذر حفظ المصروف. يرجى المحاولة مرة أخرى.', err));
      submitBtn.disabled = false;
    }
  });
}
