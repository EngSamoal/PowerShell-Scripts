/**
 * customers.js — the "العملاء" screen and the customer picker used from
 * "بيع جديد". Customer info is always optional at the point of sale (a
 * sale never requires one), but a record added from this screen needs at
 * least a name to be worth keeping.
 */

const customersScreenState = { search: '' };

// ---- Data access --------------------------------------------------------

async function listAllCustomers() {
  return dbGetAll('customers');
}

async function getCustomerById(id) {
  return dbGet('customers', id);
}

async function customerHasInvoiceHistory(customerId) {
  const invoices = await dbGetAllByIndex('invoices', 'customerId', customerId);
  return invoices.length > 0;
}

// ---- Validation / mutations -----------------------------------------------

function buildCustomer(fields) {
  const errors = [];
  const name = String(fields.name || '').trim();
  if (!name) errors.push('يرجى إدخال اسم العميل.');
  if (errors.length > 0) return { errors, customer: null };

  return {
    errors: [],
    customer: {
      name,
      phone: String(fields.phone || '').trim(),
      whatsapp: String(fields.whatsapp || '').trim(),
      address: String(fields.address || '').trim(),
      customerType: fields.customerType === 'wholesale' ? 'wholesale' : 'retail',
      notes: String(fields.notes || '').trim(),
    },
  };
}

async function addCustomer(fields) {
  const { errors, customer } = buildCustomer(fields);
  if (errors.length > 0) return { errors };
  const now = new Date().toISOString();
  const id = await dbPut('customers', { ...customer, createdAt: now, updatedAt: now });
  return { errors: [], id };
}

async function updateCustomer(id, fields) {
  const existing = await getCustomerById(id);
  if (!existing) return { errors: ['تعذر العثور على العميل.'] };
  const { errors, customer } = buildCustomer(fields);
  if (errors.length > 0) return { errors };
  await dbPut('customers', { ...existing, ...customer, updatedAt: new Date().toISOString() });
  return { errors: [] };
}

async function deleteCustomerIfUnused(id) {
  const used = await customerHasInvoiceHistory(id);
  if (used) {
    return { errors: ['لا يمكن حذف هذا العميل لأنه مرتبط بفواتير سابقة.'] };
  }
  await dbDelete('customers', id);
  return { errors: [] };
}

// ---- Rendering: "العملاء" screen -----------------------------------------

async function renderCustomersScreen() {
  const container = document.getElementById('screen-customers');

  container.innerHTML = `
    <h2 class="screen-title">العملاء</h2>
    <p class="screen-hint">بيانات العميل اختيارية دائمًا — يمكن إتمام أي عملية بيع بدون تسجيل عميل.</p>

    <div class="toolbar">
      <input type="search" id="customer-search" class="toolbar-search" placeholder="ابحث بالاسم أو رقم الجوال...">
      <div class="toolbar-spacer"></div>
      <button type="button" id="btn-add-customer" class="btn btn-primary">+ إضافة عميل</button>
    </div>

    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr><th>الاسم</th><th>الجوال</th><th>واتساب</th><th>النوع</th><th>ملاحظات</th><th>إجراءات</th></tr>
        </thead>
        <tbody id="customers-table-body"></tbody>
      </table>
      <p id="customers-empty-message" class="table-empty-message" hidden>لا يوجد عملاء مطابقون.</p>
    </div>
  `;

  document.getElementById('customer-search').value = customersScreenState.search;
  document.getElementById('customer-search').addEventListener('input', (e) => {
    customersScreenState.search = e.target.value;
    renderCustomersTableBody();
  });
  document.getElementById('btn-add-customer').addEventListener('click', () => {
    openCustomerFormModal(null, () => renderCustomersTableBody());
  });

  await renderCustomersTableBody();
}

async function renderCustomersTableBody() {
  const tbody = document.getElementById('customers-table-body');
  const emptyMessage = document.getElementById('customers-empty-message');
  if (!tbody) return;

  const all = await listAllCustomers();
  const search = customersScreenState.search.trim().toLowerCase();
  const filtered = all.filter((c) => {
    if (!search) return true;
    return c.name.toLowerCase().includes(search) || (c.phone || '').toLowerCase().includes(search);
  });
  filtered.sort((a, b) => a.name.localeCompare(b.name, 'ar'));

  if (filtered.length === 0) {
    tbody.innerHTML = '';
    emptyMessage.hidden = false;
    return;
  }
  emptyMessage.hidden = true;

  tbody.innerHTML = filtered.map((c) => `
    <tr>
      <td>${escapeHtml(c.name)}</td>
      <td>${escapeHtml(c.phone)}</td>
      <td>${escapeHtml(c.whatsapp)}</td>
      <td>${c.customerType === 'wholesale' ? 'جملة' : 'قطاعي'}</td>
      <td>${escapeHtml(c.notes)}</td>
      <td class="table-actions">
        <button type="button" class="link-btn" data-action="edit" data-id="${c.id}">تعديل</button>
        <button type="button" class="link-btn link-btn-danger" data-action="delete" data-id="${c.id}">حذف</button>
      </td>
    </tr>
  `).join('');

  tbody.querySelectorAll('[data-action="edit"]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const customer = await getCustomerById(Number(btn.dataset.id));
      openCustomerFormModal(customer, () => renderCustomersTableBody());
    });
  });
  tbody.querySelectorAll('[data-action="delete"]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const id = Number(btn.dataset.id);
      const customer = await getCustomerById(id);
      const confirmed = await UI.confirm({
        title: 'حذف العميل',
        message: `هل أنت متأكد من حذف العميل "${escapeHtml(customer.name)}"؟`,
        confirmLabel: 'حذف',
        danger: true,
      });
      if (!confirmed) return;
      const result = await deleteCustomerIfUnused(id);
      if (result.errors.length > 0) {
        UI.error(result.errors[0]);
        return;
      }
      UI.success('تم حذف العميل.');
      renderCustomersTableBody();
    });
  });
}

// ---- Rendering: add/edit customer modal (reused from "بيع جديد" too) -----------

/**
 * `onSaved(id)` runs after a successful add/edit — the Customers screen uses it
 * to refresh its table, while "بيع جديد" uses it to repopulate and select the
 * customer picker without leaving the sale in progress.
 */
function openCustomerFormModal(existingCustomer, onSaved) {
  const isEdit = Boolean(existingCustomer);
  const c = existingCustomer || { name: '', phone: '', whatsapp: '', address: '', customerType: 'retail', notes: '' };

  UI.showModal(`
    <div class="modal modal-wide">
      <h3 class="modal-title">${isEdit ? 'تعديل عميل' : 'إضافة عميل جديد'}</h3>
      <div id="customer-form-errors" class="form-errors" hidden></div>
      <form id="customer-form" class="form-grid">
        <fieldset class="form-section">
          <label class="form-field">
            <span>اسم العميل *</span>
            <input type="text" name="name" value="${escapeHtml(c.name)}" autofocus>
          </label>
          <label class="form-field">
            <span>رقم الجوال</span>
            <input type="text" name="phone" value="${escapeHtml(c.phone)}">
          </label>
          <label class="form-field">
            <span>رقم واتساب</span>
            <input type="text" name="whatsapp" value="${escapeHtml(c.whatsapp)}">
          </label>
          <label class="form-field">
            <span>نوع العميل</span>
            <select name="customerType">
              <option value="retail" ${c.customerType !== 'wholesale' ? 'selected' : ''}>قطاعي</option>
              <option value="wholesale" ${c.customerType === 'wholesale' ? 'selected' : ''}>جملة</option>
            </select>
          </label>
          <label class="form-field form-field-wide">
            <span>العنوان</span>
            <input type="text" name="address" value="${escapeHtml(c.address)}">
          </label>
          <label class="form-field form-field-wide">
            <span>ملاحظات</span>
            <textarea name="notes" rows="2">${escapeHtml(c.notes)}</textarea>
          </label>
        </fieldset>
        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" id="customer-form-cancel">إلغاء</button>
          <button type="submit" class="btn btn-primary">${isEdit ? 'حفظ التعديلات' : 'إضافة العميل'}</button>
        </div>
      </form>
    </div>
  `);

  document.getElementById('customer-form-cancel').addEventListener('click', () => UI.closeModal());

  document.getElementById('customer-form').addEventListener('submit', async (event) => {
    event.preventDefault();
    const formData = new FormData(event.target);
    const fields = Object.fromEntries(formData.entries());

    try {
      const result = isEdit ? await updateCustomer(existingCustomer.id, fields) : await addCustomer(fields);
      if (result.errors.length > 0) {
        const errorBox = document.getElementById('customer-form-errors');
        errorBox.hidden = false;
        errorBox.innerHTML = `<ul>${result.errors.map((e) => `<li>${escapeHtml(e)}</li>`).join('')}</ul>`;
        return;
      }
      UI.closeModal();
      UI.success(isEdit ? 'تم حفظ تعديلات العميل.' : 'تم إضافة العميل بنجاح.');
      if (onSaved) onSaved(isEdit ? existingCustomer.id : result.id);
    } catch (err) {
      UI.error(friendlyError('تعذر حفظ بيانات العميل. يرجى المحاولة مرة أخرى.', err));
    }
  });
}
