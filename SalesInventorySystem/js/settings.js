/**
 * settings.js — the "الإعدادات" screen.
 *
 * Settings are stored as a single record (key: 'main') in the `settings`
 * store. Every other module must read fee/currency/invoice-prefix values
 * from here rather than hardcoding them — and every invoice must copy
 * (snapshot) the values it used at creation time rather than referencing
 * this record live, so a later settings change never alters past invoices.
 */

const DEFAULT_SETTINGS = {
  key: 'main',
  businessName: '',
  phone: '',
  extraInvoiceInfo: '',
  currency: 'SAR',
  currencySymbol: 'ريال',
  feePercent: 8,
  feeFixedHalalas: 800,
  invoicePrefix: 'INV',
  lowStockDefaultThreshold: 5,
};

/** Returns the current settings, creating the default record on first run. */
async function getSettings() {
  const existing = await dbGet('settings', 'main');
  if (existing) return existing;
  await dbPut('settings', DEFAULT_SETTINGS);
  return { ...DEFAULT_SETTINGS };
}

/** Validates settings form values; returns a list of Arabic error messages (empty = valid). */
function validateSettings(values) {
  const errors = [];

  if (!values.currencySymbol || !values.currencySymbol.trim()) {
    errors.push('يرجى إدخال رمز العملة.');
  }
  if (!values.invoicePrefix || !values.invoicePrefix.trim()) {
    errors.push('يرجى إدخال بادئة رقم الفاتورة.');
  }
  if (/[^A-Za-z0-9\-]/.test(values.invoicePrefix || '')) {
    errors.push('بادئة رقم الفاتورة يجب أن تحتوي على حروف/أرقام إنجليزية فقط (مثال: INV).');
  }
  const feePercent = Number(values.feePercent);
  if (Number.isNaN(feePercent) || feePercent < 0 || feePercent > 100) {
    errors.push('نسبة رسوم الموقع يجب أن تكون رقمًا بين 0 و100.');
  }
  const feeFixed = Number(values.feeFixed);
  if (Number.isNaN(feeFixed) || feeFixed < 0) {
    errors.push('الرسم الثابت يجب أن يكون رقمًا صفر أو أكبر.');
  }
  const lowStock = Number(values.lowStockDefaultThreshold);
  if (Number.isNaN(lowStock) || lowStock < 0) {
    errors.push('الحد الأدنى الافتراضي للمخزون يجب أن يكون رقمًا صفر أو أكبر.');
  }

  return errors;
}

/** Renders the settings screen into #screen-settings and wires up saving. */
async function renderSettingsScreen() {
  const container = document.getElementById('screen-settings');
  const settings = await getSettings();

  container.innerHTML = `
    <h2 class="screen-title">الإعدادات</h2>
    <p class="screen-hint">
      التغييرات هنا تنطبق على عمليات البيع الجديدة فقط، ولا تغيّر الفواتير المحفوظة سابقًا.
    </p>
    <form id="settings-form" class="form-grid">
      <fieldset class="form-section">
        <legend>بيانات النشاط التجاري</legend>

        <label class="form-field">
          <span>اسم النشاط التجاري</span>
          <input type="text" name="businessName" value="${escapeHtml(settings.businessName)}" placeholder="مثال: متجر الأقلام">
        </label>

        <label class="form-field">
          <span>رقم الهاتف</span>
          <input type="text" name="phone" value="${escapeHtml(settings.phone)}" placeholder="05xxxxxxxx">
        </label>

        <label class="form-field form-field-wide">
          <span>معلومات إضافية تظهر على الفاتورة</span>
          <textarea name="extraInvoiceInfo" rows="2" placeholder="مثال: العنوان، ساعات العمل، شروط الاستبدال">${escapeHtml(settings.extraInvoiceInfo)}</textarea>
        </label>
      </fieldset>

      <fieldset class="form-section">
        <legend>رسوم عملية البيع</legend>

        <label class="form-field">
          <span>نسبة رسوم الموقع (%)</span>
          <input type="number" name="feePercent" step="0.01" min="0" max="100" value="${settings.feePercent}">
        </label>

        <label class="form-field">
          <span>الرسم الثابت لكل فاتورة (${escapeHtml(settings.currencySymbol)})</span>
          <input type="number" name="feeFixed" step="0.01" min="0" value="${halalasToSar(settings.feeFixedHalalas)}">
        </label>
      </fieldset>

      <fieldset class="form-section">
        <legend>العملة ورقم الفاتورة</legend>

        <label class="form-field">
          <span>رمز العملة</span>
          <input type="text" name="currencySymbol" value="${escapeHtml(settings.currencySymbol)}">
        </label>

        <label class="form-field">
          <span>بادئة رقم الفاتورة</span>
          <input type="text" name="invoicePrefix" value="${escapeHtml(settings.invoicePrefix)}" placeholder="INV">
        </label>

        <label class="form-field">
          <span>الحد الأدنى الافتراضي للمخزون</span>
          <input type="number" name="lowStockDefaultThreshold" min="0" value="${settings.lowStockDefaultThreshold}">
        </label>
      </fieldset>

      <div class="form-actions">
        <button type="submit" class="btn btn-primary btn-large">حفظ الإعدادات</button>
      </div>
    </form>
  `;

  document.getElementById('settings-form').addEventListener('submit', async (event) => {
    event.preventDefault();
    const formData = new FormData(event.target);
    const values = Object.fromEntries(formData.entries());

    const errors = validateSettings(values);
    if (errors.length > 0) {
      UI.error(errors[0]);
      return;
    }

    try {
      const updated = {
        key: 'main',
        businessName: (values.businessName || '').trim(),
        phone: (values.phone || '').trim(),
        extraInvoiceInfo: (values.extraInvoiceInfo || '').trim(),
        currency: settings.currency,
        currencySymbol: values.currencySymbol.trim(),
        feePercent: Number(values.feePercent),
        feeFixedHalalas: sarToHalalas(values.feeFixed),
        invoicePrefix: values.invoicePrefix.trim().toUpperCase(),
        lowStockDefaultThreshold: Number(values.lowStockDefaultThreshold),
      };
      await dbPut('settings', updated);
      UI.success('تم حفظ الإعدادات بنجاح.');
      renderSettingsScreen();
    } catch (err) {
      UI.error(friendlyError('تعذر حفظ الإعدادات. يرجى المحاولة مرة أخرى.', err));
    }
  });
}

/** Minimal HTML-escaping for values interpolated into templates above. */
function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (ch) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[ch]));
}
