/**
 * products.js — "المنتجات" screen: product CRUD, box/unit inventory
 * (the "المخزون" screen is merged into this one, per the approved
 * simplification), manual stock adjustments with an audit trail, and
 * Excel import/export.
 *
 * Inventory rule: a product's `totalUnits` is the single source of truth
 * for stock. Box/unit display (e.g. "8 صناديق + 9 قطع") is always derived
 * from it, never stored separately, so the two can never disagree. Stock
 * only ever changes through recordStockAdjustment()/adjustProductStock(),
 * each of which writes an auditable `stockAdjustments` entry alongside the
 * update inside one transaction.
 */

const STOCK_ADJUSTMENT_REASONS = [
  { value: 'initial', label: 'رصيد افتتاحي' },
  { value: 'count', label: 'تصحيح جرد' },
  { value: 'damage', label: 'تلف' },
  { value: 'other', label: 'أخرى' },
];

const productsScreenState = {
  search: '',
  statusFilter: 'active', // active | inactive | all
};

// ---- Data access --------------------------------------------------------

async function listAllProducts() {
  return dbGetAll('products');
}

async function getProductById(id) {
  return dbGet('products', id);
}

async function isSkuTaken(sku, excludingId) {
  const matches = await dbGetAllByIndex('products', 'sku', sku);
  return matches.some((p) => p.id !== excludingId);
}

async function productHasInvoiceHistory(productId) {
  const items = await dbGetAllByIndex('invoiceItems', 'productId', productId);
  return items.length > 0;
}

async function listDistinctCategories() {
  const products = await listAllProducts();
  const categories = new Set(products.map((p) => p.category).filter(Boolean));
  return Array.from(categories).sort((a, b) => a.localeCompare(b, 'ar'));
}

// ---- Calculations ---------------------------------------------------------

function computeUnitCostHalalas(boxPurchasePriceHalalas, unitsPerBox) {
  if (!unitsPerBox || unitsPerBox <= 0) return 0;
  return roundHalfUp(boxPurchasePriceHalalas / unitsPerBox);
}

function productStockDisplay(product) {
  const { boxes, remainder } = unitsToBoxesAndRemainder(product.totalUnits, product.unitsPerBox);
  if (boxes > 0 && remainder > 0) return `${boxes} صندوق + ${remainder} قطعة`;
  if (boxes > 0) return `${boxes} صندوق`;
  return `${remainder} قطعة`;
}

function isLowStock(product) {
  return product.totalUnits <= product.minStockUnits;
}

// ---- Validation / normalization --------------------------------------------

/**
 * Builds a normalized product record from raw field values (from either
 * the manual form or an Excel import row). Returns { errors, product }.
 * `existingId` (editing) excludes itself from the duplicate-SKU check.
 */
async function buildProduct(fields, existingId) {
  const errors = [];

  const sku = String(fields.sku || '').trim();
  const name = String(fields.name || '').trim();
  const unitsPerBox = Number(fields.unitsPerBox);
  const boxPurchasePriceHalalas = sarToHalalas(fields.boxPurchasePrice || 0);
  const retailUnitPriceHalalas = sarToHalalas(fields.retailUnitPrice || 0);
  const wholesaleBoxPriceHalalas = sarToHalalas(fields.wholesaleBoxPrice || 0);
  const minStockUnits = Number(fields.minStockUnits || 0);

  if (!sku) errors.push('يرجى إدخال كود المنتج (SKU).');
  if (!name) errors.push('يرجى إدخال اسم المنتج.');
  if (Number.isNaN(unitsPerBox) || unitsPerBox < 1) {
    errors.push('عدد الوحدات داخل الصندوق يجب أن يكون رقمًا 1 أو أكبر.');
  }
  if (Number.isNaN(boxPurchasePriceHalalas) || boxPurchasePriceHalalas < 0) {
    errors.push('سعر شراء الصندوق يجب أن يكون رقمًا صفر أو أكبر.');
  }
  if (Number.isNaN(retailUnitPriceHalalas) || retailUnitPriceHalalas < 0) {
    errors.push('سعر البيع قطاعي يجب أن يكون رقمًا صفر أو أكبر.');
  }
  if (Number.isNaN(wholesaleBoxPriceHalalas) || wholesaleBoxPriceHalalas < 0) {
    errors.push('سعر بيع الصندوق جملة يجب أن يكون رقمًا صفر أو أكبر.');
  }
  if (Number.isNaN(minStockUnits) || minStockUnits < 0) {
    errors.push('الحد الأدنى للمخزون يجب أن يكون رقمًا صفر أو أكبر.');
  }

  if (sku && (await isSkuTaken(sku, existingId))) {
    errors.push(`كود المنتج (${sku}) مستخدم من قبل لمنتج آخر.`);
  }

  if (errors.length > 0) return { errors, product: null };

  const product = {
    sku,
    barcode: String(fields.barcode || '').trim(),
    name,
    category: String(fields.category || '').trim(),
    description: String(fields.description || '').trim(),
    supplier: String(fields.supplier || '').trim(),
    boxPurchasePriceHalalas,
    unitsPerBox,
    unitCostHalalas: computeUnitCostHalalas(boxPurchasePriceHalalas, unitsPerBox),
    retailUnitPriceHalalas,
    wholesaleBoxPriceHalalas,
    minStockUnits,
    notes: String(fields.notes || '').trim(),
    status: fields.status === 'inactive' ? 'inactive' : 'active',
  };

  return { errors: [], product };
}

// ---- Mutations --------------------------------------------------------------

/** Creates a product with an optional opening stock balance (in total units), logged as an audit entry. */
async function addProduct(fields, initialTotalUnits) {
  const { errors, product } = await buildProduct(fields, null);
  if (errors.length > 0) return { errors };

  const totalUnits = Math.max(0, Math.round(Number(initialTotalUnits) || 0));
  const now = new Date().toISOString();
  const record = { ...product, totalUnits, createdAt: now, updatedAt: now };
  const id = await dbPut('products', record);

  if (totalUnits > 0) {
    await dbPut('stockAdjustments', {
      productId: id,
      date: formatDateForStorage(),
      deltaUnits: totalUnits,
      reason: 'initial',
      notes: 'رصيد افتتاحي عند إضافة المنتج',
      previousTotalUnits: 0,
      newTotalUnits: totalUnits,
      createdAt: now,
    });
  }

  return { errors: [], id };
}

/** Updates a product's details. Stock is never touched here — only through adjustProductStock(). */
async function updateProduct(id, fields) {
  const existing = await getProductById(id);
  if (!existing) return { errors: ['تعذر العثور على المنتج.'] };

  const { errors, product } = await buildProduct(fields, id);
  if (errors.length > 0) return { errors };

  const record = {
    ...existing,
    ...product,
    totalUnits: existing.totalUnits,
    updatedAt: new Date().toISOString(),
  };
  await dbPut('products', record);
  return { errors: [] };
}

async function setProductStatus(id, status) {
  const existing = await getProductById(id);
  if (!existing) return { errors: ['تعذر العثور على المنتج.'] };
  await dbPut('products', { ...existing, status, updatedAt: new Date().toISOString() });
  return { errors: [] };
}

/** Deletes a product outright, but only if it has never appeared on an invoice — otherwise the user must deactivate it instead. */
async function deleteProductIfUnused(id) {
  const used = await productHasInvoiceHistory(id);
  if (used) {
    return { errors: ['لا يمكن حذف هذا المنتج لأنه مرتبط بفواتير سابقة. يمكنك إيقافه بدلًا من ذلك.'] };
  }
  await dbDelete('products', id);
  return { errors: [] };
}

/**
 * Applies a manual stock change (deltaUnits positive or negative) atomically
 * with its audit entry, and rejects anything that would drive stock negative.
 */
async function adjustProductStock(productId, deltaUnits, reason, notes) {
  if (!deltaUnits) {
    return { errors: ['يرجى إدخال كمية تعديل غير صفرية.'] };
  }
  try {
    await dbTransaction(['products', 'stockAdjustments'], async (tx) => {
      const productsStore = tx.objectStore('products');
      const product = await requestToPromise(productsStore.get(productId));
      if (!product) throw new Error('PRODUCT_NOT_FOUND');

      const previousTotalUnits = product.totalUnits;
      const newTotalUnits = previousTotalUnits + deltaUnits;
      if (newTotalUnits < 0) throw new Error('NEGATIVE_STOCK');

      product.totalUnits = newTotalUnits;
      product.updatedAt = new Date().toISOString();
      productsStore.put(product);

      tx.objectStore('stockAdjustments').add({
        productId,
        date: formatDateForStorage(),
        deltaUnits,
        reason,
        notes: notes || '',
        previousTotalUnits,
        newTotalUnits,
        createdAt: new Date().toISOString(),
      });
    });
    return { errors: [] };
  } catch (err) {
    if (err && err.message === 'NEGATIVE_STOCK') {
      return { errors: ['الكمية المطلوب خصمها أكبر من المخزون الحالي. لا يمكن أن يصبح المخزون بالسالب.'] };
    }
    throw err;
  }
}

// ---- Excel import/export -----------------------------------------------------

const PRODUCT_EXPORT_HEADERS = [
  'SKU', 'الباركود', 'اسم المنتج', 'التصنيف', 'المورد', 'الوصف',
  'سعر شراء الصندوق', 'عدد الوحدات', 'تكلفة الوحدة', 'سعر القطاعي', 'سعر الجملة',
  'المخزون', 'الحد الأدنى للمخزون', 'الحالة', 'ملاحظات',
];

async function exportProductsToExcel() {
  const products = await listAllProducts();
  const rows = products.map((p) => [
    p.sku, p.barcode, p.name, p.category, p.supplier, p.description,
    halalasToSar(p.boxPurchasePriceHalalas), p.unitsPerBox, halalasToSar(p.unitCostHalalas),
    halalasToSar(p.retailUnitPriceHalalas), halalasToSar(p.wholesaleBoxPriceHalalas),
    p.totalUnits, p.minStockUnits, p.status === 'active' ? 'فعال' : 'غير فعال', p.notes,
  ]);
  exportRowsToExcel(`المنتجات_${formatDateForStorage()}.xlsx`, 'المنتجات', PRODUCT_EXPORT_HEADERS, rows);
}

/** Maps one raw Excel row (Arabic headers) to the field names buildProduct() expects. */
function mapImportRow(row) {
  return {
    sku: row['SKU'] ?? row['sku'] ?? '',
    barcode: row['الباركود'] ?? '',
    name: row['اسم المنتج'] ?? '',
    category: row['التصنيف'] ?? '',
    supplier: row['المورد'] ?? '',
    description: row['الوصف'] ?? '',
    notes: row['ملاحظات'] ?? '',
    boxPurchasePrice: row['سعر شراء الصندوق'] ?? '',
    unitsPerBox: row['عدد الوحدات'] ?? '',
    retailUnitPrice: row['سعر القطاعي'] ?? '',
    wholesaleBoxPrice: row['سعر الجملة'] ?? '',
    minStockUnits: row['الحد الأدنى للمخزون'] ?? 0,
    initialTotalUnits: row['المخزون'] ?? 0,
  };
}

/** Validates every row of a parsed import file against required fields and duplicate SKUs (in-file and existing). */
async function buildImportPreview(rawRows) {
  const seenSkusInFile = new Set();
  const preview = [];

  for (let i = 0; i < rawRows.length; i++) {
    const mapped = mapImportRow(rawRows[i]);
    const rowErrors = [];
    const sku = String(mapped.sku).trim();
    const name = String(mapped.name).trim();

    if (!sku) rowErrors.push('كود المنتج (SKU) مفقود');
    if (!name) rowErrors.push('اسم المنتج مفقود');

    if (sku) {
      if (seenSkusInFile.has(sku)) {
        rowErrors.push('SKU مكرر داخل الملف نفسه');
      } else {
        seenSkusInFile.add(sku);
      }
      if (await isSkuTaken(sku, null)) {
        rowErrors.push('SKU موجود مسبقًا في النظام');
      }
    }

    const unitsPerBox = Number(mapped.unitsPerBox);
    if (Number.isNaN(unitsPerBox) || unitsPerBox < 1) {
      rowErrors.push('عدد الوحدات يجب أن يكون رقمًا 1 أو أكبر');
    }

    preview.push({ rowNumber: i + 2, mapped, errors: rowErrors });
  }

  return preview;
}

async function importValidRows(preview) {
  const validRows = preview.filter((row) => row.errors.length === 0);
  let successCount = 0;
  const failures = [];

  for (const row of validRows) {
    const fields = {
      sku: row.mapped.sku,
      barcode: row.mapped.barcode,
      name: row.mapped.name,
      category: row.mapped.category,
      description: row.mapped.description,
      supplier: row.mapped.supplier,
      boxPurchasePrice: row.mapped.boxPurchasePrice,
      unitsPerBox: row.mapped.unitsPerBox,
      retailUnitPrice: row.mapped.retailUnitPrice,
      wholesaleBoxPrice: row.mapped.wholesaleBoxPrice,
      minStockUnits: row.mapped.minStockUnits,
      notes: row.mapped.notes,
      status: 'active',
    };
    const totalUnits = Math.max(0, Math.round(Number(row.mapped.initialTotalUnits) || 0));
    const result = await addProduct(fields, totalUnits);
    if (result.errors.length > 0) {
      failures.push(`صف ${row.rowNumber}: ${result.errors.join(' ')}`);
    } else {
      successCount++;
    }
  }

  return { successCount, failures };
}

// ---- Rendering: products screen ------------------------------------------------

async function renderProductsScreen() {
  const container = document.getElementById('screen-products');
  const settings = await getSettings();

  container.innerHTML = `
    <h2 class="screen-title">المنتجات</h2>
    <p class="screen-hint">تُدار الكمية دائمًا عبر زر "تعديل الكمية" حتى يبقى لكل تغيير في المخزون سبب مسجَّل.</p>

    <div class="toolbar">
      <input type="search" id="product-search" class="toolbar-search" placeholder="ابحث بالاسم أو الكود أو الباركود...">
      <select id="product-status-filter" class="toolbar-select">
        <option value="active">المنتجات الفعالة</option>
        <option value="inactive">المنتجات الموقوفة</option>
        <option value="all">كل المنتجات</option>
      </select>
      <div class="toolbar-spacer"></div>
      <button type="button" id="btn-export-products" class="btn btn-secondary">تصدير Excel</button>
      <button type="button" id="btn-import-products" class="btn btn-secondary">استيراد من Excel</button>
      <input type="file" id="product-import-input" accept=".xlsx,.xls,.csv" style="display:none">
      <button type="button" id="btn-add-product" class="btn btn-primary">+ إضافة منتج</button>
    </div>

    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>الكود</th>
            <th>الاسم</th>
            <th>التصنيف</th>
            <th>تكلفة الوحدة</th>
            <th>سعر القطاعي</th>
            <th>سعر الجملة</th>
            <th>المخزون</th>
            <th>الحالة</th>
            <th>إجراءات</th>
          </tr>
        </thead>
        <tbody id="products-table-body"></tbody>
      </table>
      <p id="products-empty-message" class="table-empty-message" hidden>لا توجد منتجات مطابقة.</p>
    </div>
  `;

  document.getElementById('product-status-filter').value = productsScreenState.statusFilter;
  document.getElementById('product-search').value = productsScreenState.search;

  document.getElementById('product-search').addEventListener('input', (e) => {
    productsScreenState.search = e.target.value;
    renderProductsTableBody(settings);
  });
  document.getElementById('product-status-filter').addEventListener('change', (e) => {
    productsScreenState.statusFilter = e.target.value;
    renderProductsTableBody(settings);
  });
  document.getElementById('btn-add-product').addEventListener('click', () => openProductFormModal(null));
  document.getElementById('btn-export-products').addEventListener('click', async () => {
    try {
      await exportProductsToExcel();
      UI.success('تم تصدير ملف المنتجات بنجاح.');
    } catch (err) {
      UI.error(friendlyError('تعذر تصدير الملف. يرجى المحاولة مرة أخرى.', err));
    }
  });
  document.getElementById('btn-import-products').addEventListener('click', () => {
    document.getElementById('product-import-input').click();
  });
  document.getElementById('product-import-input').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    await handleProductImportFile(file);
  });

  await renderProductsTableBody(settings);
}

async function renderProductsTableBody(settings) {
  const tbody = document.getElementById('products-table-body');
  const emptyMessage = document.getElementById('products-empty-message');
  if (!tbody) return;

  const all = await listAllProducts();
  const search = productsScreenState.search.trim().toLowerCase();
  const filtered = all.filter((p) => {
    if (productsScreenState.statusFilter !== 'all' && p.status !== productsScreenState.statusFilter) return false;
    if (!search) return true;
    return (
      p.name.toLowerCase().includes(search)
      || p.sku.toLowerCase().includes(search)
      || (p.barcode || '').toLowerCase().includes(search)
    );
  });

  filtered.sort((a, b) => a.name.localeCompare(b.name, 'ar'));

  if (filtered.length === 0) {
    tbody.innerHTML = '';
    emptyMessage.hidden = false;
    return;
  }
  emptyMessage.hidden = true;

  tbody.innerHTML = filtered.map((p) => {
    const low = isLowStock(p);
    return `
      <tr>
        <td>${escapeHtml(p.sku)}</td>
        <td>${escapeHtml(p.name)}</td>
        <td>${escapeHtml(p.category)}</td>
        <td>${formatCurrency(p.unitCostHalalas, settings.currencySymbol)}</td>
        <td>${formatCurrency(p.retailUnitPriceHalalas, settings.currencySymbol)}</td>
        <td>${formatCurrency(p.wholesaleBoxPriceHalalas, settings.currencySymbol)}</td>
        <td>
          ${escapeHtml(productStockDisplay(p))}
          ${low ? '<span class="badge badge-warning">منخفض</span>' : ''}
        </td>
        <td>
          <span class="badge ${p.status === 'active' ? 'badge-success' : 'badge-muted'}">
            ${p.status === 'active' ? 'فعال' : 'غير فعال'}
          </span>
        </td>
        <td class="table-actions">
          <button type="button" class="link-btn" data-action="edit" data-id="${p.id}">تعديل</button>
          <button type="button" class="link-btn" data-action="adjust" data-id="${p.id}">تعديل الكمية</button>
          <button type="button" class="link-btn" data-action="toggle-status" data-id="${p.id}">${p.status === 'active' ? 'إيقاف' : 'تفعيل'}</button>
          <button type="button" class="link-btn link-btn-danger" data-action="delete" data-id="${p.id}">حذف</button>
        </td>
      </tr>
    `;
  }).join('');

  tbody.querySelectorAll('[data-action="edit"]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const product = await getProductById(Number(btn.dataset.id));
      openProductFormModal(product);
    });
  });
  tbody.querySelectorAll('[data-action="adjust"]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const product = await getProductById(Number(btn.dataset.id));
      openStockAdjustmentModal(product);
    });
  });
  tbody.querySelectorAll('[data-action="toggle-status"]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const id = Number(btn.dataset.id);
      const product = await getProductById(id);
      const goingInactive = product.status === 'active';
      if (goingInactive) {
        const confirmed = await UI.confirm({
          title: 'إيقاف المنتج',
          message: `هل تريد إيقاف المنتج "${escapeHtml(product.name)}"؟ لن يظهر بعد ذلك في شاشة البيع الجديد.`,
          confirmLabel: 'إيقاف',
          danger: true,
        });
        if (!confirmed) return;
      }
      await setProductStatus(id, goingInactive ? 'inactive' : 'active');
      UI.success(goingInactive ? 'تم إيقاف المنتج.' : 'تم تفعيل المنتج.');
      renderProductsTableBody(settings);
    });
  });
  tbody.querySelectorAll('[data-action="delete"]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const id = Number(btn.dataset.id);
      const product = await getProductById(id);
      const confirmed = await UI.confirm({
        title: 'حذف المنتج',
        message: `هل أنت متأكد من حذف المنتج "${escapeHtml(product.name)}" نهائيًا؟ لا يمكن التراجع عن هذا الإجراء.`,
        confirmLabel: 'حذف نهائيًا',
        danger: true,
      });
      if (!confirmed) return;
      const result = await deleteProductIfUnused(id);
      if (result.errors.length > 0) {
        UI.error(result.errors[0]);
        return;
      }
      UI.success('تم حذف المنتج.');
      renderProductsTableBody(settings);
    });
  });
}

// ---- Rendering: add/edit product modal --------------------------------------------

async function openProductFormModal(existingProduct) {
  const settings = await getSettings();
  const categories = await listDistinctCategories();
  const isEdit = Boolean(existingProduct);
  const p = existingProduct || {
    sku: '', barcode: '', name: '', category: '', description: '', supplier: '',
    boxPurchasePriceHalalas: 0, unitsPerBox: '', retailUnitPriceHalalas: 0, wholesaleBoxPriceHalalas: 0,
    minStockUnits: settings.lowStockDefaultThreshold, notes: '', status: 'active', totalUnits: 0,
  };

  UI.showModal(`
    <div class="modal modal-wide">
      <h3 class="modal-title">${isEdit ? 'تعديل منتج' : 'إضافة منتج جديد'}</h3>
      <div id="product-form-errors" class="form-errors" hidden></div>
      <form id="product-form" class="form-grid">
        <fieldset class="form-section">
          <legend>البيانات الأساسية</legend>
          <label class="form-field">
            <span>كود المنتج (SKU) *</span>
            <input type="text" name="sku" value="${escapeHtml(p.sku)}" ${isEdit ? '' : 'autofocus'}>
          </label>
          <label class="form-field">
            <span>الباركود (اختياري)</span>
            <input type="text" name="barcode" value="${escapeHtml(p.barcode)}">
          </label>
          <label class="form-field">
            <span>اسم المنتج *</span>
            <input type="text" name="name" value="${escapeHtml(p.name)}">
          </label>
          <label class="form-field">
            <span>التصنيف</span>
            <input type="text" name="category" list="category-options" value="${escapeHtml(p.category)}">
            <datalist id="category-options">
              ${categories.map((c) => `<option value="${escapeHtml(c)}">`).join('')}
            </datalist>
          </label>
          <label class="form-field">
            <span>المورد (اختياري)</span>
            <input type="text" name="supplier" value="${escapeHtml(p.supplier)}">
          </label>
          <label class="form-field form-field-wide">
            <span>الوصف</span>
            <textarea name="description" rows="2">${escapeHtml(p.description)}</textarea>
          </label>
        </fieldset>

        <fieldset class="form-section">
          <legend>التسعير</legend>
          <label class="form-field">
            <span>سعر شراء الصندوق (${escapeHtml(settings.currencySymbol)}) *</span>
            <input type="number" step="0.01" min="0" name="boxPurchasePrice" id="field-box-purchase-price" value="${halalasToSar(p.boxPurchasePriceHalalas)}">
          </label>
          <label class="form-field">
            <span>عدد الوحدات داخل الصندوق *</span>
            <input type="number" step="1" min="1" name="unitsPerBox" id="field-units-per-box" value="${escapeHtml(p.unitsPerBox)}">
          </label>
          <label class="form-field">
            <span>تكلفة الوحدة (محسوبة تلقائيًا)</span>
            <input type="text" id="field-unit-cost-preview" value="${formatCurrency(p.unitCostHalalas || 0, settings.currencySymbol)}" disabled>
          </label>
          <label class="form-field">
            <span>سعر البيع قطاعي (${escapeHtml(settings.currencySymbol)}) *</span>
            <input type="number" step="0.01" min="0" name="retailUnitPrice" value="${halalasToSar(p.retailUnitPriceHalalas)}">
          </label>
          <label class="form-field">
            <span>سعر بيع الصندوق جملة (${escapeHtml(settings.currencySymbol)}) *</span>
            <input type="number" step="0.01" min="0" name="wholesaleBoxPrice" value="${halalasToSar(p.wholesaleBoxPriceHalalas)}">
          </label>
          <p class="field-hint form-field-wide">اترك السعر 0 إذا كان المنتج لا يُباع بهذه الطريقة.</p>
        </fieldset>

        <fieldset class="form-section">
          <legend>المخزون</legend>
          ${isEdit ? `
            <label class="form-field">
              <span>الكمية الحالية</span>
              <input type="text" value="${escapeHtml(productStockDisplay(p))}" disabled>
            </label>
            <p class="field-hint form-field-wide">لتغيير الكمية استخدم زر "تعديل الكمية" من قائمة المنتجات بعد الحفظ.</p>
          ` : `
            <label class="form-field">
              <span>عدد الصناديق (رصيد أولي)</span>
              <input type="number" step="1" min="0" name="initialBoxes" id="field-initial-boxes" value="0">
            </label>
            <label class="form-field">
              <span>عدد قطع إضافية</span>
              <input type="number" step="1" min="0" name="initialExtraUnits" id="field-initial-extra" value="0">
            </label>
            <label class="form-field">
              <span>الإجمالي (بالوحدة)</span>
              <input type="text" id="field-initial-total-preview" value="0" disabled>
            </label>
          `}
          <label class="form-field">
            <span>الحد الأدنى للمخزون (بالوحدة)</span>
            <input type="number" step="1" min="0" name="minStockUnits" value="${p.minStockUnits}">
          </label>
        </fieldset>

        <fieldset class="form-section">
          <legend>أخرى</legend>
          <label class="form-field form-field-wide">
            <span>ملاحظات</span>
            <textarea name="notes" rows="2">${escapeHtml(p.notes)}</textarea>
          </label>
          <label class="form-field form-checkbox">
            <input type="checkbox" name="statusActive" ${p.status === 'active' ? 'checked' : ''}>
            <span>المنتج فعال</span>
          </label>
        </fieldset>

        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" id="product-form-cancel">إلغاء</button>
          <button type="submit" class="btn btn-primary">${isEdit ? 'حفظ التعديلات' : 'إضافة المنتج'}</button>
        </div>
      </form>
    </div>
  `);

  const updateUnitCostPreview = () => {
    const boxPrice = sarToHalalas(document.getElementById('field-box-purchase-price').value || 0);
    const unitsPerBox = Number(document.getElementById('field-units-per-box').value);
    const preview = document.getElementById('field-unit-cost-preview');
    preview.value = formatCurrency(computeUnitCostHalalas(boxPrice, unitsPerBox), settings.currencySymbol);
  };
  document.getElementById('field-box-purchase-price').addEventListener('input', updateUnitCostPreview);
  document.getElementById('field-units-per-box').addEventListener('input', updateUnitCostPreview);

  if (!isEdit) {
    const updateInitialTotalPreview = () => {
      const boxes = Number(document.getElementById('field-initial-boxes').value) || 0;
      const extra = Number(document.getElementById('field-initial-extra').value) || 0;
      const unitsPerBox = Number(document.getElementById('field-units-per-box').value) || 0;
      document.getElementById('field-initial-total-preview').value = Math.max(0, boxes) * unitsPerBox + Math.max(0, extra);
    };
    document.getElementById('field-initial-boxes').addEventListener('input', updateInitialTotalPreview);
    document.getElementById('field-initial-extra').addEventListener('input', updateInitialTotalPreview);
    document.getElementById('field-units-per-box').addEventListener('input', updateInitialTotalPreview);
  }

  document.getElementById('product-form-cancel').addEventListener('click', () => UI.closeModal());

  document.getElementById('product-form').addEventListener('submit', async (event) => {
    event.preventDefault();
    const submitBtn = event.target.querySelector('button[type="submit"]');
    if (submitBtn.disabled) return; // already submitting — ignore a rapid double-click
    submitBtn.disabled = true;

    const formData = new FormData(event.target);
    const values = Object.fromEntries(formData.entries());
    const fields = {
      sku: values.sku, barcode: values.barcode, name: values.name, category: values.category,
      description: values.description, supplier: values.supplier,
      boxPurchasePrice: values.boxPurchasePrice, unitsPerBox: values.unitsPerBox,
      retailUnitPrice: values.retailUnitPrice, wholesaleBoxPrice: values.wholesaleBoxPrice,
      minStockUnits: values.minStockUnits, notes: values.notes,
      status: values.statusActive ? 'active' : 'inactive',
    };

    try {
      let result;
      if (isEdit) {
        result = await updateProduct(existingProduct.id, fields);
      } else {
        const boxes = Number(values.initialBoxes) || 0;
        const extra = Number(values.initialExtraUnits) || 0;
        const unitsPerBox = Number(values.unitsPerBox) || 0;
        const totalUnits = Math.max(0, boxes) * unitsPerBox + Math.max(0, extra);
        result = await addProduct(fields, totalUnits);
      }

      if (result.errors.length > 0) {
        const errorBox = document.getElementById('product-form-errors');
        errorBox.hidden = false;
        errorBox.innerHTML = `<ul>${result.errors.map((e) => `<li>${escapeHtml(e)}</li>`).join('')}</ul>`;
        submitBtn.disabled = false;
        return;
      }

      UI.closeModal();
      UI.success(isEdit ? 'تم حفظ تعديلات المنتج.' : 'تم إضافة المنتج بنجاح.');
      renderProductsTableBody(settings);
    } catch (err) {
      UI.error(friendlyError('تعذر حفظ المنتج. يرجى المحاولة مرة أخرى.', err));
      submitBtn.disabled = false;
    }
  });
}

// ---- Rendering: stock adjustment modal -----------------------------------------------

function openStockAdjustmentModal(product) {
  UI.showModal(`
    <div class="modal">
      <h3 class="modal-title">تعديل كمية: ${escapeHtml(product.name)}</h3>
      <p class="modal-message">الكمية الحالية: ${escapeHtml(productStockDisplay(product))}</p>
      <div id="stock-adjust-errors" class="form-errors" hidden></div>
      <form id="stock-adjust-form" class="form-grid">
        <label class="form-field">
          <span>نوع التعديل</span>
          <select name="direction">
            <option value="increase">زيادة (+)</option>
            <option value="decrease">نقص (-)</option>
          </select>
        </label>
        <label class="form-field">
          <span>الكمية (بالوحدة)</span>
          <input type="number" step="1" min="1" name="quantity" id="stock-adjust-quantity" value="1">
        </label>
        <p class="field-hint form-field-wide" id="stock-adjust-box-hint"></p>
        <label class="form-field">
          <span>السبب</span>
          <select name="reason">
            ${STOCK_ADJUSTMENT_REASONS.filter((r) => r.value !== 'initial').map((r) => `<option value="${r.value}">${r.label}</option>`).join('')}
          </select>
        </label>
        <label class="form-field form-field-wide">
          <span>ملاحظات (اختياري)</span>
          <textarea name="notes" rows="2"></textarea>
        </label>
        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" id="stock-adjust-cancel">إلغاء</button>
          <button type="submit" class="btn btn-primary">تأكيد التعديل</button>
        </div>
      </form>
    </div>
  `);

  const updateHint = () => {
    const qty = Number(document.getElementById('stock-adjust-quantity').value) || 0;
    const { boxes, remainder } = unitsToBoxesAndRemainder(qty, product.unitsPerBox);
    document.getElementById('stock-adjust-box-hint').textContent = `يعادل تقريبًا: ${boxes} صندوق و${remainder} قطعة`;
  };
  document.getElementById('stock-adjust-quantity').addEventListener('input', updateHint);
  updateHint();

  document.getElementById('stock-adjust-cancel').addEventListener('click', () => UI.closeModal());

  document.getElementById('stock-adjust-form').addEventListener('submit', async (event) => {
    event.preventDefault();
    const submitBtn = event.target.querySelector('button[type="submit"]');
    if (submitBtn.disabled) return; // already submitting — ignore a rapid double-click
    submitBtn.disabled = true;

    const formData = new FormData(event.target);
    const values = Object.fromEntries(formData.entries());
    const quantity = Math.round(Number(values.quantity));
    const errorBox = document.getElementById('stock-adjust-errors');

    if (Number.isNaN(quantity) || quantity <= 0) {
      errorBox.hidden = false;
      errorBox.innerHTML = '<ul><li>يرجى إدخال كمية أكبر من صفر.</li></ul>';
      submitBtn.disabled = false;
      return;
    }

    const deltaUnits = values.direction === 'decrease' ? -quantity : quantity;
    const result = await adjustProductStock(product.id, deltaUnits, values.reason, values.notes);

    if (result.errors.length > 0) {
      errorBox.hidden = false;
      errorBox.innerHTML = `<ul>${result.errors.map((e) => `<li>${escapeHtml(e)}</li>`).join('')}</ul>`;
      submitBtn.disabled = false;
      return;
    }

    UI.closeModal();
    UI.success('تم تحديث المخزون بنجاح.');
    const settings = await getSettings();
    renderProductsTableBody(settings);
  });
}

// ---- Excel import flow -------------------------------------------------------------

async function handleProductImportFile(file) {
  let rawRows;
  try {
    rawRows = await readExcelFileAsObjects(file);
  } catch (err) {
    UI.error(friendlyError('تعذر قراءة ملف Excel. تأكد أن الملف بصيغة صحيحة.', err));
    return;
  }
  if (rawRows.length === 0) {
    UI.error('الملف لا يحتوي على أي بيانات.');
    return;
  }

  const preview = await buildImportPreview(rawRows);
  renderImportPreviewModal(preview);
}

function renderImportPreviewModal(preview) {
  const validCount = preview.filter((row) => row.errors.length === 0).length;

  UI.showModal(`
    <div class="modal modal-wide">
      <h3 class="modal-title">معاينة استيراد المنتجات</h3>
      <p class="modal-message">سيتم استيراد ${validCount} من أصل ${preview.length} صف. الصفوف التي تحتوي أخطاء لن تُستورد.</p>
      <div class="table-wrap modal-table-scroll">
        <table class="data-table">
          <thead>
            <tr><th>الصف</th><th>SKU</th><th>الاسم</th><th>الحالة</th></tr>
          </thead>
          <tbody>
            ${preview.map((row) => `
              <tr>
                <td>${row.rowNumber}</td>
                <td>${escapeHtml(row.mapped.sku)}</td>
                <td>${escapeHtml(row.mapped.name)}</td>
                <td class="${row.errors.length > 0 ? 'text-danger' : 'text-success'}">
                  ${row.errors.length > 0 ? escapeHtml(row.errors.join('، ')) : 'صالح'}
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
      <div class="modal-actions">
        <button type="button" class="btn btn-secondary" id="import-cancel">إلغاء</button>
        <button type="button" class="btn btn-primary" id="import-confirm" ${validCount === 0 ? 'disabled' : ''}>تأكيد الاستيراد</button>
      </div>
    </div>
  `);

  document.getElementById('import-cancel').addEventListener('click', () => UI.closeModal());
  document.getElementById('import-confirm').addEventListener('click', async () => {
    const { successCount, failures } = await importValidRows(preview);
    UI.closeModal();
    if (failures.length === 0) {
      UI.success(`تم استيراد ${successCount} منتج بنجاح.`);
    } else {
      UI.toast(`تم استيراد ${successCount} منتج، وتعذر استيراد ${failures.length}.`, 'error');
      console.error('Product import failures:', failures);
    }
    const settings = await getSettings();
    renderProductsTableBody(settings);
  });
}
