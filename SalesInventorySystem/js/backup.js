/**
 * backup.js — the "النسخ الاحتياطي" screen: one button to save a full
 * JSON backup, one to restore one, with a clear warning and an automatic
 * safety snapshot of the current data taken right before any restore
 * overwrites it (so a bad restore file, or a restore started by
 * mistake, never destroys data with no way back).
 */

const BACKUP_STORE_NAMES = [
  'products', 'customers', 'invoices', 'invoiceItems',
  'expenses', 'stockAdjustments', 'settings', 'counters', 'auditLog',
];
const BACKUP_APP_NAME = 'SalesInventorySystem';
const BACKUP_VERSION = 1;

// ---- Building / reading backup payloads --------------------------------------

async function collectAllStoresData() {
  const data = {};
  for (const name of BACKUP_STORE_NAMES) {
    data[name] = await dbGetAll(name);
  }
  return data;
}

function buildBackupPayload(storesData) {
  return {
    appName: BACKUP_APP_NAME,
    backupVersion: BACKUP_VERSION,
    exportedAt: new Date().toISOString(),
    stores: storesData,
  };
}

/** Returns an Arabic error list (empty = valid) for a parsed backup file's structure. */
function validateBackupPayload(payload) {
  const errors = [];
  if (!payload || typeof payload !== 'object') {
    return ['ملف النسخة الاحتياطية غير صالح.'];
  }
  if (payload.appName !== BACKUP_APP_NAME) {
    errors.push('هذا الملف ليس نسخة احتياطية لنظام المبيعات.');
  }
  if (!payload.stores || typeof payload.stores !== 'object') {
    return [...errors, 'ملف النسخة الاحتياطية تالف أو غير مكتمل.'];
  }
  for (const name of BACKUP_STORE_NAMES) {
    if (!Array.isArray(payload.stores[name])) {
      errors.push('ملف النسخة الاحتياطية تالف أو غير مكتمل (بيانات مفقودة).');
      break;
    }
  }
  return errors;
}

function backupFilename(prefix) {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const stamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}_${pad(now.getHours())}-${pad(now.getMinutes())}`;
  return `${prefix}_${stamp}.json`;
}

function downloadJson(filename, obj) {
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// ---- Backup ---------------------------------------------------------------------

async function exportFullBackup() {
  const storesData = await collectAllStoresData();
  const payload = buildBackupPayload(storesData);
  downloadJson(backupFilename('نسخة_احتياطية'), payload);

  const settings = await getSettings();
  await dbPut('settings', { ...settings, lastBackupAt: new Date().toISOString() });
  return payload;
}

// ---- Restore ----------------------------------------------------------------------

/** Wipes every backed-up store and reloads it from payload.stores, inside one transaction. */
async function restoreFromPayload(payload) {
  await dbTransaction(BACKUP_STORE_NAMES, async (tx) => {
    for (const name of BACKUP_STORE_NAMES) {
      const store = tx.objectStore(name);
      await requestToPromise(store.clear());
      for (const record of payload.stores[name]) {
        store.put(record);
      }
    }
  });
}

async function handleRestoreFileSelected(file) {
  let payload;
  try {
    const text = await file.text();
    payload = JSON.parse(text);
  } catch (err) {
    UI.error(friendlyError('تعذر قراءة الملف. تأكد أنه ملف نسخة احتياطية صحيح.', err));
    return;
  }

  const errors = validateBackupPayload(payload);
  if (errors.length > 0) {
    UI.error(errors[0]);
    return;
  }

  const backupDate = formatDateTimeForDisplay(new Date(payload.exportedAt));
  const confirmed = await UI.confirm({
    title: 'تحذير: استعادة نسخة احتياطية',
    message: `سيتم استبدال جميع البيانات الحالية في النظام بالكامل ببيانات هذه النسخة (تاريخها: ${backupDate}). لا يمكن التراجع عن هذا الإجراء، لكن سيحفظ النظام تلقائيًا نسخة أمان من بياناتك الحالية قبل الاستبدال. هل أنت متأكد من المتابعة؟`,
    confirmLabel: 'نعم، استعادة الآن',
    danger: true,
  });
  if (!confirmed) return;

  try {
    const currentData = await collectAllStoresData();
    downloadJson(backupFilename('نسخة_أمان_قبل_الاستعادة'), buildBackupPayload(currentData));

    await restoreFromPayload(payload);
    UI.success('تمت الاستعادة بنجاح. سيُعاد تحميل النظام الآن.');
    setTimeout(() => window.location.reload(), 1200);
  } catch (err) {
    UI.error(friendlyError('تعذر إتمام الاستعادة. يرجى المحاولة مرة أخرى، أو التواصل مع الدعم الفني إذا استمرت المشكلة.', err));
  }
}

// ---- Rendering: "النسخ الاحتياطي" screen -----------------------------------------

async function renderBackupScreen() {
  const container = document.getElementById('screen-backup');
  const settings = await getSettings();

  container.innerHTML = `
    <h2 class="screen-title">النسخ الاحتياطي</h2>
    <p class="screen-hint">تشمل النسخة الاحتياطية: المنتجات، المخزون، العملاء، المبيعات، الفواتير، المصروفات، والإعدادات.</p>
    <p class="screen-hint" id="last-backup-hint"></p>

    <div class="backup-actions">
      <button type="button" id="btn-create-backup" class="btn btn-primary btn-large">حفظ نسخة احتياطية</button>
      <button type="button" id="btn-restore-backup" class="btn btn-secondary btn-large">استعادة نسخة احتياطية</button>
      <input type="file" id="restore-file-input" accept="application/json,.json" style="display:none">
    </div>
  `;

  updateLastBackupHint(settings);

  document.getElementById('btn-create-backup').addEventListener('click', async () => {
    try {
      await exportFullBackup();
      UI.success('تم حفظ النسخة الاحتياطية بنجاح.');
      updateLastBackupHint(await getSettings());
    } catch (err) {
      UI.error(friendlyError('تعذر حفظ النسخة الاحتياطية. يرجى المحاولة مرة أخرى.', err));
    }
  });

  document.getElementById('btn-restore-backup').addEventListener('click', () => {
    document.getElementById('restore-file-input').click();
  });
  document.getElementById('restore-file-input').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    await handleRestoreFileSelected(file);
  });
}

function updateLastBackupHint(settings) {
  const el = document.getElementById('last-backup-hint');
  if (!el) return;
  el.textContent = settings.lastBackupAt
    ? `آخر نسخة احتياطية: ${formatDateTimeForDisplay(new Date(settings.lastBackupAt))}`
    : 'لم يتم أخذ نسخة احتياطية بعد.';
}
