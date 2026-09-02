/**
 * app.js — application entry point: opens the database, wires up
 * navigation, and renders the initial screen.
 */

async function renderDashboardScreen() {
  const container = document.getElementById('screen-dashboard');
  const settings = await getSettings();

  const today = formatDateForStorage();
  const monthStart = `${today.slice(0, 7)}-01`;

  const [invoices, products, monthProfit] = await Promise.all([
    listAllInvoices(),
    listAllProducts(),
    computeProfitSummary(monthStart, today),
  ]);

  const completedInvoices = invoices.filter((inv) => inv.status === 'completed');
  const todaysInvoices = completedInvoices.filter((inv) => inv.date === today);
  const salesToday = todaysInvoices.reduce((sum, inv) => sum + inv.grandTotalHalalas, 0);
  const salesMonth = completedInvoices
    .filter((inv) => inv.date >= monthStart && inv.date <= today)
    .reduce((sum, inv) => sum + inv.grandTotalHalalas, 0);
  const unpaidTotal = completedInvoices.reduce((sum, inv) => sum + inv.remainingAmountHalalas, 0);
  const lowStockCount = products.filter((p) => p.status === 'active' && p.totalUnits <= p.minStockUnits).length;

  container.innerHTML = `
    <h2 class="screen-title">الرئيسية</h2>
    ${backupReminderHtml(settings)}
    <div class="dashboard-grid">
      <div class="stat-card">
        <span class="stat-label">مبيعات اليوم</span>
        <span class="stat-value">${formatCurrency(salesToday, settings.currencySymbol)}</span>
      </div>
      <div class="stat-card">
        <span class="stat-label">مبيعات الشهر</span>
        <span class="stat-value">${formatCurrency(salesMonth, settings.currencySymbol)}</span>
      </div>
      <div class="stat-card">
        <span class="stat-label">عدد الفواتير اليوم</span>
        <span class="stat-value">${todaysInvoices.length}</span>
      </div>
      <div class="stat-card">
        <span class="stat-label">الربح الصافي (هذا الشهر)</span>
        <span class="stat-value">${formatCurrency(monthProfit.netProfitHalalas, settings.currencySymbol)}</span>
      </div>
      <div class="stat-card">
        <span class="stat-label">مبالغ غير مدفوعة</span>
        <span class="stat-value">${formatCurrency(unpaidTotal, settings.currencySymbol)}</span>
      </div>
      <div class="stat-card">
        <span class="stat-label">منتجات منخفضة المخزون</span>
        <span class="stat-value">${lowStockCount}</span>
      </div>
    </div>
  `;

  const reminderBtn = document.getElementById('btn-backup-reminder-goto');
  if (reminderBtn) reminderBtn.addEventListener('click', () => showScreenById('backup'));
}

/** A gentle nudge, not a nag: only shown when there's truly no recent backup, and it's just one line + a link. */
function backupReminderHtml(settings) {
  const BACKUP_REMINDER_DAYS = 7;
  if (!settings.lastBackupAt) {
    return `
      <div class="backup-reminder">
        <span>لم تقم بأخذ أي نسخة احتياطية بعد. يُنصح بأخذ نسخة للحفاظ على بياناتك.</span>
        <button type="button" id="btn-backup-reminder-goto" class="link-btn">أخذ نسخة احتياطية الآن</button>
      </div>
    `;
  }
  const daysSince = Math.floor((Date.now() - new Date(settings.lastBackupAt).getTime()) / 86400000);
  if (daysSince < BACKUP_REMINDER_DAYS) return '';
  return `
    <div class="backup-reminder">
      <span>آخر نسخة احتياطية كانت منذ ${daysSince} يومًا. يُنصح بأخذ نسخة جديدة.</span>
      <button type="button" id="btn-backup-reminder-goto" class="link-btn">أخذ نسخة احتياطية الآن</button>
    </div>
  `;
}

async function showScreenById(screenId) {
  try {
    if (screenId === 'dashboard') await renderDashboardScreen();
    if (screenId === 'settings') await renderSettingsScreen();
    if (screenId === 'products') await renderProductsScreen();
    if (screenId === 'new-sale') await renderNewSaleScreen();
    if (screenId === 'invoices') await renderInvoicesScreen();
    if (screenId === 'customers') await renderCustomersScreen();
    if (screenId === 'expenses') await renderExpensesScreen();
    if (screenId === 'reports') await renderReportsScreen();
    if (screenId === 'backup') await renderBackupScreen();
    UI.showScreen(`screen-${screenId}`);
  } catch (err) {
    UI.error(friendlyError('تعذر فتح هذه الشاشة. يرجى المحاولة مرة أخرى.', err));
  }
}

function wireNavigation() {
  document.querySelectorAll('.nav-item').forEach((item) => {
    item.addEventListener('click', () => {
      showScreenById(item.dataset.screen);
    });
  });
}

async function initApp() {
  try {
    await openDatabase();
    wireNavigation();

    await showScreenById('dashboard');
  } catch (err) {
    document.body.innerHTML = `
      <div class="fatal-error">
        <h1>تعذر تشغيل النظام</h1>
        <p>يرجى إغلاق النافذة وإعادة فتح "نظام المبيعات" مرة أخرى. إذا استمرت المشكلة يرجى مراجعة الدعم الفني.</p>
      </div>
    `;
    console.error(err);
  }
}

/**
 * Safety net for anything a screen's own try/catch missed — the user must
 * never see a raw "Uncaught TypeError" or a silently broken screen.
 */
function installGlobalErrorHandlers() {
  window.addEventListener('error', (event) => {
    console.error('Unhandled error:', event.error || event.message);
    UI.error('تعذر إتمام العملية. يرجى المحاولة مرة أخرى.');
  });
  window.addEventListener('unhandledrejection', (event) => {
    console.error('Unhandled promise rejection:', event.reason);
    UI.error('تعذر إتمام العملية. يرجى المحاولة مرة أخرى.');
  });
}

installGlobalErrorHandlers();
document.addEventListener('DOMContentLoaded', initApp);
