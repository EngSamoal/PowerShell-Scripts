/**
 * app.js — application entry point: opens the database, wires up
 * navigation, and renders the initial screen.
 *
 * Phase 1 only implements "الرئيسية" (a placeholder) and "الإعدادات"
 * (fully working). Other menu items show a "قيد التطوير" placeholder
 * until their phase is built, so the whole navigation shell is usable
 * and testable from day one.
 */

const READY_SCREENS = new Set(['dashboard', 'settings', 'products', 'new-sale']);

async function renderDashboardScreen() {
  const container = document.getElementById('screen-dashboard');
  const settings = await getSettings();

  container.innerHTML = `
    <h2 class="screen-title">الرئيسية</h2>
    <p class="screen-hint">حساب هذه الأرقام تلقائيًا من المبيعات الفعلية سيُضاف في مرحلة لاحقة من المشروع.</p>
    <div class="dashboard-grid">
      <div class="stat-card">
        <span class="stat-label">مبيعات اليوم</span>
        <span class="stat-value">0.00 ${settings.currencySymbol}</span>
      </div>
      <div class="stat-card">
        <span class="stat-label">مبيعات الشهر</span>
        <span class="stat-value">0.00 ${settings.currencySymbol}</span>
      </div>
      <div class="stat-card">
        <span class="stat-label">عدد الفواتير اليوم</span>
        <span class="stat-value">0</span>
      </div>
      <div class="stat-card">
        <span class="stat-label">الربح</span>
        <span class="stat-value">0.00 ${settings.currencySymbol}</span>
      </div>
      <div class="stat-card">
        <span class="stat-label">مبالغ غير مدفوعة</span>
        <span class="stat-value">0.00 ${settings.currencySymbol}</span>
      </div>
      <div class="stat-card">
        <span class="stat-label">منتجات منخفضة المخزون</span>
        <span class="stat-value">0</span>
      </div>
    </div>
  `;
}

function renderPlaceholderScreen(screenId, title) {
  const container = document.getElementById(screenId);
  if (!container) return;
  container.innerHTML = `
    <h2 class="screen-title">${title}</h2>
    <div class="placeholder-box">
      <p>هذه الشاشة قيد التطوير وستُضاف في مرحلة لاحقة من المشروع.</p>
    </div>
  `;
}

async function showScreenById(screenId) {
  try {
    if (screenId === 'dashboard') await renderDashboardScreen();
    if (screenId === 'settings') await renderSettingsScreen();
    if (screenId === 'products') await renderProductsScreen();
    if (screenId === 'new-sale') await renderNewSaleScreen();
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

    const placeholders = [
      ['screen-customers', 'العملاء'],
      ['screen-invoices', 'الفواتير'],
      ['screen-expenses', 'المصروفات'],
      ['screen-reports', 'التقارير'],
      ['screen-backup', 'النسخ الاحتياطي'],
    ];
    placeholders.forEach(([id, title]) => renderPlaceholderScreen(id, title));

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

document.addEventListener('DOMContentLoaded', initApp);
