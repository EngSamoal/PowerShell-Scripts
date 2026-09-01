/**
 * utils.js — money formatting and small shared helpers.
 *
 * All monetary amounts are stored and calculated internally as integer
 * halalas (1 SAR = 100 halalas) to avoid JavaScript floating-point drift.
 * Only formatToDisplay() converts to a decimal string for the screen.
 */

/** Converts a decimal SAR amount (e.g. from a form field) to integer halalas. */
function sarToHalalas(sarAmount) {
  return Math.round(Number(sarAmount) * 100);
}

/** Converts integer halalas to a decimal SAR number. */
function halalasToSar(halalas) {
  return halalas / 100;
}

/** Formats integer halalas as a two-decimal string with the currency label, e.g. "37.16 ريال". */
function formatCurrency(halalas, currencySymbol) {
  const symbol = currencySymbol || 'ريال';
  const sar = halalasToSar(halalas);
  return `${sar.toFixed(2)} ${symbol}`;
}

/** Rounds half away from zero — the rule used for fee calculations. */
function roundHalfUp(value) {
  return Math.sign(value) * Math.round(Math.abs(value));
}

/**
 * Calculates the platform fees for a subtotal (in halalas) given the
 * fee settings snapshot in effect at the time.
 * feePercent is a plain number like 8 (meaning 8%).
 */
function calculateFees(subtotalHalalas, feePercent, feeFixedHalalas) {
  const percentFee = roundHalfUp(subtotalHalalas * (feePercent / 100));
  return {
    percentFeeHalalas: percentFee,
    fixedFeeHalalas: feeFixedHalalas,
    grandTotalHalalas: subtotalHalalas + percentFee + feeFixedHalalas,
  };
}

/** Formats a Date as YYYY-MM-DD for storage/indexing. */
function formatDateForStorage(date) {
  const d = date || new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

/** Formats a Date as a readable Arabic-friendly date + time string. */
function formatDateTimeForDisplay(date) {
  const d = date || new Date();
  return d.toLocaleString('ar-SA', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
  });
}

/** Splits a total unit count into boxes + remaining units for display, e.g. "8 صناديق + 9 قلم". */
function unitsToBoxesAndRemainder(totalUnits, unitsPerBox) {
  if (!unitsPerBox || unitsPerBox <= 0) {
    return { boxes: 0, remainder: totalUnits };
  }
  return {
    boxes: Math.floor(totalUnits / unitsPerBox),
    remainder: totalUnits % unitsPerBox,
  };
}

/** Wraps a technical error so a friendly Arabic message reaches the user, never a raw stack trace. */
function friendlyError(userMessage, technicalError) {
  if (technicalError) {
    console.error(userMessage, technicalError);
  }
  return userMessage;
}

/** Minimal HTML-escaping for values interpolated into template strings across the app. */
function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (ch) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[ch]));
}
