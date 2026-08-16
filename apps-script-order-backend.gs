/**
 * Zoco order + catering-lead backend. Deployed as a Google Apps Script Web
 * App bound to the "Zoco Orders" Google Sheet (via SHEET_ID below, not
 * container-bound — container-bound scripts kept failing to open in this
 * project).
 *
 * To update the live deployment after editing this file:
 * 1. Open script.google.com, open the "Zoco Order Backend" project.
 * 2. Replace all the code with this file's contents.
 * 3. Make sure SHEET_ID below still has the real sheet ID.
 * 4. Save (Cmd+S). Check the function dropdown near Run shows "doPost",
 *    "getOrdersSheet_", and "getCateringSheet_" as top-level entries (not
 *    nested inside anything).
 * 5. Deploy -> Manage deployments -> pencil icon -> Version: New version -> Deploy.
 *    (If that ever silently fails again, Deploy -> New deployment works too,
 *    it just hands back a different URL that has to be re-pasted into
 *    ORDERS_ENDPOINT in menu.html and CATERING_ENDPOINT in catering.html.)
 *
 * Two kinds of submissions land here, routed by the `type` field the page
 * sends:
 * - Orders from menu.html (no `type`, or type !== 'catering') go to the
 *   "Orders" tab. Items are split into five columns (Mains / Sauces / Dips /
 *   Beverages / Sides), and the optional customer-research survey answers
 *   get their own columns.
 * - Catering leads from catering.html (type: 'catering') go to a separate
 *   "Catering Leads" tab.
 */

const SHEET_ID = '16eomiZE6yrdbsiLhBc0m6SnZxxz5fV-dbbK0IM8Tjtk';

const ORDERS_HEADER = [
  'Timestamp', 'Name', 'Phone', 'Notes',
  'Mains', 'Sauces', 'Dips', 'Beverages', 'Sides', 'Total (₹)',
  'Age Group', 'Occupation', 'Order Frequency', 'Group Size', 'Spend Per Person',
  'Meal Time', 'Go-To Cuisine', 'Priorities', 'Checks Nutrition', 'Recommends to Others',
  'Meal Plan Interest (times/month)',
];

const CATERING_HEADER = [
  'Timestamp', 'Name', 'Phone', 'Number of People', 'Event Type', 'Event Date', 'Event Time',
];

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  if (data.type === 'catering') {
    const sheet = getCateringSheet_();
    sheet.appendRow([
      new Date(),
      data.name || '',
      data.phone || '',
      data.headcount || '',
      data.eventType || '',
      data.date || '',
      data.time || '',
    ]);
  } else {
    const sheet = getOrdersSheet_();
    sheet.appendRow([
      new Date(),
      data.name || '',
      data.phone || '',
      data.notes || '',
      data.mainsText || '',
      data.saucesText || '',
      data.dipsText || '',
      data.beveragesText || '',
      data.sidesText || '',
      data.total || 0,
      data.ageGroup || '',
      data.occupation || '',
      data.orderFrequency || '',
      data.groupSize || '',
      data.spendPerPerson || '',
      data.mealTime || '',
      data.cuisine || '',
      data.priorities || '',
      data.nutritionAttention || '',
      data.recommend || '',
      data.subscriptionFrequency || '',
    ]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrdersSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Orders');
  if (!sheet) {
    sheet = ss.insertSheet('Orders');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(ORDERS_HEADER);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getCateringSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Catering Leads');
  if (!sheet) {
    sheet = ss.insertSheet('Catering Leads');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(CATERING_HEADER);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
