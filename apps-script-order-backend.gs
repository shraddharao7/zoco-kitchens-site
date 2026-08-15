/**
 * Zoco order backend. Deploy this as a Google Apps Script Web App bound to
 * a Google Sheet, then paste the deployment URL into ORDERS_ENDPOINT in
 * menu.html.
 *
 * Setup:
 * 1. Go to sheets.google.com and create a new blank spreadsheet. Name it
 *    something like "Zoco Orders".
 * 2. In the sheet, go to Extensions -> Apps Script.
 * 3. Delete anything in the editor and paste this whole file in.
 * 4. Click Deploy -> New deployment.
 * 5. Click the gear icon next to "Select type" and choose "Web app".
 * 6. Set "Execute as" to Me, and "Who has access" to Anyone.
 * 7. Click Deploy, authorize it when Google asks (it's your own script,
 *    on your own sheet, so this is safe to approve).
 * 8. Copy the Web app URL it gives you.
 * 9. In menu.html, find the line `const ORDERS_ENDPOINT = '';` near the
 *    top of the <script> block and paste the URL between the quotes.
 * 10. Re-deploy this script (Deploy -> Manage deployments -> edit -> New
 *     version) any time you change this file.
 *
 * Every order placed on the site becomes one row in a sheet tab called
 * "Orders" (created automatically on the first order).
 */

function doPost(e) {
  const sheet = getOrdersSheet_();
  const data = JSON.parse(e.postData.contents);

  sheet.appendRow([
    new Date(),
    data.name || '',
    data.phone || '',
    data.notes || '',
    data.itemsText || '',
    data.total || 0,
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrdersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Orders');
  if (!sheet) {
    sheet = ss.insertSheet('Orders');
    sheet.appendRow(['Timestamp', 'Name', 'Phone', 'Notes', 'Items', 'Total (₹)']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
