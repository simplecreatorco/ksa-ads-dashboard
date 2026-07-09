/**
 * KSA Ads Dashboard — Google Apps Script
 *
 * This script powers the data pipeline for the KSA Ads Dashboard:
 * 1. Daily Meta API pull (3am Central via time-based trigger)
 * 2. Checkbox state clearing (runs after daily pull)
 * 3. Web app endpoint for dashboard writes (config, prices, checkboxes, new clients/offers)
 *
 * NAMING CONVENTION: All tab names and Price History rows use the client SLUG (e.g. "htp"),
 * not the full client name. This keeps tab names short and consistent.
 * Tab examples: "htp - Daily Meta", "htp - Config", "htp - What to Wear Tiny Offer Purchases"
 *
 * SCRIPT PROPERTIES (set once in Project Settings > Script Properties):
 *   META_TOKEN    — Meta System User access token (ads_read permission)
 *   SHEET_ID      — Google Sheet ID
 *   WEBAPP_SECRET — shared secret for verifying requests from dashboard
 *   API_KEY       — Google Sheets API key (for dashboard reads)
 *
 * TIMEZONE: Set to America/Chicago in File > Project Settings
 *
 * Sheet ID: 1OJ68ds0m-pQbV4Fp25bjBPQkNAHoRM-_Jqz3ZRUKHpg
 */


// ============================================
// CONFIGURATION
// ============================================

function getConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    META_TOKEN: props.getProperty('META_TOKEN'),
    SHEET_ID: props.getProperty('SHEET_ID'),
    WEBAPP_SECRET: props.getProperty('WEBAPP_SECRET')
  };
}

function getSheet_() {
  const config = getConfig_();
  return SpreadsheetApp.openById(config.SHEET_ID);
}


// ============================================
// DAILY META API PULL
// ============================================

/**
 * Main daily trigger function. Set up a time-based trigger to run between 3-4am Central.
 * Reads all active clients, pulls yesterday's Meta campaign data, writes to Daily Meta tabs.
 */
function runDailyPull() {
  const ss = getSheet_();
  const config = getConfig_();

  // Get all active clients
  const clientsSheet = ss.getSheetByName('Clients');
  if (!clientsSheet) {
    Logger.log('ERROR: Clients tab not found');
    return;
  }

  const clientData = clientsSheet.getDataRange().getValues();
  const headers = clientData[0];
  const nameCol = headers.indexOf('Client Name');
  const slugCol = headers.indexOf('Slug');
  const accountCol = headers.indexOf('Meta Account ID(s)');
  const statusCol = headers.indexOf('Status');

  if (nameCol === -1 || slugCol === -1 || accountCol === -1 || statusCol === -1) {
    Logger.log('ERROR: Required columns not found in Clients tab');
    return;
  }

  // Loop through active clients
  for (let i = 1; i < clientData.length; i++) {
    const row = clientData[i];
    const clientName = row[nameCol];
    const slug = String(row[slugCol]).trim();
    const accountIds = String(row[accountCol]).trim();
    const status = String(row[statusCol]).trim();

    if (status !== 'Active' || !slug || !accountIds) continue;

    Logger.log('Pulling data for: ' + clientName + ' (' + slug + ')');

    // Split multiple account IDs by comma
    const ids = accountIds.split(',').map(function(id) { return id.trim(); });

    // Get or create the Daily Meta tab using SLUG
    const tabName = slug + ' - Daily Meta';
    let metaSheet = ss.getSheetByName(tabName);
    if (!metaSheet) {
      metaSheet = createDailyMetaTab_(ss, slug);
    }

    // Pull data for each account
    for (let j = 0; j < ids.length; j++) {
      const accountId = ids[j];
      if (!accountId) continue;

      try {
        pullAccountData_(metaSheet, accountId, config.META_TOKEN, 'yesterday');
      } catch (e) {
        Logger.log('ERROR pulling account ' + accountId + ' for ' + slug + ': ' + e.message);
      }
    }
  }

  // Clear yesterday's checkbox states
  clearYesterdayCheckboxes();

  Logger.log('Daily pull complete');
}


/**
 * ONE-TIME backfill: Pull the last 30 days of Meta data for all active clients.
 * Run this manually once to seed historical data. Uses dedup so it's safe to
 * run multiple times — existing rows are never duplicated.
 * After running this once, the daily trigger (runDailyPull) handles day-by-day.
 */
function backfill30Days() {
  const ss = getSheet_();
  const config = getConfig_();

  var clientsSheet = ss.getSheetByName('Clients');
  if (!clientsSheet) { Logger.log('ERROR: Clients tab not found'); return; }

  var clientData = clientsSheet.getDataRange().getValues();
  var headers = clientData[0];
  var nameCol = headers.indexOf('Client Name');
  var slugCol = headers.indexOf('Slug');
  var accountCol = headers.indexOf('Meta Account ID(s)');
  var statusCol = headers.indexOf('Status');

  for (var i = 1; i < clientData.length; i++) {
    var row = clientData[i];
    var clientName = row[nameCol];
    var slug = String(row[slugCol]).trim();
    var accountIds = String(row[accountCol]).trim();
    var status = String(row[statusCol]).trim();

    if (status !== 'Active' || !slug || !accountIds) continue;

    Logger.log('Backfilling 30 days for: ' + clientName + ' (' + slug + ')');

    var ids = accountIds.split(',').map(function(id) { return id.trim(); });
    var tabName = slug + ' - Daily Meta';
    var metaSheet = ss.getSheetByName(tabName);
    if (!metaSheet) { metaSheet = createDailyMetaTab_(ss, slug); }

    for (var j = 0; j < ids.length; j++) {
      if (!ids[j]) continue;
      try {
        pullAccountData_(metaSheet, ids[j], config.META_TOKEN, '30d');
      } catch (e) {
        Logger.log('ERROR backfilling ' + ids[j] + ' for ' + slug + ': ' + e.message);
      }
    }
  }

  Logger.log('30-day backfill complete');
}


/**
 * Pull campaign-level insights from one Meta ad account.
 * @param {string} dateMode — 'yesterday' for daily runs, or '30d' for backfill
 * Uses dedup: only writes rows for (date, campaign) combos that don't already exist.
 */
function pullAccountData_(metaSheet, accountId, token, dateMode) {
  // Clean account ID — add 'act_' prefix if not present
  if (!accountId.startsWith('act_')) {
    accountId = 'act_' + accountId;
  }

  // Build date portion of the URL
  var datePart;
  if (dateMode === '30d') {
    var today = new Date();
    var thirtyDaysAgo = new Date(today);
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    var since = Utilities.formatDate(thirtyDaysAgo, 'America/Chicago', 'yyyy-MM-dd');
    var until = Utilities.formatDate(today, 'America/Chicago', 'yyyy-MM-dd');
    var timeRangeJson = '{"since":"' + since + '","until":"' + until + '"}';
    datePart = '&time_range=' + encodeURIComponent(timeRangeJson) + '&time_increment=1';
  } else {
    datePart = '&date_preset=yesterday';
  }

  const url = 'https://graph.facebook.com/v21.0/' + accountId + '/insights'
    + '?fields=campaign_name,spend,impressions,reach,clicks,actions,action_values,cpm,cpc,ctr,purchase_roas'
    + datePart
    + '&level=campaign'
    + '&limit=500'
    + '&access_token=' + token;

  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const statusCode = response.getResponseCode();

  if (statusCode !== 200) {
    Logger.log('Meta API error (' + statusCode + '): ' + response.getContentText());
    return;
  }

  const json = JSON.parse(response.getContentText());
  var allData = json.data || [];

  // Handle pagination — Meta may return results across multiple pages
  var nextUrl = json.paging && json.paging.next ? json.paging.next : null;
  while (nextUrl) {
    var nextResponse = UrlFetchApp.fetch(nextUrl, { muteHttpExceptions: true });
    if (nextResponse.getResponseCode() !== 200) break;
    var nextJson = JSON.parse(nextResponse.getContentText());
    allData = allData.concat(nextJson.data || []);
    nextUrl = nextJson.paging && nextJson.paging.next ? nextJson.paging.next : null;
  }

  if (allData.length === 0) {
    Logger.log('No campaign data returned for account ' + accountId);
    return;
  }

  // Build set of existing (date|campaign) keys for dedup
  var existingKeys = {};
  var existingData = metaSheet.getDataRange().getValues();
  for (var e = 1; e < existingData.length; e++) {
    var existDate = String(existingData[e][0]);
    if (existDate.length > 10) existDate = existDate.substring(0, 10);
    var existCampaign = String(existingData[e][1]);
    existingKeys[existDate + '|' + existCampaign] = true;
  }

  // Parse and collect only NEW rows
  const rows = [];
  for (let i = 0; i < allData.length; i++) {
    const campaign = allData[i];
    var dateStart = campaign.date_start || '';
    var campaignName = campaign.campaign_name || '';

    // Skip if this date+campaign combo already exists
    if (existingKeys[dateStart + '|' + campaignName]) continue;

    // Extract actions — leads and purchases
    var leads = 0;
    var purchases = 0;
    var purchaseValue = 0;

    if (campaign.actions) {
      for (var a = 0; a < campaign.actions.length; a++) {
        if (campaign.actions[a].action_type === 'lead') {
          leads = Number(campaign.actions[a].value) || 0;
        }
        if (campaign.actions[a].action_type === 'purchase' || campaign.actions[a].action_type === 'offsite_conversion.fb_pixel_purchase') {
          purchases = Number(campaign.actions[a].value) || 0;
        }
      }
    }

    if (campaign.action_values) {
      for (var v = 0; v < campaign.action_values.length; v++) {
        if (campaign.action_values[v].action_type === 'purchase' || campaign.action_values[v].action_type === 'offsite_conversion.fb_pixel_purchase') {
          purchaseValue = Number(campaign.action_values[v].value) || 0;
        }
      }
    }

    // Extract ROAS from purchase_roas array
    var roas = 0;
    if (campaign.purchase_roas && campaign.purchase_roas.length > 0) {
      roas = Number(campaign.purchase_roas[0].value) || 0;
    }

    rows.push([
      dateStart,                                     // Date
      campaignName,                                  // Campaign Name
      Number(campaign.spend) || 0,                   // Amount Spent
      Number(campaign.impressions) || 0,             // Impressions
      Number(campaign.reach) || 0,                   // Reach
      Number(campaign.clicks) || 0,                  // Clicks
      leads,                                         // Leads
      purchases,                                     // Purchases
      purchaseValue,                                 // Purchase Value
      Number(campaign.cpm) || 0,                     // CPM
      Number(campaign.cpc) || 0,                     // CPC
      Number(campaign.ctr) || 0,                     // CTR
      roas                                           // ROAS
    ]);
  }

  // Append only new rows
  if (rows.length > 0) {
    metaSheet.getRange(metaSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    Logger.log('Wrote ' + rows.length + ' NEW rows for account ' + accountId + ' (skipped ' + (allData.length - rows.length) + ' existing)');
  } else {
    Logger.log('All ' + allData.length + ' rows already exist for account ' + accountId + ' — nothing to write');
  }
}


// ============================================
// CHECKBOX STATE CLEARING
// ============================================

/**
 * Remove all rows from CHECKBOX_STATE where Date < today.
 * Called at the end of the daily pull.
 */
function clearYesterdayCheckboxes() {
  const ss = getSheet_();
  const sheet = ss.getSheetByName('CHECKBOX_STATE');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return; // only header row

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Work backwards to safely delete rows
  for (var i = data.length - 1; i >= 1; i--) {
    var rowDate = new Date(data[i][0]);
    rowDate.setHours(0, 0, 0, 0);
    if (rowDate < today) {
      sheet.deleteRow(i + 1);
    }
  }

  Logger.log('Cleared old checkbox states');
}


// ============================================
// WEB APP ENDPOINT
// ============================================

/**
 * Handle POST requests from the internal dashboard.
 * Verifies the shared secret before processing any action.
 */
function doPost(e) {
  // Parse body first, then verify secret from body or URL params
  const config = getConfig_();

  var payload;
  try {
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      payload = e.parameter;
    }
  } catch (parseErr) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: 'Invalid JSON: ' + parseErr.message
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // Check secret from body first, then URL params
  var secret = (payload && payload.secret) || e.parameter.secret || '';
  if (secret !== config.WEBAPP_SECRET) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: 'Unauthorized'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var action = payload.action;
    var result;

    switch (action) {
      case 'save_config':
        result = handleSaveConfig_(payload);
        break;
      case 'save_price_change':
        result = handleSavePriceChange_(payload);
        break;
      case 'save_checkbox_state':
        result = handleSaveCheckboxState_(payload);
        break;
      case 'setup_new_client':
        result = handleSetupNewClient_(payload);
        break;
      case 'setup_new_offer':
        result = handleSetupNewOffer_(payload);
        break;
      case 'save_purchases':
        result = handleSavePurchases_(payload);
        break;
      case 'save_client':
        result = handleSaveClient_(payload);
        break;
      case 'rename_offer':
        result = handleRenameOffer_(payload);
        break;
      case 'rename_product':
        result = handleRenameProduct_(payload);
        break;
      case 'save_product_name':
        result = handleSaveProductName_(payload);
        break;
      case 'delete_purchase':
        result = handleDeletePurchase_(payload);
        break;
      case 'log_ig_purchase':
        result = handleLogIGPurchase_(payload);
        break;
      case 'delete_ig_purchase':
        result = handleDeleteIGPurchase_(payload);
        break;
      case 'reorder_offers':
        result = handleReorderOffers_(payload);
        break;
      case 'add_product':
        result = handleAddProduct_(payload);
        break;
      case 'set_client_status':
        result = handleSetClientStatus_(payload);
        break;
      case 'push_client_snapshot':
        result = handlePushClientSnapshot_(payload);
        break;
      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle GET requests — returns dashboard config from Script Properties.
 * The dashboard fetches this on load so no secrets live in the GitHub repo.
 */
function doGet(e) {
  const props = PropertiesService.getScriptProperties();
  return ContentService.createTextOutput(JSON.stringify({
    SHEET_ID: props.getProperty('SHEET_ID'),
    API_KEY: props.getProperty('API_KEY'),
    WEBAPP_SECRET: props.getProperty('WEBAPP_SECRET')
  })).setMimeType(ContentService.MimeType.JSON);
}


// ============================================
// ACTION HANDLERS
// ============================================

/**
 * save_config — Write offer/product changes to [slug] - Config tab
 * Payload: { clientSlug, offerName, products: [{name, position, currentPrice, active}] }
 */
function handleSaveConfig_(payload) {
  const ss = getSheet_();
  const slug = payload.clientSlug;
  if (!slug) return { success: false, error: 'No clientSlug provided' };

  const tabName = slug + ' - Config';
  var configSheet = ss.getSheetByName(tabName);
  if (!configSheet) return { success: false, error: 'Config tab not found: ' + tabName };

  var products = payload.products;
  if (!products || products.length === 0) return { success: false, error: 'No products provided' };

  // Get existing data
  var data = configSheet.getDataRange().getValues();

  for (var p = 0; p < products.length; p++) {
    var product = products[p];
    var oldName = product.name;
    var newName = product.newName || oldName; // newName present only on reorder

    for (var r = 1; r < data.length; r++) {
      if (data[r][0] === payload.offerName && data[r][1] === oldName) {
        configSheet.getRange(r + 1, 2).setValue(newName);              // Product Type (rename)
        configSheet.getRange(r + 1, 3).setValue(product.position);     // Position
        configSheet.getRange(r + 1, 4).setValue(product.currentPrice); // Current Price
        configSheet.getRange(r + 1, 5).setValue(product.active);       // Active
        if (product.label !== undefined) {
          configSheet.getRange(r + 1, 6).setValue(product.label);      // Product Label
        }
        // Keep in-memory data in sync so later iterations see updated names
        data[r][1] = newName;

        // If name changed, propagate to Price History and Purchases tab header
        if (newName !== oldName) {
          var priceSheet = ss.getSheetByName('Price History');
          if (priceSheet) {
            var priceData = priceSheet.getDataRange().getValues();
            for (var j = 1; j < priceData.length; j++) {
              if (String(priceData[j][0]) === slug && priceData[j][1] === payload.offerName && priceData[j][2] === oldName) {
                priceSheet.getRange(j + 1, 3).setValue(newName);
              }
            }
          }
          var purchasesTabName = slug + ' - ' + payload.offerName + ' Purchases';
          var purchasesSheet = ss.getSheetByName(purchasesTabName);
          if (purchasesSheet) {
            var phHeaders = purchasesSheet.getRange(1, 1, 1, purchasesSheet.getLastColumn()).getValues()[0];
            for (var h = 0; h < phHeaders.length; h++) {
              if (phHeaders[h] === oldName) {
                purchasesSheet.getRange(1, h + 1).setValue(newName);
                break;
              }
            }
          }
        }
        break;
      }
    }
    // No append on reorder — only existing products are being reordered
  }

  return { success: true };
}


/**
 * add_product — Append a new product to an existing offer.
 * Adds a Config row, inserts a column in the Purchases tab before Total Units,
 * and appends the starting price to Price History.
 * Payload: { clientSlug, offerName, productName, price, label, position }
 */
function handleAddProduct_(payload) {
  const ss = getSheet_();
  var slug = payload.clientSlug;
  if (!slug) return { success: false, error: 'No clientSlug provided' };
  if (!payload.offerName) return { success: false, error: 'No offerName provided' };
  if (!payload.productName) return { success: false, error: 'No productName provided' };

  var configSheet = ss.getSheetByName(slug + ' - Config');
  if (!configSheet) return { success: false, error: 'Config tab not found: ' + slug + ' - Config' };

  var data = configSheet.getDataRange().getValues();
  var offerType = '';
  var funnelLength = '';
  var maxPosition = 0;
  for (var r = 1; r < data.length; r++) {
    if (data[r][0] === payload.offerName) {
      if (data[r][1] === payload.productName) {
        return { success: false, error: 'Product already exists: ' + payload.productName };
      }
      offerType = data[r][6] || '';
      funnelLength = data[r][7] || '';
      var pos = parseInt(data[r][2]) || 0;
      if (pos > maxPosition) maxPosition = pos;
    }
  }
  if (maxPosition === 0) return { success: false, error: 'Offer not found in Config: ' + payload.offerName };

  var position = Number(payload.position) || (maxPosition + 1);
  var price = Number(payload.price) || 0;

  configSheet.appendRow([
    payload.offerName,
    payload.productName,
    position,
    price,
    'Yes',
    payload.label || '',
    offerType,
    funnelLength
  ]);

  // Insert a column for the new product in the Purchases tab, before Total Units
  var purchasesSheet = ss.getSheetByName(slug + ' - ' + payload.offerName + ' Purchases');
  if (purchasesSheet) {
    var headers = purchasesSheet.getRange(1, 1, 1, purchasesSheet.getLastColumn()).getValues()[0];
    var totalUnitsCol = -1;
    var alreadyThere = false;
    for (var h = 0; h < headers.length; h++) {
      if (headers[h] === 'Total Units') totalUnitsCol = h + 1; // 1-indexed
      if (headers[h] === payload.productName) alreadyThere = true;
    }
    if (!alreadyThere) {
      if (totalUnitsCol > 0) {
        purchasesSheet.insertColumnBefore(totalUnitsCol);
        purchasesSheet.getRange(1, totalUnitsCol).setValue(payload.productName).setFontWeight('bold');
      } else {
        var lastCol = purchasesSheet.getLastColumn() + 1;
        purchasesSheet.getRange(1, lastCol).setValue(payload.productName).setFontWeight('bold');
      }
    }
  }

  // Record the starting price in Price History
  var priceSheet = ss.getSheetByName('Price History');
  if (priceSheet) {
    priceSheet.appendRow([
      slug,
      payload.offerName,
      payload.productName,
      price,
      payload.effectiveDate || new Date().toISOString().split('T')[0]
    ]);
  }

  return { success: true, message: 'Added ' + payload.productName + ' to ' + payload.offerName };
}


/**
 * save_price_change — Append row to Price History tab (uses SLUG as client identifier)
 * Payload: { clientSlug, offerName, productName, newPrice, effectiveDate }
 */
function handleSavePriceChange_(payload) {
  const ss = getSheet_();
  const slug = payload.clientSlug;
  if (!slug) return { success: false, error: 'No clientSlug provided' };

  var priceSheet = ss.getSheetByName('Price History');
  if (!priceSheet) return { success: false, error: 'Price History tab not found' };

  // Use SLUG in Price History, not full client name
  priceSheet.appendRow([
    slug,
    payload.offerName,
    payload.productName,
    Number(payload.newPrice),
    payload.effectiveDate || new Date().toISOString().split('T')[0]
  ]);

  // Also update current price in Config tab (uses slug for tab name)
  var configTab = slug + ' - Config';
  var configSheet = ss.getSheetByName(configTab);
  if (configSheet) {
    var configData = configSheet.getDataRange().getValues();
    for (var i = 1; i < configData.length; i++) {
      if (configData[i][0] === payload.offerName && configData[i][1] === payload.productName) {
        configSheet.getRange(i + 1, 4).setValue(Number(payload.newPrice));
        break;
      }
    }
  }

  return { success: true };
}


/**
 * save_checkbox_state — Add or remove row in CHECKBOX_STATE
 * Payload: { clientSlug, campaignName, action, checked, date }
 */
function handleSaveCheckboxState_(payload) {
  const ss = getSheet_();
  var sheet = ss.getSheetByName('CHECKBOX_STATE');
  if (!sheet) return { success: false, error: 'CHECKBOX_STATE tab not found' };

  var date = payload.date || new Date().toISOString().split('T')[0];

  if (payload.checked === true || payload.checked === 'true') {
    // Add row
    sheet.appendRow([date, payload.clientSlug, payload.campaignName, payload.action || '']);
    return { success: true };
  } else {
    // Remove matching row
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).substring(0, 10) === date &&
          data[i][1] === payload.clientSlug &&
          data[i][2] === payload.campaignName) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: true, note: 'No matching row found to remove' };
  }
}


/**
 * setup_new_client — Add to Clients tab, create Daily Meta + Config tabs using SLUG
 * Payload: { clientName, slug, metaAccountIds }
 */
function handleSetupNewClient_(payload) {
  const ss = getSheet_();

  // Check if client already exists
  var clientsSheet = ss.getSheetByName('Clients');
  if (!clientsSheet) return { success: false, error: 'Clients tab not found' };

  var existingData = clientsSheet.getDataRange().getValues();
  for (var i = 1; i < existingData.length; i++) {
    if (existingData[i][1] === payload.slug) {
      return { success: false, error: 'Client slug already exists: ' + payload.slug };
    }
  }

  // Add to Clients tab
  clientsSheet.appendRow([
    payload.clientName,
    payload.slug,
    payload.metaAccountIds || '',
    'Active',
    ''
  ]);

  // Create Daily Meta and Config tabs using SLUG
  createDailyMetaTab_(ss, payload.slug);
  createConfigTab_(ss, payload.slug);

  return { success: true, message: 'Client created: ' + payload.clientName + ' (' + payload.slug + ')' };
}


/**
 * setup_new_offer — Create Purchases tab, add to Config, add to Price History (all using SLUG)
 * Payload: { clientSlug, offerName, products: [{name, position, price}] }
 */
function handleSetupNewOffer_(payload) {
  const ss = getSheet_();
  const slug = payload.clientSlug;
  if (!slug) return { success: false, error: 'No clientSlug provided' };

  var offerType = payload.offerType || 'Product Funnel';

  // IG Profile Visits has no products — different tab structure
  if (offerType === 'IG Profile Visits') {
    return handleSetupIGOffer_(ss, payload);
  }

  var products = payload.products;
  if (!products || products.length === 0) return { success: false, error: 'No products provided' };

  // Sort products by position
  products.sort(function(a, b) { return (a.position || 0) - (b.position || 0); });

  // Create Purchases tab using SLUG
  var purchasesTabName = slug + ' - ' + payload.offerName + ' Purchases';
  var existingTab = ss.getSheetByName(purchasesTabName);
  if (existingTab) {
    return { success: false, error: 'Purchases tab already exists: ' + purchasesTabName };
  }

  var purchasesSheet = ss.insertSheet(purchasesTabName);

  // Set header row: Date | [Product names in position order] | Total Units | Total Revenue
  var headers = ['Date'];
  for (var p = 0; p < products.length; p++) {
    headers.push(products[p].name);
  }
  headers.push('Total Units');
  headers.push('Total Revenue');
  purchasesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Bold the header row
  purchasesSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  // Add to Config tab using SLUG
  var configTabName = slug + ' - Config';
  var configSheet = ss.getSheetByName(configTabName);
  if (!configSheet) {
    configSheet = createConfigTab_(ss, slug);
  }

  // Migrate header row to current 8-column schema if it was created before Offer Type / Funnel Length were added
  var currentHeaders = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
  var requiredHeaders = ['Offer Name', 'Product Type', 'Position', 'Current Price', 'Active', 'Product Name', 'Offer Type', 'Funnel Length'];
  if (currentHeaders.length < requiredHeaders.length || currentHeaders[6] !== 'Offer Type') {
    configSheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    configSheet.getRange(1, 1, 1, requiredHeaders.length).setFontWeight('bold');
  }

  var funnelLength = Number(payload.funnelLength) || 0;

  for (var c = 0; c < products.length; c++) {
    configSheet.appendRow([
      payload.offerName,
      products[c].name,
      products[c].position || (c + 1),
      Number(products[c].price) || 0,
      'Yes',
      products[c].label || '',
      offerType,
      funnelLength
    ]);
  }

  // Add to Price History using SLUG
  var priceSheet = ss.getSheetByName('Price History');
  if (priceSheet) {
    var today = new Date().toISOString().split('T')[0];
    for (var h = 0; h < products.length; h++) {
      priceSheet.appendRow([
        slug,
        payload.offerName,
        products[h].name,
        Number(products[h].price) || 0,
        today
      ]);
    }
  }

  return { success: true, message: 'Offer created: ' + payload.offerName + ' with ' + products.length + ' products' };
}


/**
 * handleSetupIGOffer_ — Create IG Profile Purchases tab and add marker row to Config
 * Called from handleSetupNewOffer_ when offerType === 'IG Profile Visits'
 */
function handleSetupIGOffer_(ss, payload) {
  var slug = payload.clientSlug;
  var offerName = payload.offerName;

  // Create IG Profile Purchases tab (one per client, shared across all IG campaigns)
  var igTabName = slug + ' - IG Profile Purchases';
  var existingTab = ss.getSheetByName(igTabName);
  if (!existingTab) {
    var igSheet = ss.insertSheet(igTabName);
    igSheet.getRange(1, 1, 1, 3).setValues([['Date', 'Description', 'Revenue']]);
    igSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }

  // Add marker row to Config so the offer appears in the dashboard
  var configTabName = slug + ' - Config';
  var configSheet = ss.getSheetByName(configTabName);
  if (!configSheet) {
    configSheet = createConfigTab_(ss, slug);
  }

  var currentHeaders = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
  var requiredHeaders = ['Offer Name', 'Product Type', 'Position', 'Current Price', 'Active', 'Product Name', 'Offer Type', 'Funnel Length'];
  if (currentHeaders.length < requiredHeaders.length || currentHeaders[6] !== 'Offer Type') {
    configSheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    configSheet.getRange(1, 1, 1, requiredHeaders.length).setFontWeight('bold');
  }

  configSheet.appendRow([offerName, '', 1, 0, 'Yes', '', 'IG Profile Visits', 0]);

  return { success: true, message: 'IG Profile Visits offer created: ' + offerName };
}


/**
 * handleLogIGPurchase_ — Append a misc purchase to the IG Profile Purchases tab
 * Payload: { clientSlug, date, description, revenue }
 */
function handleLogIGPurchase_(payload) {
  var ss = getSheet_();
  var slug = payload.clientSlug;
  if (!slug) return { success: false, error: 'No clientSlug provided' };

  var igTabName = slug + ' - IG Profile Purchases';
  var igSheet = ss.getSheetByName(igTabName);
  if (!igSheet) return { success: false, error: 'IG Profile Purchases tab not found for ' + slug };

  igSheet.appendRow([
    payload.date || new Date().toISOString().split('T')[0],
    payload.description || '',
    Number(payload.revenue) || 0
  ]);

  return { success: true };
}


/**
 * handleDeleteIGPurchase_ — Delete a row from the IG Profile Purchases tab by matching date + description + revenue
 * Payload: { clientSlug, date, description, revenue }
 */
function handleDeleteIGPurchase_(payload) {
  var ss = getSheet_();
  var slug = payload.clientSlug;
  if (!slug) return { success: false, error: 'No clientSlug provided' };

  var igTabName = slug + ' - IG Profile Purchases';
  var igSheet = ss.getSheetByName(igTabName);
  if (!igSheet) return { success: false, error: 'IG Profile Purchases tab not found for ' + slug };

  var data = igSheet.getDataRange().getValues();
  var targetDate = String(payload.date || '').trim();
  var targetDesc = String(payload.description || '').trim();
  var targetRev  = parseFloat(String(payload.revenue || '0').replace(/[$,%]/g, '')) || 0;

  for (var i = 1; i < data.length; i++) {
    var rawDate = data[i][0];
    var rowDate = (rawDate instanceof Date)
      ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(rawDate || '').trim();
    var rowDesc = String(data[i][1] || '').trim();
    var rowRev  = parseFloat(String(data[i][2] || '0').replace(/[$,%]/g, '')) || 0;
    if (rowDate === targetDate && rowDesc === targetDesc && rowRev === targetRev) {
      igSheet.deleteRow(i + 1); // sheet rows are 1-indexed; row 1 is header
      return { success: true };
    }
  }

  return { success: false, error: 'Matching row not found' };
}


/**
 * handleReorderOffers_ — Reorder offer blocks in the Config tab by rearranging rows
 * Payload: { clientSlug, offerOrder: ['Offer A', 'Offer B', ...] }
 */
function handleReorderOffers_(payload) {
  var ss = getSheet_();
  var slug = payload.clientSlug;
  if (!slug) return { success: false, error: 'No clientSlug provided' };

  var offerOrder = payload.offerOrder;
  if (!offerOrder || !offerOrder.length) return { success: false, error: 'No offerOrder provided' };

  var configTabName = slug + ' - Config';
  var configSheet = ss.getSheetByName(configTabName);
  if (!configSheet) return { success: false, error: 'Config tab not found for ' + slug };

  var allData = configSheet.getDataRange().getValues();
  if (allData.length < 2) return { success: true }; // nothing to reorder

  var headers = allData[0];
  var rows = allData.slice(1);

  // Group data rows by offer name (col A)
  var grouped = {};
  rows.forEach(function(row) {
    var name = String(row[0] || '').trim();
    if (!grouped[name]) grouped[name] = [];
    grouped[name].push(row);
  });

  // Rebuild in requested order, then append any offers not in the list
  var newRows = [];
  offerOrder.forEach(function(name) {
    (grouped[name] || []).forEach(function(row) { newRows.push(row); });
    delete grouped[name];
  });
  Object.values(grouped).forEach(function(g) {
    g.forEach(function(row) { newRows.push(row); });
  });

  // Clear existing data rows and rewrite
  var numCols = headers.length;
  if (rows.length > 0) {
    configSheet.getRange(2, 1, rows.length, numCols).clearContent();
  }
  if (newRows.length > 0) {
    configSheet.getRange(2, 1, newRows.length, numCols).setValues(newRows);
  }

  return { success: true };
}


/**
 * save_purchases — Write or update a day's purchase quantities in the Purchases tab
 * Payload: { clientSlug, offerName, date, quantities: { "Product Name": number, ... }, totalUnits, totalRevenue }
 */
function handleSavePurchases_(payload) {
  const ss = getSheet_();
  const slug = payload.clientSlug;
  if (!slug) return { success: false, error: 'No clientSlug provided' };

  var offerName = payload.offerName;
  if (!offerName) return { success: false, error: 'No offerName provided' };

  var tabName = slug + ' - ' + offerName + ' Purchases';
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) return { success: false, error: 'Purchases tab not found: ' + tabName };

  var date = payload.date;
  if (!date) return { success: false, error: 'No date provided' };

  var quantities = payload.quantities || {};

  // Read headers to determine column order
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Build the row based on header order
  var rowData = [];
  rowData.push(date); // Column A = Date
  for (var c = 1; c < headers.length; c++) {
    var header = headers[c];
    if (header === 'Total Units') {
      rowData.push(Number(payload.totalUnits) || 0);
    } else if (header === 'Total Revenue') {
      rowData.push(Number(payload.totalRevenue) || 0);
    } else {
      // Product column — look up in quantities
      rowData.push(Number(quantities[header]) || 0);
    }
  }

  // Check if a row for this date already exists — normalize both sides for M/D/YYYY vs YYYY-MM-DD
  var data = sheet.getDataRange().getValues();
  var existingRow = -1;
  var normalizedPayloadDate = normalizeDateForMatch_(date);
  for (var r = 1; r < data.length; r++) {
    if (normalizeDateForMatch_(data[r][0]) === normalizedPayloadDate) {
      existingRow = r + 1; // 1-indexed sheet row
      break;
    }
  }

  if (existingRow > 0) {
    // UPDATE existing row
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    return { success: true, message: 'Updated purchases for ' + date };
  } else {
    // APPEND new row
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, rowData.length).setValues([rowData]);
    return { success: true, message: 'Added purchases for ' + date };
  }
}


/**
 * save_client — Update a client's name and Meta account IDs in the Clients tab
 * Payload: { clientSlug, clientName, metaAccountIds }
 */
function handleSaveClient_(payload) {
  const ss = getSheet_();
  var slug = payload.clientSlug;
  if (!slug) return { success: false, error: 'No clientSlug provided' };

  var clientsSheet = ss.getSheetByName('Clients');
  if (!clientsSheet) return { success: false, error: 'Clients tab not found' };

  var data = clientsSheet.getDataRange().getValues();
  var headers = data[0];
  var slugCol   = headers.indexOf('Slug');
  var nameCol   = headers.indexOf('Client Name');
  var metaCol   = headers.indexOf('Meta Account ID(s)');

  if (slugCol === -1) return { success: false, error: 'Slug column not found in Clients tab' };
  if (metaCol === -1) return { success: false, error: 'Meta Account ID(s) column not found in Clients tab — check header name matches exactly' };

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][slugCol]).trim() === slug) {
      if (nameCol >= 0 && payload.clientName)
        clientsSheet.getRange(i + 1, nameCol + 1).setValue(payload.clientName);
      clientsSheet.getRange(i + 1, metaCol + 1).setValue(payload.metaAccountIds || '');
      return { success: true };
    }
  }

  return { success: false, error: 'Client not found with slug: ' + slug };
}


/**
 * set_client_status — Archive or reactivate a client via the Status column.
 * Archived clients are skipped by the daily pull and hidden from both dashboards;
 * all their tabs and data stay untouched, so reactivating restores everything.
 * Payload: { clientSlug, status: 'Active' | 'Archived' }
 */
function handleSetClientStatus_(payload) {
  const ss = getSheet_();
  var slug = payload.clientSlug;
  if (!slug) return { success: false, error: 'No clientSlug provided' };

  var status = payload.status;
  if (status !== 'Active' && status !== 'Archived') {
    return { success: false, error: 'Status must be Active or Archived' };
  }

  var clientsSheet = ss.getSheetByName('Clients');
  if (!clientsSheet) return { success: false, error: 'Clients tab not found' };

  var data = clientsSheet.getDataRange().getValues();
  var headers = data[0];
  var slugCol = headers.indexOf('Slug');
  var statusCol = headers.indexOf('Status');
  if (slugCol === -1) return { success: false, error: 'Slug column not found in Clients tab' };
  if (statusCol === -1) return { success: false, error: 'Status column not found in Clients tab' };

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][slugCol]).trim() === slug) {
      clientsSheet.getRange(i + 1, statusCol + 1).setValue(status);
      return { success: true, message: slug + ' set to ' + status };
    }
  }
  return { success: false, error: 'Client not found with slug: ' + slug };
}


/**
 * rename_offer — Rename an offer across Config, Price History, and Purchases tab
 * Payload: { clientSlug, oldOfferName, newOfferName }
 */
function handleRenameOffer_(payload) {
  const ss = getSheet_();
  const slug = payload.clientSlug;
  if (!slug) return { success: false, error: 'No clientSlug provided' };

  var oldName = payload.oldOfferName;
  var newName = payload.newOfferName;
  if (!oldName || !newName) return { success: false, error: 'Both oldOfferName and newOfferName are required' };
  if (oldName === newName) return { success: true, message: 'Names are the same, nothing to rename' };

  // 1. Rename in Config tab
  var configTabName = slug + ' - Config';
  var configSheet = ss.getSheetByName(configTabName);
  if (configSheet) {
    var configData = configSheet.getDataRange().getValues();
    for (var i = 1; i < configData.length; i++) {
      if (configData[i][0] === oldName) {
        configSheet.getRange(i + 1, 1).setValue(newName);
      }
    }
  }

  // 2. Rename in Price History tab
  var priceSheet = ss.getSheetByName('Price History');
  if (priceSheet) {
    var priceData = priceSheet.getDataRange().getValues();
    for (var j = 1; j < priceData.length; j++) {
      if (String(priceData[j][0]) === slug && priceData[j][1] === oldName) {
        priceSheet.getRange(j + 1, 2).setValue(newName);
      }
    }
  }

  // 3. Rename the Purchases tab itself
  var oldTabName = slug + ' - ' + oldName + ' Purchases';
  var newTabName = slug + ' - ' + newName + ' Purchases';
  var purchasesTab = ss.getSheetByName(oldTabName);
  if (purchasesTab) {
    purchasesTab.setName(newTabName);
  }

  return { success: true, message: 'Offer renamed from "' + oldName + '" to "' + newName + '"' };
}


/**
 * rename_product — Rename a product across Config, Price History, and Purchases tab header
 * Payload: { clientSlug, offerName, oldProductName, newProductName }
 */
function handleRenameProduct_(payload) {
  const ss = getSheet_();
  const slug = payload.clientSlug;
  if (!slug) return { success: false, error: 'No clientSlug provided' };

  var offerName = payload.offerName;
  var oldName = payload.oldProductName;
  var newName = payload.newProductName;
  if (!offerName || !oldName || !newName) return { success: false, error: 'offerName, oldProductName, and newProductName are all required' };
  if (oldName === newName) return { success: true, message: 'Names are the same, nothing to rename' };

  // 1. Update Config tab
  var configTabName = slug + ' - Config';
  var configSheet = ss.getSheetByName(configTabName);
  if (configSheet) {
    var configData = configSheet.getDataRange().getValues();
    for (var i = 1; i < configData.length; i++) {
      if (configData[i][0] === offerName && configData[i][1] === oldName) {
        configSheet.getRange(i + 1, 2).setValue(newName);
      }
    }
  }

  // 2. Update Price History tab
  var priceSheet = ss.getSheetByName('Price History');
  if (priceSheet) {
    var priceData = priceSheet.getDataRange().getValues();
    for (var j = 1; j < priceData.length; j++) {
      if (String(priceData[j][0]) === slug && priceData[j][1] === offerName && priceData[j][2] === oldName) {
        priceSheet.getRange(j + 1, 3).setValue(newName);
      }
    }
  }

  // 3. Update Purchases tab header
  var purchasesTabName = slug + ' - ' + offerName + ' Purchases';
  var purchasesSheet = ss.getSheetByName(purchasesTabName);
  if (purchasesSheet) {
    var headers = purchasesSheet.getRange(1, 1, 1, purchasesSheet.getLastColumn()).getValues()[0];
    for (var h = 0; h < headers.length; h++) {
      if (headers[h] === oldName) {
        purchasesSheet.getRange(1, h + 1).setValue(newName);
        break;
      }
    }
  }

  return { success: true, message: 'Product renamed from "' + oldName + '" to "' + newName + '"' };
}


/**
 * save_product_name — Write ONLY the Product Name (col 6) for one product.
 * Safe: touches no other columns, no risk of duplicating rows.
 * Payload: { clientSlug, offerName, productType, productName }
 *   productType — the role identifier in col B (e.g. "Bump", "Upsell 1")
 *   productName — the descriptive name to store in col F (e.g. "Style Guide")
 */
function handleSaveProductName_(payload) {
  var ss = getSheet_();
  var slug = payload.clientSlug;
  if (!slug) return { success: false, error: 'No clientSlug provided' };

  var configSheet = ss.getSheetByName(slug + ' - Config');
  if (!configSheet) return { success: false, error: 'Config tab not found' };

  var data = configSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === payload.offerName &&
        String(data[i][1]).trim() === payload.productType) {
      configSheet.getRange(i + 1, 6).setValue(payload.productName || '');
      return { success: true };
    }
  }

  return { success: false, error: 'Product not found: "' + payload.productType + '" in offer "' + payload.offerName + '"' };
}


/**
 * Normalize a date value from a sheet cell to YYYY-MM-DD for comparison.
 * Handles Date objects, M/D/YYYY strings, and YYYY-MM-DD strings.
 */
function normalizeDateForMatch_(d) {
  if (!d && d !== 0) return '';
  if (d instanceof Date) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(d).trim();
  // M/D/YYYY or MM/DD/YYYY
  var slashParts = s.split('/');
  if (slashParts.length === 3) {
    var mo = slashParts[0].padStart(2, '0');
    var dy = slashParts[1].padStart(2, '0');
    var yr = slashParts[2].length === 4 ? slashParts[2] : ('20' + slashParts[2]);
    return yr + '-' + mo + '-' + dy;
  }
  // YYYY-MM-DD or fallback — take first 10 chars
  return s.substring(0, 10);
}


/**
 * delete_purchase — Delete a single date row from a Purchases tab
 * Payload: { clientSlug, offerName, date }  (date in YYYY-MM-DD)
 */
function handleDeletePurchase_(payload) {
  var ss = getSheet_();
  var slug = payload.clientSlug;
  if (!slug) return { success: false, error: 'No clientSlug provided' };

  var offerName = payload.offerName;
  var date = payload.date;
  if (!offerName || !date) return { success: false, error: 'offerName and date required' };

  var tabName = slug + ' - ' + offerName + ' Purchases';
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) return { success: false, error: 'Purchases tab not found: ' + tabName };

  var normalizedTarget = normalizeDateForMatch_(date);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeDateForMatch_(data[i][0]) === normalizedTarget) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }

  return { success: false, error: 'Row not found for date: ' + date };
}


// ============================================
// TAB CREATION HELPERS
// ============================================

/**
 * Create the Daily Meta tab for a client using their SLUG.
 */
function createDailyMetaTab_(ss, slug) {
  var tabName = slug + ' - Daily Meta';
  var existing = ss.getSheetByName(tabName);
  if (existing) return existing;

  var sheet = ss.insertSheet(tabName);
  var headers = [
    'Date', 'Campaign Name', 'Amount Spent', 'Impressions', 'Reach',
    'Clicks', 'Leads', 'Purchases', 'Purchase Value', 'CPM', 'CPC', 'CTR', 'ROAS'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  return sheet;
}

/**
 * Create the Config tab for a client using their SLUG.
 */
function createConfigTab_(ss, slug) {
  var tabName = slug + ' - Config';
  var existing = ss.getSheetByName(tabName);
  if (existing) return existing;

  var sheet = ss.insertSheet(tabName);
  var headers = ['Offer Name', 'Product Type', 'Position', 'Current Price', 'Active', 'Product Name', 'Offer Type', 'Funnel Length'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  return sheet;
}


// ============================================
// UTILITY FUNCTIONS
// ============================================

/**
 * Look up a client's slug from their full name.
 */
function getSlugByClientName_(ss, clientName) {
  var clientsSheet = ss.getSheetByName('Clients');
  if (!clientsSheet) return null;

  var data = clientsSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === clientName) {
      return data[i][1];
    }
  }
  return null;
}

/**
 * Look up a client's full name from their slug.
 */
function getClientNameBySlug_(ss, slug) {
  var clientsSheet = ss.getSheetByName('Clients');
  if (!clientsSheet) return null;

  var data = clientsSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === slug) {
      return data[i][0];
    }
  }
  return null;
}


// ============================================
// INITIAL SETUP — Run once to create base tabs
// ============================================

/**
 * Run this ONCE to create the base sheet structure.
 * Creates: Clients, Price History, CHECKBOX_STATE tabs with headers.
 * Price History uses "Slug" as the first column (not "Client Name").
 */
function initialSetup() {
  const ss = getSheet_();

  // --- Clients tab ---
  var clientsSheet = ss.getSheetByName('Clients');
  if (!clientsSheet) {
    clientsSheet = ss.insertSheet('Clients');
    clientsSheet.getRange(1, 1, 1, 5).setValues([['Client Name', 'Slug', 'Meta Account ID(s)', 'Status', 'Notes']]);
    clientsSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }

  // --- Price History tab (first column is Slug, not Client Name) ---
  var priceSheet = ss.getSheetByName('Price History');
  if (!priceSheet) {
    priceSheet = ss.insertSheet('Price History');
    priceSheet.getRange(1, 1, 1, 5).setValues([['Slug', 'Offer Name', 'Product Name', 'Price', 'Effective Date']]);
    priceSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }

  // --- CHECKBOX_STATE tab ---
  var checkboxSheet = ss.getSheetByName('CHECKBOX_STATE');
  if (!checkboxSheet) {
    checkboxSheet = ss.insertSheet('CHECKBOX_STATE');
    checkboxSheet.getRange(1, 1, 1, 4).setValues([['Date', 'Client Slug', 'Campaign Name', 'Action']]);
    checkboxSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }

  Logger.log('Base tabs created. Now run seedClients() to add clients.');
}


/**
 * Seed the known KSA clients. Run ONCE after initialSetup.
 * Creates per-client tabs using SLUGS (e.g. "htp - Daily Meta", not "Hope Taylor Photography - Daily Meta").
 * Meta Account IDs are blank — Rebecca needs to fill these in.
 */
function seedClients() {
  const ss = getSheet_();
  var clientsSheet = ss.getSheetByName('Clients');
  if (!clientsSheet) {
    Logger.log('Run initialSetup() first');
    return;
  }

  var clients = [
    ['KS Agency', 'ksa', '', 'Active', 'Two Meta ad accounts — IDs needed'],
    ['Hope Taylor Photography', 'htp', '', 'Active', 'One Meta account — ID needed'],
    ['KJP', 'kjp', '', 'Active', 'One Meta account — ID needed'],
    ['Lola', 'lola', '', 'Active', 'One Meta account — ID needed']
  ];

  // Check if clients already exist
  var existing = clientsSheet.getDataRange().getValues();
  if (existing.length > 1) {
    Logger.log('Clients tab already has data. Skipping seed.');
    return;
  }

  clientsSheet.getRange(2, 1, clients.length, clients[0].length).setValues(clients);

  // Create per-client tabs using SLUGS
  for (var i = 0; i < clients.length; i++) {
    var slug = clients[i][1]; // Slug is column index 1
    createDailyMetaTab_(ss, slug);
    createConfigTab_(ss, slug);
  }

  Logger.log('Seeded ' + clients.length + ' clients with Daily Meta and Config tabs (using slugs)');
}


/**
 * Seed HTP offers (confirmed from spreadsheet). Run ONCE after seedClients.
 */
function seedHTPOffers() {
  // What to Wear Tiny Offer
  handleSetupNewOffer_({
    clientSlug: 'htp',
    offerName: 'What to Wear Tiny Offer',
    products: [
      { name: 'Tripwire', position: 1, price: 17 },
      { name: 'Bump', position: 2, price: 11 },
      { name: 'Upsell 1', position: 3, price: 47 },
      { name: 'Upsell 2', position: 4, price: 97 },
      { name: 'Upsell 3', position: 5, price: 37 },
      { name: 'Upsell 4', position: 6, price: 5 }
    ]
  });

  // New Presets — no product details yet, just the offer
  handleSetupNewOffer_({
    clientSlug: 'htp',
    offerName: 'New Presets',
    products: [
      { name: 'New Presets', position: 1, price: 0 }
    ]
  });

  Logger.log('HTP offers seeded');
}


/**
 * handlePushClientSnapshot_ — Record the most recent data date as "published through"
 * for a client so the external dashboard caps its view at that date.
 * Payload: { clientSlug }
 */
function handlePushClientSnapshot_(payload) {
  var ss = getSheet_();
  var slug = payload.clientSlug;
  if (!slug) return { success: false, error: 'No clientSlug provided' };

  // Find most recent date in this client's Daily Meta tab
  var metaSheet = ss.getSheetByName(slug + ' - Daily Meta');
  if (!metaSheet) return { success: false, error: 'Daily Meta tab not found for ' + slug };

  var metaData = metaSheet.getDataRange().getValues();
  var mostRecentDate = '';
  for (var i = 1; i < metaData.length; i++) {
    var cell = metaData[i][0];
    var dateStr = '';
    if (cell instanceof Date) {
      dateStr = Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (cell) {
      dateStr = String(cell).trim();
    }
    if (dateStr && dateStr > mostRecentDate) mostRecentDate = dateStr;
  }

  if (!mostRecentDate) return { success: false, error: 'No data found in Daily Meta for ' + slug };

  var publishedAt = new Date().toISOString();

  // Upsert into PUBLISH_STATE tab
  var publishSheet = ss.getSheetByName('PUBLISH_STATE');
  if (!publishSheet) {
    publishSheet = ss.insertSheet('PUBLISH_STATE');
    publishSheet.getRange(1, 1, 1, 3).setValues([['Client Slug', 'Published Through', 'Published At']]);
    publishSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }

  var publishData = publishSheet.getDataRange().getValues();
  var found = false;
  for (var j = 1; j < publishData.length; j++) {
    if (publishData[j][0] === slug) {
      publishSheet.getRange(j + 1, 2, 1, 2).setValues([[mostRecentDate, publishedAt]]);
      found = true;
      break;
    }
  }
  if (!found) {
    publishSheet.appendRow([slug, mostRecentDate, publishedAt]);
  }

  return { success: true, publishedThrough: mostRecentDate, publishedAt: publishedAt };
}
