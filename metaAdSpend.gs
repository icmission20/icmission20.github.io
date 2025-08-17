function getMetaAdSpend() {
  // \ud83d\udd11 REQUIRED: replace both
  var TOKEN = "EAAPrE9po26sBPNWx3nEK0qZCb8ZBe2Wme9QzfYO6sbDn47pF1JkHb1l1zm7ZAJLb8sBn2nXNrs85OxrRZCuJEwxc4ZBYQZBZCwUdoSWMaEo4faJdiRFUGK7bUbnQnirZBT6srtmTn3aOB2fZAxNTWlfCc7lNw8ZBT7lPQIzsC2e7vbPVrM4qLx1hYk1IZCUJDgJZA1oHbFch4EtcKZCktBSe7SOrZCnntfbYpZCMOF4Y0tG8aIKW4YZD";
  var AD_ACCOUNT = "act_2600900606832558";

  var base = "https://graph.facebook.com/v19.0/";
  var datePreset = "this_month";

  // -------- 1) CAMPAIGNS (budgets) --------
  var campaignFields = "id,name,daily_budget,lifetime_budget,status,effective_status,objective";
  var campaignsUrl = base + AD_ACCOUNT + "/campaigns?fields=" + campaignFields +
    "&access_token=" + encodeURIComponent(TOKEN);
  var campaigns = fetchAll_(campaignsUrl, "Campaigns");
  if (campaigns.length === 0) throw new Error("No campaigns returned. Token permissions or account access likely wrong.");

  // Build quick lookup for budgets
  var byCampaign = {};
  campaigns.forEach(function (c) {
    byCampaign[c.id] = {
      campaign_id: c.id,
      campaign_name: c.name || "",
      daily_budget: convertMinorToMajor_(c.daily_budget),
      lifetime_budget: convertMinorToMajor_(c.lifetime_budget),
      status: c.status || "",
      effective_status: c.effective_status || "",
      objective: c.objective || "",
      spend: 0
    };
  });

  // -------- 2) INSIGHTS (spend) --------
  var insightsUrl = base + AD_ACCOUNT +
    "/insights?level=campaign&fields=campaign_id,campaign_name,spend" +
    "&date_preset=" + encodeURIComponent(datePreset) +
    "&access_token=" + encodeURIComponent(TOKEN);
  var insights = fetchAll_(insightsUrl, "Insights");

  // Join spend into campaign map
  insights.forEach(function (row) {
    var id = row.campaign_id;
    if (!id || !byCampaign[id]) return;
    byCampaign[id].campaign_name = byCampaign[id].campaign_name || row.campaign_name || "";
    byCampaign[id].spend += parseFloat(row.spend || "0");
  });

  // -------- 3) Write to Sheets --------
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Summary sheet (home screen)
  var summarySheet = ss.getSheetByName("Home");
  if (!summarySheet) summarySheet = ss.insertSheet("Home");
  summarySheet.clear();
  var summaryHeaders = ["Campaign ID", "Campaign Name", "Spend (this month)"];
  summarySheet.appendRow(summaryHeaders);
  var summaryRows = Object.keys(byCampaign).sort().map(function (id) {
    var r = byCampaign[id];
    return [r.campaign_id, r.campaign_name, safeMoney_(r.spend)];
  });
  if (summaryRows.length) summarySheet.getRange(2, 1, summaryRows.length, summaryHeaders.length).setValues(summaryRows);
  if (summaryRows.length) summarySheet.getRange(2, 3, summaryRows.length, 1).setNumberFormat("$#,##0.00");

  // Details sheet
  var detailSheet = ss.getSheetByName("CampaignDetails");
  if (!detailSheet) detailSheet = ss.insertSheet("CampaignDetails");
  detailSheet.clear();
  var detailHeaders = [
    "Campaign ID","Campaign Name","Spend (this month)","Daily Budget","Lifetime Budget",
    "Status","Effective Status","Objective"
  ];
  detailSheet.appendRow(detailHeaders);
  var detailRows = Object.keys(byCampaign).sort().map(function (id) {
    var r = byCampaign[id];
    return [
      r.campaign_id,
      r.campaign_name,
      safeMoney_(r.spend),
      safeMoney_(r.daily_budget),
      safeMoney_(r.lifetime_budget),
      r.status,
      r.effective_status,
      r.objective
    ];
  });
  if (detailRows.length) detailSheet.getRange(2, 1, detailRows.length, detailHeaders.length).setValues(detailRows);
  if (detailRows.length) detailSheet.getRange(2, 3, detailRows.length, 3).setNumberFormat("$#,##0.00");

  // Diagnostics
  var totalSpend = 0, totalDaily = 0, totalLifetime = 0;
  Object.keys(byCampaign).forEach(function (id) {
    var r = byCampaign[id];
    totalSpend += r.spend || 0;
    totalDaily += r.daily_budget || 0;
    totalLifetime += r.lifetime_budget || 0;
  });
  summarySheet.getRange(1, 5).setValue("Totals");
  summarySheet.getRange(2, 5).setValue("Campaigns fetched: " + campaigns.length);
  summarySheet.getRange(3, 5).setValue("Total spend: " + safeMoney_(totalSpend));
  summarySheet.getRange(4, 5).setValue("Total daily budget: " + safeMoney_(totalDaily));
  summarySheet.getRange(5, 5).setValue("Total lifetime budget: " + safeMoney_(totalLifetime));
  summarySheet.getRange(3, 5, 3, 1).setNumberFormat("$#,##0.00");
}

// ---- Helpers ----
function assertOk_(label, resp) {
  var code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(label + " call failed (" + code + "): " + resp.getContentText());
  }
}

function convertMinorToMajor_(v) {
  if (v === null || v === undefined || v === "") return null;
  var n = Number(v);
  if (isNaN(n)) return null;
  return n / 100.0; // Meta budgets are minor units
}

function safeMoney_(n) {
  return (n === null || n === undefined || isNaN(n)) ? "" : Number(n);
}

function fetchAll_(url, label) {
  var all = [];
  while (url) {
    var resp = UrlFetchApp.fetch(url + "&limit=200", { muteHttpExceptions: true });
    assertOk_(label, resp);
    var json = JSON.parse(resp.getContentText());
    all = all.concat(json.data || []);
    url = (json.paging && json.paging.next) ? json.paging.next : null;
  }
  return all;
}
