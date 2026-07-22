/**
 * GadgetVilla Dashboard - runtime configuration.
 * This file is git-ignored (see .gitignore). Do NOT commit real URLs/tokens.
 *
 * Linked Google Sheet:
 *   https://docs.google.com/spreadsheets/d/15u6K6fl4VukQc8SdKUvnlNcq5PhrLqPzp78L3hc3Cgw/edit
 *
 * TO GO LIVE: paste the Web App /exec URL into webAppUrl below. That is the ONLY
 * thing you need to change here (useLive is already true).
 */
window.GV_CONFIG = {
  // <<< PASTE the Apps Script Web App URL here (ends with /exec) >>>
  webAppUrl: "https://script.google.com/macros/s/AKfycbytb3MR9ITs-z-4_CmoKwCXI75soOPcrVqwxmNyQUu-GzaUHuA9mBY_PurVFVJ691TiWA/exec",

  // Optional shared token. Leave "" if you did not set API_TOKEN in Apps Script.
  token: "",

  // Feature toggle: true = live Google Sheets data, false = built-in fallback.
  useLive: true,

  // Real-time auto-refresh interval in milliseconds. 0 = off (load once).
  // 15000 = refresh every 15 seconds.
  refreshMs: 15000
};
