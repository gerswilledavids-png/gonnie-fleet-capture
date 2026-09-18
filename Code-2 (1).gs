/**
 * GONNIE FLEET MASTER — Trip Capture Handler (v2)
 * Connects the FleetCapture web app to the 'Trip Log' tab.
 *
 * ── WHY THIS VERSION IS DIFFERENT ────────────────────────────────────────
 * The old script hard-coded ~39 column formulas as JS strings. The sheet has
 * since grown to 44 live input/formula columns (A–AR) plus fuel-return and
 * route-check formulas out to BD, and two UNRELATED reference tables
 * (Depot Register BF:BI, Route Distance Database BK:BN) that happen to sit
 * in the same row range but must NEVER be touched by a trip write.
 *
 * So instead of re-typing every formula, this script:
 *   1. Copies the *formulas* from a known-good template row (row 3) down
 *      into the new row, for formula columns only.
 *   2. Writes the raw values the capture app collected into the input
 *      columns only.
 *   3. Leaves the Depot Register / Route Distance Database columns alone.
 *
 * If you ever add a column to Trip Log, add its number to FORMULA_COLS
 * (if it's a formula) or the INPUT map (if the app should fill it) — you
 * do not need to touch the calculation logic itself.
 *
 * ── HOW TO MERGE INTO YOUR EXISTING PROJECT ──────────────────────────────
 * 1. Open the spreadsheet → Extensions → Apps Script.
 * 2. If you already have a Code.gs, REPLACE its doPost/doGet functions with
 *    the ones below (keep any other functions you've added, e.g. triggers).
 * 3. Confirm the sheet tab is literally named "Trip Log" (with the space).
 *    If yours is "Trip_Log" (underscore), change SHEET_NAME below.
 * 4. Row 3 must keep its formulas intact — it is the template every new
 *    trip is copied from. Never delete row 3.
 * 5. Deploy → Manage deployments → New deployment (or edit existing) →
 *    Web app → Execute as: Me, Who has access: Anyone. Copy the /exec URL
 *    into your capture app's config.
 * 6. IMPORTANT (CORS): send the POST body with
 *    headers: { "Content-Type": "text/plain;charset=utf-8" }
 *    and JSON.stringify(payload) as the body. Apps Script web apps don't
 *    handle the CORS preflight that "application/json" triggers, so
 *    text/plain avoids the preflight while the body is still parsed as
 *    JSON server-side (see doPost below).
 */

var SHEET_NAME = "Trip Log";
var TEMPLATE_ROW = 3; // first data row — source of every formula copy-down

// Columns that hold FORMULAS and must be copied down from TEMPLATE_ROW.
// (Deliberately excludes AS/BE spacer cols and the Depot Register /
// Route Distance Database reference tables at BF:BI and BK:BN.)
var FORMULA_COLS = [
  1,  // A  Trip ID
  3,  // C  Fleet No (lookup)
  12, // L  Distance KM
  13, // M  Fuel Cost (R)
  14, // N  KM/L Actual
  15, // O  Expected KM/L (lookup)
  16, // P  Fuel Variance %
  17, // Q  Theft Alert
  21, // U  WhatsApp Report link
  22, // V  Cost/KM
  25, // Y  END DESTINATION (lookup) — overwritten below if app supplies one
  30, // AD Driver Score
  31, // AE Maint Flag (lookup)
  32, // AF GPS Distance KM
  33, // AG Maint Alloc (R) (lookup)
  34, // AH Insurance Alloc (R) (lookup)
  35, // AI License Alloc (R) (lookup)
  36, // AJ Total Fixed Cost (R)
  37, // AK Total Variable Cost (R)
  38, // AL Net Profit (R)
  39, // AM Profit Margin %
  42, // AP Cost Data Check (lookup)
  43, // AQ Insurer (lookup)
  47, // AU Return Fuel Variance %
  48, // AV Return Fuel Check
  50, // AX Fuel Used - Best Estimate (L)
  51, // AY Est. Basis
  52, // AZ Route Key (auto)
  53, // BA Expected Route KM
  54, // BB Route KM Variance
  55, // BC Route Deviation %
  56, // BD Route Check
  68  // BP Odo/Route Check (auto)
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No POST body received.");
    }
    var data = JSON.parse(e.postData.contents);

    var lastRow = sheet.getLastRow();
    var targetRow = lastRow + 1;
    if (targetRow < TEMPLATE_ROW) targetRow = TEMPLATE_ROW;
    var tripId = targetRow - 2; // matches sheet's own "=ROW()-2" Trip ID formula

    // 1. Copy formulas down from the template row (no-op on the template
    //    row itself, e.g. when this is the very first trip logged).
    FORMULA_COLS.forEach(function (col) {
      sheet.getRange(TEMPLATE_ROW, col).copyTo(
        sheet.getRange(targetRow, col),
        SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
        false
      );
    });

    // 2. Write the raw values the capture app collected.
    var INPUT = {
      2:  data.date ? new Date(data.date) : new Date(),                  // B  Date
      4:  data.registration || "",                                       // D  Registration
      5:  data.driver_name || "",                                        // E  Driver Name
      6:  data.driver_phone || "",                                       // F  Driver Phone
      7:  numOrBlank(data.odo_start),                                    // G  Odo Start
      8:  numOrBlank(data.odo_end),                                      // H  Odo End
      9:  numOrBlank(data.fuel_used_l),                                  // I  Fuel Used (L)
      10: numOrBlank(data.cost_per_liter),                               // J  Cost/Liter (R)
      11: data.no_trips !== undefined ? Number(data.no_trips) : 1,       // K  No.Trips
      18: data.notes || "",                                              // R  Trip Notes
      19: data.manager || "",                                            // S  Manager
      20: data.manager_phone || "",                                      // T  Manager Phone
      23: data.start_destination || "",                                  // W  START DESTINATION
      24: numOrBlank(data.revenue),                                      // X  Revenue (R)
      26: data.cargo_type || "",                                         // Z  Cargo Type
      27: numOrBlank(data.load_kg),                                      // AA Load KG
      28: data.fuel_station || "",                                       // AB Fuel Station
      29: data.gps_verified || "PENDING",                                // AC GPS Verified
      40: data.paid_status || "Pending",                                 // AN Paid Status
      41: data.shopify_customer_id || "",                                // AO Shopify customer ID
      44: data.customer || "",                                           // AR Customer
      46: numOrBlank(data.fuel_filled_return_l),                         // AT Fuel Filled on Return (L)
      49: data.fill_type || "Manual"                                     // AW Fill Type: Metered | Full-to-Full | Manual
    };

    // Optional: only overwrite the END DESTINATION lookup formula if the
    // capture app actually supplies one (e.g. from the Scan/OCR flow).
    if (data.end_destination) {
      INPUT[25] = data.end_destination; // Y
    }

    Object.keys(INPUT).forEach(function (col) {
      var v = INPUT[col];
      if (v !== "") sheet.getRange(targetRow, Number(col)).setValue(v);
    });

    SpreadsheetApp.flush();

    return jsonOut({
      status: "success",
      trip_id: tripId,
      row: targetRow,
      message: "Trip logged to Gonnie Fleet Master"
    });

  } catch (err) {
    return jsonOut({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    var action = (e && e.parameter) ? e.parameter.action : "";
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "recent") {
      return jsonOut({ status: "success", trips: getRecentTrips(ss) });
    }

    if (action === "vehicles") {
      return jsonOut({ status: "success", vehicles: getVehicles(ss) });
    }

    if (action === "drivers") {
      return jsonOut({ status: "success", drivers: getDrivers(ss) });
    }

    return jsonOut({
      status: "error",
      message: "Unknown action. Use ?action=recent | vehicles | drivers"
    });

  } catch (err) {
    return jsonOut({ status: "error", message: err.toString() });
  }
}

// ── Helpers ────────────────────────────────────────────────────────────

function numOrBlank(v) {
  if (v === undefined || v === null || v === "") return "";
  var n = Number(v);
  return isNaN(n) ? "" : n;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getRecentTrips(ss) {
  var sheet = ss.getSheetByName(SHEET_NAME);
  var rows = sheet.getDataRange().getValues();
  var trips = [];
  var startIdx = Math.max(TEMPLATE_ROW - 1, rows.length - 50); // last 50
  for (var i = startIdx; i < rows.length; i++) {
    var r = rows[i];
    if (r[1] === "") continue; // skip blank Date rows
    trips.push({
      trip_id: r[0],            // A
      date: r[1],                // B
      registration: r[3],        // D
      driver_name: r[4],         // E
      odo_start: r[6],           // G
      odo_end: r[7],             // H
      distance_km: r[11],        // L
      km_l_actual: r[13],        // N
      theft_alert: r[16],        // Q
      revenue: r[23],            // X
      net_profit: r[37],         // AL
      paid_status: r[39]         // AN
    });
  }
  return trips;
}

function getVehicles(ss) {
  var sheet = ss.getSheetByName("Vehicle Register");
  var rows = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 2; i < rows.length; i++) { // header is row 2 (index 1)
    var r = rows[i];
    if (!r[0]) continue; // Registration blank
    out.push({
      registration: r[0],
      make: r[2],
      model: r[3],
      year: r[4],
      status: r[7],
      expected_kml: r[8],
      current_odometer: r[14]
    });
  }
  return out;
}

function getDrivers(ss) {
  var sheet = ss.getSheetByName("Driver Register");
  var rows = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 2; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue; // Driver Name blank
    out.push({
      driver_name: r[0],
      phone: r[1],
      license_expiry: r[3],
      avg_score: r[9]
    });
  }
  return out;
}
