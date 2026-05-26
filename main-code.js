/***** CONFIG *****/
const MASTER_SHEET = "Master";
const CATEGORY_COL_NAME = "Category Tags";
const SUBCATEGORY_COL_NAME = "Subcategory Tags";
const RECORD_ID_COL_NAME = "RecordID";

const CATEGORY_SHEET_HEADERS = [
  "RecordID",
  "Organization Name",
  "Category Tags",
  "Subcategory Tags",
  "Phone",
  "Address",
  "Services Offered",
  "Website",
  "Notes",
  "Latitude",
  "Longitude"
];

/***** MENU *****/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Resource Search')
    .addItem('Open Search Sidebar', 'showSearchSidebar')
    .addToUi();

  ui.createMenu('Resource Directory')
    .addItem('Create Packet', 'showPacketSidebar')
    .addToUi();
  //showSearchSidebar();
}

function showSearchSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SearchSidebar')
    .setTitle('Search Resources');
  SpreadsheetApp.getUi().showSidebar(html);
}

/***** HELPERS *****/
function getHeaderMap_(sheet) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  header.forEach((name, idx) => (map[name] = idx));
  return { header, map };
}

function ensureRecordID_(sheet, rowNum, rowValues) {
  const { map } = getHeaderMap_(sheet);
  const idCol = map[RECORD_ID_COL_NAME] + 1;
  const existing = rowValues[idCol - 1];
  if (existing) return existing;

  const newID = "ID-" + new Date().getTime() + "-" + Math.floor(Math.random() * 99999);
  sheet.getRange(rowNum, idCol).setValue(newID);
  return newID;
}

/***** MAIN SYNC HANDLER *****/
function onEdit(e) {
  if (!e) return;

  const sheet = e.range.getSheet();
  const ss = sheet.getParent();

  if (sheet.getName() !== MASTER_SHEET) return;

  const { map } = getHeaderMap_(sheet);
  const categoryCol = map[CATEGORY_COL_NAME] + 1;
  const addrCol    = map["Address"] + 1;
  const latCol     = map["Latitude"] + 1;
  const lngCol     = map["Longitude"] + 1;

  const rowNum = e.range.getRow();
  if (rowNum === 1) return;

  let row = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];

  // --- Ensure RecordID ---
  const recordID = ensureRecordID_(sheet, rowNum, row);
  row[map[RECORD_ID_COL_NAME]] = recordID;

  // --- Geocode if needed ---
  let lat = row[latCol - 1];
  let lng = row[lngCol - 1];
  const addr = row[addrCol - 1];

  if ((!lat || !lng) && addr) {
    const coords = geocodeAddress_(addr);
    if (coords) {
      sheet.getRange(rowNum, latCol).setValue(coords.lat);
      sheet.getRange(rowNum, lngCol).setValue(coords.lng);
      lat = coords.lat;
      lng = coords.lng;
      row[latCol - 1] = lat;
      row[lngCol - 1] = lng;
    }
  }

  // --- Determine categories ---
  const categories = (row[categoryCol - 1] || "")
    .toString()
    .split(",")
    .map(c => c.trim())
    .filter(Boolean);

  // --- Sync to category sheets ---
  syncToCategorySheets_(ss, row, categories, recordID);
}

/***** CORE SYNC ENGINE *****/
function syncToCategorySheets_(ss, row, categories, recordID) {
  // First: remove this row from ALL category sheets it no longer belongs to
  const allSheets = ss.getSheets();
  allSheets.forEach(sh => {
    const name = sh.getName();
    if (name === MASTER_SHEET) return;
    if (!categories.includes(name)) {
      const data = sh.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === recordID) {
          sh.deleteRow(i + 1);
          break;
        }
      }
    }
  });

  // Next: add/update in sheets it SHOULD belong to
  categories.forEach(cat => {
    let catSheet = ss.getSheetByName(cat);
    if (!catSheet) {
      catSheet = ss.insertSheet(cat);
      catSheet.getRange(1, 1, 1, CATEGORY_SHEET_HEADERS.length)
        .setValues([CATEGORY_SHEET_HEADERS]);
    }

    const data = catSheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === recordID) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow > 0) {
      // update
      catSheet.getRange(foundRow, 1, 1, row.length).setValues([row]);
    } else {
      // append new
      catSheet.appendRow(row);
    }
  });
}

/***** CATEGORY + SUBCATEGORY FOR SIDEBAR *****/
function getCategoryToSubcategoryMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(MASTER_SHEET);

  const { map } = getHeaderMap_(master);
  const data = master.getDataRange().getValues();

  const mapping = {}; // { category: Set(subcategories...) }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const cats = (row[map[CATEGORY_COL_NAME]] || "")
      .toString().split(",").map(s => s.trim()).filter(Boolean);
    const subs = (row[map[SUBCATEGORY_COL_NAME]] || "")
      .toString().split(",").map(s => s.trim()).filter(Boolean);

    cats.forEach(cat => {
      if (!mapping[cat]) mapping[cat] = new Set();
      subs.forEach(sub => mapping[cat].add(sub));
    });
  }

  // Convert sets → arrays for HTML
  Object.keys(mapping).forEach(c => mapping[c] = Array.from(mapping[c]));
  return mapping;
}

/***** SEARCH FUNCTIONS (unchanged except ID support) *****/
function runSearchFromSidebar(form) {
  const radius = parseFloat(form.radius);
  const cats = form.categories || [];
  const subs = form.subcategories || [];
  searchResources(form.address, radius, cats, subs);
  return "Search complete! Check the 'Search Results' sheet.";
}

function searchResources(clientAddress, radiusMiles, categories, subcategories) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(MASTER_SHEET);
  const results = ss.getSheetByName("Search Results") || ss.insertSheet("Search Results");

  const { header, map } = getHeaderMap_(master);
  const data = master.getDataRange().getValues();

  const outHeader = header.concat(["Distance (mi)", "Google Maps Link", "DEBUG"]);
  results.clear();
  results.getRange(1, 1, 1, outHeader.length).setValues([outHeader]);

  const client = geocodeAddress_(clientAddress);
  if (!client) {
    results.appendRow(["", "", "", "", "", "", "", "", "", "", "", "", "CLIENT ADDRESS NOT FOUND"]);
    return;
  }

  const addrIdx = map["Address"];
  const latIdx = map["Latitude"];
  const lngIdx = map["Longitude"];

  let written = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i].slice();
    const debug = [];

    const rowCats = (row[map[CATEGORY_COL_NAME]] || "")
      .toString()
      .toLowerCase()
      .split(",")
      .map(s => s.trim())
      .filter(Boolean);
    const rowSubs = (row[map[SUBCATEGORY_COL_NAME]] || "")
      .toString()
      .toLowerCase()
      .split(",")
      .map(s => s.trim())
      .filter(Boolean);

    if (categories.length && !categories.every(c => rowCats.includes(c.toLowerCase()))) {
      debug.push("Category mismatch");
    }
    if (!debug.length && subcategories.length) {
  // Normalize and test: a selected subcategory matches if either
  // - resource explicitly lists that subcategory, OR
  // - resource has NO subcategories listed (treated as inclusive)
  const rowHasAnySub = rowSubs.length > 0;
  const allSelectedMatch = subcategories.every(sel => {
    const selLower = sel.toLowerCase();
    // if resource has no subcategories, treat as match (inclusive)
    if (!rowHasAnySub) return true;
    return rowSubs.includes(selLower);
  });

  if (!allSelectedMatch) {
    debug.push("Subcategory mismatch");
  }
}
    let lat = row[latIdx];
    let lng = row[lngIdx];

    if (!debug.length && (!lat || !lng)) {
      const addr = row[addrIdx];
      if (addr) {
        const coords = geocodeAddress_(addr);
        if (coords) {
          lat = coords.lat;
          lng = coords.lng;
          master.getRange(i + 1, latIdx + 1).setValue(lat);
          master.getRange(i + 1, lngIdx + 1).setValue(lng);
        } else {
          debug.push("Address not geocodable");
        }
      } else {
        debug.push("No address");
      }
    }

    if (!debug.length) {
      const dist = haversineMiles_(client.lat, client.lng, lat, lng);
      if (isNaN(dist)) {
        debug.push("Bad coordinates");
      } else if (dist > radiusMiles) {
        debug.push(`Outside radius (${dist.toFixed(2)} > ${radiusMiles})`);
      } else {
        const mapsLink = `https://www.google.com/maps/dir/?api=1&destination=${lat},${lng}`;
        const outRow = row.concat([dist.toFixed(2), mapsLink, "MATCH ✅"]);
        results.appendRow(outRow);
        written++;
        continue;
      }
    }

    const mapsLink = (lat && lng)
      ? `https://www.google.com/maps/dir/?api=1&destination=${lat},${lng}`
      : "";
    results.appendRow(row.concat(["", mapsLink, debug.join("; ")]));
  }

  SpreadsheetApp.getActive().toast(`Search done: ${written} match(es).`);
}

/***** GEO + DISTANCE *****/
function geocodeAddress_(address) {
  if (!address) return null;
  try {
    const res = Maps.newGeocoder().geocode(address);
    if (res.status === "OK" && res.results && res.results.length) {
      const loc = res.results[0].geometry.location;
      return { lat: loc.lat, lng: loc.lng };
    }
  } catch (err) {
    Logger.log("Geocode error: " + err);
  }
  return null;
}

function haversineMiles_(lat1, lon1, lat2, lon2) {
  const R = 3958.8;
  const toRad = d => (d * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(lat1)) *
      Math.cos(toRad(lat2)) *
      Math.sin(dLon / 2) ** 2;
  return R * 2 * Math.asin(Math.sqrt(a));
}


/**
 * Returns object:
 * {
 *   categories: [...],               // sorted list of categories
 *   subcategoriesFlat: [...],        // sorted list of all subcategories
 *   categoryToSubcategories: {       // mapping category -> [subcategories]
 *     "Housing": ["Male Only", "Sober Living"],
 *     ...
 *   }
 * }
 */
function getCategoriesAndSubcategories() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(MASTER_SHEET);
  if (!master) return { categories: [], subcategoriesFlat: [], categoryToSubcategories: {} };

  const { map } = getHeaderMap_(master);
  const data = master.getDataRange().getValues();

  const catSet = new Set();
  const subSet = new Set();
  const mapping = {}; // category -> Set(subcategories)

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rawCats = (row[map[CATEGORY_COL_NAME]] || "").toString();
    const rawSubs = (row[map[SUBCATEGORY_COL_NAME]] || "").toString();

    const cats = rawCats.split(",").map(s => s.trim()).filter(Boolean);
    const subs = rawSubs.split(",").map(s => s.trim()).filter(Boolean);

    // Add to flat sets
    cats.forEach(c => catSet.add(c));
    subs.forEach(s => subSet.add(s));

    // Build mapping
    cats.forEach(c => {
      if (!mapping[c]) mapping[c] = new Set();
      subs.forEach(s => mapping[c].add(s));
    });
  }

  // Convert Sets to sorted arrays
  const categories = [...catSet].sort();
  const subcategoriesFlat = [...subSet].sort();

  const categoryToSubcategories = {};
  Object.keys(mapping).forEach(c => {
    categoryToSubcategories[c] = [...mapping[c]].sort();
  });

  return {
    categories,
    subcategoriesFlat,
    categoryToSubcategories
  };
}



function installTriggers() {
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}
