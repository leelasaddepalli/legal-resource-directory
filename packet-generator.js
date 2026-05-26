/***** Code.gs - Generate Google Doc resource packet from Search Results *****/

const TEMPLATE_OVERVIEW = `This resource directory provides an overview of community services available in the locality you expressed interest in. These organizations are independent entities, and the Harris County Public Defender’s Office does not sponsor, endorse, or promote any of them. They are offered solely for your information, and it is your choice and responsibility whether to contact or utilize any services listed.

If you have any questions or concerns about the directory or need additional guidance, please contact your advocate at (phone number) or email (email address).`;

function showPacketSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('PacketSidebar')
    .setTitle('Create Resource Packet');
  SpreadsheetApp.getUi().showSidebar(html);
}

/* Helper: build map from header row */
function getHeaderMapFromSheet(sheet) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  header.forEach((h, idx) => map[h] = idx);
  return { header, map };
}

/* Returns array of matched (only MATCH rows) from "Search Results" sheet */
function getMatchedResults() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Search Results");
  if (!sheet) throw new Error("Search Results sheet not found.");

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const rows = data.slice(1);

  // Filter only rows where the last column contains "MATCH"
  const matchRows = rows.filter(row => {
    const debugCol = row[row.length - 1]; // Last column (DEBUG)
    return typeof debugCol === "string" && debugCol.toUpperCase().includes("MATCH ✅");
  });

  // Map into objects keyed by header names
  const objects = matchRows.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });

  return objects;
}

/* Return count for sidebar preview */
function getMatchesCount() {
  return getMatchedResults().length;
}

/* Main: create Google Doc using template-like layout and insert matched resources */
function generateResourcePacket(clientName, advocateName, advocatePhone, advocateEmail) {
  clientName = clientName || "";
  advocateName = advocateName || "";
  advocatePhone = advocatePhone || "";
  advocateEmail = advocateEmail || "";

  const matches = getMatchedResults();
  if (!matches || matches.length === 0) {
    throw new Error("No matched results found in the 'Search Results' sheet. Run a search first.");
  }

  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT", "MMMM d, yyyy");
  const docName = `Individualized Resource Directory - ${clientName || 'Client'} - ${dateStr}`;
  const doc = DocumentApp.create(docName);
  const body = doc.getBody();

  // ----------------------------------------------------------
  // Styling defaults
  // ----------------------------------------------------------
  body.setAttributes({
    [DocumentApp.Attribute.FONT_FAMILY]: "Book Antiqua",
    [DocumentApp.Attribute.FONT_SIZE]: 12
  });

  // ----------------------------------------------------------
  // Intro Table
  // ----------------------------------------------------------
  const metaTable = body.appendTable();
  const metaRow = metaTable.appendTableRow();
  metaRow.appendTableCell(`Date: ${dateStr}`);
  metaRow.appendTableCell(`Client: ${clientName}`);
  metaRow.appendTableCell(`Writer: ${advocateName}`);

  body.appendParagraph("");

  // ----------------------------------------------------------
  // OVERVIEW SECTION
  // ----------------------------------------------------------
  const overviewHeader = body.appendParagraph("Overview");
  overviewHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const overviewText =
    `This resource directory provides an overview of community services available in the locality you expressed interest in. These organizations are independent entities, and the Harris County Public Defender’s Office does not sponsor, endorse, or promote any of them. They are offered solely for your information, and it is your choice and responsibility whether to contact or utilize any services listed.\n\n` +
    `If you have any questions or concerns about the directory or need additional guidance, please contact your advocate at ${advocatePhone || "(phone number)"} or email ${advocateEmail || "(email address)"}.`;

   body.appendParagraph(overviewText).setSpacingAfter(12);

  // ----------------------------------------------------------
  // RESOURCE BLOCKS
  // ----------------------------------------------------------
  matches.forEach((r, idx) => {
    body.appendHorizontalRule();

    const section = body.appendParagraph(`Resource ${idx + 1}`);
    section.setHeading(DocumentApp.ParagraphHeading.HEADING3);

    const org  = r['Organization Name'] || r['Organization'] || "";
    const cat  = r['Category Tags'] || "";
    const sub  = r['Subcategory Tags'] || "";
    const serv = r['Services Offered'] || "";
    const phone = r['Phone'] || "";
    const addr  = r['Address'] || "";
    const dist  = r['Distance (mi)'] || "";
    const maps  = r['Google Maps Link'] || "";

    // ORG NAME
    body.appendParagraph(org).setBold(true);

    // CATEGORY
    if (cat || sub) {
      body.appendParagraph(`Category: ${cat}${sub ? " — " + sub : ""}`)
        .setBold(false);
    }

    // SERVICES
    if (serv) {
      body.appendParagraph("Services Offered:").setBold(true);
      body.appendParagraph(serv).setBold(false);
    }

    // CONTACT
    if (phone || addr) {
      body.appendParagraph("Contact / Address:").setBold(true);
      if (phone) body.appendParagraph(`Phone: ${phone}`).setBold(false);
      if (addr)  body.appendParagraph(`Address: ${addr}`).setBold(false);
    }

    // DISTANCE
    if (dist) {
      body.appendParagraph(`Distance from Client: ${dist} miles`)
        .setBold(false);
    }

    // MAP LINK
    if (maps) {
      const mapsP = body.appendParagraph("Google Maps Link: ");
      mapsP.appendText("Open in Maps").setLinkUrl(maps);
    }

    body.appendParagraph("").setSpacingAfter(6);
  });

  doc.saveAndClose();
  return doc.getUrl();
}
