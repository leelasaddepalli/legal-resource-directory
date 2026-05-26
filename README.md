# Legal Resource Directory
**Harris County Public Defender's Office | Partners for Justice Fellowship, 2025–26**

A Google Sheets-integrated tool built to help client advocates at a county public defender's office connect justice-impacted individuals with community resources across housing, employment, healthcare, benefits, and behavioral health services.

---

## Background

Client advocates at the Harris County Public Defender's Office had no centralized, searchable directory of community resources. Advocates were manually tracking services across disconnected lists, making it difficult to efficiently identify relevant resources for clients during needs assessments.

This tool was **designed and built from scratch** during a Partners for Justice Fellowship to solve that problem. It normalizes organizational data, dynamically indexes resources by category, performs radius-based geospatial search, and generates formatted resource packets for client use — all within the Google Sheets environment the office already used.

---

## Features

### Resource Search
- Advocate enters a client address and selects service categories and subcategories
- Tool geocodes the address and performs a **radius-based geospatial search** using the Haversine formula to calculate real-world distances
- Returns matching resources sorted by proximity, with Google Maps directions links
- Handles edge cases: missing coordinates are geocoded on-the-fly; resources with no subcategories are treated as inclusive matches

### Dynamic Category Indexing
- Master sheet drives all category and subcategory data
- `onEdit` trigger automatically syncs records to category-specific sheets when entries are added or updated
- Removes records from sheets they no longer belong to when categories change
- Auto-generates unique `RecordID` for each entry on creation

### Packet Generator
- Advocates can select resources from search results and generate a formatted client-facing resource packet
- Accessible via a custom Google Sheets menu (`Resource Directory → Create Packet`)

### Data Normalization
- Standardizes inconsistently formatted organizational data from multiple source lists
- Handles messy inputs: comma-separated tags, missing addresses, duplicate entries
- Results written to a dedicated `Search Results` sheet with match status and debug flags for transparency

---

## Tech Stack

- **Google Apps Script** (JavaScript) — core application logic
- **HTML/CSS** — sidebar UI rendered inside Google Sheets
- **Google Maps Geocoding API** — address-to-coordinate conversion
- **Google Sheets** — backend data store and user interface host

---

## File Structure

```
legal-resource-directory/
│
├── main-code.js          # Core logic: onEdit trigger, category sync engine,
│                         # geospatial search, geocoding, Haversine distance
│                         # calculation, category/subcategory indexing
│
├── packet-generator.js   # Packet generation logic: compiles selected resources
│                         # into a formatted client-facing document
│
├── search-sidebar.html   # Search UI: address input, category/subcategory
│                         # selection, radius filter, search trigger
│
├── packet-sidebar.html   # Packet UI: resource selection interface and
│                         # packet generation controls
│
└── README.md
```

---

## How It Works

```
Advocate opens sidebar
        ↓
Enters client address + selects categories/subcategories + sets radius
        ↓
Tool geocodes client address via Google Maps API
        ↓
Iterates Master sheet → filters by category → geocodes missing coordinates
        ↓
Calculates Haversine distance for each resource from client location
        ↓
Writes matches to Search Results sheet with distance + Maps link
        ↓
Advocate selects relevant resources → generates formatted client packet
```

---

## Key Technical Details

**Haversine Distance Calculation**
Implements the Haversine formula from scratch to compute great-circle distances between the client's location and each resource, enabling accurate radius-based filtering without an external routing API.

**Trigger-Based Auto-Sync**
An `onEdit` trigger fires whenever the Master sheet is updated, automatically routing records to the correct category sheets and removing stale entries — keeping the directory consistent without manual maintenance.

**Subcategory Matching Logic**
Resources with no subcategories listed are treated as inclusive matches for any subcategory filter, ensuring broad-category resources aren't excluded from searches unnecessarily.

**Geocoding Fallback**
If a resource is missing latitude/longitude, the tool geocodes its address on-the-fly during search and writes the coordinates back to the Master sheet for future efficiency.

---

## Context

Built during a Partners for Justice Research Fellowship at the **Harris County Public Defender's Office**, Houston, TX (2025–26). The office serves one of the largest public defense systems in the United States. This tool was used by client advocates conducting needs assessments for justice-impacted individuals navigating housing instability, lack of identification, barriers to employment, and access to healthcare and behavioral health services.
