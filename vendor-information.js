const GOOGLE_CLIENT_ID = "154634144934-9hg9o4ra7uriu5hrivaaj73mduj7udf4.apps.googleusercontent.com";
const GOOGLE_API_KEY = "AIzaSyCh8ia27PwiWJkPCypoUyvj5TD8YJVjJSc";
const SHEETS_DISCOVERY_DOC = "https://sheets.googleapis.com/$discovery/rest?version=v4";
const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets.readonly";
const VENDOR_SPREADSHEET_ID = "1lL1WO9Py1V_lHa9SSy5gh9EanP6zOQuiu6RvWsVd5Mk";
const VENDOR_TAB = "VENDOR - ALL INFORMATION";

const $ = (id) => document.getElementById(id);
const state = { tokenClient: null, gapiReady: false, gisReady: false, rows: [], headers: [], vendorColumn: -1 };

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>'"]/g, (char) => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;","\"":"&quot;"}[char]));
}

function normalizeHeader(value) {
  return String(value || "").toLowerCase().replace(/&/g, "and").replace(/[^a-z0-9]+/g, " ").trim();
}

function setStatus(message, kind = "") {
  $("vendorStatus").textContent = message;
  $("connectionDot").className = ("connection-dot " + kind).trim();
}

function sheetRange() {
  return "'" + VENDOR_TAB.replaceAll("'", "''") + "'!A:ZZ";
}

function initializeGoogle() {
  if (state.tokenClient || state.gapiReady) return;
  if (!window.gapi || !window.google?.accounts?.oauth2) { setTimeout(initializeGoogle, 250); return; }
  gapi.load("client", async () => {
    try {
      await gapi.client.init({ apiKey: GOOGLE_API_KEY, discoveryDocs: [SHEETS_DISCOVERY_DOC] });
      state.gapiReady = true;
      maybeEnableConnection();
    } catch (error) { setStatus("Google API setup error: " + error.message, "error"); }
  });
  state.tokenClient = google.accounts.oauth2.initTokenClient({ client_id: GOOGLE_CLIENT_ID, scope: SHEETS_SCOPE, callback: "" });
  state.gisReady = true;
  maybeEnableConnection();
}

function maybeEnableConnection() { $("connectButton").disabled = !(state.gapiReady && state.gisReady); }

function requestGoogleAccess() {
  state.tokenClient.callback = async (response) => {
    if (response.error) { setStatus("Google authorization failed: " + response.error, "error"); return; }
    gapi.client.setToken({ access_token: response.access_token });
    await loadVendorGrid();
  };
  state.tokenClient.requestAccessToken({ prompt: gapi.client.getToken() ? "" : "consent" });
}

async function loadVendorGrid() {
  try {
    setStatus("Reading the live vendor grid…", "working");
    const response = await gapi.client.sheets.spreadsheets.values.get({ spreadsheetId: VENDOR_SPREADSHEET_ID, range: sheetRange(), valueRenderOption: "FORMATTED_VALUE" });
    const values = response.result.values || [];
    if (!values.length) throw new Error("The vendor grid is empty.");
    state.headers = values[0].map((header) => String(header || "").trim());
    state.vendorColumn = findColumn(["vendor", "vendor name", "source", "vendor source"]);
    if (state.vendorColumn < 0) throw new Error("Could not find a Vendor column in the sheet.");
    state.rows = values.slice(1).map((cells, index) => ({ rowNumber: index + 2, cells })).filter((row) => String(row.cells[state.vendorColumn] || "").trim());
    $("searchPanel").classList.remove("hidden");
    $("vendorSearch").disabled = false;
    $("refreshButton").disabled = false;
    $("connectButton").textContent = "Reconnect Vendor Grid";
    setStatus("Connected · " + state.rows.length + " vendors loaded", "connected");
    renderResults("");
  } catch (error) { setStatus("Could not read the vendor grid: " + (error.result?.error?.message || error.message), "error"); }
}

function findColumn(names) {
  const wanted = names.map(normalizeHeader);
  return state.headers.findIndex((header) => wanted.includes(normalizeHeader(header)));
}

function valueFor(names) {
  const index = findColumn(names);
  return index < 0 ? "" : String(state.selected?.cells[index] || "").trim();
}

function renderResults(query) {
  const normalized = normalizeHeader(query);
  const matches = state.rows.filter((row) => !normalized || normalizeHeader(row.cells[state.vendorColumn]).includes(normalized)).slice(0, 50);
  $("searchMeta").textContent = normalized ? `${matches.length}${matches.length === 50 ? "+" : ""} matching vendors` : `${state.rows.length} vendors available`;
  $("vendorResults").innerHTML = matches.length ? matches.map((row) => `<button class="vendor-result" type="button" data-row="${row.rowNumber}">${escapeHtml(row.cells[state.vendorColumn])}</button>`).join("") : '<div class="vendor-empty">No matching vendors found.</div>';
  $("vendorResults").querySelectorAll("[data-row]").forEach((button) => button.addEventListener("click", () => showVendor(Number(button.dataset.row))));
}

function showVendor(rowNumber) {
  state.selected = state.rows.find((row) => row.rowNumber === rowNumber);
  if (!state.selected) return;
  const boldFields = [
    ["Agreement Status", ["Agreement Status"]], ["Vendor Status", ["Vendor Status"]], ["Link to the Agreement", ["Link to the Agreement", "Agreement Link"]],
    ["Counsel", ["Counsel"]], ["Date Last Approved", ["Date Last Approved", "Date last approved"]], ["Vendor Notes", ["Vendor Notes"]]
  ];
  const detailFields = [
    ["Legal Flag Information (Licensing Status)", ["Legal Flag Information (Licensing Status)", "Licensing Status", "Legal Flag Information"]], ["License Rights", ["License Rights"]], ["License Restrictions", ["License Restrictions"]], ["License Usage", ["License Usage"]], ["Standard Credit", ["Standard Credit"]], ["Rep Name", ["Rep Name"]], ["Rep Email", ["Rep Email"]], ["Contact Phone", ["Contact Phone"]], ["Footage Archival Rate", ["Footage Archival Rate"]], ["Still Archival Rate", ["Still Archival Rate"]]
  ];
  $("vendorDetail").innerHTML = `<div class="vendor-detail-heading"><div class="eyebrow">VENDOR PROFILE</div><h2>${escapeHtml(valueFor(["Vendor", "Vendor Name", "Source"]))}</h2><span>Live row ${rowNumber} from ${escapeHtml(VENDOR_TAB)}</span></div>${renderSection("Agreement & Approval", boldFields, true)}${renderSection("Licensing, Contact & Rates", detailFields, false)}`;
  $("vendorDetail").classList.remove("hidden");
}

function renderSection(title, fields, prominent) {
  return `<section class="vendor-section ${prominent ? "prominent" : ""}"><h3>${title}</h3><div class="vendor-field-grid">${fields.map(([label, names]) => `<div class="vendor-field"><span>${label}</span><strong>${formatValue(valueFor(names), label)}</strong></div>`).join("")}</div></section>`;
}

function formatValue(value, label) {
  if (!value) return '<em>Not listed</em>';
  if (/link/i.test(label) && /^https?:\/\//i.test(value)) return `<a href="${escapeHtml(value)}" target="_blank" rel="noopener">Open agreement ↗</a>`;
  return escapeHtml(value).replace(/\n/g, "<br>");
}

$("connectButton").addEventListener("click", requestGoogleAccess);
$("refreshButton").addEventListener("click", loadVendorGrid);
$("vendorSearch").addEventListener("input", (event) => renderResults(event.target.value));
initializeGoogle();
