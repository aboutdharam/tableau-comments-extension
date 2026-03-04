(function () {
  const debugEl = document.getElementById('debug');
  const statusEl = document.getElementById('status');
  const btn = document.getElementById('openFormBtn');
  const frame = document.getElementById('powerAppsFrame');
  const container = document.getElementById('formContainer');

  // Added: Header element so we can hide it
  const titleSection = document.getElementById('titleSection');

  function log(msg, obj) {
    const line = `[${new Date().toLocaleTimeString()}] ${msg}`;
    console.log(line, obj ?? "");
    if (debugEl) debugEl.textContent += `\n${line}${obj ? " " + JSON.stringify(obj) : ""}`;
    if (statusEl) statusEl.textContent = msg;
  }

  if (!btn || !frame || !container) {
    log("ERROR: Missing HTML elements.");
    return;
  }

  // Base PowerApps URL (must contain & not &amp;)
  const baseUrl =
    "https://apps.powerapps.com/play/e/default-dc265699-74fc-490e-b9d0-f41eb1055450/a/a5f0a652-a3d3-4676-b992-8c1e894b2b6c?tenantId=dc265699-74fc-490e-b9d0-f41eb1055450&hint=6e87fd1b-7859-4e05-9922-75b435dc5802&source=sharebutton&source=iframe&hideNavBar=true";

  // ---------- Helper functions ----------

  async function listDashboardWorksheets() {
    const ws = tableau.extensions.dashboardContent.dashboard.worksheets || [];
    const names = ws.map(w => w.name);
    log("Worksheets present:", names);
    return names;
  }

  async function getParam(name) {
    const dashboard = tableau.extensions.dashboardContent.dashboard;
    const params = await dashboard.getParametersAsync();
    const p = params.find(x => x.name === name);
    if (!p) log(`WARN: Tableau parameter missing: ${name}`);
    const v = p?.currentValue;
    return v?.formattedValue ?? v?.value ?? "";
  }

  function getDashboardName() {
    return tableau.extensions.dashboardContent.dashboard.name || "";
  }

  function getViewName() {
    const sheets = tableau.extensions.dashboardContent.dashboard.worksheets || [];
    return sheets.length > 0 ? sheets[0].name : "";
  }

  // -------- USERNAME SHEET FETCH --------
  async function getUsernameFromSheet() {
    try {
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const ws = dashboard.worksheets.find(w => w.name === "Username");
      if (!ws) return "";

      const summary = await ws.getSummaryDataAsync({ maxRows: 1, ignoreSelection: true });
      const cell = summary.data?.[0]?.[0];
      const raw = cell?.formattedValue ?? cell?.value ?? "";
      return raw || "";
    } catch (e) {
      return "";
    }
  }

  function getEnvironmentUserFallback() {
    try {
      const uid = tableau.extensions.environment?.uniqueUserId;
      return typeof uid === "string" ? uid : "";
    } catch {
      return "";
    }
  }

  async function resolveUsername() {
    const sheetVal = await getUsernameFromSheet();
    return sheetVal || getEnvironmentUserFallback();
  }

  // -------- CARDNAME SHEET FETCH (NEW) --------
  async function getCardNameFromSheet() {
    try {
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const ws = dashboard.worksheets.find(w => w.name === "CardName");
      if (!ws) {
        log('WARN: Worksheet "CardName" NOT found.');
        return "";
      }

      const summary = await ws.getSummaryDataAsync({ maxRows: 1, ignoreSelection: true });
      const cell = summary.data?.[0]?.[0];
      const raw = cell?.formattedValue ?? cell?.value ?? "";

      log("CardName fetched:", raw);
      return raw || "";
    } catch (e) {
      log("ERROR reading CardName sheet", e);
      return "";
    }
  }

  // -------- BUILD PowerApps URL --------
  async function buildPowerAppsUrl() {
    const DashboardName = getDashboardName();
    const ViewName = getViewName();
    const Username = await resolveUsername();

    const IssuerName = await getParam("Issuer Name Param");
    const StartDate = await getParam("Start Date");
    const EndDate = await getParam("End Date");
    const IssuerCountry = await getParam("Issuer Country");

    // NEW: fetch CardName from sheet
    const CardName = await getCardNameFromSheet();

    const qp = new URLSearchParams({
      DashboardName,
      ViewName,
      IssuerName,
      StartDate,
      EndDate,
      IssuerCountry,
      Username,
      CardName   // <--- NEW FIELD passed to Power Apps
    });

    const finalUrl = `${baseUrl}&${qp.toString()}`;
    log("Final PowerApps URL:", finalUrl);
    return finalUrl;
  }

  // -------- Init and UI hide/show logic --------
  async function waitForExtensionsApi(timeoutMs = 10000) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (window.tableau?.extensions?.initializeAsync) return true;
      await new Promise(r => setTimeout(r, 100));
    }
    return false;
  }

  async function init() {
    log("Initializing extension…");

    if (!(await waitForExtensionsApi())) {
      log("ERROR: Tableau Extensions API missing.");
      return;
    }

    try {
      await tableau.extensions.initializeAsync();
    } catch (e) {
      log("ERROR init:", e);
      return;
    }

    await listDashboardWorksheets();
    btn.disabled = false;

    btn.addEventListener("click", async () => {
      log('Open Comment Form clicked');

      // Hide top UI
      btn.style.display = "none";
      if (titleSection) titleSection.style.display = "none";
      if (statusEl) statusEl.style.display = "none";
      if (debugEl) debugEl.style.display = "none";

      try {
        const url = await buildPowerAppsUrl();
        if (!url.startsWith("https://apps.powerapps.com/")) {
          log("Invalid URL generated", url);

          // Restore UI if error
          btn.style.display = "";
          if (titleSection) titleSection.style.display = "";
          if (statusEl) statusEl.style.display = "";
          if (debugEl) debugEl.style.display = "";
          return;
        }

        frame.src = url;
        container.style.display = "block";
      } catch (e) {
        log("ERROR while opening form", e);

        // Restore UI
        btn.style.display = "";
        if (titleSection) titleSection.style.display = "";
        if (statusEl) statusEl.style.display = "";
        if (debugEl) debugEl.style.display = "";
      }
    });

    log('Ready. Click "Open Comment Form".');
  }

  init();
})();