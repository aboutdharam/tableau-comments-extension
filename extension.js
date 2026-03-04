(function () {
  const debugEl = document.getElementById('debug');
  const statusEl = document.getElementById('status');
  const btn = document.getElementById('openFormBtn');
  const frame = document.getElementById('powerAppsFrame');
  const container = document.getElementById('formContainer');

  function log(msg, obj) {
    const line = `[${new Date().toLocaleTimeString()}] ${msg}`;
    console.log(line, obj ?? '');
    if (debugEl) debugEl.textContent += `\n${line}${obj ? ' ' + JSON.stringify(obj) : ''}`;
    if (statusEl) statusEl.textContent = msg;
  }

  // NOTE: Keep the sanity check, but use logical OR (||) for readability.
  if (!btn || !frame || !container) {
    log('ERROR: Missing HTML elements (openFormBtn / powerAppsFrame / formContainer).');
    return;
  }

  // Power Apps base URL — must be plain & (not &amp;)
  // NOTE: This URL is already using plain '&'. No replace() needed.
  const baseUrl =
    'https://apps.powerapps.com/play/e/default-dc265699-74fc-490e-b9d0-f41eb1055450/a/a5f0a652-a3d3-4676-b992-8c1e894b2b6c?tenantId=dc265699-74fc-490e-b9d0-f41eb1055450&hint=6e87fd1b-7859-4e05-9922-75b435dc5802&source=sharebutton&source=iframe&hideNavBar=true';

  // ---------- Tableau helpers ----------
  async function listDashboardWorksheets() {
    const ws = tableau.extensions.dashboardContent.dashboard.worksheets || [];
    const names = ws.map(w => w.name);
    log('Worksheets present on this dashboard:', names);
    return names;
  }

  async function getParam(name) {
    const dashboard = tableau.extensions.dashboardContent.dashboard;
    const params = await dashboard.getParametersAsync();
    const p = params.find(x => x.name === name);
    if (!p) log(`WARN: Tableau parameter not found: "${name}"`);
    // NOTE: prefer formattedValue when present; fallback to raw value
    const v = p?.currentValue;
    return v?.formattedValue ?? v?.value ?? '';
  }

  function getDashboardName() {
    return tableau.extensions.dashboardContent.dashboard.name || '';
  }

  function getViewName() {
    const sheets = tableau.extensions.dashboardContent.dashboard.worksheets || [];
    return sheets.length > 0 ? sheets[0].name : '';
  }

  // Preferred: USERNAME() from worksheet named EXACTLY "Username"
  async function getUsernameFromSheet() {
    try {
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const ws = dashboard.worksheets.find(w => w.name === 'Username');
      if (!ws) { log('WARN: Worksheet "Username" NOT found on this dashboard.'); return ''; }

      // ignoreSelection prevents mark selections from blanking the cell;
      // global dashboard filters can still blank it, so exclude this sheet from those filters.
      const summary = await ws.getSummaryDataAsync({ maxRows: 1, ignoreSelection: true });
      const cell = summary.data?.[0]?.[0];
      const raw = cell?.formattedValue ?? cell?.value ?? '';
      log('Username read from "Username" sheet:', { raw });
      return raw || '';
    } catch (e) {
      log('ERROR reading USERNAME() from sheet "Username"', e);
      return '';
    }
  }

  // Fallback: environment.uniqueUserId (when exposed on your site)
  function getEnvironmentUserFallback() {
    try {
      const env = tableau.extensions.environment;
      const uid = (env && typeof env.uniqueUserId === 'string') ? env.uniqueUserId : '';
      if (uid) log('Using fallback uniqueUserId as Username.', { uniqueUserId: uid });
      return uid || '';
    } catch (e) {
      log('ERROR reading environment.uniqueUserId', e);
      return '';
    }
  }

  async function resolveUsername() {
    const fromSheet = await getUsernameFromSheet();
    if (fromSheet) return fromSheet;
    return getEnvironmentUserFallback();
  }

  async function buildPowerAppsUrl() {
    const DashboardName = getDashboardName();
    const ViewName      = getViewName();
    const Username      = await resolveUsername();

    const IssuerName    = await getParam('Issuer Name Param'); // exact Tableau parameter names
    const StartDate     = await getParam('Start Date');
    const EndDate       = await getParam('End Date');
    const IssuerCountry = await getParam('Issuer Country');    // make sure the Tableau parameter matches exactly

    const qp = new URLSearchParams({
      DashboardName: DashboardName ?? '',
      ViewName:      ViewName ?? '',
      IssuerName:    IssuerName ?? '',
      StartDate:     StartDate ?? '',
      EndDate:       EndDate ?? '',
      IssuerCountry: IssuerCountry ?? '', // send key that Power Apps expects
      Username:      Username ?? ''
    });

    // NOTE: Use real '&' concatenation. Power Apps expects standard query string.
    const finalUrl = `${baseUrl}&${qp.toString()}`;
    log('Final Power Apps URL built.', { finalUrl });
    return finalUrl;
  }

  // Wait for Extensions API (up to 10s)
  async function waitForExtensionsApi(timeoutMs = 10000) {
    const t0 = Date.now();
    while (Date.now() - t0 < timeoutMs) {
      if (window.tableau && tableau.extensions && typeof tableau.extensions.initializeAsync === 'function') {
        return true;
      }
      await new Promise(r => setTimeout(r, 100));
    }
    return false;
  }

  async function init() {
    log('Initializing extension…');

    const ok = await waitForExtensionsApi();
    if (!ok) {
      log('ERROR: Tableau Extensions API not detected. Ensure this is added as an Extension (not Web Page) and API script loaded.');
      return;
    }

    try {
      await tableau.extensions.initializeAsync();
      log('Tableau extensions initialized.');
    } catch (e) {
      log('ERROR initializing Tableau extensions.', e);
      return;
    }

    // List the worksheets ON THIS DASHBOARD so you can verify "Username" appears
    await listDashboardWorksheets();

    btn.disabled = false;

    btn.addEventListener('click', async () => {
      log('"Open Comment Form" clicked.');
      try {
        // CHANGE: Hide the button immediately after it’s clicked,
        // so only the form (iframe) remains visible.
        btn.style.display = 'none';

        const url = await buildPowerAppsUrl();
        if (!url || !url.startsWith('https://apps.powerapps.com/')) {
          log('ERROR: Built URL is empty or invalid.', { url });

          // CHANGE: If URL is invalid, show the button again so user can retry.
          btn.style.display = ''; // revert to default (visible)
          return;
        }

        frame.src = url;
        container.style.display = 'block';
        log('Iframe src set and container shown.');
      } catch (e) {
        log('ERROR while building or setting iframe URL.', e);

        // CHANGE: On unexpected error, also restore the button so user can retry.
        btn.style.display = '';
      }
    });

    log('Ready. Click "Open Comment Form".');
  }

  init();
})();