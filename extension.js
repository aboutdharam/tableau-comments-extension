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

  // Hard-fail if required elements are missing
  if (!btn || !frame || !container) {
    log('ERROR: Missing HTML elements (openFormBtn / powerAppsFrame / formContainer).');
    return;
  }

  // Your Power Apps base URL (IMPORTANT: use plain &)
  const baseUrl =
    'https://apps.powerapps.com/play/e/default-dc265699-74fc-490e-b9d0-f41eb1055450/a/a5f0a652-a3d3-4676-b992-8c1e894b2b6c?tenantId=dc265699-74fc-490e-b9d0-f41eb1055450&hint=6e87fd1b-7859-4e05-9922-75b435dc5802&source=sharebutton&source=iframe&hideNavBar=true';

  // --- Tableau helpers (Tableau mode only) ---
  async function getParam(name) {
    const dashboard = tableau.extensions.dashboardContent.dashboard;
    const params = await dashboard.getParametersAsync();
    const p = params.find(x => x.name === name);
    if (!p) log(`WARN: Tableau parameter not found: "${name}"`);
    return p ? p.currentValue.value : '';
  }
  function getDashboardName() {
    return tableau.extensions.dashboardContent.dashboard.name || '';
  }
  function getViewName() {
    const sheets = tableau.extensions.dashboardContent.dashboard.worksheets || [];
    return sheets.length > 0 ? sheets[0].name : '';
  }
  function getUsername() {
    return tableau.extensions.environment.username || '';
  }

  async function buildPowerAppsUrl_Tableau() {
    const DashboardName = getDashboardName();
    const ViewName      = getViewName();
    const Username      = getUsername();

    const IssuerName    = await getParam('Issuer Name Param'); // exact Tableau parameter name
    const StartDate     = await getParam('Start Date');        // exact Tableau parameter name
    const EndDate       = await getParam('End Date');          // exact Tableau parameter name

    const qp = new URLSearchParams({
      DashboardName: DashboardName ?? '',
      ViewName:      ViewName ?? '',
      IssuerName:    IssuerName ?? '',
      StartDate:     StartDate ?? '',
      EndDate:       EndDate ?? '',
      Username:      Username ?? ''
    });

    const finalUrl = `${baseUrl}&${qp.toString()}`;
    log('Final Power Apps URL built.', { finalUrl });
    return finalUrl;
  }

  // Main init flow — Tableau **only**
  async function init() {
    // 1) Verify the Extensions API is present
    if (typeof window.tableau === 'undefined' ||
        !tableau.extensions ||
        typeof tableau.extensions.initializeAsync !== 'function') {
      log('ERROR: Tableau Extensions API not detected. The page is not running as an extension, or the API failed to load.');
      // Keep the button disabled; nothing to do
      return;
    }

    // 2) Initialize the extension
    try {
      await tableau.extensions.initializeAsync();
      log('Tableau extensions initialized.');
    } catch (e) {
      log('ERROR initializing Tableau extensions.', e);
      return;
    }

    // 3) Enable the button now that initialization succeeded
    btn.disabled = false;

    // 4) Wire click handler
    btn.addEventListener('click', async () => {
      log('"Open Comment Form" clicked.');
      try {
        const url = await buildPowerAppsUrl_Tableau();
        if (!url || !url.startsWith('https://apps.powerapps.com/')) {
          log('ERROR: Built URL is empty or invalid.', { url });
          return;
        }
        frame.src = url;
        container.style.display = 'block';
        log('Iframe src set and container shown.');
      } catch (e) {
        log('ERROR while building or setting iframe URL.', e);
      }
    });

    log('Ready. Click "Open Comment Form".');
  }

  // Start init
  init();
})();