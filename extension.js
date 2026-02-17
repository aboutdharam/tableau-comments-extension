(function () {
  const debugEl = document.getElementById('debug');
  const btn = document.getElementById('openFormBtn');
  const frame = document.getElementById('powerAppsFrame');
  const container = document.getElementById('formContainer');

  function log(msg, obj) {
    const line = `[${new Date().toLocaleTimeString()}] ${msg}`;
    console.log(line, obj ?? '');
    if (debugEl) debugEl.textContent += `\n${line}${obj ? ' ' + JSON.stringify(obj) : ''}`;
  }

  // --- 0) Sanity: elements exist
  if (!btn || !frame || !container) {
    log('ERROR: Missing required HTML elements (openFormBtn / powerAppsFrame / formContainer).');
    return;
  }

  // --- 1) Determine if we are inside Tableau (Extensions API available)
  const inTableau = !!(window.tableau &&
                       tableau.extensions &&
                       typeof tableau.extensions.initializeAsync === 'function');

  // --- 2) Utilities (used in Tableau mode)
  async function getParam(name) {
    try {
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const params = await dashboard.getParametersAsync();
      const p = params.find(x => x.name === name);
      if (!p) log(`WARN: Tableau parameter not found: "${name}"`);
      return p ? p.currentValue.value : '';
    } catch (e) {
      log(`ERROR while reading parameter "${name}"`, e);
      return '';
    }
  }
  function getDashboardName() {
    try { return tableau.extensions.dashboardContent.dashboard.name || ''; }
    catch (e) { log('ERROR reading dashboard name', e); return ''; }
  }
  function getViewName() {
    try {
      const sheets = tableau.extensions.dashboardContent.dashboard.worksheets || [];
      return sheets.length > 0 ? sheets[0].name : '';
    } catch (e) { log('ERROR reading view name', e); return ''; }
  }
  function getUsername() {
    try { return tableau.extensions.environment.username || ''; }
    catch (e) { log('ERROR reading environment username', e); return ''; }
  }

  // --- 3) Your Power Apps base URL (IMPORTANT: use plain & not &amp;)
  const baseUrl =
    'https://apps.powerapps.com/play/e/default-dc265699-74fc-490e-b9d0-f41eb1055450/a/a5f0a652-a3d3-4676-b992-8c1e894b2b6c?tenantId=dc265699-74fc-490e-b9d0-f41eb1055450&hint=6e87fd1b-7859-4e05-9922-75b435dc5802&source=sharebutton';

  // Safety: if someone pasted &amp; by mistake, normalize it:
  const normalizedBaseUrl = baseUrl.replace(/&amp;/g, '&');

  async function buildPowerAppsUrl_Tableau() {
    const DashboardName = getDashboardName();
    const ViewName      = getViewName();
    const Username      = getUsername();
    const IssuerName    = await getParam('Issuer Name Param');
    const StartDate     = await getParam('Start Date');
    const EndDate       = await getParam('End Date');

    const qp = new URLSearchParams({
      DashboardName: DashboardName ?? '',
      ViewName:      ViewName ?? '',
      IssuerName:    IssuerName ?? '',
      StartDate:     StartDate ?? '',
      EndDate:       EndDate ?? '',
      Username:      Username ?? ''
    });

    const finalUrl = `${normalizedBaseUrl}&${qp.toString()}`;
    log('Final Power Apps URL (Tableau mode) built.', { finalUrl });
    return finalUrl;
  }

  // Fallback for plain browser (no Tableau): open base URL only
  async function buildPowerAppsUrl_Browser() {
    const finalUrl = normalizedBaseUrl;
    log('Final Power Apps URL (Browser fallback) built.', { finalUrl });
    return finalUrl;
  }

  // --- 4) Initialize (if in Tableau), then wire up the button
  async function init() {
    if (inTableau) {
      try {
        await tableau.extensions.initializeAsync();
        log('Tableau extensions initialized.');
      } catch (e) {
        log('ERROR initializing Tableau extensions.', e);
        return;
      }
    } else {
      log('Running in plain browser (no Tableau). Using fallback mode.');
    }

    btn.addEventListener('click', async () => {
      log('"Open Comment Form" clicked.');
      try {
        const url = inTableau ? await buildPowerAppsUrl_Tableau()
                              : await buildPowerAppsUrl_Browser();

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

  // Start
  init();
})();