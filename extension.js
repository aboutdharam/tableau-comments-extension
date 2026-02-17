(async function () {

    // 1. Initialize Tableau Extension
    await tableau.extensions.initializeAsync();

    const frame = document.getElementById("powerAppsFrame");
    const button = document.getElementById("openFormBtn");

    // Get Tableau Parameter Value
    async function getParam(name) {
        const params = await tableau.extensions.dashboardContent.dashboard.getParametersAsync();
        const p = params.find(x => x.name === name);
        return p ? p.currentValue.value : "";
    }

    // Get Dashboard Name
    function getDashboardName() {
        return tableau.extensions.dashboardContent.dashboard.name;
    }

    // Get View (Worksheet) Name
    function getViewName() {
        const sheets = tableau.extensions.dashboardContent.dashboard.worksheets;
        return sheets.length > 0 ? sheets[0].name : "";
    }

    // Get Tableau Username (viewer)
    function getUsername() {
        return tableau.extensions.environment.username;
    }

    async function buildPowerAppsUrl() {

        // Auto-detected
        const DashboardName = getDashboardName();
        const ViewName = getViewName();
        const Username = getUsername();

        // From Tableau Parameters
        const IssuerName = await getParam("Issuer Name Param");
        const StartDate  = await getParam("Start Date");
        const EndDate    = await getParam("End Date");

        // Your Power Apps URL (clean)
        const baseUrl = "https://apps.powerapps.com/play/e/default-dc265699-74fc-490e-b9d0-f41eb1055450/a/a5f0a652-a3d3-4676-b992-8c1e894b2b6c?tenantId=dc265699-74fc-490e-b9d0-f41eb1055450&source=sharebutton";

        // Build final URL
        const finalUrl =
              baseUrl
            + `&DashboardName=${encodeURIComponent(DashboardName)}`
            + `&ViewName=${encodeURIComponent(ViewName)}`
            + `&IssuerName=${encodeURIComponent(IssuerName)}`
            + `&StartDate=${encodeURIComponent(StartDate)}`
            + `&EndDate=${encodeURIComponent(EndDate)}`
            + `&Username=${encodeURIComponent(Username)}`;

        return finalUrl;
    }

    button.addEventListener("click", async () => {
        const url = await buildPowerAppsUrl();
        frame.src = url;
        document.getElementById("formContainer").style.display = "block";
    });

})();

