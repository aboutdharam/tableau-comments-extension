// Wait until DOM loads
document.addEventListener("DOMContentLoaded", function () {

  // Initialize Tableau Extension
  tableau.extensions.initializeAsync().then(function () {

    console.log("Tableau Extension Initialized");

    // Get username from Tableau
    const tableauUser = tableau.extensions.environment.userName || "Unknown User";

    // Button click event
    document.getElementById("openForm").addEventListener("click", function () {

      // 👉 Replace with your actual values or Tableau parameter logic later
      const clientName = "Sample Client";
      const period = "2026 Q1";
      const startDate = "2026-01-01";
      const endDate = "2026-03-31";

      // 👉 Replace with your real Power Apps URL
      const basePowerAppsUrl = "PASTE_YOUR_POWERAPPS_URL_HERE";

      // Build URL with parameters
      const fullUrl =
        basePowerAppsUrl +
        "&client=" + encodeURIComponent(clientName) +
        "&user=" + encodeURIComponent(tableauUser) +
        "&period=" + encodeURIComponent(period) +
        "&startDate=" + encodeURIComponent(startDate) +
        "&endDate=" + encodeURIComponent(endDate);

      // Open Power Apps in new tab
      window.open(fullUrl, "_blank");

    });

  }).catch(function (error) {
    console.error("Error initializing Tableau Extension:", error);
  });

});
