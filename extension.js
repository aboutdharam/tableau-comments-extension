$(document).ready(function () {
    $('#openFormBtn').click(function () {
        $('#formContainer').show();

        // 👉 Replace with your actual Power Apps URL
        const powerAppsUrl = "https://apps.powerapps.com/play/XYZ…";

        $('#powerAppsFrame').attr("src", powerAppsUrl);
    });
});