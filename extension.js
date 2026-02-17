$(document).ready(function () {
    $('#openFormBtn').click(function () {
        $('#formContainer').show();

        // 👉 Replace with your actual Power Apps URL
        const powerAppsUrl = "https://apps.powerapps.com/play/e/default-dc265699-74fc-490e-b9d0-f41eb1055450/a/a5f0a652-a3d3-4676-b992-8c1e894b2b6c?tenantId=dc265699-74fc-490e-b9d0-f41eb1055450&hint=6e87fd1b-7859-4e05-9922-75b435dc5802&sourcetime=1771337483239";

        $('#powerAppsFrame').attr("src", powerAppsUrl);
    });
});