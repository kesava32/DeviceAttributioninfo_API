function callSingularAPI() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiUrl = "https://api.singular.net/api/attribution/attribution_details?keyspace=&device_id=&api_key=";

  const keyspace = sheet.getRange("D3").getValue();
  const deviceId = sheet.getRange("D5").getValue();
  const apiKey = sheet.getRange("D1").getValue();

  const apiCallUrl = `${apiUrl}${apiKey}&keyspace=${keyspace}&device_id=${deviceId}`;

  // Make the API call
  const response = UrlFetchApp.fetch(apiCallUrl);
  const jsonResponse = JSON.parse(response.getContentText());

  // Clear previous output
  sheet.getRange(10, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent().clearFormat();

  let startRow = 10;
  let startColumn = 1; // Start at column A for the first app

  jsonResponse.forEach((appData, index) => {
    let currentRow = startRow;
    let currentColumn = startColumn + (index * 3); // Leave 1 column gap (e.g., A-B for app1, D-E for app2, etc.)

    // Header for each app
    const headerRange = sheet.getRange(currentRow, currentColumn, 1, 2);
    headerRange.merge();
    headerRange.setValue("Attribution Information")
               .setBackground('#FFD700') // Gold background for header
               .setFontWeight('bold')
               .setHorizontalAlignment('center')
               .setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
    currentRow++;

    // App Information
    sheet.getRange(currentRow, currentColumn).setValue("App Name").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(currentRow, currentColumn + 1).setValue(appData.app_name).setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
    currentRow++;

    sheet.getRange(currentRow, currentColumn).setValue("Bundle ID").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(currentRow, currentColumn + 1).setValue(appData.app_long_name).setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
    currentRow++;

    sheet.getRange(currentRow, currentColumn).setValue("Source").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(currentRow, currentColumn + 1).setValue(appData.install_info.network).setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
    currentRow++;

    sheet.getRange(currentRow, currentColumn).setValue("Install Time").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(currentRow, currentColumn + 1).setValue(appData.install_info.install_time).setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
    currentRow++;

    // Optional: Campaign Name, Fingerprint Attribution, View-Through Attribution
    if (appData.install_info.campaign_name) {
      sheet.getRange(currentRow, currentColumn).setValue("Tracking Link Name").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(currentRow, currentColumn + 1).setValue(appData.install_info.campaign_name).setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      currentRow++;
    }

    sheet.getRange(currentRow, currentColumn).setValue("Fingerprint Attribution").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(currentRow, currentColumn + 1).setValue(appData.install_info.fingerprint_attribution ? "True" : "False").setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
    currentRow++;

    sheet.getRange(currentRow, currentColumn).setValue("View-Through Attribution").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(currentRow, currentColumn + 1).setValue(appData.install_info.view_through_attribution ? "True" : "False").setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
    currentRow++;

    // Additional Parameters (only if available)
    const additionalParams = appData.install_info.additional_parameters;
    if (additionalParams) {
      let additionalParamsText = "";
      for (let paramKey in additionalParams) {
        additionalParamsText += `${paramKey}: ${additionalParams[paramKey]}\n`;
      }

      sheet.getRange(currentRow, currentColumn).setValue("Additional Parameters").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(currentRow, currentColumn + 1).setValue(additionalParamsText).setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID).setWrap(true);
      currentRow++;
    }


    // Re-engagement Information (only if it has a non-empty install_time)
    if (appData.re_engagement_info) {
      currentRow++;
      const reEngagementHeaderRange = sheet.getRange(currentRow, currentColumn, 1, 2);
      reEngagementHeaderRange.merge();
      reEngagementHeaderRange.setValue("Re-engagement Information")
                            .setBackground('#FFD700')
                            .setFontWeight('bold')
                            .setHorizontalAlignment('center')
                            .setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      currentRow++;

      sheet.getRange(currentRow, currentColumn).setValue("Re-engagement Source").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(currentRow, currentColumn + 1).setValue(appData.re_engagement_info.network).setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      currentRow++;

      sheet.getRange(currentRow, currentColumn).setValue("Re-engagement Campaign Name").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(currentRow, currentColumn + 1).setValue(appData.re_engagement_info.campaign_name).setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      currentRow++;

      sheet.getRange(currentRow, currentColumn).setValue("Fingerprint Attribution").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(currentRow, currentColumn + 1).setValue(appData.re_engagement_info.fingerprint_attribution ? "True" : "False").setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      currentRow++;

      sheet.getRange(currentRow, currentColumn).setValue("View-Through Attribution").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(currentRow, currentColumn + 1).setValue(appData.re_engagement_info.view_through_attribution ? "True" : "False").setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      currentRow++;

      sheet.getRange(currentRow, currentColumn).setValue("Re-engagement Additional Parameters").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      let reEngagementParamsText = "";
      const reEngagementParams = appData.re_engagement_info.additional_parameters;
      if (reEngagementParams) {
        for (let paramKey in reEngagementParams) {
          reEngagementParamsText += `${paramKey}: ${reEngagementParams[paramKey]}\n`;
        }
      }
      sheet.getRange(currentRow, currentColumn + 1).setValue(reEngagementParamsText).setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID).setWrap(true);
      currentRow++;

      sheet.getRange(currentRow, currentColumn).setValue("Re-engagement Install Time").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(currentRow, currentColumn + 1).setValue(appData.re_engagement_info.install_time).setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      currentRow++;
    }



    // Leave a row and add "Events" section if events are available
    if (appData.events && appData.events.length > 0) {
      currentRow++;
      const eventHeaderRange = sheet.getRange(currentRow, currentColumn, 1, 2);
      eventHeaderRange.merge();
      eventHeaderRange.setValue("Events")
                     .setBackground('#FFD700')
                     .setFontWeight('bold')
                     .setHorizontalAlignment('center')
                     .setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
      currentRow++;

      // Process events
      appData.events.forEach((event) => {
        sheet.getRange(currentRow, currentColumn).setValue("Event Name").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(currentRow, currentColumn + 1).setValue(event.event_name).setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
        currentRow++;

        sheet.getRange(currentRow, currentColumn).setValue("Event Count").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(currentRow, currentColumn + 1).setValue(event.event_count).setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
        currentRow++;

        sheet.getRange(currentRow, currentColumn).setValue("First Event Time").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(currentRow, currentColumn + 1).setValue(event.first_event_time).setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
        currentRow++;

        sheet.getRange(currentRow, currentColumn).setValue("Last Event Time").setBackground('#F0F8FF').setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(currentRow, currentColumn + 1).setValue(event.last_event_time).setBorder(true, true, true, true, null, null, "lightgrey", SpreadsheetApp.BorderStyle.SOLID);
        currentRow++;
      });
    }
  });
}



function clearDataOnSheetLoad() {
  // Get the active sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Specify the range to clear starting from row 10 onwards
  sheet.getRange('A10:Z').clear(); // This clears both content and formatting

  // // Clear content (but keep formatting) for cells E1, E3, and E5
  // sheet.getRange('D1').clearContent();
  // sheet.getRange('D3').clearContent();
  // sheet.getRange('D5').clearContent();
}


function onOpen() {
  clearDataOnSheetLoad();
}





