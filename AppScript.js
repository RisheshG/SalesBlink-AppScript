const API_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJfaWQiOiI2NjUwOWYwZmRhNWI4ODk0MGQzMDViNTUiLCJhY2NvdW50SWQiOiIxNjQ3NWI5NS04NzE5LTRlM2EtOWIwOC02N2RmNjZiNzAyNWUiLCJ1c2VySWQiOiJhMGJlOWIyOS01NzA4LTQ2ODktOGFmOS05MjQ0ODcwMWUyYzMiLCJ1c2VyRW1haWwiOiJidXNpbmVzc0B4ZW1haWx2ZXJpZnkuY29tIiwicGVybWlzc2lvbiI6Im93bmVyIiwicHJvdmlkZXIiOiJhcHAiLCJpYXQiOjE3MjAxOTU4MDAsImV4cCI6MTcyNTM3OTgwMH0.ntxik1KJeGf58RF6dB-Q75gjDLBJO7V47lugDQ2-V3w';

function fetchEmailAddresses() {
  const baseUrl = 'https://run.salesblink.io/api/senders';
  const limit = 50;
  let allEmails = [];
  let skip = 0;
  let hasMore = true;

  while (hasMore) {
    const url = `${baseUrl}?limit=${limit}&skip=${skip}&search=`;
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${API_KEY}`
      },
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      const data = JSON.parse(response.getContentText());

      if (response.getResponseCode() === 200 && data.success) {
        const emails = data.data.flatMap(entry => [
          entry.google_email,
          entry.microsoft_email,
          entry.outlook_mail,
          entry.smtpUsername
        ].filter(email => email)); // Filter out any null or undefined values

        allEmails = allEmails.concat(emails);

        if (data.data.length < limit) {
          hasMore = false;
        } else {
          skip += 1; // Increment the skip value correctly for next page
        }
      } else {
        Logger.log(`Failed to fetch email addresses. Response code: ${response.getResponseCode()}, Response content: ${response.getContentText()}`);
        hasMore = false;
      }
    } catch (err) {
      Logger.log(`Error fetching email addresses: ${err}`);
      hasMore = false;
    }
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('EmailAddresses') || spreadsheet.getActiveSheet();

  // Rename sheet to 'EmailAddresses'
  if (sheet.getName() !== 'EmailAddresses') {
    sheet.setName('EmailAddresses');
  }

  // Clear any existing data validation rules
  sheet.getRange('A2').clearDataValidations();
  
  // Set the dropdown menu in cell A2
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(allEmails).build();
  sheet.getRange('A2').setDataValidation(rule);

  // Update cell A1
  sheet.getRange('A1').setValue('Select Email Address');

  // Apply font style, center alignment, and cell styling to the sheet
  const range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  range.setFontFamily('Roboto');
  range.setFontSize(12);
  range.setHorizontalAlignment('center');

  // Style A1
  const cellA1 = sheet.getRange('A1');
  cellA1.setBackground('black');
  cellA1.setFontColor('white');
}

function fetchWarmupDetailsForSelectedEmail() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('EmailAddresses') || spreadsheet.getActiveSheet();
  const selectedEmail = sheet.getRange('A2').getValue();

  if (!selectedEmail) {
    SpreadsheetApp.getUi().alert('Please select an email address.');
    return;
  }

  const url = `https://run.salesblink.io/api/senders?limit=50&skip=0&search=${selectedEmail}`;
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${API_KEY}`
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    if (response.getResponseCode() === 200 && data.success && data.data.length > 0) {
      const user = data.data[0];
      const user_id = user.id;  // Use the 'id' field

      // Calculate the start and end times for the last 7 days
      const end = new Date().getTime();
      const start = end - (7 * 24 * 60 * 60 * 1000);

      const warmupUrl = `https://run.salesblink.io/api/warmuplog/${user_id}/chart?type=daily&start=${start}&end=${end}`;

      // Log the endpoint to the sheet for verification
      const warmupSheet = spreadsheet.getSheetByName('WarmupDetails') || spreadsheet.insertSheet('WarmupDetails');
      warmupSheet.getRange('C1').setValue(`API Endpoint: ${warmupUrl}`);

      const warmupResponse = UrlFetchApp.fetch(warmupUrl, options);
      const warmupData = JSON.parse(warmupResponse.getContentText());

      if (warmupResponse.getResponseCode() === 200 && warmupData.success) {
        displayWarmupDetails(warmupData);
      } else {
        Logger.log(`Failed to fetch warmup details. Response code: ${warmupResponse.getResponseCode()}, Response content: ${warmupResponse.getContentText()}`);
      }
    } else {
      Logger.log(`Failed to fetch user data. Response code: ${response.getResponseCode()}, Response content: ${response.getContentText()}`);
    }
  } catch (err) {
    Logger.log(`Error fetching warmup details for selected email: ${err}`);
  }
}

function displayWarmupDetails(data) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('WarmupDetails') || spreadsheet.insertSheet('WarmupDetails');
  sheet.clear();

  // Clear existing charts
  clearCharts(sheet);

  // Populate the sheet with data
  sheet.getRange(1, 1).setValue('Date');
  sheet.getRange(1, 2).setValue('Sent');
  sheet.getRange(1, 3).setValue('Received');
  sheet.getRange(1, 4).setValue('Replied');
  sheet.getRange(1, 5).setValue('Spam');

  const entries = data.data.labels.map((label, index) => {
    return [
      label,
      data.data.datasets[0].data[index],  // Sent
      data.data.datasets[1].data[index],  // Received
      data.data.datasets[2].data[index],  // Replied
      data.data.datasets[3].data[index]   // Spam
    ];
  });

  entries.forEach((entry, index) => {
    sheet.getRange(index + 2, 1, 1, 5).setValues([entry]);
  });

  // Add chart for visualization
  createChart(sheet, data.data.labels.length);

  // Apply font style, center alignment, and cell styling to the sheet
  const range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  range.setFontFamily('Roboto');
  range.setFontSize(12);
  range.setHorizontalAlignment('center');

  // Style header cells
  const headerCells = sheet.getRange('B1:E1');
  headerCells.setBackground('black');
  headerCells.setFontColor('white');

  // Style A1
  const cellA1 = sheet.getRange('A1');
  cellA1.setBackground('black');
  cellA1.setFontColor('white');
}

function createChart(sheet, dataLength) {
  const range = sheet.getRange(1, 1, dataLength + 1, 5); // Adjust the range to include headers and data
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(range)
    .setPosition(1, 7, 0, 0)
    .build();
  sheet.insertChart(chart);
}

function clearCharts(sheet) {
  const charts = sheet.getCharts();
  charts.forEach(chart => sheet.removeChart(chart));
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Custom Menu')
    .addItem('Fetch Email Addresses', 'fetchEmailAddresses')
    .addItem('Fetch Warmup Details', 'fetchWarmupDetailsForSelectedEmail')
    .addToUi();
}
