// ================================================================
// CHEADLE JUNIORS — COMPETITION SIGNUP — GOOGLE APPS SCRIPT
// ================================================================
// SETUP INSTRUCTIONS:
//
// 1. Open your Google Sheet (leaguedata)
// 2. Click Extensions → Apps Script
// 3. Delete any existing code in the editor
// 4. Paste ALL of this code in
// 5. Click Save (floppy disk icon)
// 6. Click Deploy → New Deployment
// 7. Click the gear icon next to "Select type" → choose Web App
// 8. Set these options:
//      Description:    Cheadle Juniors Signup
//      Execute as:     Me (your Google account)
//      Who has access: Anyone
// 9. Click Deploy
// 10. Copy the Web App URL it gives you
// 11. Paste that URL into signup.html replacing YOUR_APPS_SCRIPT_URL_HERE
//
// That's it! Every signup will now appear in the Signups sheet.
// ================================================================

function doPost(e) {
  try {
    // Parse the incoming JSON payload
    var data = JSON.parse(e.postData.contents);

    // Get or create the Signups sheet
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Signups');

    if (!sheet) {
      sheet = ss.insertSheet('Signups');

      // Write header row on first use
      sheet.appendRow([
        'Timestamp',
        'Competition',
        'Date',
        'Time',
        'Venue',
        'Series',
        'Age Group',
        'Player Name',
        'Player First Name',
        'Player Last Name',
        'Parent / Guardian',
        'Relationship',
        'Email',
        'Phone',
        'Notes'
      ]);

      // Style the header row
      var headerRange = sheet.getRange(1, 1, 1, 15);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#00c896');
      headerRange.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    // Append the signup row
    sheet.appendRow([
      data.timestamp    || new Date().toLocaleString(),
      data.competition  || '',
      data.date         || '',
      data.time         || '',
      data.venue        || '',
      data.series       || '',
      data.ageGroup     || '',
      data.playerName   || '',
      data.playerFirst  || '',
      data.playerLast   || '',
      data.parentName   || '',
      data.relationship || '',
      data.email        || '',
      data.phone        || '',
      data.notes        || ''
    ]);

    // Optional: send a notification email to the club
    // Uncomment the lines below and replace the email address
    // to get an email every time someone signs up
    //
    // MailApp.sendEmail({
    //   to:      'yourname@cheadlegolf.com',
    //   subject: '⛳ New Signup: ' + data.playerName + ' — ' + data.competition,
    //   body:    'New competition signup received:\n\n'
    //          + 'Player:      ' + data.playerName   + '\n'
    //          + 'Age Group:   ' + data.ageGroup      + '\n'
    //          + 'Competition: ' + data.competition   + '\n'
    //          + 'Date:        ' + data.date          + '\n'
    //          + 'Venue:       ' + data.venue         + '\n'
    //          + 'Parent:      ' + data.parentName    + '\n'
    //          + 'Email:       ' + data.email         + '\n'
    //          + 'Phone:       ' + data.phone         + '\n'
    //          + 'Notes:       ' + data.notes         + '\n'
    // });

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Test function — run this manually in the Apps Script editor
// to check the sheet is being created and written to correctly
function testSignup() {
  var testData = {
    postData: {
      contents: JSON.stringify({
        timestamp:    '21/05/2026, 10:00:00',
        competition:  'Summer Series – June 14th',
        date:         'June 14th 2026',
        time:         '13:00–14:00',
        venue:        'Cheadle Golf Club',
        series:       'Summer Series',
        ageGroup:     '9s',
        playerName:   'Jamie Smith',
        playerFirst:  'Jamie',
        playerLast:   'Smith',
        parentName:   'Sarah Smith',
        relationship: 'Parent',
        email:        'sarah@test.com',
        phone:        '07700 900000',
        notes:        'Test entry'
      })
    }
  };
  doPost(testData);
  Logger.log('Test signup written to sheet.');
}
