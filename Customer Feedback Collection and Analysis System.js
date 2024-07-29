function sendDataToNotion(customerEmail, rating, feedback, type, sentiment) {
  var notionApiKey = PropertiesService.getScriptProperties().getProperty('NOTION_API_KEY');
  var databaseId = PropertiesService.getScriptProperties().getProperty('DATABASE_ID');

  var url = 'https://api.notion.com/v1/pages';
  var headers = {
    'Authorization': 'Bearer ' + notionApiKey,
    'Content-Type': 'application/json',
    'Notion-Version': '2022-06-28' // Latest Notion API version
  };

  var payload = JSON.stringify({
    parent: { database_id: databaseId },
    properties: {
      'Customer Email': {
        title: [
          {
            text: {
              content: customerEmail
            }
          }
        ]
      },
      'Rating': {
        multi_select: [
          {
            name: rating.toString()
          }
        ]
      },
      'Feedback': {
        rich_text: [
          {
            text: {
              content: feedback
            }
          }
        ]
      },
      'Type': {
        rich_text: [
          {
            text: {
              content: type
            }
          }
        ]
      },
      'Sentiment': {
        rich_text: [
          {
            text: {
              content: sentiment
            }
          }
        ]
      }
    }
  });

  var options = {
    'method': 'post',
    'headers': headers,
    'payload': payload,
    'muteHttpExceptions': true // To get the full response for debugging
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText()); // Log the response for debugging
}

function categorizeFeedback() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var data = sheet.getDataRange().getValues();
  var existingEntries = new Set(); // To track existing entries

  for (var i = 1; i < data.length; i++) {
    var customerEmail = data[i][1]; // Assuming the email is in the second column
    var rating = data[i][2];        // Assuming the rating is in the third column
    var feedback = data[i][3];      // Assuming the feedback is in the fourth column
    var type = data[i][4];          // Assuming the type is in the fifth column
    var sentiment = analyzeSentiment(rating); // Determine sentiment based on rating

    // Create a unique key for each entry to avoid duplicates
    var entryKey = `${customerEmail}-${rating}-${feedback}-${type}-${sentiment}`;

    if (!existingEntries.has(entryKey)) {
      existingEntries.add(entryKey);

      if (feedback) {
        Logger.log(`Row ${i + 1}: Feedback: ${feedback}, Sentiment: ${sentiment}`);
        sheet.getRange(i + 1, 6).setValue(sentiment); // Set sentiment in the sheet

        // Send data to Notion
        sendDataToNotion(customerEmail, rating, feedback, type, sentiment);
      } else {
        Logger.log(`Row ${i + 1}: Feedback is undefined.`);
      }
    } else {
      Logger.log(`Row ${i + 1}: Duplicate entry skipped.`);
    }
  }
}

function analyzeSentiment(rating) {
  if (rating <= 2) {
    return 'Poor';
  } else if (rating === 3) {
    return 'Neutral';
  } else if (rating >= 4) {
    return 'Positive';
  } else {
    return 'Neutral';
  }
}

function generateReports() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var data = sheet.getDataRange().getValues();
  var positive = 0, neutral = 0, negative = 0;

  for (var i = 1; i < data.length; i++) {
    var sentiment = data[i][5]; // Assuming the sentiment is in the sixth column
    if (sentiment === 'Positive') {
      positive++;
    } else if (sentiment === 'Neutral') {
      neutral++;
    } else if (sentiment === 'Poor') {
      negative++;
    }
  }

  var reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Summary Report');
  if (!reportSheet) {
    reportSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Summary Report');
  } else {
    reportSheet.clear(); // Clear previous data
  }

  reportSheet.appendRow(['Category', 'Count']);
  reportSheet.appendRow(['Positive', positive]);
  reportSheet.appendRow(['Neutral', neutral]);
  reportSheet.appendRow(['Poor', negative]);

  var chart = reportSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(reportSheet.getRange('A1:B4'))
    .setPosition(5, 5, 0, 0)
    .build();

  reportSheet.insertChart(chart);
}

function sendFollowUpEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var email = data[i][1]; // Assuming the email is in the second column
    var sentiment = data[i][5]; // Assuming the sentiment is in the sixth column

    if (validateEmail(email)) {
      var subject, body;

      if (sentiment === 'Positive') {
        subject = 'Thank you for your positive feedback!';
        body = 'We appreciate your positive feedback and are glad you had a great experience!';
      } else if (sentiment === 'Poor') {
        subject = 'We are sorry for your experience';
        body = 'We apologize for any inconvenience caused. Our support team will reach out to you shortly.';
      } else {
        subject = 'Thank you for your feedback';
        body = 'We appreciate your feedback and are continuously working to improve our services.';
      }

      MailApp.sendEmail(email, subject, body);
    } else {
      Logger.log(`Invalid email: ${email}`);
    }
  }
}

function validateEmail(email) {
  var re = /\S+@\S+\.\S+/;
  return re.test(email);
}

function createTriggers() {
  ScriptApp.newTrigger('categorizeFeedback')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger('sendFollowUpEmails')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create();
}
