Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
      document.getElementById("reportPhishingButton").onclick = reportPhishing;
      document.getElementById("reportJunkButton").onclick = reportJunk;
  }
});

function reportPhishing() {
  var item = Office.context.mailbox.item;
  createForwardEmail(item, "reporting@mxtest365.com");
}

function reportJunk() {
  var item = Office.context.mailbox.item;
  createForwardEmail(item, "reporting@mxtest365.com");
}

function createForwardEmail(item, reportAddress) {
  // Load the email item
  item.loadCustomPropertiesAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          var properties = result.value;
          var itemId = properties.get('itemId');

          // Create the body for the forwarded email
          var emailBody = {
              "subject": "Reported Email: " + item.subject,
              "body": {
                  "contentType": "HTML",
                  "content": "<p>This email was reported as phishing/junk.</p>"
              },
              "toRecipients": [
                  {
                      "emailAddress": {
                          "address": reportAddress
                      }
                  }
              ],
              "attachments": [
                  {
                      "@odata.type": "#Microsoft.OutlookServices.ItemAttachment",
                      "item": {
                          "subject": item.subject,
                          "body": {
                              "contentType": item.body.contentType,
                              "content": item.body.content
                          }
                      }
                  }
              ]
          };

          // Display the new email form
          Office.context.mailbox.displayNewMessageForm(emailBody);
      } else {
          console.error("Error loading item properties: " + result.error.message);
      }
  });
}
