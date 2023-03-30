/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import * as OfficeHelpers from "@microsoft/office-js-helpers";
const Office = require('@microsoft/office-js');

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

async function run() {
  try {
    await Word.run(async (context) => {
      // Get the custom placeholders from the input fields
      const propertyNumberPlaceholder = document.getElementById("propertyNumberPlaceholder").value;
      const ownerNamePlaceholder = document.getElementById("ownerNamePlaceholder").value;

      // Load the data from your data source (e.g., Excel, a JSON file, or an API)
      const data = [
        { propertyNumber: "123", ownerName: "John Doe" },
        { propertyNumber: "456", ownerName: "Jane Smith" },
      ];

      // Perform mail merge and save each iteration with a different filename
      for (const record of data) {
        // Replace the placeholders in the document with the data from the current record
        const body = context.document.body;
        body.replaceText(propertyNumberPlaceholder, record.propertyNumber);
        body.replaceText(ownerNamePlaceholder, record.ownerName);

        // Save the document as a new file with a filename based on the data
        const fileName = `Property-${record.propertyNumber}-${record.ownerName}.docx`;
        const fileFormat = Word.FileFormat.docx;
        const base64File = await context.document.saveAsBase64(fileFormat);

        // Use the OfficeHelpers.Utilities.downloadFile helper function to download the file
        OfficeHelpers.Utilities.downloadFile(base64File, fileName, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");

        // Revert the changes in the document for the next iteration
        body.replaceText(record.propertyNumber, propertyNumberPlaceholder);
        body.replaceText(record.ownerName, ownerNamePlaceholder);
      }

      await context.sync();
    });
  } catch (error) {
    OfficeHelpers.UI.notify(error);
    OfficeHelpers.Utilities.log(error);
  }
}
