/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

Office.initialize = function () {
  Office.context.onse;
};

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

async function abbrv(event) {
  try {
    Word.run(async (context) => {
      var sentences = context.document
        .getSelection()
        .getTextRanges([" "] /* Using the "." as delimiter */, false /*means without trimming spaces*/);
      context.load(sentences);
      return context.sync().then(function () {
        //  expands the range to the end of the paragraph to get all the complete sentences.
        var sentecesToTheEndOfParagraph = sentences.items[0]
          .getRange()
          .expandTo(
            context.document
              .getSelection()
              .paragraphs.getFirst()
              .getRange("end") /* Expanding the range all the way to the end of the paragraph */
          )
          .getTextRanges([" "], false);
        context.load(sentecesToTheEndOfParagraph);
        return context.sync().then(function () {
          var word = sentecesToTheEndOfParagraph.items[0].text;
          console.log(word);
          // document.getElementById("app-body").innerText = word;
        });
      });
    });
  } catch (error) {
    console.error(error);
  }
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;
