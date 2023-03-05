/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

const { GIPHY_API_KEY } = process.env;

let isOfficeInitialized = false;

Office.onReady(() => {
  isOfficeInitialized = true;
});

/**
 * Fetchs a random GIF and inserts into the message body
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function insertRandomGIF(event) {
  if (!isOfficeInitialized) return;

  let message;

  fetch(`https://api.giphy.com/v1/gifs/random?api_key=${GIPHY_API_KEY}&tag=&rating=g`)
    .then((response) => response.json())
    .then(({ data, meta }) => {
      if (meta.status !== 200) {
        message = {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: "An error occurred while inserting the GIF. Please try again later.",
          icon: "Icon.80x80",
          persistent: true,
        };
      } else {
        Office.context.mailbox.item.body.setSelectedDataAsync(
          `<div><img src="${data.images.original.url}" /></div>`,
          { coercionType: Office.CoercionType.Html },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
              console.error(result.error.message);
              message = {
                type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                message: result.error.message,
                icon: "Icon.80x80",
                persistent: true,
              };
            } else {
              message = {
                type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
                message: "GIF inserted successfully.",
                icon: "Icon.80x80",
                persistent: true,
              };
            }
          }
        );
      }
    })
    .catch((err) => {
      console.error(err);
      message = {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: err.message || "An error occurred while inserting the GIF. Please try again later.",
        icon: "Icon.80x80",
        persistent: true,
      };
    })
    .finally(() => {
      // Show a notification message
      Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

      event.completed();
    });
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
g.insertRandomGIF = insertRandomGIF;
