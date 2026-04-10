/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item?.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);

/**
 * Open a modal dialog displaying the given label.
 */
function showDialog(label: string, event: Office.AddinCommands.Event) {
  const url = `https://localhost:3000/dialog.html?label=${encodeURIComponent(label)}`;
  Office.context.ui.displayDialogAsync(
    url,
    { height: 30, width: 30, displayInIframe: false },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
          dialog.close();
        });
      }
      event.completed();
    }
  );
}

function showTest1(event: Office.AddinCommands.Event) {
  showDialog("test1", event);
}

function showTest2(event: Office.AddinCommands.Event) {
  showDialog("test2", event);
}

function showTest3(event: Office.AddinCommands.Event) {
  showDialog("test3", event);
}

Office.actions.associate("showTest1", showTest1);
Office.actions.associate("showTest2", showTest2);
Office.actions.associate("showTest3", showTest3);
