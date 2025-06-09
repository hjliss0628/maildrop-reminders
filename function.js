Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    // Only show if we're in compose
    if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
      Office.context.mailbox.item.notificationMessages.addAsync("maildropReminder", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Please add the matter's maildrop address or maildrop@polarislawyers.com as a recipient, CC, or BCC.",
        persistent: true
      });
    }
  }
});
