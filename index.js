function showReminderDialog(event) {
  Office.context.ui.displayDialogAsync(
    "https://hjliss0628.github.io/maildrop-reminders/index",
    { height: 40, width: 30 },
    function (result) {
      event.completed();
    }
  );
}
