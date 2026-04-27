Office.onReady(() => {
  Office.actions.associate("openTaskpane", () => {
    // The manifest binds this action to ShowTaskpane; the handler is present for
    // Office hosts that require a registered command function.
    return;
  });
});
