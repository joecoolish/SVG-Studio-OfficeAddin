/* global Office */

Office.onReady(() => {
  // Registered function commands would be initialized here.
});

/**
 * Placeholder action so the manifest's FunctionFile has a valid target.
 * @param event
 */
function action(event: Office.AddinCommands.Event): void {
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
