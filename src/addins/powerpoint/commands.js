/* global Office */

Office.onReady(() => {})

function onNanoMode(event) {
  Office.context.document.setSelectedDataAsync(
    'NIS: Nano Mode started ✓',
    { coercionType: Office.CoercionType.Text },
    () => { event.completed() }
  )
}

if (typeof window !== 'undefined') {
  window.onNanoMode = onNanoMode
}
