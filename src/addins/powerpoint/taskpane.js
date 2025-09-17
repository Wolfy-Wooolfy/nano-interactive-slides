Office.onReady(() => {
  const btn = document.getElementById('nano-btn')
  if (!btn) return
  btn.addEventListener('click', () => {
    Office.context.document.setSelectedDataAsync(
      'NIS: Nano Mode started ✓',
      { coercionType: Office.CoercionType.Text },
      () => {}
    )
  })
})
