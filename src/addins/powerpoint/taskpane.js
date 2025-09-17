const btn = document.getElementById('btn');
const log = document.getElementById('log');
btn.addEventListener('click', () => {
  const t = new Date().toLocaleTimeString();
  log.textContent = 'Nano Mode clicked at ' + t;
});
