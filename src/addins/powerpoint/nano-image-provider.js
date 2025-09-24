(function () {
  function sizeFor(aspect){ if(aspect==='4:3')return{w:1024,h:768}; if(aspect==='1:1')return{w:1024,h:1024}; return{w:1280,h:720}; }
  window.NIS_generateImage = async ({ theme='', prompt='', seed=0, aspect='16:9' }) => {
    const { w, h } = sizeFor(aspect);
    const c = document.createElement('canvas'); c.width = w; c.height = h;
    const ctx = c.getContext('2d');
    ctx.fillStyle = '#223'; ctx.fillRect(0,0,w,h);
    ctx.fillStyle = '#4da3ff'; ctx.fillRect(0,0,w,h*0.55);
    ctx.fillStyle = '#fff'; ctx.font = 'bold 36px system-ui,Segoe UI,Arial';
    ctx.fillText('API IMAGE', 32, 56);
    ctx.font = '20px system-ui,Segoe UI,Arial';
    ctx.fillText('Theme: '+theme, 32, 100);
    ctx.fillText('Prompt: '+prompt, 32, 130);
    ctx.fillText('Seed: '+String(seed)+' | '+aspect, 32, 160);
    return { base64: c.toDataURL('image/png').split(',')[1] };
  };
})();
