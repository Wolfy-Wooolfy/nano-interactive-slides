import * as sim from '../../../packages/simulation-engine/index.js'

async function loadScene(){
  const res = await fetch('/examples/supply-chain/scene.json')
  return res.json()
}

function render(state){
  const t = document.getElementById('time')
  const v = document.getElementById('view')
  if(!t||!v) return
  t.textContent = 't = ' + state.t + 's'
  const nodesRows = state.nodes.map(n=>(
    '<tr><td>'+n.id+'</td><td>'+n.type+'</td><td>'+n.stock+'</td><td>'+ (n.rate??'') +'</td><td>'+ (n.capacity??'∞') +'</td></tr>'
  )).join('')
  const edgesRows = state.edges.map(e=>(
    '<tr><td>'+e.from+' → '+e.to+'</td><td>'+e.inTransit+'</td></tr>'
  )).join('')
  v.innerHTML =
    '<h3>Nodes</h3>'+
    '<table><thead><tr><th>ID</th><th>Type</th><th>Stock</th><th>Rate</th><th>Capacity</th></tr></thead><tbody>'+nodesRows+'</tbody></table>'+
    '<h3>Edges</h3>'+
    '<table><thead><tr><th>Link</th><th>In Transit</th></tr></thead><tbody>'+edgesRows+'</tbody></table>'
}

Office.onReady(async ()=>{
  const scene = await loadScene()
  sim.init(scene)
  render(sim.getState())

  const sBtn = document.getElementById('start')
  const pBtn = document.getElementById('stop')
  if(sBtn) sBtn.onclick = ()=> sim.start(render)
  if(pBtn) pBtn.onclick = ()=> sim.stop()

  const nano = document.getElementById('nano-btn')
  if(nano) nano.onclick = ()=>{
    Office.context.document.setSelectedDataAsync(
      'NIS: Nano Mode started ✓',
      { coercionType: Office.CoercionType.Text },
      ()=>{}
    )
  }
})
