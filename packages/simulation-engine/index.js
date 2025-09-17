let S = null
let timer = null

function clone(x){ return JSON.parse(JSON.stringify(x)) }

export function init(scene){
  const sc = clone(scene)
  const nodes = {}
  sc.nodes.forEach(n=>{
    nodes[n.id] = {
      id:n.id, type:n.type,
      rate: n.rate||0,
      capacity: typeof n.capacity==='number'? n.capacity: Infinity,
      stock: 0
    }
  })
  const initial = sc.params && sc.params.initialStock ? sc.params.initialStock : {}
  Object.keys(initial).forEach(k=>{
    if(nodes[k]) nodes[k].stock = initial[k]
  })
  const edges = sc.edges.map(e=>({
    from:e.from, to:e.to,
    delay: Math.max(0, Number(e.delay)||0),
    q:[]
  }))
  S = {
    tickMs: sc.params && sc.params.tickMs ? sc.params.tickMs : 100,
    now: 0,
    nodes,
    edges
  }
}

function step(){
  const dt = S.tickMs/1000
  Object.values(S.nodes).forEach(n=>{
    if(n.type==='producer'){
      n.stock += n.rate*dt
      if(n.stock>n.capacity) n.stock=n.capacity
    }
  })
  S.edges.forEach(e=>{
    const src = S.nodes[e.from]
    if(!src) return
    const outRate = src.rate>0? src.rate: Infinity
    const send = Math.min(src.stock, isFinite(outRate)? outRate*dt: src.stock)
    if(send>0){
      src.stock -= send
      e.q.push({eta:S.now+e.delay, qty:send})
    }
  })
  S.edges.forEach(e=>{
    const dst = S.nodes[e.to]
    if(!dst) return
    const arrived = []
    for(let i=0;i<e.q.length;i++){
      if(e.q[i].eta<=S.now) arrived.push(i)
    }
    let add = 0
    arrived.reverse().forEach(idx=>{
      add += e.q[idx].qty
      e.q.splice(idx,1)
    })
    if(add>0){
      dst.stock = Math.min(dst.capacity, dst.stock+add)
    }
  })
  Object.values(S.nodes).forEach(n=>{
    if(n.type==='consumer'){
      const need = n.rate*dt
      const use = Math.min(need, n.stock)
      n.stock -= use
    }
  })
  S.now += dt
}

export function start(onUpdate){
  if(!S) return
  if(timer) clearInterval(timer)
  timer = setInterval(()=>{
    step()
    if(onUpdate) onUpdate(getState())
  }, S.tickMs)
}

export function stop(){
  if(timer){ clearInterval(timer); timer=null }
}

export function getState(){
  const nodes = Object.values(S.nodes).map(n=>({
    id:n.id, type:n.type, stock: Number(n.stock.toFixed(2)),
    rate:n.rate, capacity: isFinite(n.capacity)? n.capacity: null
  }))
  const edges = S.edges.map(e=>({
    from:e.from, to:e.to, inTransit: Number(e.q.reduce((a,b)=>a+b.qty,0).toFixed(2))
  }))
  return { t: Number(S.now.toFixed(2)), nodes, edges }
}
