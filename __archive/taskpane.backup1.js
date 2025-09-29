/* ================== Core simulation engine ================== */
function clamp(n){return Number.isFinite(n)?n:0}
function computeNext(state,dt,params){
  var s=JSON.parse(JSON.stringify(state));
  var rateFactory=clamp(params.factoryRate||0);
  var rateStore=clamp(params.storeRate||0);
  var capacityWarehouse=clamp(params.warehouseCapacity||0);
  var delayFW=clamp(params.delayFactoryToWarehouse||0);
  var delayWS=clamp(params.delayWarehouseToStore||0);
  var tick=dt/1000;
  var produce=rateFactory*tick;
  var consume=rateStore*tick;
  var canStore=Math.max(0,capacityWarehouse-s.stock.warehouse);
  var toWarehouse=Math.min(produce,canStore);
  var pipeFW=Math.max(0,(toWarehouse-delayFW>0)?toWarehouse:0);
  s.stock.warehouse+=pipeFW;
  var fromWarehouse=Math.min(consume,s.stock.warehouse);
  var pipeWS=Math.max(0,(fromWarehouse-delayWS>0)?fromWarehouse:0);
  s.stock.warehouse-=pipeWS;
  var delivered=Math.min(pipeWS,consume);
  s.stock.store+=delivered;
  return s;
}
function createEngine(initialState,initialParams){
  var state=JSON.parse(JSON.stringify(initialState||{stock:{factory:0,warehouse:10,store:0}}));
  var params=Object.assign({
    tickMs:100,factoryRate:5,warehouseCapacity:50,storeRate:4,
    delayFactoryToWarehouse:2,delayWarehouseToStore:1,speed:1
  },initialParams||{});
  var timer=null; var lastTs=0;
  var listeners=new Set();
  function emit(){ listeners.forEach(function(fn){ fn(getState()); }); }
  function loopTick(){
    var now=performance.now();
    var dt=(now-lastTs)*(params.speed||1);
    lastTs=now;
    state=computeNext(state,dt,params);
    emit();
  }
  function start(){ if(timer) return; lastTs=performance.now(); timer=setInterval(loopTick,Math.max(10,params.tickMs||100)); }
  function stop(){ if(timer){ clearInterval(timer); timer=null; } }
  function step(dt){ var d=Math.max(1,dt||params.tickMs||100); state=computeNext(state,d,params); emit(); }
  function setParams(next){
    params=Object.assign({},params,next||{});
    if(timer){ clearInterval(timer); timer=setInterval(loopTick,Math.max(10,params.tickMs||100)); }
  }
  function onUpdate(fn){ listeners.add(fn); return function(){listeners.delete(fn)} }
  function getState(){ return JSON.parse(JSON.stringify({state:state,params:params})) }
  function findNode(nodes,type){
    for(var i=0;i<(nodes||[]).length;i++){ if((nodes[i]||{}).type===type) return nodes[i]; }
    return null;
  }
  function findEdge(edges,from,to){
    for(var i=0;i<(edges||[]).length;i++){ var e=edges[i]||{}; if(e.from===from && e.to===to) return e; }
    return null;
  }
  function loadScenario(scn){
    if(!scn) return;
    var p=scn.params||{};
    var st=(p.initialStock || scn.initialStock || {});
    state={stock:{
      factory:clamp(st.factory||0),
      warehouse:clamp(st.warehouse||10),
      store:clamp(st.store||0)
    }};
    var prod=findNode(scn.nodes,"producer")||{};
    var buf =findNode(scn.nodes,"buffer")  ||{};
    var cons=findNode(scn.nodes,"consumer")||{};
    var eFW =findEdge(scn.edges,"factory","warehouse")||{};
    var eWS =findEdge(scn.edges,"warehouse","store")  ||{};
    params=Object.assign({},params,{
      tickMs:clamp(p.tickMs||params.tickMs),
      factoryRate:clamp(prod.rate||params.factoryRate),
      warehouseCapacity:clamp(buf.capacity||params.warehouseCapacity),
      storeRate:clamp(cons.rate||params.storeRate),
      delayFactoryToWarehouse:clamp(eFW.delay||params.delayFactoryToWarehouse),
      delayWarehouseToStore:clamp(eWS.delay||params.delayWarehouseToStore)
    });
    emit();
  }
  return {start:start,stop:stop,step:step,setParams:setParams,onUpdate:onUpdate,getState:getState,loadScenario:loadScenario}
}

/* ================== UI helpers ================== */
function el(tag,attrs,children){
  var e=document.createElement(tag);
  if(attrs){
    Object.keys(attrs).forEach(function(k){
      if(k==="style" && typeof attrs[k]==="object"){ Object.assign(e.style,attrs[k]); }
      else if(k==="class"){ e.className=attrs[k]; }
      else if(k.indexOf("on")===0 && typeof attrs[k]==="function"){ e.addEventListener(k.slice(2).toLowerCase(),attrs[k]); }
      else{ e.setAttribute(k,attrs[k]); }
    });
  }
  (children||[]).forEach(function(c){
    if(typeof c==="string") e.appendChild(document.createTextNode(c));
    else if(c) e.appendChild(c);
  });
  return e;
}
var q=function(s){ return document.querySelector(s); };

/* ================== Scenarios & Controls ================== */
var SETTINGS_KEY="nis.slides";
function defaultScenario(){
  return {
    nodes:[
      {id:"factory",type:"producer",rate:5},
      {id:"warehouse",type:"buffer",capacity:50},
      {id:"store",type:"consumer",rate:4}
    ],
    edges:[
      {from:"factory",to:"warehouse",delay:2},
      {from:"warehouse",to:"store",delay:1}
    ],
    params:{tickMs:100,initialStock:{factory:0,warehouse:10,store:0}}
  };
}
function scenarioFromInputs(){
  return {
    nodes:[
      {id:"factory",type:"producer",rate:Number(q("#inp-factory").value)},
      {id:"warehouse",type:"buffer",capacity:Number(q("#inp-warehouse").value)},
      {id:"store",type:"consumer",rate:Number(q("#inp-store").value)}
    ],
    edges:[
      {from:"factory",to:"warehouse",delay:Number(q("#inp-dfw").value)},
      {from:"warehouse",to:"store",delay:Number(q("#inp-dws").value)}
    ],
    params:{tickMs:Number(q("#inp-tick").value),initialStock:{factory:0,warehouse:10,store:0}}
  };
}
function applyInputs(engine){
  engine.setParams({
    tickMs:Number(q("#inp-tick").value),
    speed:Number(q("#inp-speed").value),
    factoryRate:Number(q("#inp-factory").value),
    warehouseCapacity:Number(q("#inp-warehouse").value),
    storeRate:Number(q("#inp-store").value),
    delayFactoryToWarehouse:Number(q("#inp-dfw").value),
    delayWarehouseToStore:Number(q("#inp-dws").value)
  });
}
function ui(root){
  root.innerHTML="";
  var h=el("h2",null,["Nano Interactive Slides"]);
  var row=el("div",{class:"row"},[]);
  var bStart=el("button",{id:"btn-start"},["Start"]);
  var bStop =el("button",{id:"btn-stop"},["Stop"]);
  var bNano =el("button",{id:"btn-nano"},["Nano Mode"]);
  var bReset=el("button",{id:"btn-reset"},["Reset"]);
  row.append(bStart,bStop,bNano,bReset);

  var panel=el("div",{class:"panel"},[el("h3",null,["Scenario Controls"])]);
  function ctrl(label,id,def,step,min){
    var w=el("div",{class:"ctrl"},[]);
    w.append(el("label",null,[label]), el("input",{id:id,type:"number",step:String(step||1),value:String(def),min:String(min||0)},[]));
    return w;
  }
  panel.append(
    ctrl("Tick (ms)","inp-tick",100,10,10),
    ctrl("Speed x","inp-speed",1,0.1,0.1),
    ctrl("Factory Rate","inp-factory",5,0.1,0),
    ctrl("Warehouse Capacity","inp-warehouse",50,1,0),
    ctrl("Store Rate","inp-store",4,0.1,0),
    ctrl("Delay F→W","inp-dfw",2,0.1,0),
    ctrl("Delay W→S","inp-dws",1,0.1,0)
  );
  var row2=el("div",{class:"row"},[]);
  var bLoad=el("button",{id:"btn-load"},["Load"]);
  var bSave=el("button",{id:"btn-save"},["Save"]);
  row2.append(bLoad,bSave); panel.append(row2);

  var live=el("div",{class:"panel"},[el("h3",null,["Live State"]), el("pre",{id:"state-box"},[""])]);
  var nano=el("div",{id:"nano-panel",class:"panel",style:{display:"none"}},[
    el("h3",null,["Nano Reality Layers"]),
    el("div",{id:"nano-stage"},[])
  ]);

  root.append(h,row,panel,live,nano);
}

/* ================== Nano parallax ================== */
var rafId=null;
function startParallax(){
  var st=q("#nano-stage");
  st.innerHTML="";
  var a=el("div"), b=el("div");
  [a,b].forEach(function(x){
    x.style.position="absolute"; x.style.top="0"; x.style.left="0"; x.style.right="0"; x.style.bottom="0"; st.appendChild(x);
  });
  a.style.background="linear-gradient(90deg,#f0f0f0,#fafafa)";
  b.style.backgroundImage="radial-gradient(#999 2px, transparent 2px)";
  b.style.backgroundSize="20px 20px";
  st.onmousemove=function(e){
    var r=st.getBoundingClientRect();
    var mx=(e.clientX-r.left)/r.width-0.5;
    var my=(e.clientY-r.top)/r.height-0.5;
    a.style.transform="translate("+(mx*6)+"px,"+(my*4)+"px)";
    b.style.transform="translate("+(mx*12)+"px,"+(my*8)+"px)";
  };
  var loop=function(){ rafId=requestAnimationFrame(loop); }; loop();
}
function stopParallax(){
  var st=q("#nano-stage");
  if(st) st.onmousemove=null;
  if(rafId){ cancelAnimationFrame(rafId); rafId=null; }
  if(st) st.innerHTML="";
}

/* ================== Persistence (document.settings) ================== */
function debounce(fn,ms){ var t; return function(){ var args=arguments; clearTimeout(t); t=setTimeout(function(){ fn.apply(null,args); },ms); }; }
function loadSlidesSettings(){
  try{ return Office.context.document.settings.get(SETTINGS_KEY) || {}; }
  catch(e){ return {}; }
}
function saveSlidesSettings(obj,cb){
  try{
    Office.context.document.settings.set(SETTINGS_KEY,obj);
    Office.context.document.settings.saveAsync(function(){ if(cb) cb(); });
  }catch(e){ if(cb) cb(); }
}
function scenarioFromEngine(engine){
  var gs=engine.getState(); var p=gs.params||{};
  function num(v,d){ v=Number(v); return Number.isFinite(v)?v:d; }
  return {
    nodes:[
      {id:"factory",type:"producer",rate:num(p.factoryRate,5)},
      {id:"warehouse",type:"buffer",capacity:num(p.warehouseCapacity,50)},
      {id:"store",type:"consumer",rate:num(p.storeRate,4)}
    ],
    edges:[
      {from:"factory",to:"warehouse",delay:num(p.delayFactoryToWarehouse,2)},
      {from:"warehouse",to:"store",delay:num(p.delayWarehouseToStore,1)}
    ],
    params:{
      tickMs:num(p.tickMs,100),
      initialStock:{
        factory:clamp(((gs.state||{}).stock||{}).factory||0),
        warehouse:clamp(((gs.state||{}).stock||{}).warehouse||10),
        store:clamp(((gs.state||{}).stock||{}).store||0)
      }
    }
  };
}

/* ================== Slide engines switching ================== */
var slideEngines=new Map();
var currentSlideId=null;
var persistDebounced=debounce(function(slideId){
  if(!slideId) return;
  var eng=slideEngines.get(slideId); if(!eng) return;
  var slides=loadSlidesSettings();
  slides[slideId]=scenarioFromEngine(eng);
  saveSlidesSettings(slides);
},400);

function bindEngineToUI(engine){
  engine.onUpdate(function(payload){
    var state=payload.state, params=payload.params;
    var box=q("#state-box");
    if(box) box.textContent=JSON.stringify({stock:state.stock,params:params},null,2);
    persistDebounced(currentSlideId);
  });

  ["inp-tick","inp-speed","inp-factory","inp-warehouse","inp-store","inp-dfw","inp-dws"].forEach(function(id){
    var e=q("#"+id);
    if(e){ e.oninput=function(){ applyInputs(engine); persistDebounced(currentSlideId); }; }
  });

  var bStart=q("#btn-start"),bStop=q("#btn-stop"),bNano=q("#btn-nano"),
      bSave=q("#btn-save"), bLoad=q("#btn-load"), bReset=q("#btn-reset");

  if(bStart) bStart.onclick=function(){ engine.start(); };
  if(bStop)  bStop.onclick =function(){ engine.stop();  };
  if(bNano)  bNano.onclick =function(){
    var p=q("#nano-panel");
    var show = p && p.style.display!=="block";
    if(show){ p.style.display="block"; startParallax(); }
    else    { p.style.display="none";  stopParallax(); }
  };
  if(bSave)  bSave.onclick=function(){ persistDebounced(currentSlideId); };
  if(bLoad)  bLoad.onclick=function(){
    var slides=loadSlidesSettings();
    var scn=slides[currentSlideId];
    if(!scn) return;
    engine.loadScenario(scn);
    var getNode=function(type){ var arr=scn.nodes||[]; for(var i=0;i<arr.length;i++){ if(arr[i].type===type) return arr[i]; } return {}; };
    var getEdge=function(from,to){ var arr=scn.edges||[]; for(var i=0;i<arr.length;i++){ var e=arr[i]; if(e.from===from&&e.to===to) return e; } return {}; };
    var setVal=function(sel,val){ var elx=q(sel); if(elx) elx.value=String(val); };
    setVal("#inp-tick",(scn.params||{}).tickMs||100);
    setVal("#inp-factory",getNode("producer").rate || 5);
    setVal("#inp-warehouse",getNode("buffer").capacity || 50);
    setVal("#inp-store",getNode("consumer").rate || 4);
    setVal("#inp-dfw",getEdge("factory","warehouse").delay || 2);
    setVal("#inp-dws",getEdge("warehouse","store").delay || 1);
    applyInputs(engine);
    persistDebounced(currentSlideId);
  };
  if(bReset) bReset.onclick=function(){
    engine.loadScenario(defaultScenario());
    ["#inp-tick","#inp-speed","#inp-factory","#inp-warehouse","#inp-store","#inp-dfw","#inp-dws"]
      .forEach(function(sel,i){ var v=[100,1,5,50,4,2,1][i]; var elx=q(sel); if(elx) elx.value=String(v); });
    applyInputs(engine);
    persistDebounced(currentSlideId);
  };
}

function scnSafe(x){ try{ if(!x) return null; return JSON.parse(JSON.stringify(x)); } catch(e){ return null; } }
function getOrCreateEngineForSlide(slideId){
  if(!slideEngines.has(slideId)){
    var eng=createEngine({stock:{factory:0,warehouse:10,store:0}},{});
    var slides=loadSlidesSettings();
    var saved=scnSafe(slides[slideId]);
    if(saved) eng.loadScenario(saved); else eng.loadScenario(defaultScenario());
    slideEngines.set(slideId,eng);
  }
  return slideEngines.get(slideId);
}

function switchToSelectedSlideEngine(){
  if(!(window.Office && window.PowerPoint)) return;
  return PowerPoint.run(function(ctx){
    var slides=ctx.presentation.getSelectedSlides();
    slides.load("items");
    return ctx.sync().then(function(){
      if(!slides.items || slides.items.length===0) return;
      var sid=slides.items[0].id;
      if(currentSlideId===sid) return;
      var prev=currentSlideId?slideEngines.get(currentSlideId):null;
      if(prev) prev.stop();
      currentSlideId=sid;
      var eng=getOrCreateEngineForSlide(sid);
      var p=(eng.getState().params)||{};
      var setVal=function(sel,val){ var elx=q(sel); if(elx) elx.value=String(val); };
      setVal("#inp-tick",p.tickMs||100);
      setVal("#inp-speed",p.speed||1);
      setVal("#inp-factory",p.factoryRate||5);
      setVal("#inp-warehouse",p.warehouseCapacity||50);
      setVal("#inp-store",p.storeRate||4);
      setVal("#inp-dfw",p.delayFactoryToWarehouse||2);
      setVal("#inp-dws",p.delayWarehouseToStore||1);
      bindEngineToUI(eng);
    });
  });
}

/* ================== Bootstrap ================== */
function main(){
  var root=document.getElementById("nis-root");
  if(!root || root.dataset.rendered==="1") return;
  root.dataset.rendered="1";
  ui(root);

  if(window.Office && window.PowerPoint){
    // داخل PowerPoint
    switchToSelectedSlideEngine();
    Office.context.document.addHandlerAsync(Office.EventType.SelectionChanged,function(){
      switchToSelectedSlideEngine();
    });
  }else{
    // تشغيل من المتصفح العادي
    currentSlideId="browser";
    var eng=getOrCreateEngineForSlide(currentSlideId);
    bindEngineToUI(eng);
  }
}

// اشتغل داخل أو خارج Office بدون كراش
if(window.Office && typeof Office.onReady==="function"){
  Office.onReady().then(function(){ main(); }).catch(function(){ document.addEventListener("DOMContentLoaded",main); });
}else{
  document.addEventListener("DOMContentLoaded", main);
}
