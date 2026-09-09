'use strict';
const test=require('node:test'),assert=require('node:assert/strict'),vm=require('node:vm'),fs=require('node:fs');
const source=fs.readFileSync(require('node:path').join(__dirname,'../taskpane.js'),'utf8');
async function setup(host='Word',platform='Mac') {
  const elements={};for(const id of ['status','identity','case-label','capabilities','observe','export','result'])elements[id]={value:'',disabled:true,events:{},addEventListener(name,fn){this.events[name]=fn;}};
  const events={};let reads=0;
  const pane={id:'stub-ephemeral',live:true,retire(){this.live=false;}};
  const context={document:{getElementById:id=>elements[id]},window:{Office:{onReady:async()=>({host,platform})},WordCapabilityProbe:{createPane:()=>pane,capabilities:()=>({nativeAccepted:false}),observe:async()=>{reads++;return {status:'read-acknowledged',saved:false};}},addEventListener:(name,fn)=>{events[name]=fn;}},setTimeout,clearTimeout};
  vm.runInNewContext(source,context);await new Promise(resolve=>setImmediate(resolve));
  return {elements,events,pane,get reads(){return reads;}};
}
test('initialization never auto-submits a document read',async()=>{
  const s=await setup();assert.equal(s.reads,0);assert.equal(s.elements.observe.disabled,false);assert.equal(JSON.parse(s.elements.result.value).observation.status,'not-run');
  await s.elements.observe.events.click();assert.equal(s.reads,1);assert.equal(JSON.parse(s.elements.result.value).observation.saved,false);
});
test('wrong host has no enabled document read',async()=>{
  const s=await setup('Excel');assert.equal(s.elements.observe.disabled,true);assert.equal(s.reads,0);
});
test('pagehide retires pane and rejects further read clicks',async()=>{
  const s=await setup();s.events.pagehide();await s.elements.observe.events.click();assert.equal(s.reads,0);assert.equal(s.pane.live,false);
});
