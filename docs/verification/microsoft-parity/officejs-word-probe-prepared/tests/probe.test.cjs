'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto').webcrypto;
const p = require('../probe.js');
function env({host='Word', platform='Mac', support=true, badSync=false, badSaved=false}={}) {
  let syncs=0;
  const range={load(fields){assert.equal(fields,'start,end,storyType');},start:12,end:19,storyType:'MainText'};
  const doc={saved:badSaved?'true':false,load(fields){assert.equal(fields,'saved');},getSelection(){return range;}};
  return {Office:{context:{diagnostics:{host,platform,version:'stub-only'},requirements:{isSetSupported(){return support;}},document:{mode:'readWrite',url:null,getFileAsync(){throw Error('Snapshot must not be invoked');}}}},Word:{run:async fn=>fn({document:doc,sync:async()=>{syncs++;if(badSync)throw {code:'AccessDenied',message:'private document path'};}})},crypto,get syncs(){return syncs;}};
}
test('pane identity is ephemeral, unique, and retired explicitly',()=>{
  const a=p.createPane(crypto), b=p.createPane(crypto);assert.notEqual(a.id,b.id);assert.equal(a.live,true);a.retire();assert.equal(a.live,false);
});
test('no secure randomness fails closed',()=>assert.throws(()=>p.createPane({}),/SECURE_RANDOM/));
test('requirement false/missing/error is distinct and never fabricated',()=>{
  let e=env({support:false});assert.equal(p.capabilities(e).requirements.WordApi_1_1.supported,false);
  delete e.Office.context.requirements;assert.equal(p.capabilities(e).requirements.WordApi_1_1.supported,null);
  e.Office.context.requirements={isSetSupported(){throw Error('private');}};assert.equal(p.capabilities(e).requirements.WordApi_1_1.status,'error');
});
test('nine gaps remain unproven and snapshot is not executed',()=>{
  const c=p.capabilities(env());assert.equal(c.gaps.length,9);assert(c.gaps.every(g=>g.nativeAccepted===false));assert.equal(c.snapshot.executed,false);assert.equal(c.snapshot.getFileAsyncPresent,true);
});
test('saved flag and optional coordinates are acknowledged reads only',async()=>{
  const e=env();const r=await p.observe(e,p.createPane(crypto));assert.equal(r.status,'read-acknowledged');assert.equal(r.saved,false);assert.deepEqual(r.selection,{start:12,end:19,storyType:'MainText'});assert.equal(e.syncs,1);
});
test('non-desktop and wrong host never call Word.run',async()=>{
  for(const opts of [{host:'Excel'},{platform:'OfficeOnline'}, {platform:'iOS'}]) {const e=env(opts);assert.equal((await p.observe(e,p.createPane(crypto))).status,'not-run');assert.equal(e.syncs,0);}
});
test('false baseline and missing Word.run fail closed',async()=>{
  const e=env({support:false});assert.equal((await p.observe(e,p.createPane(crypto))).status,'not-run');assert.equal(e.syncs,0);
  delete e.Word.run;assert.equal((await p.observe(e,p.createPane(crypto))).status,'not-run');
});
test('unsupported desktop coordinates do not access range',async()=>{
  const e=env();e.Office.context.requirements.isSetSupported=(name)=>name==='WordApi';e.Word.run=async fn=>fn({document:{saved:true,load(){},getSelection(){throw Error('Not gated');}},sync:async()=>{}});
  const r=await p.observe(e,p.createPane(crypto));assert.equal(r.saved,true);assert.equal(r.selection,null);
});
test('sync failure exposes code but not message/document data',async()=>{
  const r=await p.observe(env({badSync:true}),p.createPane(crypto));assert.equal(r.status,'read-error');assert.equal(r.errorCode,'AccessDenied');assert(!JSON.stringify(r).includes('private'));
});
test('mistyped native saved value is not accepted',async()=>assert.equal((await p.observe(env({badSaved:true}),p.createPane(crypto))).status,'read-error'));
test('pane retirement rejects late read result and stale reuse',async()=>{
  const e=env(),pane=p.createPane(crypto);e.Word.run=async fn=>{pane.retire();return fn({document:{saved:true,load(){}},sync:async()=>{}});};
  assert.equal((await p.observe(e,pane)).status,'stale');assert.equal((await p.observe(e,pane)).status,'stale');
});
test('timeout retires pane; late resolution cannot become an ACK',async()=>{
  const e=env(),pane=p.createPane(crypto);let finish;e.Word.run=()=>new Promise(resolve=>{finish=resolve;});
  const r=await p.observe(e,pane,{timeoutMs:5});assert.equal(r.status,'timeout');assert.equal(pane.live,false);finish({saved:true});
  assert.equal((await p.observe(e,pane)).status,'stale');
});
test('one pending read per pane',async()=>{
  const e=env(),pane=p.createPane(crypto);let finish;e.Word.run=()=>new Promise(resolve=>{finish=resolve;});const a=p.observe(e,pane);
  assert.equal((await p.observe(e,pane)).status,'busy');finish({saved:false,selection:null});await a;
});
test('no Office context remains explicitly unavailable',()=>{
  const r=p.capabilities({});assert.equal(r.host,null);assert.equal(r.requirements.WordApi_1_1.supported,null);
});
test('non-boolean requirement result remains invalid rather than true',()=>{
  const e=env();e.Office.context.requirements.isSetSupported=()=>1;
  assert.deepEqual(p.capabilities(e).requirements.File_1_1,{name:'File',version:'1.1',supported:null,status:'invalid-result'});
});
test('readOnly mode is reported without declaring protection semantics',()=>{
  const e=env();e.Office.context.document.mode='readOnly';assert.equal(p.capabilities(e).readOnly,true);
  e.Office.context.document.mode='unknown';assert.equal(p.capabilities(e).readOnly,null);
});
test('invalid range scalar causes error rather than acceptance',async()=>{
  const e=env();e.Word.run=async fn=>fn({document:{saved:true,load(){},getSelection(){return {load(){},start:19,end:12,storyType:'MainText'};}},sync:async()=>{}});
  assert.equal((await p.observe(e,p.createPane(crypto))).errorCode,'INVALID_RANGE_VALUE');
});
test('gap keys match the frozen nine-method contract exactly',()=>{
  assert.deepEqual(p.capabilities(env()).gaps.map(g=>g.method).sort(),[
    'add_captioned_figure_fallback',
    'add_captioned_figure_native',
    'add_equation_native',
    'add_equation_native_fallback',
    'add_heading_level_native',
    'add_horizontal_line',
    'add_semantic_table_fallback',
    'add_semantic_table_native',
    'add_wordart'
  ].sort());
});
