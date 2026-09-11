/* Read-only, pane-local capability probe. No command transport or snapshot collector. */
(function (root, factory) {
  'use strict';
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.WordCapabilityProbe = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const SETS = [['WordApi','1.1'],['WordApi','1.2'],['WordApi','1.3'],['WordApi','1.4'],['WordApi','1.5'],
    ['WordApiDesktop','1.2'],['WordApiDesktop','1.3'],['WordApiDesktop','1.4'],['File','1.1'],['CompressedFile','1.1']];
  const GAPS = [
    ['add_semantic_table_native',['WordApi_1_1','WordApi_1_3','WordApi_1_4'],'Range.insertOoxml / insertTable; Table.mergeCells; replacement and rollback unproven'],
    ['add_semantic_table_fallback',['WordApi_1_1','WordApi_1_3','WordApi_1_4'],'Same primitives; native recovery and issue contract unproven'],
    ['add_captioned_figure_native',['WordApi_1_2','WordApi_1_3','WordApi_1_5'],'InlinePicture, container table and fields; caption ownership unproven'],
    ['add_captioned_figure_fallback',['WordApi_1_2','WordApi_1_3','WordApi_1_5'],'Same primitives; fallback/recovery unproven'],
    ['add_equation_native',['WordApi_1_1','WordApi_1_5'],'OMML/OOXML candidate; no direct OMath BuildUp equivalent established'],
    ['add_equation_native_fallback',['WordApi_1_1','WordApi_1_5'],'OMML/OOXML candidate; fallback/recovery unproven'],
    ['add_heading_level_native',['WordApi_1_1','WordApiDesktop_1_3'],'Complex-script font candidate; full style/numbering contract unproven'],
    ['add_horizontal_line',['WordApi_1_1','WordApiDesktop_1_2'],'Drawing/OOXML candidates only; exact native inline horizontal-line type unproven'],
    ['add_wordart',['WordApi_1_1','WordApiDesktop_1_2'],'Drawing/OOXML candidates only; WordArt preset semantics unproven']
  ];
  const METHODS = [['Range','insertOoxml'],['Range','insertTable'],['Range','insertInlinePictureFromBase64'],
    ['Range','insertField'],['Table','mergeCells']];
  function safe(getter) { try { return getter(); } catch (_) { return undefined; } }
  function stringOrNull(value) { return typeof value === 'string' && value.length ? value : null; }
  function code(error) { const value=safe(()=>error.code);return typeof value==='string' && /^[A-Za-z0-9_.-]{1,80}$/.test(value)?value:'READ_FAILED'; }
  function createPane(crypto) {
    if (!crypto || typeof crypto.getRandomValues !== 'function') throw new Error('SECURE_RANDOM_REQUIRED');
    const bytes=new Uint8Array(16);crypto.getRandomValues(bytes);
    return {id:Array.from(bytes,b=>b.toString(16).padStart(2,'0')).join(''),live:true,busy:false,sequence:0,
      retire(){this.live=false;this.sequence++;}};
  }
  function capabilities(env) {
    const context=safe(()=>env.Office.context), diagnostics=safe(()=>context.diagnostics);
    const requirements={};
    for (const [name,version] of SETS) {
      const key=name+'_'+version.replace('.','_');let supported=null,status='unavailable';
      try {
        if(context && context.requirements && typeof context.requirements.isSetSupported==='function') {
          const value=context.requirements.isSetSupported(name,version);
          if(typeof value==='boolean') {supported=value;status='reported';} else status='invalid-result';
        }
      } catch (_) {status='error';}
      requirements[key]={name,version,supported,status};
    }
    const methodPresence={};
    for(const [type,method] of METHODS) {
      const value=safe(()=>env.Word[type].prototype[method]);
      methodPresence[type+'.'+method]=typeof value==='function'?'exposed':'not-exposed-or-unavailable';
    }
    const mode=stringOrNull(safe(()=>context.document.mode));
    return {schema:1,scope:'capability/binding observation only; no parity acceptance',
      host:stringOrNull(safe(()=>diagnostics.host)),platform:stringOrNull(safe(()=>diagnostics.platform)),
      version:stringOrNull(safe(()=>diagnostics.version)),documentMode:mode,
      readOnly:mode==='readOnly'?true:mode==='readWrite'?false:null,
      documentUrlAvailable:typeof safe(()=>context.document.url)==='string' && Boolean(safe(()=>context.document.url)),
      requirements,methodPresence,
      snapshot:{executed:false,getFileAsyncPresent:typeof safe(()=>context.document.getFileAsync)==='function',
        candidateSets:['File_1_1','CompressedFile_1_1'],nativeAccepted:false},
      gaps:GAPS.map(([method,candidateSets,note])=>({method,candidateSets,note,nativeAccepted:false}))};
  }
  async function observe(env,pane,{timeoutMs=10000}={}) {
    if(!pane.live)return {status:'stale'};
    if(pane.busy)return {status:'busy'};
    const c=capabilities(env);
    if(c.host!=='Word'||!['Mac','PC'].includes(c.platform)||c.requirements.WordApi_1_1.supported!==true||typeof safe(()=>env.Word.run)!=='function')
      return {status:'not-run',reason:'Requires reported Word desktop, WordApi 1.1 and Word.run'};
    if(!Number.isFinite(timeoutMs)||timeoutMs<=0||timeoutMs>30000)throw new Error('INVALID_DEADLINE');
    pane.busy=true;const sequence=++pane.sequence;let timer;
    const work=(async()=>{
      try {
        const result=await env.Word.run(async context=>{
          const doc=context.document;doc.load('saved');let selected=null;
          if(c.requirements.WordApiDesktop_1_4.supported===true) {
            selected=doc.getSelection();selected.load('start,end,storyType');
          }
          await context.sync();
          if(typeof doc.saved!=='boolean')throw {code:'INVALID_SAVED_VALUE'};
          let selection=null;
          if(selected) {
            const {start,end,storyType}=selected;
            if(!Number.isInteger(start)||!Number.isInteger(end)||start<0||end<start||typeof storyType!=='string'||!storyType)throw {code:'INVALID_RANGE_VALUE'};
            selection={start,end,storyType};
          }
          return {saved:doc.saved,selection};
        });
        if(!pane.live||pane.sequence!==sequence)return {status:'stale'};
        return {status:'read-acknowledged',...result};
      } catch(error) {
        if(!pane.live||pane.sequence!==sequence)return {status:'stale'};
        pane.retire();return {status:'read-error',errorCode:code(error)};
      }
    })();
    try {
      return await Promise.race([work,new Promise(resolve=>{timer=setTimeout(()=>{
        pane.retire();resolve({status:'timeout',note:'Not canceled at host; pane retired. Reconcile before a new run.'});
      },timeoutMs);})]);
    } finally {clearTimeout(timer);pane.busy=false;}
  }
  return {createPane,capabilities,observe};
});
