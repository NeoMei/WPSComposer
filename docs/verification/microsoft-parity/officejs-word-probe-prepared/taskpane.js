/* UI-only wiring; no bridge, telemetry, file snapshot, or persisted identity. */
(function () {
  'use strict';
  const el=id=>document.getElementById(id);
  let pane=null, result=null, ready=false, initializationEnded=false;
  const environment=()=>({Office:window.Office,Word:window.Word,crypto:window.crypto});
  function display(observation) {
    result={schema:1,capturedAt:new Date().toISOString(),testLabel:el('case-label').value,
      paneInstanceId:pane.id,paneLive:pane.live,paneIdentityScope:'page lifetime only; document identity NOT proven',
      capabilities:window.WordCapabilityProbe.capabilities(environment()),observation};
    el('result').value=JSON.stringify(result,null,2);el('export').disabled=false;
  }
  const timeout=setTimeout(()=>{
    if(!initializationEnded){initializationEnded=true;el('status').textContent='Office.js readiness timed out. No host read submitted. Reconcile setup before reloading.';}
  },10000);
  if(!window.Office || typeof window.Office.onReady!=='function') {
    clearTimeout(timeout);initializationEnded=true;el('status').textContent='Office.js unavailable. This page has not run in an acknowledged Office host.';return;
  }
  window.Office.onReady().then(info=>{
    if(initializationEnded)return;initializationEnded=true;clearTimeout(timeout);
    if(!info||info.host!=='Word'||!['Mac','PC'].includes(info.platform)) {
      el('status').textContent='Not an acknowledged desktop Word host. No document reads enabled.';return;
    }
    try {pane=window.WordCapabilityProbe.createPane(window.crypto);}
    catch(_){el('status').textContent='Secure ephemeral identity unavailable; stopped.';return;}
    ready=true;el('identity').textContent='Pane instance: '+pane.id;
    el('capabilities').disabled=false;el('observe').disabled=false;
    el('status').textContent='Ready for a manually initiated observation. No document read submitted.';
    display({status:'not-run'});
  }).catch(()=>{
    clearTimeout(timeout);initializationEnded=true;el('status').textContent='Office.js readiness failed. No document read submitted.';
  });
  el('capabilities').addEventListener('click',()=>{if(ready&&pane.live)display({status:'not-run'});});
  el('observe').addEventListener('click',async()=>{
    if(!ready||!pane.live)return;
    el('observe').disabled=true;el('capabilities').disabled=true;
    el('status').textContent='Read pending; no second read will be submitted.';
    const observation=await window.WordCapabilityProbe.observe(environment(),pane);
    display(observation);el('status').textContent='Observation: '+observation.status+'. This is not parity acceptance.';
    el('observe').disabled=!pane.live;el('capabilities').disabled=!pane.live;
  });
  el('export').addEventListener('click',()=>{
    if(!result)return;
    const url=URL.createObjectURL(new Blob([JSON.stringify(result,null,2)],{type:'application/json'}));
    const link=document.createElement('a');link.href=url;link.download='word-probe-'+pane.id+'.json';
    document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
  });
  window.addEventListener('pagehide',()=>{ready=false;if(pane)pane.retire();});
})();
