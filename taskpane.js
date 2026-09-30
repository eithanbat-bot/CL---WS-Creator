/* global Office, Excel, CLParser */
var BRIDGE='http://127.0.0.1:17832', sheets=[];
function $(id){return document.getElementById(id)}
function esc(v){return String(v||'').replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})}
function pill(t,k){$('statusPill').textContent=t;$('statusPill').className='pill '+(k||'neutral')}
async function bridge(path,opts){
  opts=opts||{};
  var r=await fetch(BRIDGE+path,{
    method:opts.method||'GET',
    headers:{'Content-Type':'application/json'},
    body:opts.body,
    targetAddressSpace:'loopback'
  });
  var d={};try{d=await r.json()}catch(e){}
  if(!r.ok)throw new Error(d.error||('Bridge request failed: '+r.status));
  return d;
}
async function loadSheets(){
  await Excel.run(async function(ctx){
    var all=ctx.workbook.worksheets;all.load('items/name');await ctx.sync();
    sheets=all.items.map(function(x,i){return {name:x.name,index:i,selected:/^BAT/i.test(x.name)}})
  });
  renderSheets()
}
function renderSheets(){
  var box=$('sheetList');box.innerHTML='';box.className='sheetList';
  sheets.forEach(function(s,i){
    var row=document.createElement('label');row.className='sheetItem';
    row.innerHTML='<input type="checkbox" data-index="'+i+'" '+(s.selected?'checked':'')+'> <span>'+esc(s.name)+'</span>';
    box.appendChild(row)
  });
  updateCount()
}
function updateCount(){$('sheetCount').textContent=$('sheetList').querySelectorAll('input:checked').length+' selected'}
async function readCL(){
  var selected=[].slice.call($('sheetList').querySelectorAll('input:checked')).map(function(x){return sheets[Number(x.dataset.index)]});
  if(!selected.length)throw new Error('Select at least one CL sheet.');
  var all=[],job='';
  await Excel.run(async function(ctx){
    var refs=selected.map(function(s){return ctx.workbook.worksheets.getItem(s.name)});
    refs.forEach(function(ws,i){var used=ws.getUsedRangeOrNullObject(true);used.load('values,isNullObject');selected[i].used=used});
    await ctx.sync();
    selected.forEach(function(s){
      if(!s.used.isNullObject){
        all=all.concat(CLParser.parseSheet(s.used.values,{sheet:s.name}));
        job=job||CLParser.parseJobNumber(s.used.values)
      }
    })
  });
  if(!$('jobName').value.trim()&&job)$('jobName').value=job;
  return CLParser.consolidate(all)
}
function showParts(parts,statusMode){
  var body=$('resultTable').querySelector('tbody');body.innerHTML='';
  $('partsFound').textContent=parts.length;
  $('totalQty').textContent=parts.reduce(function(a,x){return a+x.qty},0);
  parts.forEach(function(p){
    var tr=document.createElement('tr');
    var st=statusMode==='preview'?'—':(p.statusLabel||p.status||'');
    var cls=p.status==='READY'?'okText':(p.status==='REVIEW'?'warnText':'badText');
    tr.innerHTML='<td>'+esc(p.part)+'</td><td>'+p.qty+'</td><td>'+esc(p.material)+'</td><td>'+esc(p.thickness)+'</td><td class="'+cls+'">'+esc(st)+'</td>';
    body.appendChild(tr)
  })
}
async function preview(){try{var p=await readCL();showParts(p,'preview');$('buildStatus').textContent=p.length+' consolidated CL lines ready.';pill('CL ready','ok')}catch(e){$('buildStatus').textContent=e.message;pill('CL error','bad')}}
async function scan(){
  try{
    var root=$('libraryPath').value.trim()||'S:\\SNDataX1\\PARTS';$('libraryPath').value=root;
    var r=await bridge('/api/scan',{method:'POST',body:JSON.stringify({root:root})});
    $('libraryStatus').textContent=r.count+' .PRS files indexed.';pill('Library ready','ok')
  }catch(e){$('libraryStatus').textContent=e.message;pill('Bridge offline','bad')}
}
async function build(){
  try{
    var parts=await readCL(),root=$('libraryPath').value.trim()||'S:\\SNDataX1\\PARTS',job=$('jobName').value.trim()||('CL_'+Date.now());
    var r=await bridge('/api/build-job',{method:'POST',body:JSON.stringify({jobName:job,libraryRoot:root,parts:parts})});
    showParts(r.parts,'built');
    $('geometryFound').textContent=r.parts.filter(function(x){return x.status==='READY'}).length;
    $('geometryMissing').textContent=r.parts.filter(function(x){return x.status!=='READY'}).length;
    $('buildStatus').textContent=r.wsPath?(r.message+'\nWS: '+r.wsPath):r.message;
    pill(r.reviewCount?'Review required':'SigmaNEST WS ready',r.reviewCount?'warn':'ok')
  }catch(e){$('buildStatus').textContent=e.message;pill('Build failed','bad')}
}
Office.onReady(async function(info){
  if(info.host!==Office.HostType.Excel){pill('Excel only','bad');return}
  $('refreshSheets').onclick=loadSheets;$('scanLibrary').onclick=scan;$('preview').onclick=preview;$('build').onclick=build;
  $('sheetList').onchange=updateCount;
  $('libraryPath').value=localStorage.getItem('clwsc_libraryRoot')||'S:\\SNDataX1\\PARTS';
  $('libraryPath').onchange=function(){localStorage.setItem('clwsc_libraryRoot',$('libraryPath').value.trim())};
  try{
    var h=await bridge('/api/health');
    $('libraryStatus').textContent='Bridge connected. Library: '+h.libraryRoot;pill('Connected','ok')
  }catch(e){
    $('libraryStatus').textContent='Start start-bridge.bat on this PC.';pill('Bridge offline','warn')
  }
  await loadSheets()
});
