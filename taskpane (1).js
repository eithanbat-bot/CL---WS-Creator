/* global Office, XLSX, CLParser */
var BRIDGE='http://127.0.0.1:17832';
var clFile=null, clWorkbook=null, sheets=[];

function $(id){return document.getElementById(id)}
function esc(v){return String(v==null?'':v).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})}
function pill(t,k){$('statusPill').textContent=t;$('statusPill').className='pill '+(k||'neutral')}

async function bridge(path,opts){
  opts=opts||{};
  var url=BRIDGE+path;
  var timeout=opts.timeoutMs||30000;
  var controller=window.AbortController?new AbortController():null;
  var timer=controller?window.setTimeout(function(){controller.abort()},timeout):null;
  try{
    var init={
      method:opts.method||'GET',
      headers:{'Content-Type':'application/json'},
      body:opts.body,
      targetAddressSpace:'loopback',
      cache:'no-store'
    };
    if(controller)init.signal=controller.signal;
    var r=await fetch(url,init);
    var d={};try{d=await r.json()}catch(e){}
    if(!r.ok)throw new Error(d.error||('Bridge request failed: '+r.status));
    return d;
  }catch(e){
    if(e&&e.name==='AbortError')throw new Error('The local bridge did not finish within '+Math.round(timeout/1000)+' seconds. The requested operation is still running on the local PC.');
    if(e&&e.message==='Failed to fetch')throw new Error('Cannot connect to the local CL-WS bridge at '+url+'. Make sure start-bridge.bat is running on this PC.');
    throw e;
  }finally{
    if(timer)window.clearTimeout(timer);
  }
}

function setBuildEnabled(enabled){
  $('preview').disabled=!enabled;
  $('build').disabled=!enabled;
  $('refreshSheets').disabled=!enabled;
}

function clearResults(){
  $('resultTable').querySelector('tbody').innerHTML='';
  $('partsFound').textContent='0';
  $('geometryFound').textContent='0';
  $('geometryMissing').textContent='0';
  $('totalQty').textContent='0';
}

function sumQty(parts){
  return (parts||[]).reduce(function(a,x){return a+(Number(x.qty)||0)},0);
}

function statusCount(parts, label){
  return (parts||[]).filter(function(x){return String(x.statusLabel||'')===label}).length;
}

/* ---------- Report sheets (hardened: text-safe cells, chunked writes, labelled errors) ---------- */
function safeCell(v){
  if(v==null)return '';
  if(typeof v==='number')return isFinite(v)?v:'';
  if(typeof v==='boolean')return v?'TRUE':'FALSE';
  if(typeof v==='object'){try{v=Array.isArray(v)?v.join(', '):JSON.stringify(v)}catch(e){v=String(v)}}
  v=String(v);
  if(v.length>30000)v=v.slice(0,30000)+'...';
  if(/^[=+\-@]/.test(v))v="'"+v;      /* stop Excel treating text as a formula */
  return v;
}
function safeMatrix(rows,cols){
  return rows.map(function(r){
    var a=r.slice(0,cols).map(safeCell);
    while(a.length<cols)a.push('');
    return a;
  });
}
async function reportStep(label,fatal,fn){
  try{await fn()}
  catch(e){
    var loc=(e&&e.debugInfo&&e.debugInfo.errorLocation)?(' ['+e.debugInfo.errorLocation+']'):'';
    var err=new Error(label+': '+((e&&e.message)||e)+loc);
    if(fatal)throw err;
    window.__reportWarnings=(window.__reportWarnings||[]);window.__reportWarnings.push(err.message);
  }
}

async function writeReportSheets(result,job,selectedSheetNames){
  window.__reportWarnings=[];
  var parts=result.parts||[];
  var review=parts.filter(function(x){return x.status!=='READY'});
  var ready=parts.filter(function(x){return x.status==='READY'});
  var prs=parts.filter(function(x){return String(x.sourceType||'')==='PRS'});
  var dxf=parts.filter(function(x){return String(x.sourceType||'')==='DXF'});
  var missing=parts.filter(function(x){return x.status==='MISSING'});
  var now=new Date().toLocaleString();

  var summaryData=[
    ['CL - WS Creator',''],
    ['Job',job||''],
    ['Generated',now],
    ['CL workbook',clFile?clFile.name:''],
    ['Selected CL sheets',(selectedSheetNames||[]).join(', ')],
    ['PRS geometry root',$('libraryPath').value.trim()],
    ['DXF geometry root',$('dxfPath').value.trim()],
    ['',''],
    ['RESULT','VALUE'],
    ['Consolidated CL lines',parts.length],
    ['Total required quantity',sumQty(parts)],
    ['Ready for SigmaNEST',ready.length],
    ['Review required',review.length],
    ['PRS geometry matches',prs.length],
    ['DXF geometry matches',dxf.length],
    ['Missing geometry',missing.length],
    ['Material mismatches',statusCount(parts,'MATERIAL MISMATCH')],
    ['Thickness mismatches',statusCount(parts,'THICKNESS MISMATCH')],
    ['Variation review',review.filter(function(x){return String(x.statusLabel||'').indexOf('VARIATION')>=0}).length],
    ['Ambiguous matches',review.filter(function(x){return String(x.statusLabel||'').indexOf('AMBIGUOUS')===0}).length],
    ['DXF still indexing',statusCount(parts,'DXF INDEXING')],
    ['',''],
    ['ACTION / HAND-OFF','DETAIL'],
    ['Geometry search','Recursive .PRS search under the PRS root and recursive .DXF search under the DXF root, including all subfolders'],
    ['SigmaNEST WS',result.wsPath||'Not created - review items remain'],
    ['Staging folder',result.outputDir||''],
    ['Part review sheet','See "Part Review" worksheet for every item requiring confirmation']
  ];
  var resultRow=summaryData.findIndex(function(r){return r[0]==='RESULT'});
  var actionRow=summaryData.findIndex(function(r){return r[0]==='ACTION / HAND-OFF'});

  var reviewHeaders=['Part','Qty','Material','Thickness','Status','Reason','Match Type','Source Type','Source File','Source Path','PRS Material','PRS Thickness','CL Sheets'];
  var reviewRows=review.map(function(x){
    return [
      x.part,x.qty,x.material,x.thickness,x.statusLabel||x.status,x.reviewReason||'',
      x.matchType||'',x.sourceType||'',x.sourcePath?String(x.sourcePath).split(/[\/\\]/).pop():'',
      x.sourcePath||'',x.libraryMaterial||'',x.libraryThickness||'',
      Array.isArray(x.sourceSheets)?x.sourceSheets.join(', '):(x.sheet||'')
    ];
  });
  if(!reviewRows.length)reviewRows=[['No parts require review.','','','','','','','','','','','','']];
  var summaryName='CL WS Summary', reviewName='Part Review';

  /* 1. recreate both sheets fresh (avoids merged-cell / stale-format problems) */
  await reportStep('Creating report sheets',true,async function(){
    await Excel.run(async function(context){
      var wb=context.workbook;
      var a=wb.worksheets.getItemOrNullObject(summaryName);
      var b=wb.worksheets.getItemOrNullObject(reviewName);
      await context.sync();
      if(!a.isNullObject)a.delete();
      if(!b.isNullObject)b.delete();
      await context.sync();
      wb.worksheets.add(summaryName);
      wb.worksheets.add(reviewName);
      await context.sync();
    });
  });

  /* 2. summary values */
  await reportStep('Writing summary values',true,async function(){
    await Excel.run(async function(context){
      var sh=context.workbook.worksheets.getItem(summaryName);
      var m=safeMatrix(summaryData,2);
      sh.getRangeByIndexes(0,0,m.length,2).values=m;
      await context.sync();
    });
  });

  /* 3. review values in chunks of 150 rows */
  await reportStep('Writing Part Review values',true,async function(){
    await Excel.run(async function(context){
      var sh=context.workbook.worksheets.getItem(reviewName);
      sh.getRangeByIndexes(0,0,1,reviewHeaders.length).values=[reviewHeaders];
      await context.sync();
      var m=safeMatrix(reviewRows,reviewHeaders.length);
      for(var i=0;i<m.length;i+=150){
        var chunk=m.slice(i,i+150);
        sh.getRangeByIndexes(1+i,0,chunk.length,reviewHeaders.length).values=chunk;
        await context.sync();
      }
    });
  });

  /* 4. formatting (non-fatal: a formatting problem never loses the data) */
  await reportStep('Formatting summary',false,async function(){
    await Excel.run(async function(context){
      var sh=context.workbook.worksheets.getItem(summaryName);
      var all=sh.getRangeByIndexes(0,0,summaryData.length,2);
      all.format.font.name='Segoe UI';all.format.font.size=10;all.format.wrapText=true;
      sh.getRange('A1').format.font.size=18;sh.getRange('A1').format.font.bold=true;
      var h1=sh.getRangeByIndexes(resultRow,0,1,2);
      h1.format.font.bold=true;h1.format.fill.color='#0B2942';h1.format.font.color='#FFFFFF';
      var h2=sh.getRangeByIndexes(actionRow,0,1,2);
      h2.format.font.bold=true;h2.format.fill.color='#E7EEF5';
      sh.getRange('A:A').format.columnWidth=210;
      sh.getRange('B:B').format.columnWidth=420;
      await context.sync();
    });
  });
  await reportStep('Formatting Part Review',false,async function(){
    await Excel.run(async function(context){
      var sh=context.workbook.worksheets.getItem(reviewName);
      var n=reviewRows.length+1;
      var all=sh.getRangeByIndexes(0,0,n,reviewHeaders.length);
      all.format.font.name='Segoe UI';all.format.font.size=10;
      var hd=sh.getRangeByIndexes(0,0,1,reviewHeaders.length);
      hd.format.font.bold=true;hd.format.fill.color='#0B2942';hd.format.font.color='#FFFFFF';
      var widths=[150,70,130,80,150,260,100,90,170,360,130,100,180];
      widths.forEach(function(w,i){sh.getRangeByIndexes(0,i,1,1).format.columnWidth=w});
      await context.sync();
      try{sh.freezePanes.freezeRows(1);await context.sync()}catch(e){}
    });
  });
}

function resetCL(){
  clFile=null;
  clWorkbook=null;
  sheets=[];
  $('clFile').value='';
  $('clFileName').textContent='No CL workbook selected.';
  $('clFileStatus').textContent='Select the cutting-list .xlsx/.xls file from this PC.';
  $('sheetList').innerHTML='Choose a CL workbook first.';
  $('sheetList').className='sheetList empty';
  $('sheetCount').textContent='0 selected';
  $('jobName').value='';
  $('buildStatus').textContent='Choose a CL workbook to begin.';
  $('clearCL').disabled=true;
  setBuildEnabled(false);
  clearResults();
}


function renderSheets(){
  var box=$('sheetList');
  box.innerHTML='';
  box.className='sheetList';
  if(!sheets.length){
    box.className='sheetList empty';
    box.textContent='No worksheets were found in this workbook.';
    $('sheetCount').textContent='0 selected';
    setBuildEnabled(false);
    return;
  }
  sheets.forEach(function(s,i){
    var row=document.createElement('label');
    row.className='sheetItem';
    row.innerHTML='<input type="checkbox" data-index="'+i+'" '+(s.selected?'checked':'')+'> <span>'+esc(s.name)+'</span>';
    box.appendChild(row);
  });
  updateCount();
  setBuildEnabled(true);
}

function updateCount(){
  var n=$('sheetList').querySelectorAll('input:checked').length;
  $('sheetCount').textContent=n+' selected';
  $('preview').disabled=!clWorkbook||n===0;
  $('build').disabled=!clWorkbook||n===0;
}

function parseCLFile(file){
  return new Promise(function(resolve,reject){
    if(!window.XLSX)return reject(new Error('The Excel file parser did not load. Refresh the add-in and try again.'));
    var reader=new FileReader();
    reader.onerror=function(){reject(new Error('Could not read the selected CL workbook.'))};
    reader.onload=function(evt){
      try{
        var wb=XLSX.read(evt.target.result,{type:'array',cellDates:false,cellNF:false,cellStyles:false});
        var parsed=wb.SheetNames.map(function(name){
          var ws=wb.Sheets[name];
          var values=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true,blankrows:true});
          return {name:name,values:values,selected:/^BAT/i.test(name)};
        });
        resolve({file:file,sheets:parsed});
      }catch(e){reject(new Error('The selected file could not be parsed as an Excel workbook: '+e.message))}
    };
    reader.readAsArrayBuffer(file);
  });
}

async function chooseCL(file){
  if(!file)return;
  $('clFileStatus').textContent='Reading '+file.name+'...';
  $('clFileName').textContent=file.name;
  pill('Reading CL','neutral');
  try{
    clFile=file;
    clWorkbook=await parseCLFile(file);
    sheets=clWorkbook.sheets;
    $('clearCL').disabled=false;
    renderSheets();

    var batCount=sheets.filter(function(s){return s.selected}).length;
    var job='';
    for(var i=0;i<sheets.length&&!job;i++)job=CLParser.parseJobNumber(sheets[i].values);
    if(!$('jobName').value.trim()&&job)$('jobName').value=job;

    $('clFileStatus').textContent=sheets.length+' worksheet(s) loaded. '+batCount+' BAT worksheet(s) preselected.';
    $('buildStatus').textContent=batCount?'Select/confirm the CL sheets, then Preview CL or Create SigmaNEST Job.':'No BAT sheets were found; select the required worksheet(s) manually.';
    pill('CL loaded','ok');
  }catch(e){
    resetCL();
    $('clFileName').textContent=file.name;
    $('clFileStatus').textContent=e.message;
    pill('CL load failed','bad');
  }
}

async function loadSheets(){
  if(!clFile){resetCL();return}
  await chooseCL(clFile);
}

function selectedSheets(){
  return [].slice.call($('sheetList').querySelectorAll('input:checked')).map(function(x){return sheets[Number(x.dataset.index)]});
}

async function readCL(){
  if(!clWorkbook)throw new Error('Choose a CL workbook first.');
  var selected=selectedSheets();
  if(!selected.length)throw new Error('Select at least one CL sheet.');
  var all=[],job='';
  selected.forEach(function(s){
    all=all.concat(CLParser.parseSheet(s.values,{sheet:s.name}));
    job=job||CLParser.parseJobNumber(s.values);
  });
  if(!$('jobName').value.trim()&&job)$('jobName').value=job;
  if(!all.length)throw new Error('No usable CL part rows were found in the selected worksheets.');
  return CLParser.consolidate(all);
}

function showParts(parts,statusMode){
  var body=$('resultTable').querySelector('tbody');
  body.innerHTML='';
  $('partsFound').textContent=parts.length;
  $('totalQty').textContent=parts.reduce(function(a,x){return a+x.qty},0);
  parts.forEach(function(p){
    var tr=document.createElement('tr');
    var st=statusMode==='preview'?'—':(p.statusLabel||p.status||'');
    var cls=p.status==='READY'?'okText':(p.status==='REVIEW'?'warnText':'badText');
    tr.innerHTML='<td>'+esc(p.part)+'</td><td>'+p.qty+'</td><td>'+esc(p.material)+'</td><td>'+esc(p.thickness)+'</td><td>'+esc(p.sourceType||'')+'</td><td>'+esc(p.matchType||'')+'</td><td class="'+cls+'">'+esc(st)+'</td>';
    body.appendChild(tr);
  });
}

async function preview(){
  try{
    var p=await readCL();
    showParts(p,'preview');
    $('geometryFound').textContent='—';
    $('geometryMissing').textContent='—';
    $('buildStatus').textContent=p.length+' consolidated CL lines ready.';
    pill('CL ready','ok');
  }catch(e){
    $('buildStatus').textContent=e.message;
    pill('CL error','bad');
  }
}

var dxfMonitorTimer=null;

function monitorDxfIndex(root){
  if(dxfMonitorTimer)window.clearInterval(dxfMonitorTimer);
  async function check(){
    try{
      var st=await bridge('/api/dxf-status');
      if(String(st.root||'')!==String(root))return;
      if(st.state==='RUNNING'){
        $('libraryStatus').textContent='DXF indexing in progress under '+root+' — '+(st.filesFound||0)+' DXF files indexed so far.'+(st.currentPath?' Current: '+st.currentPath:'');
        pill('DXF indexing','neutral');
      }else if(st.state==='COMPLETE'){
        $('libraryStatus').textContent=(st.filesFound||0)+' .DXF files indexed recursively under '+root+'. PRS library is indexed separately.';
        pill('Libraries ready','ok');
        if(dxfMonitorTimer){window.clearInterval(dxfMonitorTimer);dxfMonitorTimer=null;}
      }else if(st.state==='FAILED'){
        $('libraryStatus').textContent='DXF indexing failed: '+(st.message||'Unknown error');
        pill('DXF index failed','bad');
        if(dxfMonitorTimer){window.clearInterval(dxfMonitorTimer);dxfMonitorTimer=null;}
      }
    }catch(e){
      /* Keep polling while the bridge remains temporarily unavailable. */
    }
  }
  check();
  dxfMonitorTimer=window.setInterval(check,3000);
}

async function scan(){
  var btn=$('scanLibrary');
  var started=Date.now();
  btn.disabled=true;
  $('libraryStatus').textContent='Checking PRS and starting background DXF indexing...';
  pill('Scanning geometry','neutral');
  try{
    var prsRoot=$('libraryPath').value.trim()||'S:\\SNDataX1\\PARTS';
    var dxfRoot=$('dxfPath').value.trim()||'Y:\\';
    $('libraryPath').value=prsRoot;
    $('dxfPath').value=dxfRoot;
    localStorage.setItem('clwsc_prsRoot',prsRoot);
    localStorage.setItem('clwsc_dxfRoot',dxfRoot);
    var r=await bridge('/api/scan',{method:'POST',body:JSON.stringify({prsRoot:prsRoot,dxfRoot:dxfRoot})});
    var scanErrors=(r.scanErrors||[]).length, inspectErrors=(r.inspectErrors||[]).length;
    var elapsed=Math.round((Date.now()-started)/1000);
    if(scanErrors||inspectErrors){
      var first=(r.scanErrors&&r.scanErrors[0])?(r.scanErrors[0].path+': '+r.scanErrors[0].error):(r.inspectErrors[0]?r.inspectErrors[0].path+': '+r.inspectErrors[0].error:'Unknown scan issue');
      $('libraryStatus').textContent=(r.prsCount||0)+' .PRS indexed; DXF background indexing status: '+((r.dxfStatus&&r.dxfStatus.state)||'UNKNOWN')+'. First issue: '+first+' ('+elapsed+'s)';
      pill('Library partial','warn');
    }else{
      var ds=r.dxfStatus||{};
      $('libraryStatus').textContent=(r.prsCount||0)+' .PRS indexed under '+prsRoot+'. DXF indexing: '+(ds.state||'STARTING')+' under '+dxfRoot+'. This continues in the background.';
      if(ds.state==='COMPLETE'){
        $('libraryStatus').textContent=(r.prsCount||0)+' .PRS + '+(ds.filesFound||r.dxfCount||0)+' .DXF indexed recursively. Libraries ready.';
        pill('Libraries ready','ok');
      }else if(ds.state==='FAILED'){
        pill('DXF index failed','bad');
      }else{
        pill('DXF indexing','neutral');
        monitorDxfIndex(dxfRoot);
      }
    }
  }catch(e){
    $('libraryStatus').textContent=e.message;
    pill(e.message.indexOf('did not finish')>=0?'Bridge timeout':'Bridge error','bad');
  }finally{
    btn.disabled=false;
  }
}

function breakdownText(b){
  if(!b)return '';
  var k=Object.keys(b);
  return k.length?k.map(function(x){return x+': '+b[x]}).join(' | '):'';
}

async function build(){
  var btn=$('build');
  btn.disabled=true;
  pill('Building job','neutral');
  $('buildStatus').textContent='Matching geometry and building the SigmaNEST job...';
  try{
    var parts=await readCL();
    var selectedNames=selectedSheets().map(function(s){return s.name});
    var prsRoot=$('libraryPath').value.trim()||'S:\\SNDataX1\\PARTS';
    var dxfRoot=$('dxfPath').value.trim()||'Y:\\';
    var job=$('jobName').value.trim()||'CL_JOB';

    /* Do not build against a half-finished DXF index: it makes parts show as review/missing. */
    try{
      var st=await bridge('/api/dxf-status',{timeoutMs:5000});
      if(st.state==='RUNNING'){
        $('buildStatus').textContent='DXF indexing is still running ('+(st.filesFound||0)+' files so far). Wait until the status says "Libraries ready", then click Create again.';
        pill('DXF indexing','neutral');
        monitorDxfIndex(dxfRoot);
        return;
      }
    }catch(e){}

    var r=await bridge('/api/build-job',{
      method:'POST',
      timeoutMs:600000,
      body:JSON.stringify({prsRoot:prsRoot,dxfRoot:dxfRoot,jobName:job,parts:parts})
    });
    var out=r.parts||[];
    showParts(out,'build');
    var ready=out.filter(function(x){return x.status==='READY'}).length;
    $('geometryFound').textContent=ready;
    $('geometryMissing').textContent=out.length-ready;
    var msg=r.message||'Job staged.';
    var bd=breakdownText(r.reviewBreakdown);
    if(bd)msg+=' ['+bd+']';
    try{
      await writeReportSheets(r,job,selectedNames);
      msg+=' Summary and Part Review sheets written to this workbook.';
      if(window.__reportWarnings&&window.__reportWarnings.length)msg+=' (Formatting skipped: '+window.__reportWarnings.join('; ')+')';
    }catch(e){
      msg+=' (Could not write report sheets - '+e.message+')';
    }
    $('buildStatus').textContent=msg;
    pill(r.reviewCount?'Review required':'Job created',r.reviewCount?'warn':'ok');
  }catch(e){
    $('buildStatus').textContent=e.message;
    pill('Build error','bad');
  }finally{
    updateCount();
  }
}

/* Checks the local bridge, never hangs: always ends on a clear status. */
async function checkBridge(){
  try{
    var h=await bridge('/api/health',{timeoutMs:4000});
    pill('Bridge connected','ok');
    var v=$('bridgeVersion'); if(v)v.textContent='Local bridge '+(h.bridgeVersion||'')+' at 127.0.0.1:17832';
    return true;
  }catch(e){
    pill('Bridge not running','bad');
    $('buildStatus').textContent='The local bridge is not running. Start start-bridge.bat on this PC, then click Retry. ('+e.message+')';
    return false;
  }
}

function init(){
  try{
    $('libraryPath').value=localStorage.getItem('clwsc_prsRoot')||'S:\\SNDataX1\\PARTS';
    $('dxfPath').value=localStorage.getItem('clwsc_dxfRoot')||'Y:\\';
  }catch(e){
    $('libraryPath').value='S:\\SNDataX1\\PARTS';
    $('dxfPath').value='Y:\\';
  }
  $('chooseCL').addEventListener('click',function(){$('clFile').click()});
  $('clFile').addEventListener('change',function(){chooseCL($('clFile').files[0])});
  $('clearCL').addEventListener('click',resetCL);
  $('refreshSheets').addEventListener('click',loadSheets);
  $('sheetList').addEventListener('change',updateCount);
  $('scanLibrary').addEventListener('click',scan);
  $('preview').addEventListener('click',preview);
  $('build').addEventListener('click',build);
  $('statusPill').addEventListener('click',checkBridge);
  setBuildEnabled(false);
  pill('Starting...','neutral');
  checkBridge();
}

/* Start even if office.js fails to load or never calls back (e.g. opened in a browser). */
var started=false;
function startOnce(){ if(started)return; started=true; init(); }
window.addEventListener('error',function(ev){
  try{ if(!started) pill('Script error','bad'); $('buildStatus').textContent='Script error: '+ev.message; }catch(e){}
});
if(window.Office&&Office.onReady){ Office.onReady(function(){startOnce()}); }
window.setTimeout(startOnce,2500);
