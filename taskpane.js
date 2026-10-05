/* global Office, XLSX, CLParser */
var BRIDGE='http://127.0.0.1:17832';
var clFile=null, clWorkbook=null, sheets=[], currentWsPath='', lastActionParts=[];
var dxfMonitorTimer=null;
try{currentWsPath=localStorage.getItem('clwsc_lastWsPath')||''}catch(e){}

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
    if(!r.ok){
      var message=d.error||('Bridge request failed: '+r.status);
      if(d.state)message+=' State: '+d.state+'.';
      if(d.message && d.message!==d.error)message+=' '+d.message;
      if(d.currentPath)message+=' Current path: '+d.currentPath;
      if(d.filesFound!=null)message+=' Indexed: '+d.filesFound+'.';
      throw new Error(message);
    }
    return d;
  }catch(e){
    if(e&&e.name==='AbortError')throw new Error('The local bridge did not finish within '+Math.round(timeout/1000)+' seconds. The requested operation is still running on the local PC.');
    if(e&&e.message==='Failed to fetch')throw new Error('Cannot connect to the local CL-WS bridge at '+url+'. Make sure Start Bridge Fixed.bat is running on this PC.');
    throw e;
  }finally{
    if(timer)window.clearTimeout(timer);
  }
}

function setActionEnabled(enabled){
  $('preview').disabled=!enabled;
  $('refreshSheets').disabled=!enabled;
  $('importGeometry').disabled=!enabled;
  $('autoTaskOrder').disabled=!enabled||!currentWsPath;
}
function updateWorkspacePath(path){
  currentWsPath=String(path||'').trim();
  $('wsPath').value=currentWsPath;
  $('autoTaskOrder').disabled=!clWorkbook||!currentWsPath;
  try{localStorage.setItem('clwsc_lastWsPath',currentWsPath)}catch(e){}
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

async function writeReportSheets(result,job,selectedSheetNames,actionName){
  actionName=actionName||'Operation';
  var parts=result.parts||[];
  var review=parts.filter(function(x){return x.status!=='READY'});
  var ready=parts.filter(function(x){return x.status==='READY'});
  var prs=parts.filter(function(x){return String(x.sourceType||'')==='PRS'});
  var dxf=parts.filter(function(x){return String(x.sourceType||'')==='DXF'});
  var missing=parts.filter(function(x){return x.status==='MISSING'});
  var updates=Array.isArray(result.partUpdates)?result.partUpdates:[];
  var tasks=Array.isArray(result.taskData)?result.taskData:[];
  var warnings=Array.isArray(result.warnings)?result.warnings:[];
  var now=new Date().toLocaleString();

  function rectangular(rows,cols){
    return rows.map(function(r){
      var a=[];for(var i=0;i<cols;i++)a.push(r[i]==null?'':r[i]);return a;
    });
  }

  var summaryData=[
    ['CL - WS Creator','', '', ''],
    [actionName,now,'',''],
    ['Job',job||'','Workspace',result.wsPath||currentWsPath||''],
    ['CL workbook',clFile?clFile.name:'','Selected sheets',(selectedSheetNames||[]).join(', ')],
    ['','','',''],
    ['PARTS','IMPORTED','TASKS','WARNINGS'],
    [parts.length,result.importedCount!=null?result.importedCount:(actionName==='Import Geometry'?parts.length:0),result.tasksCreated!=null?result.tasksCreated:0,warnings.length],
    ['','','',''],
    ['RESULT DETAIL','VALUE','RESULT DETAIL','VALUE'],
    ['Ready',ready.length,'Review required',review.length],
    ['Total required quantity',sumQty(parts),'Missing geometry',result.missingCount!=null?result.missingCount:missing.length],
    ['PRS geometry',prs.length,'DXF geometry',dxf.length],
    ['Material mismatches',statusCount(parts,'MATERIAL MISMATCH'),'Thickness mismatches',statusCount(parts,'THICKNESS MISMATCH')],
    ['Best candidate selected',parts.filter(function(x){return String(x.selectionDecision||'')==='BEST CANDIDATE'}).length,'Candidate reviews',parts.filter(function(x){return String(x.statusLabel||'').indexOf('AMBIGUOUS')===0}).length],
    ['CL overrides applied',updates.filter(function(x){return !(x.warnings||[]).length}).length,'Override warnings',updates.reduce(function(n,x){return n+(x.warnings||[]).length},0)],
    ['Batch values applied',tasks.filter(function(x){return x.batchApplied}).length,'Task labels applied',tasks.filter(function(x){return x.labelApplied}).length],
    ['','','',''],
    ['HAND-OFF','DETAIL','',''],
    ['Geometry source','PRS preferred; recursive DXF index used when PRS is unavailable.','',''],
    ['CL overrides','Material, thickness and quantity are applied to the imported workspace part; the master PRS file is not overwritten.','',''],
    ['Workspace',result.wsPath||currentWsPath||'Not available','',''],
    ['Warnings',warnings.join(' | '),'','']
  ];

  var reviewHeaders=['Part','Qty','Material','Thickness','Status','Reason','Match Type','Source Type','Source File','Source Path','PRS Material','PRS Thickness','CL Sheets','Candidate Decision','Alternatives'];
  var reviewRows=review.map(function(x){
    return [
      x.part,x.qty,x.material,x.thickness,x.statusLabel||x.status,x.reviewReason||'',
      x.matchType||'',x.sourceType||'',x.sourcePath?String(x.sourcePath).split(/[/\\]/).pop():'',
      x.sourcePath||'',x.libraryMaterial||'',x.libraryThickness||'',
      Array.isArray(x.sourceSheets)?x.sourceSheets.join(', '):'',
      x.selectionDecision||'BEST CANDIDATE',
      Array.isArray(x.candidateAlternatives)?x.candidateAlternatives.map(function(a){return String(a.file||'').split(/[/\\]/).pop()+' ['+a.score+']'}).join(' | '):''
    ];
  });
  if(!reviewRows.length)reviewRows=[['No parts require review.','','','','','','','','','','','','','','']];

  var taskHeaders=['Task','Material','Thickness','Batch','Label Applied','Batch Applied','Parts','Property'];
  var taskRows=tasks.map(function(x){
    return [x.taskIndex,x.material,x.thickness,(x.batchMultiplier||1),'YES'===String(x.labelApplied).toUpperCase()?'YES':(x.labelApplied?'YES':'NO'),x.batchApplied?'YES':'NO',x.partCount||'',x.batchProperty||''];
  });
  if(!taskRows.length)taskRows=[['No tasks created','','','','','','','']];

  var summaryName='CL WS Summary',reviewName='Part Review';
  await Excel.run(async function(context){
    var wb=context.workbook;
    var oldSummary=wb.worksheets.getItemOrNullObject(summaryName);
    var oldReview=wb.worksheets.getItemOrNullObject(reviewName);
    await context.sync();
    if(!oldSummary.isNullObject)oldSummary.delete();
    if(!oldReview.isNullObject)oldReview.delete();
    await context.sync();

    var summarySheet=wb.worksheets.add(summaryName);
    var reviewSheet=wb.worksheets.add(reviewName);
    await context.sync();

    summarySheet.getRangeByIndexes(0,0,summaryData.length,4).values=rectangular(summaryData,4);
    summarySheet.getRange('A1:D1').merge();
    summarySheet.getRange('A2:D2').merge();
    summarySheet.getRange('A1:D1').format.font.bold=true;
    summarySheet.getRange('A1:D1').format.font.size=17;
    summarySheet.getRange('A1:D1').format.font.color='#FFFFFF';
    summarySheet.getRange('A1:D1').format.fill.color='#0B2942';
    summarySheet.getRange('A2:D2').format.font.bold=true;
    summarySheet.getRange('A2:D2').format.fill.color='#E7EEF5';
    summarySheet.getRange('A6:D6').format.font.bold=true;
    summarySheet.getRange('A6:D6').format.fill.color='#0B5CAB';
    summarySheet.getRange('A6:D6').format.font.color='#FFFFFF';
    summarySheet.getRange('A9:D9').format.font.bold=true;
    summarySheet.getRange('A9:D9').format.fill.color='#E7EEF5';
    summarySheet.getRange('A18:D18').format.font.bold=true;
    summarySheet.getRange('A18:D18').format.fill.color='#E7EEF5';
    summarySheet.getRange('A:A').format.columnWidth=165;
    summarySheet.getRange('B:B').format.columnWidth=220;
    summarySheet.getRange('C:C').format.columnWidth=165;
    summarySheet.getRange('D:D').format.columnWidth=220;
    summarySheet.getUsedRange().format.wrapText=true;
    summarySheet.freezePanes.freezeRows(2);

    var rm=[reviewHeaders].concat(reviewRows);
    reviewSheet.getRangeByIndexes(0,0,rm.length,reviewHeaders.length).values=rectangular(rm,reviewHeaders.length);
    var rh=reviewSheet.getRangeByIndexes(0,0,1,reviewHeaders.length);
    rh.format.font.bold=true;rh.format.font.color='#FFFFFF';rh.format.fill.color='#0B2942';
    reviewSheet.getUsedRange().format.wrapText=true;
    reviewSheet.getRangeByIndexes(0,0,rm.length,reviewHeaders.length).format.autofitColumns();
    reviewSheet.getRange('F:F').format.columnWidth=260;
    reviewSheet.getRange('J:J').format.columnWidth=240;
    reviewSheet.getRange('O:O').format.columnWidth=260;
    reviewSheet.freezePanes.freezeRows(1);

    if(tasks.length){
      var taskSheet=wb.worksheets.add('Task Release');
      var tr=[taskHeaders].concat(taskRows);
      taskSheet.getRangeByIndexes(0,0,tr.length,taskHeaders.length).values=rectangular(tr,taskHeaders.length);
      var th=taskSheet.getRangeByIndexes(0,0,1,taskHeaders.length);
      th.format.font.bold=true;th.format.font.color='#FFFFFF';th.format.fill.color='#166534';
      taskSheet.getUsedRange().format.wrapText=true;
      taskSheet.getUsedRange().format.autofitColumns();
      taskSheet.freezePanes.freezeRows(1);
    }
    await context.sync();
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
  updateWorkspacePath('');
  $('buildStatus').textContent='Choose a CL workbook to begin.';
  $('clearCL').disabled=true;
  setActionEnabled(false);
  $('wsPath').value=currentWsPath;
  updateCount();
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
    setActionEnabled(false);
    return;
  }
  sheets.forEach(function(s,i){
    var row=document.createElement('label');
    row.className='sheetItem';
    row.innerHTML='<input type="checkbox" data-index="'+i+'" '+(s.selected?'checked':'')+'> <span>'+esc(s.name)+'</span>';
    box.appendChild(row);
  });
  updateCount();
  setActionEnabled(true);
}

function updateCount(){
  var n=$('sheetList').querySelectorAll('input:checked').length;
  $('sheetCount').textContent=n+' selected';
  $('preview').disabled=!clWorkbook||n===0;
  $('importGeometry').disabled=!clWorkbook||n===0;
  $('autoTaskOrder').disabled=!clWorkbook||n===0||!currentWsPath;
  $('refreshSheets').disabled=!clWorkbook;
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

function monitorDxfIndex(root){
  if(dxfMonitorTimer)window.clearInterval(dxfMonitorTimer);
  async function check(){
    try{
      var st=await bridge('/api/dxf-status');
      if(String(st.root||'')!==String(root))return;
      if(st.state==='RUNNING'){
        $('libraryStatus').textContent=(st.mode||'DXF')+' indexing under '+root+' — '+(st.filesFound||0)+' DXF files found; '+(st.workersCompleted||0)+'/'+(st.workers||0)+' workers complete.'+(st.currentPath?' Current: '+st.currentPath:'');
        pill('DXF indexing','neutral');
      }else if(st.state==='COMPLETE'){
        $('libraryStatus').textContent=(st.filesFound||0)+' .DXF files indexed under '+root+' using '+(st.workers||1)+' worker(s). Next refresh uses '+(st.workers||12)+' worker(s). Automatic overnight refresh: '+(st.nextRefreshLocal||'scheduled')+'. PRS library is indexed separately.';
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
        var detail=ds.message||'Unknown DXF indexer failure.';
        $('libraryStatus').textContent='DXF indexing FAILED under '+dxfRoot+'. '+detail;
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

async function refreshDxf(){
  var btn=$('refreshDxf');
  btn.disabled=true;
  $('libraryStatus').textContent='Starting an immediate parallel DXF refresh under '+($('dxfPath').value.trim()||'Y:\\')+'...';
  pill('DXF refresh','neutral');
  try{
    var dxfRoot=$('dxfPath').value.trim()||'Y:\\';
    var r=await bridge('/api/dxf-refresh',{method:'POST',body:JSON.stringify({dxfRoot:dxfRoot})});
    var st=r.dxfStatus||{};
    $('libraryStatus').textContent=(st.mode||'REFRESH')+' started under '+dxfRoot+' — '+(st.workers||0)+' worker(s). The existing index remains available until the refreshed index is complete.';
    monitorDxfIndex(dxfRoot);
  }catch(e){
    $('libraryStatus').textContent=e.message;
    pill('DXF refresh failed','bad');
  }finally{
    btn.disabled=false;
  }
}

function reviewBreakdownText(breakdown){
  var entries=[];
  Object.keys(breakdown||{}).forEach(function(k){entries.push(k+': '+breakdown[k])});
  return entries.join(' | ');
}

function sleep(ms){return new Promise(function(resolve){window.setTimeout(resolve,ms)})}

async function waitForBuild(jobId,job,selectedNames,actionName){
  var reportWritten=false;
  while(true){
    var s=await bridge('/api/build-status/'+encodeURIComponent(jobId),{timeoutMs:10000});
    var state=String(s.state||'').toUpperCase(), phase=s.phase||'WORKING';
    var elapsed=Number(s.elapsedSeconds)||0;

    if(state==='RUNNING'||state==='STARTING'){
      var detail=s.message||actionName+' is running.';
      if(!reportWritten&&String(phase).toUpperCase()==='PREPARED'&&Array.isArray(s.parts)){
        var interim=s.result||{};interim.parts=s.parts;interim.outputDir=s.outputDir||'';interim.wsPath=s.wsPath||currentWsPath;
        try{await writeReportSheets(interim,job,selectedNames,actionName);reportWritten=true;detail+=' Summary updated.';}catch(e){}
      }
      $('buildStatus').textContent=detail+' Elapsed: '+Math.floor(elapsed)+'s.';
      pill(actionName+' running','neutral');
      await sleep(2500);
      continue;
    }

    var result=s.result||{};
    result.parts=s.parts||result.parts||[];
    result.outputDir=s.outputDir||result.outputDir||'';
    result.wsPath=result.wsPath||s.wsPath||currentWsPath||'';
    result.importedCount=result.importedCount!=null?result.importedCount:(s.importedCount!=null?s.importedCount:0);
    result.missingCount=result.missingCount!=null?result.missingCount:(s.missingCount!=null?s.missingCount:0);
    result.tasksCreated=result.tasksCreated!=null?result.tasksCreated:(s.tasksCreated!=null?s.tasksCreated:0);
    result.taskData=result.taskData||s.taskData||[];
    result.partUpdates=result.partUpdates||s.partUpdates||[];
    result.warnings=result.warnings||s.warnings||[];
    result.reviewCount=result.reviewCount!=null?result.reviewCount:(s.reviewCount!=null?s.reviewCount:0);
    result.reviewBreakdown=result.reviewBreakdown||s.reviewBreakdown||{};

    lastActionParts=Array.isArray(result.parts)?result.parts:[];
    showParts(result.parts,'build');
    if(result.wsPath)updateWorkspacePath(result.wsPath);

    $('geometryFound').textContent=Number(result.importedCount)||0;
    $('geometryMissing').textContent=Number(result.reviewCount)||result.parts.filter(function(x){return x.status!=='READY'}).length;
    $('totalQty').textContent=sumQty(result.parts);

    if(state==='COMPLETE'){
      $('releaseTitle').textContent=actionName+' complete';
      $('releaseDetail').textContent=(result.message||'Operation completed.')+' '+(result.wsPath||'');
      try{await writeReportSheets(result,job,selectedNames,actionName);}catch(e){}
      pill(result.warnings&&result.warnings.length?'Completed with warnings':'Completed','ok');
      $('buildStatus').textContent=(result.message||actionName+' completed.')+' Release Summary updated.';
      return result;
    }

    $('releaseTitle').textContent=actionName+' failed';
    $('releaseDetail').textContent=s.message||result.error||'Operation failed.';
    try{await writeReportSheets(result,job,selectedNames,actionName);}catch(e){}
    pill('Operation failed','bad');
    throw new Error(s.message||result.error||actionName+' failed.');
  }
}

async function runAction(endpoint,body,actionName){
  $('buildStatus').textContent=actionName+' accepted. Starting background operation...';
  var r=await bridge(endpoint,{method:'POST',timeoutMs:30000,body:JSON.stringify(body)});
  if(!r.accepted||!r.jobId)throw new Error(r.message||('The bridge did not accept '+actionName+'.'));
  return await waitForBuild(r.jobId,body.jobName||$('jobName').value.trim()||'CL_JOB',body.selectedSheetNames||selectedSheets().map(function(x){return x.name}),actionName);
}

async function importGeometry(){
  var btn=$('importGeometry');btn.disabled=true;
  try{
    var dxfRoot=$('dxfPath').value.trim()||'Y:\';
    var ds=await bridge('/api/dxf-status',{timeoutMs:4000});
    if(String(ds.root||'').toUpperCase()===String(dxfRoot).toUpperCase()&&String(ds.state||'').toUpperCase()!=='COMPLETE'){
      throw new Error('DXF index is not complete yet. '+(ds.filesFound||0)+' DXFs indexed so far.');
    }
    var parts=lastActionParts.length?lastActionParts:await readCL();
    var selectedNames=selectedSheets().map(function(x){return x.name});
    var prsRoot=$('libraryPath').value.trim()||'S:\SNDataX1\PARTS';
    var job=$('jobName').value.trim()||'CL_JOB';
    updateWorkspacePath('');
    var result=await runAction('/api/import-geometry',{prsRoot:prsRoot,dxfRoot:dxfRoot,jobName:job,selectedSheetNames:selectedNames,parts:parts},'Import Geometry');
    if(result.wsPath)updateWorkspacePath(result.wsPath);
  }catch(e){
    $('buildStatus').textContent=e.message;pill('Import failed','bad');
  }finally{updateCount();}
}

async function autoTaskOrder(){
  var btn=$('autoTaskOrder');btn.disabled=true;
  try{
    if(!currentWsPath)throw new Error('Import Geometry first or enter the path of an existing SigmaNEST .ws.');
    var parts=await readCL();
    var selectedNames=selectedSheets().map(function(x){return x.name});
    var prsRoot=$('libraryPath').value.trim()||'S:\SNDataX1\PARTS';
    var job=$('jobName').value.trim()||'CL_JOB';
    await runAction('/api/autotask-label',{prsRoot:prsRoot,wsPath:currentWsPath,jobName:job,selectedSheetNames:selectedNames,parts:parts},'AutoTask + Order Label');
  }catch(e){
    $('buildStatus').textContent=e.message;pill('AutoTask failed','bad');
  }finally{updateCount();}
}

async function checkBridge(){
  try{
    var h=await bridge('/api/health',{timeoutMs:4000});
    pill('Bridge connected','ok');
    var v=$('bridgeVersion'); if(v)v.textContent='Local bridge '+(h.bridgeVersion||'')+' at 127.0.0.1:17832';
    return true;
  }catch(e){
    pill('Bridge not running','bad');
    $('buildStatus').textContent='The local bridge is not running. Start Start Bridge Fixed.bat on this PC, then click Retry. ('+e.message+')';
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
  $('refreshDxf').addEventListener('click',refreshDxf);
  $('preview').addEventListener('click',preview);
  $('importGeometry').addEventListener('click',importGeometry);
  $('autoTaskOrder').addEventListener('click',autoTaskOrder);
  $('wsPath').addEventListener('change',function(){updateWorkspacePath($('wsPath').value)});
  $('statusPill').addEventListener('click',checkBridge);
  setActionEnabled(false);
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
