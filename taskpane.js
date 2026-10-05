/* global Office, XLSX, CLParser */
var BRIDGE='http://127.0.0.1:17832';
var clFile=null, clWorkbook=null, sheets=[];
var dxfMonitorTimer=null;

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

async function writeReportSheets(result,job,selectedSheetNames){
  var parts=result.parts||[];
  var review=parts.filter(function(x){return x.status!=='READY'});
  var ready=parts.filter(function(x){return x.status==='READY'});
  var prs=parts.filter(function(x){return String(x.sourceType||'')==='PRS'});
  var dxf=parts.filter(function(x){return String(x.sourceType||'')==='DXF'});
  var missing=parts.filter(function(x){return x.status==='MISSING'});
  var now=new Date().toLocaleString();

  function rectangular(rows,cols){
    return rows.map(function(r){
      var a=[];
      for(var i=0;i<cols;i++)a.push(r[i]==null?'':r[i]);
      return a;
    });
  }

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
    ['Ready status',ready.length],
    ['Review required',review.length],
    ['Imported into WS',(result.importedCount!=null?result.importedCount:ready.length)+" of "+parts.length],
    ['Missing geometry',(result.missingCount!=null?result.missingCount:missing.length)],
    ['PRS geometry matches',prs.length],
    ['DXF geometry matches',dxf.length],
    ['Missing geometry',missing.length],
    ['Material mismatches',statusCount(parts,'MATERIAL MISMATCH')],
    ['Thickness mismatches',statusCount(parts,'THICKNESS MISMATCH')],
    ['Variation review',review.filter(function(x){return String(x.statusLabel||'').indexOf('VARIATION')>=0}).length],
    ['Ambiguous matches',review.filter(function(x){return String(x.statusLabel||'').indexOf('AMBIGUOUS')===0}).length],
    ['',''],
    ['ACTION / HAND-OFF','DETAIL'],
    ['Geometry search','Recursive .PRS search under the PRS root and indexed .DXF search under the DXF root, including all subfolders'],
    ['SigmaNEST WS',result.wsPath||'Not created — no geometry was available to import'],
    ['Staging folder',result.outputDir||''],
    ['Part review sheet','See "Part Review" worksheet for every item requiring confirmation']
  ];

  var reviewHeaders=['Part','Qty','Material','Thickness','Status','Reason','Match Type','Source Type','Source File','Source Path','PRS Material','PRS Thickness','CL Sheets'];
  var reviewRows=review.map(function(x){
    return [
      x.part,x.qty,x.material,x.thickness,x.statusLabel||x.status,x.reviewReason||'',
      x.matchType||'',x.sourceType||'',x.sourcePath?String(x.sourcePath).split(/[/\\\\]/).pop():'',
      x.sourcePath||'',x.libraryMaterial||'',x.libraryThickness||'',
      Array.isArray(x.sourceSheets)?x.sourceSheets.join(', '):''
    ];
  });
  if(!reviewRows.length)reviewRows=[['No parts require review.','','','','','','','','','','','','']];

  var summaryName='CL WS Summary';
  var reviewName='Part Review';

  // Recreate only the two report sheets. This avoids UsedRange/clear/merge
  // edge cases in different Excel desktop builds.
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

    summarySheet.getRangeByIndexes(0,0,summaryData.length,2).values=rectangular(summaryData,2);
    var rm=[reviewHeaders].concat(reviewRows);
    reviewSheet.getRangeByIndexes(0,0,rm.length,reviewHeaders.length).values=rectangular(rm,reviewHeaders.length);
    await context.sync();

    summarySheet.getRange('A1:B1').format.font.bold=true;
    summarySheet.getRange('A8:B8').format.font.bold=true;
    summarySheet.getRange('A23:B23').format.font.bold=true;
    reviewSheet.getRangeByIndexes(0,0,1,reviewHeaders.length).format.font.bold=true;
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

async function waitForBuild(jobId,job,selectedNames){
  var reportWritten=false;
  while(true){
    var s=await bridge('/api/build-status/'+encodeURIComponent(jobId),{timeoutMs:10000});
    var elapsed=Number(s.elapsedSeconds)||0;
    if(String(s.state||'').toUpperCase()==='RUNNING' || String(s.state||'').toUpperCase()==='STARTING'){
      var phase=s.phase||'WORKING';
      var detail=s.message||'SigmaNEST background worker is running.';
      if(String(phase).toUpperCase()==='PREPARED' && !reportWritten && Array.isArray(s.parts)){
        var interim={
          parts:s.parts,
          outputDir:s.outputDir||'',
          wsPath:s.wsPath||'',
          reviewCount:s.reviewCount||0,
          reviewBreakdown:s.reviewBreakdown||{},
          importedCount:0,
          missingCount:s.missingCount||0
        };
        try{
          await writeReportSheets(interim,job,selectedNames);
          reportWritten=true;
          detail+=' Summary and Part Review are now written to this workbook.';
        }catch(e){
          detail+=' Report sheets will be retried: '+(e&&e.message?e.message:String(e));
        }
      }
      $('buildStatus').textContent=detail+' Elapsed: '+Math.floor(elapsed)+'s. Phase: '+phase+'.';
      pill('Building job','neutral');
      await sleep(2500);
      continue;
    }

    var finalResult=s.result||{};
    finalResult.parts=s.parts||finalResult.parts||[];
    finalResult.outputDir=s.outputDir||finalResult.outputDir||'';
    finalResult.wsPath=finalResult.wsPath||s.wsPath||'';
    finalResult.importedCount=finalResult.importedCount!=null?finalResult.importedCount:(s.importedCount!=null?s.importedCount:(finalResult.partCount!=null?finalResult.partCount:0));
    finalResult.missingCount=finalResult.missingCount!=null?finalResult.missingCount:(s.missingCount!=null?s.missingCount:0);
    finalResult.reviewCount=finalResult.reviewCount!=null?finalResult.reviewCount:(s.reviewCount!=null?s.reviewCount:0);
    finalResult.reviewBreakdown=finalResult.reviewBreakdown||s.reviewBreakdown||{};

    var out=finalResult.parts||[];
    showParts(out,'build');

    var imported=Number(finalResult.importedCount)||0;
    var missing=Number(finalResult.missingCount)||out.filter(function(x){return x.status==='MISSING'}).length;
    var reviewCount=Number(finalResult.reviewCount)||out.filter(function(x){return x.status!=='READY'}).length;
    $('geometryFound').textContent=imported;
    $('geometryMissing').textContent=reviewCount;
    $('totalQty').textContent=sumQty(out);

    var state=String(s.state||'').toUpperCase();
    if(state==='COMPLETE'){
      var msg=s.message||finalResult.message||'SigmaNEST job completed.';
      if(finalResult.wsPath)msg+=' WS: '+finalResult.wsPath+'.';
      msg+=' Imported into WS: '+imported+' of '+out.length+'.';
      if(missing>0)msg+=' Missing geometry: '+missing+'.';
      if(reviewCount>0 && Object.keys(finalResult.reviewBreakdown||{}).length){
        msg+=' Reasons: '+reviewBreakdownText(finalResult.reviewBreakdown)+'.';
      }

      try{
        await writeReportSheets(finalResult,job,selectedNames);
        msg+=' Summary and Part Review sheets written to this workbook.';
      }catch(e){
        var reportError=e&&e.message?e.message:String(e);
        msg+=' (Could not write report sheets: '+reportError+')';
      }

      $('buildStatus').textContent=msg+' Completed in '+Math.floor(Number(s.elapsedSeconds)||0)+'s.';
      pill(reviewCount?'Review required':'Job created',reviewCount?'warn':'ok');
      return finalResult;
    }

    // A failed worker may still have created geometry before its final error.
    var failMsg=s.message||finalResult.error||'SigmaNEST background build failed.';
    try{
      await writeReportSheets(finalResult,job,selectedNames);
      failMsg+=' A Summary and Part Review were written from the completed diagnostic state.';
    }catch(e){
      failMsg+=' (Could not write report sheets: '+(e&&e.message?e.message:String(e))+')';
    }
    $('buildStatus').textContent=failMsg+' Elapsed: '+Math.floor(Number(s.elapsedSeconds)||0)+'s.';
    pill('Build failed','bad');
    throw new Error(failMsg);
  }
}

async function build(){
  var btn=$('build');
  btn.disabled=true;
  pill('Starting job','neutral');
  $('buildStatus').textContent='Matching geometry and starting the SigmaNEST background job...';
  try{
    var dxfRoot=$('dxfPath').value.trim()||'Y:\\';
    var dxfStatus=await bridge('/api/dxf-status',{timeoutMs:4000});
    if(String(dxfStatus.root||'').toUpperCase()===String(dxfRoot).toUpperCase() && String(dxfStatus.state||'').toUpperCase()!=='COMPLETE'){
      var dsCount=Number(dxfStatus.filesFound)||0;
      $('buildStatus').textContent='DXF index is still '+String(dxfStatus.state||'not ready').toLowerCase()+' under '+dxfRoot+'. '+dsCount+' DXF files indexed so far. Wait for the index to complete, then create the SigmaNEST job.';
      pill('DXF index not ready','warn');
      return;
    }

    var parts=await readCL();
    var selectedNames=selectedSheets().map(function(s){return s.name});
    var prsRoot=$('libraryPath').value.trim()||'S:\\SNDataX1\\PARTS';
    var job=$('jobName').value.trim()||'CL_JOB';

    var r=await bridge('/api/build-job',{
      method:'POST',
      timeoutMs:30000,
      body:JSON.stringify({
        prsRoot:prsRoot,
        dxfRoot:dxfRoot,
        jobName:job,
        selectedSheetNames:selectedNames,
        parts:parts
      })
    });

    if(!r.accepted || !r.jobId){
      throw new Error(r.message||'The bridge did not accept the SigmaNEST background job.');
    }

    $('buildStatus').textContent=(r.message||'SigmaNEST job accepted.')+' Job ID: '+r.jobId+'. The build will continue independently of the Excel request timeout.';
    pill('Building job','neutral');

    // The HTTP request is now complete. From this point onward Excel only
    // polls short status requests, so a long SigmaNEST build cannot time out
    // the original build operation.
    await waitForBuild(r.jobId,job,selectedNames);
  }catch(e){
    $('buildStatus').textContent=e.message;
    if(e&&e.message&&e.message.indexOf('DXF index')>=0)pill('DXF index not ready','warn');
    else if(e&&e.message&&e.message.indexOf('Cannot connect')>=0)pill('Bridge error','bad');
    else if(String($('statusPill').textContent||'')!=='Build failed')pill('Build error','bad');
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
