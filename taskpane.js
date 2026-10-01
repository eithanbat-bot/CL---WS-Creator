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
      var a=r.slice(0,cols);
      while(a.length<cols)a.push('');
      return a;
    });
  }

  var summary=[
    ['CL - WS Creator',''],
    ['Job',job||''],
    ['Generated',now],
    ['CL workbook',clFile?clFile.name:''],
    ['Selected CL sheets',(selectedSheetNames||[]).join(', ')],
    ['Geometry library root',$('libraryPath').value.trim()],
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
    ['',''],
    ['ACTION / HAND-OFF','DETAIL'],
    ['Geometry search','Recursive .PRS + .DXF search through the configured server folder and all subfolders'],
    ['SigmaNEST WS',result.wsPath||'Not created — review items remain'],
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

  await Excel.run(async function(context){
    var wb=context.workbook;
    var summary=wb.worksheets.getItemOrNullObject(summaryName);
    var reviewSheet=wb.worksheets.getItemOrNullObject(reviewName);
    await context.sync();

    if(summary.isNullObject)summary=wb.worksheets.add(summaryName);
    if(reviewSheet.isNullObject)reviewSheet=wb.worksheets.add(reviewName);
    await context.sync();

    var oldSummary=summary.getUsedRangeOrNullObject();
    var oldReview=reviewSheet.getUsedRangeOrNullObject();
    await context.sync();
    if(!oldSummary.isNullObject)oldSummary.clear('All');
    if(!oldReview.isNullObject)oldReview.clear('All');

    var sm=rectangular(summary,2);
    var sr=summary.getRangeByIndexes(0,0,sm.length,2);
    sr.values=sm;
    sr.format.font.name='Segoe UI';
    sr.format.font.size=10;

    var title=summary.getRange('A1:B1');
    title.merge();
    title.format.font.size=18;
    title.format.font.bold=true;

    var summaryHeaderRow=summary.getRange('A8:B8');
    summaryHeaderRow.format.font.bold=true;
    summaryHeaderRow.format.fill.color='#0B2942';
    summaryHeaderRow.format.font.color='#FFFFFF';

    var actionHeaderRow=summary.getRange('A22:B22');
    actionHeaderRow.format.font.bold=true;
    actionHeaderRow.format.fill.color='#E7EEF5';

    summary.getRange('A1:B'+sm.length).format.wrapText=true;
    summary.getRange('A1:A'+sm.length).format.columnWidth=210;
    summary.getRange('B1:B'+sm.length).format.columnWidth=420;

    var rm=[reviewHeaders].concat(reviewRows);
    var rr=reviewSheet.getRangeByIndexes(0,0,rm.length,reviewHeaders.length);
    rr.values=rectangular(rm,reviewHeaders.length);
    rr.format.font.name='Segoe UI';
    rr.format.font.size=10;
    reviewSheet.getRangeByIndexes(0,0,1,reviewHeaders.length).format.font.bold=true;
    reviewSheet.getRangeByIndexes(0,0,1,reviewHeaders.length).format.fill.color='#0B2942';
    reviewSheet.getRangeByIndexes(0,0,1,reviewHeaders.length).format.font.color='#FFFFFF';
    reviewSheet.getUsedRange().format.wrapText=true;

    var widths=[150,70,130,80,150,260,100,90,170,360,130,100,180];
    widths.forEach(function(w,i){
      reviewSheet.getRangeByIndexes(0,i,Math.min(rm.length,1),1).format.columnWidth=w;
    });
    try{reviewSheet.freezePanes.freezeRows(1)}catch(e){}
    try{summary.freezePanes.freezeRows(8)}catch(e){}

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

async function scan(){
  var btn=$('scanLibrary');
  var started=Date.now();
  btn.disabled=true;
  $('libraryStatus').textContent='Scanning S:\\SNDataX1\\PARTS ...';
  pill('Scanning geometry','neutral');
  try{
    var root=$('libraryPath').value.trim()||'S:\\SNDataX1\\PARTS';
    $('libraryPath').value=root;
    var r=await bridge('/api/scan',{method:'POST',body:JSON.stringify({root:root})});
    var scanErrors=(r.scanErrors||[]).length, inspectErrors=(r.inspectErrors||[]).length;
    var elapsed=Math.round((Date.now()-started)/1000);
    if(!r.count && !r.discoveredFiles){
      $('libraryStatus').textContent='No .PRS or .DXF geometry files were found in the selected folder tree. Scan completed in '+elapsed+'s.';
      pill('No PRS files','warn');
    }else if(scanErrors||inspectErrors){
      var first=(r.scanErrors&&r.scanErrors[0])?(r.scanErrors[0].path+': '+r.scanErrors[0].error):(r.inspectErrors[0]?r.inspectErrors[0].path+': '+r.inspectErrors[0].error:'Unknown scan issue');
      $('libraryStatus').textContent=(r.prsCount||0)+' .PRS + '+(r.dxfCount||0)+' .DXF files indexed; '+(scanErrors+inspectErrors)+' issue(s). First: '+first+' ('+elapsed+'s)';
      pill('Library partial','warn');
    }else{
      $('libraryStatus').textContent=(r.prsCount||0)+' .PRS + '+(r.dxfCount||0)+' .DXF files indexed recursively in '+elapsed+'s.';
      pill('Library ready','ok');
    }
  }catch(e){
    $('libraryStatus').textContent=e.message;
    pill(e.message.indexOf('did not finish')>=0?'Scan timeout':'Bridge error','bad');
  }finally{
    btn.disabled=false;
  }
}

async function build(){
  var btn=$('build');
  btn.disabled=true;
  $('preview').disabled=true;
  $('buildStatus').textContent='Preparing '+$('jobName').value.trim()+'...';
  pill('Creating job','neutral');
  try{
    var parts=await readCL();
    $('buildStatus').textContent='Matching '+parts.length+' CL lines against the PRS library...';
    var root=$('libraryPath').value.trim()||'S:\\SNDataX1\\PARTS';
    var job=$('jobName').value.trim()||('CL_'+Date.now());
    $('buildStatus').textContent='Building '+job+' — please leave the bridge window open...';
    var r=await bridge('/api/build-job',{
      method:'POST',
      body:JSON.stringify({jobName:job,libraryRoot:root,parts:parts}),
      timeoutMs:300000
    });
    showParts(r.parts,'built');
    try{await writeReportSheets(r,job,selectedSheets().map(function(x){return x.name}))}catch(reportErr){$('buildStatus').textContent='Job processed, but Excel report sheets could not be updated: '+reportErr.message;}
    $('geometryFound').textContent=r.parts.filter(function(x){return x.status==='READY'}).length;
    $('geometryMissing').textContent=r.parts.filter(function(x){return x.status!=='READY'}).length;
    if(r.wsPath){
      $('buildStatus').textContent=r.message+'\nWS: '+r.wsPath;
    }else{
      var breakdown=r.reviewBreakdown||{};
      var details=Object.keys(breakdown).map(function(k){return k+': '+breakdown[k]}).join(' | ');
      $('buildStatus').textContent=details?r.message+'\n'+details:r.message;
    }
    pill(r.reviewCount?'Review required':'SigmaNEST WS ready',r.reviewCount?'warn':'ok');
  }catch(e){
    $('buildStatus').textContent=e.message;
    pill(e.message.indexOf('did not finish')>=0?'Build timeout':'Build failed','bad');
  }finally{
    btn.disabled=false;
    updateCount();
  }
}

Office.onReady(async function(info){
  if(info.host!==Office.HostType.Excel){pill('Excel only','bad');return}

  $('chooseCL').onclick=function(){$('clFile').click()};
  $('clFile').onchange=function(){chooseCL(this.files&&this.files[0])};
  $('clearCL').onclick=resetCL;
  $('refreshSheets').onclick=loadSheets;
  $('preview').onclick=preview;
  $('scanLibrary').onclick=scan;
  $('build').onclick=build;
  $('sheetList').onchange=updateCount;

  $('libraryPath').value=localStorage.getItem('clwsc_libraryRoot')||'S:\\SNDataX1\\PARTS';
  $('libraryPath').onchange=function(){localStorage.setItem('clwsc_libraryRoot',$('libraryPath').value.trim())};

  resetCL();

  try{
    var h=await bridge('/api/health');
    $('libraryStatus').textContent='Bridge connected (v'+(h.bridgeVersion||'?')+'). Library: '+h.libraryRoot+(h.prsCount!=null?' | PRS '+h.prsCount+' | DXF '+(h.dxfCount||0):'');
    pill('Connected','ok');
  }catch(e){
    $('libraryStatus').textContent='Start start-bridge.bat on this PC.';
    pill('Bridge offline','warn');
  }
});
