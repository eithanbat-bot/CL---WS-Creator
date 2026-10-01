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
    tr.innerHTML='<td>'+esc(p.part)+'</td><td>'+p.qty+'</td><td>'+esc(p.material)+'</td><td>'+esc(p.thickness)+'</td><td class="'+cls+'">'+esc(st)+'</td>';
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
  pill('Scanning PRS','neutral');
  try{
    var root=$('libraryPath').value.trim()||'S:\\SNDataX1\\PARTS';
    $('libraryPath').value=root;
    var r=await bridge('/api/scan',{method:'POST',body:JSON.stringify({root:root})});
    var scanErrors=(r.scanErrors||[]).length, inspectErrors=(r.inspectErrors||[]).length;
    var elapsed=Math.round((Date.now()-started)/1000);
    if(!r.count && !r.discoveredFiles){
      $('libraryStatus').textContent='No .PRS files were found in the selected folder. Scan completed in '+elapsed+'s.';
      pill('No PRS files','warn');
    }else if(scanErrors||inspectErrors){
      var first=(r.scanErrors&&r.scanErrors[0])?(r.scanErrors[0].path+': '+r.scanErrors[0].error):(r.inspectErrors[0]?r.inspectErrors[0].path+': '+r.inspectErrors[0].error:'Unknown scan issue');
      $('libraryStatus').textContent=r.count+' .PRS files indexed; '+(scanErrors+inspectErrors)+' file/folder read issue(s). First: '+first+' ('+elapsed+'s)';
      pill('Library partial','warn');
    }else{
      $('libraryStatus').textContent=r.count+' .PRS files indexed in '+elapsed+'s.';
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
    $('geometryFound').textContent=r.parts.filter(function(x){return x.status==='READY'}).length;
    $('geometryMissing').textContent=r.parts.filter(function(x){return x.status!=='READY'}).length;
    $('buildStatus').textContent=r.wsPath?(r.message+'\nWS: '+r.wsPath):r.message;
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
    $('libraryStatus').textContent='Bridge connected. Library: '+h.libraryRoot;
    pill('Connected','ok');
  }catch(e){
    $('libraryStatus').textContent='Start start-bridge.bat on this PC.';
    pill('Bridge offline','warn');
  }
});
