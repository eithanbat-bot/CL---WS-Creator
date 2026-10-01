const http=require('http'),fs=require('fs'),path=require('path'),cp=require('child_process');const {inspectPrs}=require('./prsInspector');
const PORT=17832,DEFAULT_LIBRARY=process.env.SN_PARTS||'S:\\SNDataX1\\PARTS',ROOT=__dirname;let index=[];const cfgFile=path.join(__dirname,'config.json');let cfg={libraryRoot:DEFAULT_LIBRARY,lastScan:null,count:0};try{if(fs.existsSync(cfgFile))cfg=Object.assign(cfg,JSON.parse(fs.readFileSync(cfgFile,'utf8')))}catch(e){}
function reply(res,status,data){const body=JSON.stringify(data,null,2);res.writeHead(status,{'Content-Type':'application/json; charset=utf-8','Access-Control-Allow-Origin':'*','Access-Control-Allow-Headers':'Content-Type, Accept','Access-Control-Allow-Methods':'GET, POST, OPTIONS','Access-Control-Allow-Private-Network':'true','Cache-Control':'no-store'});res.end(body)}
function readBody(req){return new Promise(function(resolve,reject){let b='';req.on('data',function(c){b+=c});req.on('end',function(){try{resolve(b?JSON.parse(b):{})}catch(e){reject(e)}})})}
function collect(dir,out,errors){out=out||[];errors=errors||[];let entries;
  try{entries=fs.readdirSync(dir,{withFileTypes:true});}
  catch(e){errors.push({path:dir,error:e.message});return out}
  for(const ent of entries){
    const full=path.join(dir,ent.name);
    try{
      if(ent.isDirectory())collect(full,out,errors);
      else if(/\.prs$/i.test(ent.name))out.push(full);
    }catch(e){errors.push({path:full,error:e.message});}
  }
  return out;
}
function norm(s){return String(s||'').toUpperCase().replace(/[^A-Z0-9]+/g,'').replace(/PRS$/,'')}
function vkey(s){return norm(s).replace(/([A-Z])$/,'')}
function ensure(root){
  if(index.length&&String(cfg.libraryRoot).toUpperCase()===root.toUpperCase())return {scanErrors:cfg.scanErrors||[],inspectErrors:cfg.inspectErrors||[]};
  const scanErrors=[],files=collect(root,[],scanErrors),inspectErrors=[];
  const parsed=[];
  for(const file of files){
    try{parsed.push(inspectPrs(file));}
    catch(e){inspectErrors.push({path:file,error:e.message});}
  }
  index=parsed;
  cfg={libraryRoot:root,lastScan:new Date().toISOString(),count:index.length,discoveredFiles:files.length,scanErrors:scanErrors,inspectErrors:inspectErrors};
  try{fs.writeFileSync(cfgFile,JSON.stringify(cfg,null,2));}catch(e){}
  return {scanErrors:scanErrors,inspectErrors:inspectErrors};
}
function findPart(part){const n=norm(part);let h=index.find(function(x){return norm(x.partName)===n});if(h)return Object.assign({},h,{matchType:'EXACT'});h=index.find(function(x){return norm(x.embeddedPartName)===n||norm(x.embeddedPartName).indexOf(n+'-')===0});if(h)return Object.assign({},h,{matchType:'EMBEDDED'});const vk=vkey(part),vars=index.filter(function(x){return vkey(x.partName)===vk});if(vars.length===1)return Object.assign({},vars[0],{matchType:'VARIATION'});return vars.length?{ambiguous:vars}:null}
function normMat(s){return String(s||'').replace(/Armoxt|Amoxt/ig,'Armox').replace(/Ramort/ig,'Ramor').replace(/\s+/g,' ').trim()}
function matEq(a,b){const A=normMat(a).toUpperCase(),B=normMat(b).toUpperCase();if(!A||!B||A===B)return true;if(/4\s*MM\s*ARMOX/.test(A)&&/RAMOR\s*500/.test(B))return true;if(/4\s*MM\s*ARMOX/.test(B)&&/RAMOR\s*500/.test(A))return true;return A.indexOf(B)>=0||B.indexOf(A)>=0}
function thk(s){const m=String(s||'').match(/(\d+(?:\.\d+)?)\s*mm/i);return m?Number(m[1]):NaN}
function sigmaMaterial(clMat,libMat){let s=normMat(libMat||clMat);s=s.replace(/^\d+(?:\.\d+)?\s*mm\s*/i,'').replace(/\s+(sheet|plate)$/i,'').trim();return s}
function csv(v){const s=String(v==null?'':v);return /[",\n]/.test(s)?('"'+s.replace(/"/g,'""')+'"'):s}
function writeJob(root,name,parts){const out=path.join(root,'_CL_WS_BUILDER',name);fs.mkdirSync(path.join(out,'parts'),{recursive:true});parts.filter(function(p){return p.file}).forEach(function(p){fs.copyFileSync(p.file,path.join(out,'parts',path.basename(p.file)))});const rows=['Part,Qty,Material,Thickness,PRS,MatchType,PRS_Material,PRS_Thickness,SourceDXF,Status'];parts.forEach(function(p){rows.push([p.part,p.qty,p.material,p.thickness,p.prs?path.basename(p.prs):'',p.matchType||'',p.libraryMaterial||'',p.libraryThickness||'',p.sourceDxf||'',p.status].map(csv).join(','))});fs.writeFileSync(path.join(out,'WS_PARTS.csv'),rows.join('\n'));fs.writeFileSync(path.join(out,'SIGMANEST_JOB.json'),JSON.stringify({schema:'cl-ws-creator/0.3',jobName:name,libraryRoot:root,created:new Date().toISOString(),parts:parts},null,2));return out}
function createSigmaNest(parts,jobName,libraryRoot,wsDirectory){const reqPath=path.join(ROOT,'_request-'+process.pid+'-'+Date.now()+'.json');const req={jobName:jobName,libraryRoot:libraryRoot,wsDirectory:wsDirectory||'',parts:parts.map(function(p){return {part:p.part,qty:p.qty,batchMultiplier:p.batchMultiplier||1,taskSheet:p.sheet||'',prsPath:p.file,sigmaMaterial:sigmaMaterial(p.material,p.libraryMaterial),thicknessMm:thk(p.thickness)}})};fs.writeFileSync(reqPath,JSON.stringify(req,null,2));const ps=path.join(ROOT,'create-sigmanest-ws.ps1');const result=cp.spawnSync('powershell.exe',['-NoProfile','-ExecutionPolicy','Bypass','-File',ps,'-RequestFile',reqPath],{encoding:'utf8',timeout:120000,windowsHide:true});try{fs.unlinkSync(reqPath)}catch(e){}const stdout=String(result.stdout||'').trim(),stderr=String(result.stderr||'').trim();let data=null;try{data=JSON.parse(stdout)}catch(e){}if(result.error)throw new Error(result.error.message);if(result.status!==0||!data||!data.ok)throw new Error((data&&data.error)||stderr||stdout||('SigmaNEST worker exited with code '+result.status));return data}
async function handle(req,res){if(req.method==='OPTIONS'){res.writeHead(204,{'Access-Control-Allow-Origin':'*','Access-Control-Allow-Headers':'Content-Type, Accept, X-Requested-With','Access-Control-Allow-Methods':'GET, POST, OPTIONS','Access-Control-Allow-Private-Network':'true','Access-Control-Max-Age':'600','Content-Length':'0'});return res.end()};try{
if(req.url==='/api/health')return reply(res,200,{ok:true,port:PORT,libraryRoot:cfg.libraryRoot,lastScan:cfg.lastScan,count:cfg.count,discoveredFiles:cfg.discoveredFiles||cfg.count,scanErrors:cfg.scanErrors||[],inspectErrors:cfg.inspectErrors||[],sigmaNestCom:true});
if(req.url==='/api/scan'&&req.method==='POST'){
  const b=await readBody(req),rawRoot=String(b.root||DEFAULT_LIBRARY).trim();
  if(!rawRoot)return reply(res,400,{error:'Enter the SigmaNEST .PRS library folder path.'});
  const root=path.resolve(rawRoot);
  let st;
  try{st=fs.statSync(root);}catch(e){
    return reply(res,400,{error:'Cannot access the PRS library folder: '+root+' | '+e.message,root:root});
  }
  if(!st.isDirectory())return reply(res,400,{error:'Library path is not a folder: '+root,root:root});
  const diag=ensure(root);
  return reply(res,200,{
    count:index.length,
    discoveredFiles:cfg.discoveredFiles||index.length,
    scanErrors:diag.scanErrors,
    inspectErrors:diag.inspectErrors,
    root:root,
    parts:index.map(function(x){return {partName:x.partName,embeddedPartName:x.embeddedPartName,likelyMaterial:x.likelyMaterial,thickness:x.thickness,sourceDxf:x.sourceDxf}})
  });
}
if(req.url==='/api/build-job'&&req.method==='POST'){const b=await readBody(req),root=path.resolve(b.libraryRoot||DEFAULT_LIBRARY);if(!fs.existsSync(root))return reply(res,400,{error:'Library folder does not exist: '+root});ensure(root);const name=(String(b.jobName||'CL_JOB').replace(/[^A-Za-z0-9._ -]/g,'_').trim()||'CL_JOB');const parts=(b.parts||[]).map(function(p){const f=findPart(p.part);if(!f)return Object.assign({},p,{status:'MISSING',statusLabel:'GEOMETRY MISSING'});if(f.ambiguous)return Object.assign({},p,{status:'REVIEW',statusLabel:'AMBIGUOUS ('+f.ambiguous.length+')'});const libMat=f.likelyMaterial||'',libThk=f.thickness||'',matKnown=!!libMat.trim(),thkKnown=!!libThk.trim(),mok=matKnown&&matEq(p.material,libMat),tok=thkKnown&&!!p.thickness&&thk(p.thickness)===thk(libThk),variation=f.matchType==='VARIATION',status=matKnown&&thkKnown&&mok&&tok&&!variation?'READY':'REVIEW',label=status==='READY'?'FOUND':variation?'VARIATION - REVIEW':(!matKnown?'MATERIAL NOT CONFIRMED':(!mok?'MATERIAL MISMATCH':(!thkKnown?'THICKNESS NOT CONFIRMED':(!tok?'THICKNESS MISMATCH':'REVIEW'))));return Object.assign({},p,{status:status,statusLabel:label,file:f.file,prs:f.file,sourceDxf:f.sourceDxf,matchType:f.matchType,libraryMaterial:libMat,libraryThickness:libThk})});const review=parts.filter(function(p){return p.status!=='READY'});const staging=writeJob(root,name,parts);if(review.length)return reply(res,200,{outputDir:staging,message:'Job staged, but '+review.length+' part(s) require review before SigmaNEST creation.',parts:parts,reviewCount:review.length,sigmaNestCreated:false});let sigma=null;if(b.createSigmaNEST!==false){sigma=createSigmaNest(parts,name,b.libraryRoot||DEFAULT_LIBRARY,b.wsDirectory||'')}return reply(res,200,{outputDir:staging,message:sigma?sigma.message:('Job staged with '+parts.length+' part(s).'),parts:parts,reviewCount:0,sigmaNestCreated:!!sigma,wsPath:sigma?sigma.wsPath:null,sigmaPartCount:sigma?sigma.partCount:0,taskPlan:parts.map(function(p){return {sheet:p.sheet,batchMultiplier:p.batchMultiplier||1,part:p.part,qty:p.qty}})})}
return reply(res,404,{error:'Not found'});}catch(e){return reply(res,500,{error:e.message})}}
http.createServer(handle).listen(PORT,'127.0.0.1',function(){console.log('CL-WS-Creator bridge listening on http://127.0.0.1:'+PORT)});
