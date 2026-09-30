const fs=require('fs'),path=require('path');
function strings8(b){return b.toString('latin1').match(/[\x20-\x7e]{4,}/g)||[]}
function strings16(b){const out=[];for(let i=0;i<b.length-3;i++){let j=i;while(j+1<b.length&&b[j+1]===0&&b[j]>=32&&b[j]<=126)j+=2;if(j-i>=6){out.push(b.subarray(i,j+1).toString('utf16le').replace(/\0+$/,'').trim());i=j}}return out}
function inspectPrs(file){
  const b=fs.readFileSync(file),strings=[...new Set([...strings8(b),...strings16(b)].filter(Boolean))],stem=path.basename(file,path.extname(file));
  const embedded=strings.find(function(s){return s.indexOf(stem)===0})||stem;
  const sourceDxf=strings.find(function(s){return /\.(dxf|dwg)$/i.test(s)})||'';
  const material=strings.find(function(s){return /armox|ramor|s355|s690|hardox|chromodeck|mild steel|strenx|aluminium|aluminum|stainless/i.test(s)})||'';
  const m=material.match(/(\d+(?:\.\d+)?)\s*mm/i)||embedded.match(/-\s*(\d+(?:\.\d+)?)\s*mm/i)||[];
  const rotations=strings.find(function(s){return /^0\s*,?\s*90\s*,?\s*180\s*,?\s*270$/i.test(s)})||'';
  return {file:path.resolve(file),fileName:path.basename(file),partName:stem,embeddedPartName:embedded,likelyMaterial:material,thickness:m[1]?(m[1]+'mm'):'',sourceDxf:sourceDxf,rotations:rotations};
}
module.exports={inspectPrs};
