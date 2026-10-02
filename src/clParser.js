(function(){
  const ALIASES={
    part:['part','part no','part number','part no.','partname','part name','pn','item','item no','drawing no','dwg no'],
    qty:['qty','quantity','required qty','req qty','amount','qty required'],
    material:['material','material grade','grade','steel'],
    thickness:['thickness','thk','sheet thickness','gauge'],
    qtyPerUnit:['qty/vehicle','qty per vehicle','qty per body','qty per unit','qty per kit'],
    units:['no. of kits','no of kits','no of bodies','number of bodies','kits','bodies']
  };
  function clean(v){return String(v??'').replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim();}
  function normalize(v){return clean(v).toLowerCase().replace(/[^a-z0-9.]+/g,' ').trim();}
  function normalizeMaterial(v){
    let s=clean(v);
    const replacements=[
      [/armoxt\s+advance/ig,'Armox Advance'],[/armoxt/ig,'Armox'],[/amoxt\s+advance/ig,'Armox Advance'],[/amoxt/ig,'Armox'],
      [/ramort/ig,'Ramor'],[/mild\s+steel\s+plate/ig,'Mild Steel'],[/mild\s+steel\s+sheet/ig,'Mild Steel']
    ];
    for(const [rx,repl] of replacements)s=s.replace(rx,repl);
    return s.replace(/\s+/g,' ').trim();
  }
  function thicknessFromMaterial(v){const m=clean(v).match(/(\d+(?:\.\d+)?)\s*mm/i);return m?m[1]+'mm':'';}
  function findColumn(headers,key){
    const exact=headers.map(clean).map(normalize), names=ALIASES[key]||[key];
    for(const a of names){const idx=exact.indexOf(normalize(a));if(idx>=0)return idx;}
    return -1;
  }
  function locateHeaderRow(matrix){
    for(let i=0;i<Math.min(matrix.length,80);i++){
      const headers=(matrix[i]||[]).map(clean);
      const pi=findColumn(headers,'part'), qi=findColumn(headers,'qty'), mi=findColumn(headers,'material');
      if(pi>=0&&(qi>=0||findColumn(headers,'qtyPerUnit')>=0)&&mi>=0)return {row:i,headers};
    }
    return {row:0,headers:(matrix[0]||[]).map(clean)};
  }
  function parseJobNumber(matrix){
    for(const row of matrix){for(const cell of (row||[])){
      const s=clean(cell), m=s.match(/job\s*(?:no\.?|number)?\s*[:#-]?\s*([A-Za-z0-9._/-]+)/i);
      if(m)return m[1];
    }} return '';
  }
  function toNumber(v){
    if(typeof v==='number'&&Number.isFinite(v))return v;
    const m=clean(v).replace(/,/g,'').match(/-?\d+(?:\.\d+)?/);return m?Number(m[0]):NaN;
  }
  function sheetBatchMultiplier(sheetName){const m=clean(sheetName).match(/\(\s*x\s*(\d+(?:\.\d+)?)\s*\)/i);return m?Number(m[1]):1;}
  function parseSheet(matrix,mapping={}){
    if(!matrix||!matrix.length)return [];
    const located=locateHeaderRow(matrix), headers=located.headers, partIdx=findColumn(headers,'part'), qtyIdx=findColumn(headers,'qty');
    const matIdx=findColumn(headers,'material'), thkIdx=findColumn(headers,'thickness'), qpuIdx=findColumn(headers,'qtyPerUnit'), unitsIdx=findColumn(headers,'units');
    if(partIdx<0)throw new Error('Part column not found.');
    const rows=[];
    for(let i=located.row+1;i<matrix.length;i++){
      const r=matrix[i]||[], part=clean(r[partIdx]); if(!part)continue;
      const material=normalizeMaterial(r[matIdx]); let qty=qtyIdx>=0?toNumber(r[qtyIdx]):NaN;
      const qpu=qpuIdx>=0?toNumber(r[qpuIdx]):NaN;
      if(!Number.isFinite(qty)&&Number.isFinite(qpu))qty=qpu;
      const thickness=clean(r[thkIdx])||thicknessFromMaterial(material);
      if(!Number.isFinite(qty))continue;
      const batchMultiplier=sheetBatchMultiplier(mapping.sheet||'');
      const baseQty=qty;
      // qty is always the quantity required for ONE vehicle. The sheet's
      // (xN) multiplier belongs to the task/batch, not the WS part quantity.
      const effectiveQty=baseQty*batchMultiplier;
      rows.push({sheet:mapping.sheet||'',sourceRow:i+1,part,qty:baseQty,baseQty,batchMultiplier,effectiveQty,material,thickness,rawMaterial:clean(r[matIdx]),status:baseQty>0?'READY':(baseQty===0?'ZERO_QTY':'NEGATIVE_QTY')});
    }
    return rows;
  }
  function consolidate(rows){
    const map=new Map();
    for(const r of rows){
      if(r.status!=='READY')continue;
      // Same part + same material + same thickness consolidates normally.
      // A part appearing with different material or thickness remains separate,
      // and is explicitly flagged so the build/review step cannot combine it silently.
      const key=[normalize(r.part).toUpperCase(),normalizeMaterial(r.material).toUpperCase(),normalize(r.thickness).toUpperCase()].join('|');
      if(!map.has(key))map.set(key,{...r,sourceSheets:[r.sheet],sourceRows:[r.sourceRow],taskBatches:[{sheet:r.sheet,batchMultiplier:r.batchMultiplier,vehicleQty:r.baseQty}]});
      else{
        const x=map.get(key);
        x.qty+=r.qty;
        x.sourceSheets.push(r.sheet);
        x.sourceRows.push(r.sourceRow);
        x.taskBatches.push({sheet:r.sheet,batchMultiplier:r.batchMultiplier,vehicleQty:r.baseQty});
      }
    }

    const groups=[...map.values()];
    const byPart=new Map();
    for(const x of groups){
      const partKey=normalize(x.part).toUpperCase();
      if(!byPart.has(partKey))byPart.set(partKey,[]);
      byPart.get(partKey).push(x);
    }

    for(const items of byPart.values()){
      const materialKeys=[...new Set(items.map(x=>normalizeMaterial(x.material).toUpperCase()))].filter(Boolean);
      const thicknessKeys=[...new Set(items.map(x=>normalize(x.thickness).toUpperCase()))].filter(Boolean);
      if(materialKeys.length>1){
        const detail='CL contains this part with multiple materials; quantities were not combined.';
        items.forEach(x=>{x.clReview=true;x.clReviewReason='CL MATERIAL VARIANTS';x.clReviewDetail=detail;});
      }else if(thicknessKeys.length>1){
        const detail='CL contains this part with multiple thicknesses; quantities were not combined.';
        items.forEach(x=>{x.clReview=true;x.clReviewReason='CL THICKNESS VARIANTS';x.clReviewDetail=detail;});
      }
    }
    return groups;
  }

  window.CLParser={parseSheet,consolidate,findColumn,normalizeMaterial,thicknessFromMaterial,parseJobNumber,sheetBatchMultiplier,clean};
})();

