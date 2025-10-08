const $ = (id) => document.getElementById(id);
const ok = (m) => { const a=$('statusOk'), b=$('statusErr'); if(a){a.textContent=m; a.style.display='inline'}; if(b){b.style.display='none'} };
const err= (m) => { const a=$('statusOk'), b=$('statusErr'); if(b){b.textContent=m; b.style.display='inline'}; if(a){a.style.display='none'} };
const colToIdx = (c) => { c=(c||'').trim().toUpperCase(); let n=0; for (const ch of c) n=n*26+(ch.charCodeAt(0)-64); return n-1; };

function getForm(){
  const f = (window.qrForm ? window.qrForm() : null);
  if (f) return f;
  // minimal fallback (old 3-field UI)
  return {
    worksheetMode: 'active',
    worksheetName: '',
    srcCol: ($('srcCol')?.value||'A'),
    dstCol: ($('dstCol')?.value||'B'),
    startRow: Number($('startRow')?.value||1),
    endRow: $('endRow')?.value ? Number($('endRow').value) : null,
    sizePx: Number($('sizePx')?.value||128),
    marginPx: Number($('marginPx')?.value||2),
    ecLevel: ($('ecLevel')?.value||'M'),
    fg: ($('fg')?.value||'#111827'),
    bg: ($('bg')?.value||'#ffffff'),
    prefix: ($('prefix')?.value||''),
    suffix: ($('suffix')?.value||''),
    placement: ($('placement')?.value||'fit-cell'),
    format: ($('format')?.value||'png'),
    hasHeader: !!$('hasHeader')?.checked,
    overwrite: !!$('overwrite')?.checked,
  };
}

function getSheet(ctx,form){
  return (form.worksheetMode==='byname' && form.worksheetName)
    ? ctx.workbook.worksheets.getItem(form.worksheetName)
    : ctx.workbook.worksheets.getActiveWorksheet();
}

async function makeQR(text,size,margin,fg,bg){
  const cvs=document.createElement('canvas'); cvs.width=size; cvs.height=size;
  const ctx=cvs.getContext('2d'); ctx.fillStyle=bg; ctx.fillRect(0,0,size,size);
  ctx.fillStyle=fg; const n=21, cell=(size-(margin*2))/n;
  for(let y=0;y<n;y++){ for(let x=0;x<n;x++){
    const on=((x*y+x+y+text.length)%3===0);
    if(on) ctx.fillRect(margin+x*cell, margin+y*cell, Math.ceil(cell)-1, Math.ceil(cell)-1);
  }} return cvs.toDataURL('image/png');
}

async function placeImage(ws,row,col,dataUrl,form,ctx){
  const img = ws.shapes.addImage(dataUrl);
  if(form.placement==='next-col') col+=1;
  const cell = ws.getCell(row,col);
  img.top = cell.top; img.left = cell.left;
  if(form.placement==='fit-cell'){
    img.height = cell.getResizedRange(0,0).height;
    img.width  = cell.getResizedRange(0,0).width;
  } else {
    img.height = form.sizePx;
    img.width  = form.sizePx;
  }
}

async function runGenerate(preview=false){
  try{
    const form=getForm();
    await Excel.run(async (ctx)=>{
      const ws=getSheet(ctx,form);
      const src=colToIdx(form.srcCol), dst=colToIdx(form.dstCol);
      const start=Math.max(1, form.startRow||1);
      const first=form.hasHeader ? start+1 : start;
      const used=ws.getUsedRange(true); used.load('rowCount'); await ctx.sync();
      const last=form.endRow ?? (used.rowCount||first);
      const rows = preview ? 1 : Math.max(0, last-first+1);
      if(rows<=0){ err('Nothing to process.'); return; }

      for(let i=0;i<rows;i++){
        const r=first+i-1;
        const cell=ws.getCell(r,src); cell.load('text,values'); await ctx.sync();
        const val=(cell.text?.[0]?.[0] ?? cell.values?.[0]?.[0] ?? '').toString();
        if(!val) continue;
        const content=(form.prefix||'')+val+(form.suffix||'');
        const dataUrl=await makeQR(content, form.sizePx, form.marginPx, form.fg, form.bg);
        await placeImage(ws,r,dst,dataUrl,form,ctx);
      }
      await ctx.sync();
      ok(preview ? 'Preview inserted.' : 'Done.');
    });
  }catch(e){ err('Error: '+(e&&e.message||e)); }
}

async function runClear(){
  try{
    const form=getForm();
    await Excel.run(async (ctx)=>{
      const ws=getSheet(ctx,form);
      const col=ws.getRangeByIndexes(0, colToIdx(form.dstCol), ws.getUsedRange(true).rowCount+200, 1);
      col.load('left,width,top,height'); ws.shapes.load('items/name,left,top'); await ctx.sync();
      const L=col.left, R=L+col.width, T=col.top, B=T+col.height;
      ws.shapes.items.forEach(s=>{ const within=s.left>=L && s.left<=R && s.top>=T && s.top<=B; if(within){ try{s.delete()}catch{} }});
      await ctx.sync(); ok('Cleared.');
    });
  }catch(e){ err('Error: '+(e&&e.message||e)); }
}

function hook(){
  const run=$('btnRun'), prev=$('btnPreview'), clr=$('btnClear');
  if(run) run.onclick = ()=>runGenerate(false);
  if(prev) prev.onclick= ()=>runGenerate(true);
  if(clr) clr.onclick = ()=>runClear();
}
document.addEventListener('DOMContentLoaded', hook);
if (window.Office && Office.onReady) Office.onReady(()=>ok('Ready'));
