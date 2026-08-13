const statusEl = document.getElementById("libraryStatus");

let pdfjsLib = null;
let selectedFile = null;
let deferredInstallPrompt = null;

const $ = id => document.getElementById(id);
const pdfInput = $("pdfInput"), dropzone = $("dropzone"), convertBtn = $("convertBtn");
const fileLabel = $("fileLabel"), progressWrap = $("progressWrap");
const progressBar = $("progressBar"), progressPct = $("progressPct"), statusText = $("statusText");
const result = $("result"), errorBox = $("error"), installBtn = $("installBtn");

window.addEventListener("beforeinstallprompt", e => {
  e.preventDefault();
  deferredInstallPrompt = e;
  installBtn.classList.remove("hidden");
});
installBtn.addEventListener("click", async () => {
  if (!deferredInstallPrompt) return;
  deferredInstallPrompt.prompt();
  await deferredInstallPrompt.userChoice;
  deferredInstallPrompt = null;
  installBtn.classList.add("hidden");
});

if ("serviceWorker" in navigator && location.protocol !== "file:") {
  navigator.serviceWorker.register("./service-worker.js").catch(err => console.warn("Service worker:", err));
}

async function loadPdfEngine() {
  try {
    pdfjsLib = await import("https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.min.mjs");
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.min.mjs";
    statusEl.textContent = "PDF and Excel engines are ready.";
    statusEl.className = "status ok";
  } catch (e) {
    statusEl.textContent = "PDF engine could not load. Check internet/CDN access and reload.";
    statusEl.className = "status bad";
    console.error(e);
  }
}
loadPdfEngine();

function clean(v){ return String(v ?? "").trim(); }
function norm(v){ return String(v ?? "").replace(/\s+/g," ").trim(); }
const NUMERIC_RE = /^-?\d[\d,]*\.?\d*$/;

function breakAfterComma(v){
  const s = clean(v);
  if (!s || NUMERIC_RE.test(s)) return s;
  return s.split(",").map(x=>x.trim()).filter(Boolean).join(",\n");
}
function isEnglishRow(row){
  const visible = row.filter(Boolean).join(" ").replace(/\s/g,"");
  if (!visible) return false;
  let ascii=0;
  for(const ch of visible) if(ch.charCodeAt(0)<128) ascii++;
  return ascii/visible.length >= .60;
}
function looksLikeHeaderBlock(rows){
  if(rows.length<2) return false;
  const second=rows[1];
  return second.every(c=>/^\(\d+\)$/.test(c)||!clean(c)) && second.some(c=>clean(c));
}
function sameRow(a,b){
  return a.length===b.length && a.every((x,i)=>norm(x)===norm(b[i]));
}

/* Group PDF.js text items into visual lines. */
function groupTextLines(items){
  const cleanItems=items.map(it=>({
    text:norm(it.str),
    x:it.transform[4],
    y:it.transform[5],
    w:it.width||0,
    h:Math.abs(it.transform[3])||10
  })).filter(x=>x.text);

  cleanItems.sort((a,b)=>b.y-a.y || a.x-b.x);
  const lines=[];
  for(const item of cleanItems){
    let line=null;
    for(const candidate of lines){
      const tolerance=Math.max(2.5, Math.min(6,item.h*.45));
      if(Math.abs(candidate.y-item.y)<=tolerance){ line=candidate; break; }
    }
    if(!line){ line={y:item.y,items:[]}; lines.push(line); }
    line.items.push(item);
  }
  lines.sort((a,b)=>b.y-a.y);
  for(const line of lines) line.items.sort((a,b)=>a.x-b.x);
  return lines;
}

/*
  Estimate column anchors from a header/index pair.
  The index row is especially useful because (1),(2),(3) items normally
  sit near the center of their columns.
*/
function anchorsFromHeaderAndIndex(headerItems,indexItems){
  const anchors=[];
  for(const it of indexItems){
    const x=it.x;
    const nearest=headerItems.reduce((best,h)=>{
      const d=Math.abs(h.x-x);
      return !best || d<best.d ? {d,x:h.x} : best;
    },null);
    anchors.push(nearest ? nearest.x : x);
  }
  const unique=[];
  for(const x of anchors.sort((a,b)=>a-b)){
    if(!unique.length || Math.abs(x-unique.at(-1))>7) unique.push(x);
  }
  return unique;
}

function cellsByAnchors(items,anchors){
  if(!anchors.length) return [items.map(i=>i.text).join(" ")];
  const cells=Array(anchors.length).fill("");
  for(const it of items){
    let best=0, dist=Infinity;
    anchors.forEach((a,i)=>{ const d=Math.abs(it.x-a); if(d<dist){dist=d;best=i;} });
    cells[best]=cells[best] ? cells[best]+" "+it.text : it.text;
  }
  return cells.map(clean);
}

function detectPageTables(lines){
  const tables=[];
  let i=0;
  while(i<lines.length){
    const current=lines[i].items;
    const next=lines[i+1]?.items || [];
    const nextIndexes=next.filter(x=>/^\(\d+\)$/.test(clean(x.text)));
    if(nextIndexes.length){
      const header=current.map(x=>x.text);
      const anchors=anchorsFromHeaderAndIndex(current,next);
      const rows=[cellsByAnchors(current,anchors), cellsByAnchors(next,anchors)];
      i+=2;
      while(i<lines.length){
        const items=lines[i].items;
        const text=items.map(x=>x.text).join(" ").trim();
        if(!text){ i++; continue; }
        /* A new header+index pair starts a new table. */
        const following=lines[i+1]?.items||[];
        if(following.filter(x=>/^\(\d+\)$/.test(clean(x.text))).length){
          break;
        }
        rows.push(cellsByAnchors(items,anchors));
        i++;
      }
      tables.push(rows);
    }else{
      i++;
    }
  }
  return tables;
}

function normalizeRows(table){
  const n=Math.max(...table.map(r=>r.length));
  return table.map(r=>Array.from({length:n},(_,i)=>clean(r[i]||"")));
}

function mergeContinuation(rows){
  if(rows.length<2) return rows;
  const out=[rows[0]];
  for(let i=1;i<rows.length;i++){
    const row=rows[i];
    const firstBlank=!clean(row[0]);
    const hasOther=row.slice(1).some(x=>clean(x));
    if(firstBlank && hasOther && out.length>1){
      const prev=out.at(-1);
      for(let c=0;c<prev.length;c++){
        if(clean(row[c])) prev[c]=clean(prev[c]) ? prev[c]+"\n"+clean(row[c]) : clean(row[c]);
      }
    }else if(row.some(x=>clean(x))) out.push(row);
  }
  return out;
}

function mergeAcrossPages(pageTables){
  const result=[];
  for(const table of pageTables){
    if(!table.length) continue;
    const normalized=normalizeRows(table);
    const header=normalized[0];
    const existing=result.at(-1);
    if(existing && sameRow(existing[0],header)){
      existing.push(...normalized.slice(1));
    }else{
      result.push(normalized);
    }
  }
  return result;
}

async function extractTables(file){
  if(!pdfjsLib) throw new Error("PDF engine is not ready yet. Please wait a moment and try again.");
  const pdf=await pdfjsLib.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise;
  const all=[];
  for(let p=1;p<=pdf.numPages;p++){
    setProgress(5+(p-1)/pdf.numPages*58,`Reading PDF page ${p} of ${pdf.numPages}…`);
    const page=await pdf.getPage(p);
    const content=await page.getTextContent({includeMarkedContent:false,disableNormalization:false});
    const lines=groupTextLines(content.items);
    all.push(...detectPageTables(lines));
  }

  let tables=mergeAcrossPages(all);
  if($("englishOnly").checked) tables=tables.filter(t=>isEnglishRow(t[0]));
  if($("mergePages").checked) tables=tables.map(mergeContinuation);

  tables=tables.map(t=>{
    const n=normalizeRows(t);
    return looksLikeHeaderBlock(n) ? [n[0],...n.slice(2)] : n;
  });

  return tables.filter(t=>t.length && t[0].some(x=>clean(x)));
}

function setProgress(percent,text){
  progressWrap.classList.remove("hidden");
  progressBar.style.width=Math.max(0,Math.min(100,percent))+"%";
  progressPct.textContent=Math.round(percent)+"%";
  statusText.textContent=text;
}
function resetMessages(){result.classList.add("hidden");errorBox.classList.add("hidden");}

function safeSheetName(s){
  return String(s).replace(/[\\/*?:[\]]/g," ").slice(0,31)||"Table";
}

async function writeExcel(tables,fileName){
  if(!window.ExcelJS) throw new Error("Excel engine is not loaded. Reload the page.");
  const wb=new ExcelJS.Workbook();
  wb.creator="PDF Table to Excel PWA";
  wb.created=new Date(); wb.modified=new Date();

  tables.forEach((table,index)=>{
    const ws=wb.addWorksheet(safeSheetName(tables.length>1?`Table ${index+1}`:"Table"));
    const header=table[0];
    const hr=ws.addRow(header);
    hr.height=30;
    hr.eachCell(cell=>{
      cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF2563EB"}};
      cell.font={bold:true,color:{argb:"FFFFFFFF"}};
      cell.alignment={wrapText:true,vertical:"middle",horizontal:"center"};
      cell.border=border();
    });

    for(let r=1;r<table.length;r++){
      const values=table[r].map(v=>$("commaBreaks").checked?breakAfterComma(v):clean(v));
      const er=ws.addRow(values);
      er.eachCell(cell=>{
        cell.alignment={wrapText:true,vertical:"top",horizontal:"left"};
        cell.border=border();
      });
    }

    for(let c=1;c<=header.length;c++){
      let max=10;
      for(let r=1;r<=ws.rowCount;r++){
        String(ws.getCell(r,c).value??"").split("\n").forEach(line=>max=Math.max(max,line.length));
      }
      ws.getColumn(c).width=Math.min(max+2,45);
    }
    ws.views=[{state:"frozen",ySplit:1}];
  });

  const buffer=await wb.xlsx.writeBuffer();
  const blob=new Blob([buffer],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const base=fileName.replace(/\.pdf$/i,"")||"converted";
  download(blob,`${base}_converted.xlsx`);
}
function border(){
  const side={style:"thin",color:{argb:"FFD1D5DB"}};
  return {left:side,right:side,top:side,bottom:side};
}
function download(blob,name){
  const url=URL.createObjectURL(blob),a=document.createElement("a");
  a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),1500);
}

function updateFile(){
  const valid=selectedFile && /\.pdf$/i.test(selectedFile.name);
  fileLabel.textContent=valid?selectedFile.name:"Choose a PDF file";
  convertBtn.disabled=!valid;
  resetMessages();
}

pdfInput.addEventListener("change",()=>{selectedFile=pdfInput.files[0]||null;updateFile();});
["dragenter","dragover"].forEach(type=>dropzone.addEventListener(type,e=>{e.preventDefault();dropzone.classList.add("drag");}));
["dragleave","drop"].forEach(type=>dropzone.addEventListener(type,e=>{e.preventDefault();dropzone.classList.remove("drag");}));
dropzone.addEventListener("drop",e=>{
  selectedFile=[...e.dataTransfer.files].find(f=>/\.pdf$/i.test(f.name))||null;
  updateFile();
});

convertBtn.addEventListener("click",async()=>{
  if(!selectedFile)return;
  resetMessages();convertBtn.disabled=true;
  try{
    setProgress(2,"Opening PDF…");
    const tables=await extractTables(selectedFile);
    if(!tables.length) throw new Error("No suitable English table was detected. The PDF may be scanned, or its table structure may not match the browser table detector.");
    setProgress(72,`Creating ${tables.length} Excel sheet(s)…`);
    await writeExcel(tables,selectedFile.name);
    const rows=tables.reduce((n,t)=>n+Math.max(0,t.length-1),0);
    setProgress(100,"Conversion completed.");
    result.textContent=`Success — ${tables.length} table(s), ${rows} data row(s). The Excel file has been downloaded.`;
    result.classList.remove("hidden");
  }catch(err){
    console.error(err);
    errorBox.textContent=`Conversion failed: ${err.message||err}`;
    errorBox.classList.remove("hidden");
    setProgress(0,"Ready");
  }finally{
    convertBtn.disabled=!selectedFile;
  }
});
