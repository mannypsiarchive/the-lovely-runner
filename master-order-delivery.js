/* Master Order Helper 1.16 test — browser-only Google Sheets and local video workflow. */
const MOD_FFMPEG_VERSION="0.12.10", MOD_FFMPEG_SCRIPT=`https://cdn.jsdelivr.net/npm/@ffmpeg/ffmpeg@0.12.10/dist/umd/ffmpeg.js`, MOD_FFMPEG_CORE=`https://cdn.jsdelivr.net/npm/@ffmpeg/core@0.12.10/dist/umd`, MOD_FFMPEG_WORKER=`https://cdn.jsdelivr.net/npm/@ffmpeg/ffmpeg@0.12.10/dist/umd/814.ffmpeg.js`;
const MOD_CLIENT_ID = "154634144934-9hg9o4ra7uriu5hrivaaj73mduj7udf4.apps.googleusercontent.com";
const MOD_DISCOVERY = "https://sheets.googleapis.com/$discovery/rest?version=v4";
const MOD_SCOPE = "https://www.googleapis.com/auth/spreadsheets";
const modState = { id:"", title:"", tabs:[], sourceHeaders:[], sourceRows:[], vendors:[], rates:new Map(), assignments:new Map(), activeSource:"", tokenClient:null, gapiReady:false, gisReady:false, accessToken:"", mapping:{}, cut:null, cutObjectUrl:"", activeRangeVideo:null, screeners:[], screenerMetadata:new Map(), screenerMetadataPromise:null, shots:[], frameObserver:null, frameCleanupObserver:null, ensureShotFrames:null, releaseShotFrames:null, sourceDurationIndex:-1, exportingAe:false, excelJsPromise:null, worker:null, paddleOcr:null, paddleLoadPromise:null, ffmpeg:null, mediaMetadata:null, metadataPromise:null, ffmpegLogs:[], aeTapeTitles:new Map(), aeReviewMode:false, aeReviewIndex:0 };
const mod$ = id => document.getElementById(id);
const MOD_FIELDS = [
  {key:"vendor",label:"Source / Vendor",aliases:["source","vendor","vendor/source","archive source"]},
  {key:"type",label:"EDL Type",aliases:["type","archival class","media type","class"]},
  {key:"filename",label:"Original File Name",aliases:["original file name","original filename","source file name","source filename","original file"]},
  {key:"sftype",label:"Still / Footage (S/F)",aliases:["still(s)/ftg(f)","still/footage","still footage","s/f","media type"]},
  {key:"srcIn",label:"Source In",aliases:["source in","source in tc","src in","source in timecode"]},
  {key:"srcOut",label:"Source Out",aliases:["source out","source out tc","src out","source out timecode"]},
  {key:"timeIn",label:"Cut / Timeline In",aliases:["time in","timeline in","program in","record in"]},
  {key:"timeOut",label:"Cut / Timeline Out",aliases:["time out","timeline out","program out","record out"]},
  {key:"description",label:"Description / Notes",aliases:["description","file description","notes","comments"]},
  {key:"arc",label:"ARC / Final File Name",aliases:["final file name","arc number","arc","file name"]}
];
function modStatus(message,kind=""){const el=mod$("modConnectionStatus");if(el){el.textContent=message;el.className="mod-status "+kind;}mod$("modStatusPill").textContent=kind==="error"?"Needs attention":kind==="good"?"Connected":"Ready";}
function modErr(e){return e?.result?.error?.message||e?.message||String(e);}
function modCellText(v){if(v==null)return "";if(typeof v==="object")return String(v.text??v.result??"").trim();return String(v).trim();}
function modTcNorm(v){return modCellText(v).replace(/\s+/g,"").replace(/\./g,":");}
function modExtractTc(v){return modCellText(v).match(/\d{2}:\d{2}:\d{2}[:;]\d{2}/)?.[0]||"";}
function modRate(v){const s=modCellText(v);const f=s.match(/(\d+(?:\.\d+)?)\s*\/\s*(\d+(?:\.\d+)?)/);if(f&&Number(f[2]))return Number(f[1])/Number(f[2]);const n=Number(s);return Number.isFinite(n)&&n>0?n:null;}
function modNominal(fps){return Math.abs(fps-23.976)<.02?24:Math.abs(fps-29.97)<.02?30:Math.abs(fps-59.94)<.02?60:Math.round(fps);}
function modParseTcFrames(tc,fps){const raw=modTcNorm(tc),p=raw.replace(";",":").split(":").map(Number),nom=modNominal(fps);if(p.length!==4||p.some(x=>!Number.isFinite(x))||p[1]>59||p[2]>59||p[3]>=nom)throw Error(`Invalid timecode: ${tc}`);let n=((p[0]*60+p[1])*60+p[2])*nom+p[3];if(raw.includes(";")&&(Math.abs(fps-29.97)<.02||Math.abs(fps-59.94)<.02))n-=(nom===60?4:2)*(p[0]*60+p[1]-Math.floor((p[0]*60+p[1])/10));return n;}
function modMakeMetadata(fps,start,source){fps=Math.abs(fps-23.98)<.02?23.976:Math.abs(fps-29.98)<.03?29.97:Math.abs(fps-59.96)<.03?59.94:fps;start=modExtractTc(start);if(!fps||!start)throw Error("The file metadata did not include both a starting timecode and frame rate. The helper will not guess these values.");return{fps,startTimecode:start,source};}
function modMetadataFromProbe(p){const streams=p?.streams||[],v=streams.find(x=>x.codec_type==="video")||{},tags={...(p?.format?.tags||{}),...(v.tags||{})},tc=Object.entries(tags).filter(([k])=>/time.?code|start.?time/i.test(k)).map(([,x])=>modExtractTc(x)).find(Boolean);return modMakeMetadata(modRate(v.avg_frame_rate)||modRate(v.r_frame_rate),tc,"FFmpeg");}
function modMediaInfoMetadata(result){let tracks=result?.media?.track||result?.Media?.track||[];if(!Array.isArray(tracks))tracks=[tracks];const find=t=>tracks.find(x=>modCellText(x?.["@type"]).toLowerCase()===t.toLowerCase())||{},vid=find("Video"),gen=find("General"),get=(o,names)=>{for(const name of names){const k=Object.keys(o).find(x=>x.toLowerCase()===name.toLowerCase());if(k&&modCellText(o[k]))return o[k];}return "";};const a=get(vid,["FrameRate_Num","FrameRateNumerator"]),b=get(vid,["FrameRate_Den","FrameRateDenominator"]),fps=a&&b?Number(a)/Number(b):modRate(get(vid,["FrameRate","FrameRate_Original","FrameRate_Nominal"]));const tc=[get(vid,["TimeCode_FirstFrame","TimeCode_FirstFrame_Original","TimeCode_Start"]),get(gen,["TimeCode_FirstFrame","TimeCode","TimeCode_Start"]),get(vid,["TimeCode"]),...tracks.flatMap(x=>Object.entries(x||{}).filter(([k,v])=>/time.?code|start.?time/i.test(k)&&modCellText(v)).map(([,v])=>v))].map(modExtractTc).find(Boolean);return modMakeMetadata(fps,tc,"MediaInfo");}
function modFFmpegClass(){return window.FFmpegWASM?.FFmpeg||window.FFmpeg?.FFmpeg||(typeof window.FFmpeg==="function"?window.FFmpeg:null);}
function modLoadScript(url){return new Promise((resolve,reject)=>{const s=document.createElement("script");s.src=url;s.onload=resolve;s.onerror=()=>reject(Error("Could not load FFmpeg metadata tools."));document.head.append(s);});}
async function modBlobUrl(url,type){const r=await fetch(url);if(!r.ok)throw Error(`Could not load FFmpeg core (${r.status}).`);return URL.createObjectURL(new Blob([await r.blob()],{type}));}
async function modMediaInfoScan(file){let factory=window.MediaInfo?.default||window.MediaInfo?.mediaInfoFactory||window.MediaInfo||window.mediaInfoFactory;if(!factory){await modLoadScript("https://unpkg.com/mediainfo.js@0.3.7/dist/umd/index.min.js");factory=window.MediaInfo?.default||window.MediaInfo?.mediaInfoFactory||window.MediaInfo||window.mediaInfoFactory;}if(!factory)throw Error("MediaInfo did not load.");const mi=await factory({format:"object",full:true}),result=await mi.analyzeData(file.size,async(size,offset)=>new Uint8Array(await file.slice(offset,offset+size).arrayBuffer()));return modMediaInfoMetadata(result);}
async function modDetectMetadata(file,isScreener=false){
  modState.mediaMetadata=isScreener?modState.mediaMetadata:null;
  const status=isScreener?null:mod$("modMetadataStatus");
  if(status){status.className="mod-status";status.textContent="Reading embedded start timecode and frame rate with MediaInfo…";}
  let metadata=null,mediaInfoError=null,ffmpegError=null;
  // MediaInfo reads metadata in small file slices; try it first so multi-gigabyte cuts
  // do not need to be copied into one in-memory FFmpeg buffer for the common path.
  try{metadata=await modMediaInfoScan(file);}catch(e){mediaInfoError=e;}
  if(!metadata){
    try{
      if(!modFFmpegClass())await modLoadScript(MOD_FFMPEG_SCRIPT);
      const FFmpeg=modFFmpegClass();if(!FFmpeg)throw Error("FFmpeg library did not load.");
      if(!modState.ffmpeg){const urls=await Promise.all([modBlobUrl(MOD_FFMPEG_CORE+"/ffmpeg-core.js","text/javascript"),modBlobUrl(MOD_FFMPEG_CORE+"/ffmpeg-core.wasm","application/wasm"),modBlobUrl(MOD_FFMPEG_WORKER,"text/javascript")]);modState.ffmpeg=new FFmpeg();await modState.ffmpeg.load({coreURL:urls[0],wasmURL:urls[1],classWorkerURL:urls[2]});modState.ffmpeg.on?.("log",({message})=>modState.ffmpegLogs.push(message));}
      const ff=modState.ffmpeg,name="cut-"+Date.now()+"-"+file.name.replace(/[^\w.-]/g,"_");modState.ffmpegLogs=[];
      await ff.writeFile(name,new Uint8Array(await file.arrayBuffer()));
      try{if(typeof ff.ffprobe==="function"){await ff.ffprobe(["-v","error","-select_streams","v:0","-show_entries","stream=avg_frame_rate,r_frame_rate,width,height:stream_tags=timecode:format=duration:format_tags=timecode","-of","json",name,"-o","mod-probe.json"]);metadata=modMetadataFromProbe(JSON.parse(new TextDecoder().decode(await ff.readFile("mod-probe.json"))));}}catch(e){ffmpegError=e;}
      if(!metadata){modState.ffmpegLogs=[];try{await ff.exec(["-hide_banner","-i",name]);}catch{}const report=modState.ffmpegLogs.join("\n"),tc=report.match(/time.?code\s*[:=]\s*([^\s,]+)/i)?.[1]||report.match(/time.?code[^\n]*?(\d{2}:\d{2}:\d{2}[:;]\d{2})/i)?.[1],rate=report.match(/(\d+(?:\.\d+)?)\s+(?:fps|tbr)\b/i)?.[1];try{metadata=modMakeMetadata(modRate(rate),tc,"FFmpeg report");}catch(e){ffmpegError=e;}}
      try{await ff.deleteFile(name);}catch{}try{await ff.deleteFile("mod-probe.json");}catch{}
    }catch(e){ffmpegError=e;}
  }
  if(!metadata){const failure=mediaInfoError||ffmpegError||Error("Could not detect the file metadata.");if(status){status.className="mod-status error";status.textContent="Automatic detection failed: "+modErr(failure)+" Please confirm the file contains embedded start timecode metadata.";}throw failure;}
  if(isScreener)return metadata;
  modState.mediaMetadata=metadata;
  if(status){status.className="mod-status good";status.textContent=`Detected from ${metadata.source}: file starts at ${metadata.startTimecode} · ${metadata.fps} fps. EDL timecodes will be matched frame for frame.`;}
  mod$("modTimelineStart").value=metadata.startTimecode;mod$("modFps").value=String(metadata.fps);document.querySelector(".mod-cut-options")?.classList.remove("hidden");return metadata;
}
function modScreenerKey(file){return file.webkitRelativePath||file.name;}
async function modDetectScreenerMetadata(files){modState.screenerMetadata.clear();mod$("modScreenerMetadataDetails").classList.add("hidden");mod$("modScreenerMetadataList").replaceChildren();const status=mod$("modScreenerMetadataStatus");if(!files.length){status.textContent="Choose the raw screener folder to detect each file’s start timecode and frame rate.";return;}status.className="mod-status";status.textContent=`Detecting metadata for ${files.length} screener file${files.length===1?"":"s"}…`;const errors=[];for(let i=0;i<files.length;i++){const file=files[i];try{let meta;try{meta=await modMediaInfoScan(file);}catch{meta=await modDetectMetadata(file,true);}modState.screenerMetadata.set(modScreenerKey(file),meta);}catch(e){errors.push(file.name+": "+modErr(e));}status.textContent=`Detected screener metadata for ${modState.screenerMetadata.size} of ${files.length} files (${i+1}/${files.length} scanned).`;}const list=mod$("modScreenerMetadataList"),details=mod$("modScreenerMetadataDetails"),table=document.createElement("table");table.className="mod-table";const head=document.createElement("tr");["Screener file","Start timecode","Frame rate","Read by"].forEach(text=>{const th=document.createElement("th");th.textContent=text;head.append(th);});const thead=document.createElement("thead");thead.append(head);const tbody=document.createElement("tbody");for(const file of files){const meta=modState.screenerMetadata.get(modScreenerKey(file)),tr=document.createElement("tr");[file.webkitRelativePath||file.name,meta?.startTimecode||"Not detected",meta?`${meta.fps} fps`:"",meta?.source||""].forEach(text=>{const td=document.createElement("td");td.textContent=text;tr.append(td);});tbody.append(tr);}table.append(thead,tbody);list.replaceChildren(table);details.classList.toggle("hidden",files.length===0);status.className=errors.length?"mod-status":"mod-status good";status.textContent=errors.length?`Detected start timecode and frame rate for ${modState.screenerMetadata.size} of ${files.length} screener files. ${errors.length} could not be read; those frames will be marked for review.`:`Detected start timecode and frame rate for all ${files.length} screener files. Expand the list to verify each value.`;}

function modQuote(s){return "'"+String(s).replaceAll("'","''")+"'";}
function modParseId(url){const m=String(url).match(/docs\.google\.com\/spreadsheets\/d\/([\w-]+)/);if(!m)throw Error("Please paste a valid Google Sheets URL.");return m[1];}
function modNorm(s){return String(s??"").toLowerCase().replace(/[\s_\-()]/g,"").replace(/[^a-z0-9]/g,"");}
function modOptions(select,items){select.replaceChildren(...items.map(x=>{const o=document.createElement("option");o.value=x.value;o.textContent=x.label;return o;}));}
function modInitGoogle(){if(modState.gapiReady&&modState.gisReady){mod$("modConnect").disabled=false;return;}if(!window.gapi||!window.google?.accounts?.oauth2){setTimeout(modInitGoogle,250);return;}gapi.load("client",async()=>{try{await gapi.client.init({discoveryDocs:[MOD_DISCOVERY]});modState.gapiReady=true;modMaybeReady();}catch(e){modStatus("Google Sheets API could not initialize: "+modErr(e),"error");}});modState.tokenClient=google.accounts.oauth2.initTokenClient({client_id:MOD_CLIENT_ID,scope:MOD_SCOPE,callback:""});modState.gisReady=true;modMaybeReady();}
function modMaybeReady(){if(modState.gapiReady&&modState.gisReady){mod$("modConnect").disabled=false;modStatus("Ready. Paste a tracker link, then connect your Google account.");}}
function modAuthorize(){return new Promise((resolve,reject)=>{if(!modState.tokenClient)return reject(Error("Google sign-in is not ready yet."));modState.tokenClient.callback=r=>{if(r.error)return reject(Error(r.error));modState.accessToken=r.access_token;gapi.client.setToken({access_token:r.access_token});resolve();};modState.tokenClient.requestAccessToken({prompt:gapi.client.getToken()?"":"consent"});});}
async function modConnect(){try{modState.id=modParseId(mod$("modSheetUrl").value);modStatus("Connecting and reading available tabs…");await modAuthorize();const r=await gapi.client.sheets.spreadsheets.get({spreadsheetId:modState.id,fields:"properties(title),sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))"});modState.title=r.result.properties.title;modState.tabs=r.result.sheets.map(s=>s.properties);const opts=modState.tabs.map(t=>({value:t.title,label:t.title}));modOptions(mod$("modSourceTab"),opts);modOptions(mod$("modRatesTab"),opts);const found=modState.tabs.find(t=>/cost|legal|flag|tracker/i.test(t.title));if(found)mod$("modSourceTab").value=found.title;const rates=modState.tabs.find(t=>/vendor info.*rates|vendor.*rates/i.test(t.title));if(rates)mod$("modRatesTab").value=rates.title;mod$("modTabs").classList.remove("hidden");mod$("modDisconnect").disabled=false;modStatus(`Connected to “${modState.title}”. Choose the cost tracker tab and vendor rates tab.`,"good");await modLoadRateTab();}catch(e){modStatus("Could not connect: "+modErr(e),"error");}}
async function modDisconnect(){const t=gapi.client.getToken();if(t?.access_token)google.accounts.oauth2.revoke(t.access_token);gapi.client.setToken("");modState.id="";modState.sourceRows=[];modState.vendors=[];modState.assignments.clear();mod$("modTabs").classList.add("hidden");mod$("modMappingCard").classList.add("hidden");mod$("modOrderCard").classList.add("hidden");mod$("modSecondCard").classList.add("hidden");modStatus("Disconnected.");mod$("modDisconnect").disabled=true;}
async function modReadTab(title){const r=await gapi.client.sheets.spreadsheets.values.get({spreadsheetId:modState.id,range:modQuote(title)+"!A1:AZ",valueRenderOption:"FORMATTED_VALUE"});return r.result.values||[];}
async function modLoadRateTab(){const data=await modReadTab(mod$("modRatesTab").value);if(!data.length)return;const h=data[0]||[];let vendorIx=h.findIndex(x=>/vendor|source/i.test(String(x)));if(vendorIx<0)vendorIx=0;modState.rates.clear();modState.vendors=[];data.slice(1).forEach(row=>{const name=String(row[vendorIx]||"").trim();if(!name)return;const ftg=String(row[14]||"").trim();const still=String(row[16]||"").trim();modState.rates.set(name,{footage:ftg,still});modState.vendors.push(name);});modState.vendors=[...new Set(modState.vendors)].sort((a,b)=>a.localeCompare(b));}
function modLooksLikeTc(value){return /^\d{2}[:;]\d{2}[:;]\d{2}[:;]\d{2}$/.test(String(value??"").trim());}
function modTimeHeaderScore(key,header){
  const raw=String(header??"").toLowerCase(),h=modNorm(header),tokens=raw.split(/[^a-z0-9]+/).filter(Boolean),side=key.endsWith("In")?"in":"out",isSource=key.startsWith("src");
  const hasSide=tokens.includes(side)||tokens.some(t=>[`${side}tc`,`${side}time`,`${side}timecode`,`time${side}`,`timecode${side}`,`tc${side}`,`source${side}`,`src${side}`,`cut${side}`,`timeline${side}`,`program${side}`,`record${side}`].includes(t));
  if(!hasSide){if(!isSource&&/timecode|tc/.test(h)&&((side==="in"&&/start|begin/.test(h))||(side==="out"&&/end|finish/.test(h))))return 95;return -100;}
  if(isSource){if(/source|src/.test(h))return h===`source${side}`||h===`src${side}`?150:120;if(/raw|original/.test(h))return 70;return -100;}
  if(/source|src|raw/.test(h))return -100;
  if(/cut|timeline|show|program|record|sequence|edit/.test(h))return 130;
  if(h===`time${side}`)return 115;
  if(/timecode|tc/.test(h))return 90;
  if(/time/.test(h))return 85;
  if(tokens.length===1&&tokens[0]===side)return 30;
  return -100;
}
function modResolveTimeColumns(headers,data,current){
  const resolved={...current},adjusted=[];
  for(const key of ["srcIn","srcOut","timeIn","timeOut"]){
    const candidates=headers.map((header,index)=>{
      const nonempty=data.map(r=>String(r[index]??"").trim()).filter(Boolean),valid=nonempty.filter(modLooksLikeTc).length;
      const ratio=nonempty.length?valid/nonempty.length:0,headerScore=modTimeHeaderScore(key,header);
      return{index,header,valid,ratio,headerScore,score:headerScore+ratio*45};
    }).filter(c=>c.headerScore>0&&c.valid>0&&c.ratio>=.5).sort((a,b)=>b.score-a.score||b.valid-a.valid);
    if(!candidates.length)continue;
    const selected=Number(resolved[key]),chosen=candidates.find(c=>c.index===selected&&c.headerScore>=85)||candidates[0];
    if(selected!==chosen.index){resolved[key]=chosen.index;adjusted.push(`${key} → ${String.fromCharCode(65+chosen.index)} (${chosen.header})`);}
  }
  return{mapping:resolved,adjusted};
}
function modBuildMap(headers){const holder=mod$("modColumnMap");holder.replaceChildren();MOD_FIELDS.forEach(field=>{const box=document.createElement("div");box.className="field";const lab=document.createElement("label");lab.textContent=field.label;lab.htmlFor="modMap_"+field.key;const sel=document.createElement("select");sel.id="modMap_"+field.key;const options=[{value:"",label:"Not mapped"},...headers.map((h,i)=>({value:String(i),label:`${String.fromCharCode(65+i)} · ${h||"(blank header)"}`}))];modOptions(sel,options);const alias=field.aliases.map(modNorm);let index=headers.findIndex(h=>alias.includes(modNorm(h)));if(index<0)index=headers.findIndex(h=>alias.some(a=>a.length>3&&modNorm(h).includes(a)));if(index>=0)sel.value=String(index);const sample=document.createElement("div");sample.className="mod-note mod-map-sample";sample.id=`modMapSample_${field.key}`;const update=()=>{const i=Number(sel.value);sample.textContent=i>=0&&headers[i]!==undefined?`Column sample: ${String(headers[i]||"(blank header)")} · first value: ${String(modState.mappingPreview?.[i]??"")||"(blank)"}`:"Column sample: not mapped";};sel.addEventListener("change",update);box.append(lab,sel,sample);holder.append(box);update();});mod$("modMappingCard").classList.remove("hidden");}
async function modLoadRows(){try{modStatus("Loading source tab…");const rows=await modReadTab(mod$("modSourceTab").value);if(rows.length<2)throw Error("The selected tab needs a header row and data rows.");modState.sourceHeaders=rows[0];modState.sourceDurationIndex=modFindSourceDurationColumn(rows[0],rows.slice(1));modState.mapping={};MOD_FIELDS.forEach(f=>{const v=mod$("modMap_"+f.key).value;modState.mapping[f.key]=v===""?null:Number(v);});const tcResolution=modResolveTimeColumns(rows[0],rows.slice(1),modState.mapping);modState.mapping=tcResolution.mapping;for(const key of ["srcIn","srcOut","timeIn","timeOut"]){const sel=mod$("modMap_"+key);if(sel&&modState.mapping[key]!=null)sel.value=String(modState.mapping[key]);}modState.mappingPreview=rows[1]||[];for(const key of ["srcIn","srcOut","timeIn","timeOut"]){const sel=mod$("modMap_"+key);if(sel)sel.dispatchEvent(new Event("change"));}const get=(row,key)=>{const i=modState.mapping[key];return i==null?"":String(row[i]??"").trim();};modState.sourceRows=rows.slice(1).map((r,i)=>({rowNumber:i+2,values:r,vendor:get(r,"vendor"),type:get(r,"type"),filename:get(r,"filename"),sftype:get(r,"sftype"),srcIn:get(r,"srcIn"),srcOut:get(r,"srcOut"),sourceDuration:modState.sourceDurationIndex>=0?modCellText(r[modState.sourceDurationIndex]):"",timeIn:get(r,"timeIn"),timeOut:get(r,"timeOut"),description:get(r,"description"),arc:get(r,"arc")})).filter(r=>r.vendor&&(r.filename||r.timeIn||r.timeOut||r.srcIn||r.description));if(!modState.sourceRows.length)throw Error("No source rows found. Check the source, filename, and timecode column mappings.");const sources=[...new Set(modState.sourceRows.map(r=>r.vendor))].sort((a,b)=>a.localeCompare(b));modState.vendors=sources;modState.assignments.clear();sources.forEach(name=>{const rs=modState.sourceRows.filter(r=>r.vendor===name),kinds=new Set(rs.map(r=>{const rate=modState.rates.get(name)||{};return modRateType(modIsStill(r)?rate.still:rate.footage);}).filter(Boolean));modState.assignments.set(name,kinds.size===1?[...kinds][0]:"");});mod$("modOrderCard").classList.remove("hidden");modRenderSourceAssignments();const correction=tcResolution.adjusted.length?` Timecode columns were corrected from tracker values: ${tcResolution.adjusted.join("; ")}.`:" Timecode column samples are shown above for review.";mod$("modOrderStatus").textContent=`Loaded ${modState.sourceRows.length} EDL rows across ${sources.length} sources. Choose how each source should be handled.`;mod$("modLoadStatus").textContent=correction;modStatus("Source list loaded.","good");}catch(e){mod$("modLoadStatus").textContent=modErr(e);modStatus("Could not load rows: "+modErr(e),"error");}}
function modFindSourceDurationColumn(headers,data=[]){const normalized=headers.map(modNorm),plain=normalized.findIndex(h=>h==="duration");if(plain>=0)return plain;const candidates=[];for(let i=0;i<normalized.length;i++){const h=normalized[i],sourceSpecific=h.includes("source")||h.includes("src"),durationHeader=h.includes("duration")&&(sourceSpecific||h.startsWith("duration")||h.includes("edlduration"));if(!durationHeader)continue;const samples=data.map(r=>String(r?.[i]??"").trim()).filter(Boolean),tcCount=samples.filter(v=>/^\d{1,3}:\d{2}:\d{2}[:;]\d{2}$/.test(v)).length,ratio=samples.length?tcCount/samples.length:0,exact=["sourceduration","sourcedur","srcduration","sourcelength","edlduration"].includes(h);candidates.push({i,score:ratio*100+(sourceSpecific?30:0)+(exact?20:0)+(samples.length?1:0)});}return candidates.sort((a,b)=>b.score-a.score)[0]?.i??-1;}
function modBuildAeDurationColumn(){const select=mod$("modAeDurationColumn"),headers=modState.sourceHeaders||[];if(!select)return;const sampleRow=modState.sourceRows.find(r=>r.values)?.values||[],options=[{value:"-1",label:"Not mapped"},...headers.map((h,i)=>{const sample=modCellText(sampleRow[i]);return {value:String(i),label:`${String.fromCharCode(65+i)} · ${h||"(blank header)"}${sample?` · sample: ${sample}`:""}`};})];modOptions(select,options);const detected=modFindSourceDurationColumn(headers,modState.sourceRows.map(r=>r.values));if(modState.sourceDurationIndex<0&&detected>=0)modState.sourceDurationIndex=detected;select.value=String(modState.sourceDurationIndex>=0?modState.sourceDurationIndex:detected);if(modState.sourceDurationIndex>=0)for(const row of modState.sourceRows)row.sourceDuration=modCellText(row.values?.[modState.sourceDurationIndex]);}
function modUpdateAeDurationValues(){const select=mod$("modAeDurationColumn");if(!select)return;const value=Number(select.value);modState.sourceDurationIndex=Number.isInteger(value)?value:-1;for(const row of modState.sourceRows)row.sourceDuration=modState.sourceDurationIndex>=0?modCellText(row.values?.[modState.sourceDurationIndex]):"";for(const shot of modState.shots){const next=modState.sourceDurationIndex>=0?modCellText(shot.values?.[modState.sourceDurationIndex]):"";if(next!==shot.sourceDuration)shot.confirmed=false;shot.sourceDuration=next;if(shot.inputDuration)shot.inputDuration.value=shot.sourceDuration;}modUpdateAeOrderState();}
function modIsAEMediaResources(name){return String(name||"").trim().toUpperCase()==="A+E MEDIA RESOURCES";}
function modDurationSeconds(value,fps){
  const s=String(value??"").trim();if(!s)return null;
  if(s.includes(":")||s.includes(";")){
    const tc=s.replace(/;/g,":").split(":");
    if(tc.length===4&&tc.every(x=>/^\d+$/.test(x))){const [h,m,sec,frames]=tc.map(Number);if(m>59||sec>59||!fps||frames>=modNominal(fps))return null;return h*3600+m*60+sec+frames/fps;}
    if(tc.length===3&&tc.slice(0,2).every(x=>/^\d+$/.test(x))&&/^\d+(?:\.\d+)?$/.test(tc[2])){const h=Number(tc[0]),m=Number(tc[1]),sec=Number(tc[2]);if(m>59||sec>=60)return null;return h*3600+m*60+sec;}
    if(tc.length===2&&tc.every(x=>/^\d+(?:\.\d+)?$/.test(x))){const m=Number(tc[0]),sec=Number(tc[1]);if(sec>=60)return null;return m*60+sec;}
    return null;
  }
  const n=Number(s.replace(/,/g,""));return Number.isFinite(n)&&n>=0?n:null;
}
function modTcExcelDay(tc,fps){const s=String(tc||"").trim().replace(/;/g,":").replace(/\./g,":");const p=s.split(":").map(Number);if(p.length!==4||p.some(x=>!Number.isFinite(x))||!fps)return null;const [h,m,sec,frames]=p;if(m>59||sec>59||frames>=modNominal(fps))return null;return(h*3600+m*60+sec+frames/fps)/86400;}
function modUpdateAeOrderState(){
  const button=mod$("modExportAeOrder"),status=mod$("modAeOrderStatus"),setup=mod$("modAeOrderSetup");if(!button||!status)return;
  const isAe=modIsAEMediaResources(modState.activeSource),allConfirmed=modState.shots.length>0&&modState.shots.every(shot=>shot.confirmed);button.classList.toggle("hidden",!isAe||!allConfirmed);if(setup)setup.classList.toggle("hidden",!isAe);
  mod$("modWriteSecond")?.classList.toggle("hidden",isAe);modUpdateAeReviewControls();
  if(!isAe){button.disabled=true;status.classList.add("hidden");return;}
  status.classList.remove("hidden");
  const issues=[];
  if(!String(mod$("modAeProject")?.value||"").trim())issues.push("enter a Project / Episode");
  if(!modState.shots.length)issues.push("build the shot review");
  const confirmed=modState.shots.filter(shot=>shot.confirmed).length;
  for(const shot of modState.shots){
    if(!shot.confirmed)issues.push(`confirm shot from tracker row ${shot.rowNumber}`);
    if(!shot.processedOnce)issues.push(`prepare screenshots for tracker row ${shot.rowNumber}`);
    if(!modLooksLikeTc(shot.ocrIn)||!modLooksLikeTc(shot.ocrOut))issues.push(`verify burnt-in timecodes on tracker row ${shot.rowNumber}`);
    if(!String(shot.aeTapeNumber||"").trim()||!String(shot.aeProgramTitle||"").trim()||!String(shot.aeDescription||"").trim())issues.push(`complete the A+E fields on tracker row ${shot.rowNumber}`);
    const durationFps=modState.screenerMetadata.get(shot.screener?modScreenerKey(shot.screener):"")?.fps||modState.mediaMetadata?.fps||24;if(!(modDurationSeconds(shot.sourceDuration,durationFps)>0))issues.push(`check Source Duration on tracker row ${shot.rowNumber}`);
    if(!shot.screener||shot.sourceInAvailable!==true)issues.push(`Source In image is unavailable on tracker row ${shot.rowNumber}`);
  }
  button.disabled=issues.length>0||modState.exportingAe;
  status.textContent=issues.length?`${confirmed} of ${modState.shots.length} shots confirmed. Complete and confirm every shot in the one-by-one review before downloading the order form. ${issues[0]||""}`:`Ready. All ${modState.shots.length} shots are confirmed; the final form uses the Source In screenshot and exact EDL Source Duration text.`;
}
function modIsStill(r){const s=r.sftype.toLowerCase();return s==="s"||/still|photo|image/.test(s);}
function modRateType(value){const v=String(value||"").toLowerCase();if(/per\s*second|per\s*sec|second|sec\b/.test(v))return"second";if(/per\s*asset|per\s*item|per\s*clip|asset|item|flat fee/.test(v))return"asset";return"";}
function modVendorRows(){return modState.sourceRows.filter(r=>r.vendor===modState.activeSource);}
function modRenderSourceAssignments(){const root=mod$("modSourceAssignments");root.replaceChildren();for(const name of modState.vendors){const rows=modState.sourceRows.filter(r=>r.vendor===name),line=document.createElement("div");line.className="mod-source-assignment";const title=document.createElement("div");title.className="mod-source-name";title.textContent=name;const detail=document.createElement("div");detail.className="mod-note";detail.textContent=`${rows.length} EDL row${rows.length===1?"":"s"}`;const select=document.createElement("select");select.setAttribute("aria-label",`Order type for ${name}`);select.dataset.source=name;modOptions(select,[{value:"",label:"Select an order type"},{value:"second",label:"Per Second"},{value:"asset",label:"Per Asset"},{value:"fairuse",label:"Fair Use"},{value:"pd",label:"Public Domain"},{value:"na",label:"N/A"}]);select.value=modState.assignments.get(name)||"na";select.addEventListener("change",()=>{modState.assignments.set(name,select.value);modUpdateOutputControls();});line.append(title,detail,select);root.append(line);}modUpdateOutputControls();}
function modUpdateOutputControls(){const second=[...modState.assignments].filter(([,kind])=>kind==="second").map(([name])=>name);mod$("modSecondSources").replaceChildren();for(const name of second){const row=document.createElement("div");row.className="mod-review-source";const label=document.createElement("span");label.textContent=name;const button=document.createElement("button");button.type="button";button.className="button secondary";button.textContent="Build Shot Review";button.addEventListener("click",()=>{modState.activeSource=name;mod$("modWriteSecond").disabled=true;mod$("modExportAeOrder").disabled=true;const isAe=modIsAEMediaResources(name);modState.aeReviewMode=false;mod$("modWriteSecond").classList.toggle("hidden",isAe);mod$("modExportAeOrder").classList.add("hidden");mod$("modAeOrderSetup").classList.toggle("hidden",!isAe);mod$("modAeTapePage").classList.add("hidden");mod$("modAeReviewControls").classList.add("hidden");mod$("modShots").classList.remove("hidden");mod$("modShots").classList.remove("mod-ae-click-review");mod$("modReviewSourceName").textContent=`${name} · Per Second`;mod$("modSecondCard").classList.remove("hidden");mod$("modSecondCard").scrollIntoView({behavior:"smooth",block:"start"});modBuildFrames();});row.append(label,button);mod$("modSecondSources").append(row);}mod$("modSecondSetup").classList.toggle("hidden",!second.length);mod$("modOutputSummary").textContent=`${second.length} Per Second · ${[...modState.assignments.values()].filter(x=>x==="asset").length} Per Asset · ${[...modState.assignments.values()].filter(x=>x==="fairuse").length} Fair Use · ${[...modState.assignments.values()].filter(x=>x==="pd").length} Public Domain · ${[...modState.assignments.values()].filter(x=>x==="na").length} N/A · ${[...modState.assignments.values()].filter(x=>!x).length} Unassigned`;
}
function modAddTabRequest(title){return gapi.client.sheets.spreadsheets.batchUpdate({spreadsheetId:modState.id,resource:{requests:[{addSheet:{properties:{title}}}]}});}
function modUniqueTabName(base){const stamp=new Date().toISOString().slice(0,19).replace(/[T:]/g,"-");return (base+" "+stamp).slice(0,95);}
async function modWriteTab(title,values,styleHeader=false){const created=await gapi.client.sheets.spreadsheets.batchUpdate({spreadsheetId:modState.id,resource:{requests:[{addSheet:{properties:{title,gridProperties:{rowCount:Math.max(1000,values.length+20),columnCount:Math.max(26,values[0]?.length||26)}}}}]}});await gapi.client.sheets.spreadsheets.values.update({spreadsheetId:modState.id,range:modQuote(title)+"!A1",valueInputOption:"RAW",resource:{values}});if(styleHeader){const sheetId=created.result.replies[0].addSheet.properties.sheetId;await gapi.client.sheets.spreadsheets.batchUpdate({spreadsheetId:modState.id,resource:{requests:[{repeatCell:{range:{sheetId,startRowIndex:0,endRowIndex:1,startColumnIndex:0,endColumnIndex:values[0].length},cell:{userEnteredFormat:{backgroundColor:{red:0.84,green:0.92,blue:0.98},textFormat:{bold:true}}},fields:"userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.bold"}}]}});}}
function modVendorBlock(vendor){const ids=[...new Set(modState.sourceRows.filter(r=>r.vendor===vendor).map(r=>r.filename.trim()).filter(Boolean))];return {vendor,ids};}
function modAssetMatrix(blocks){const width=blocks.length*2-1,max=Math.max(1,...blocks.map(b=>b.ids.length));const grid=Array.from({length:max+1},()=>Array(width).fill(""));blocks.forEach((b,i)=>{const c=i*2;grid[0][c]=b.vendor;b.ids.forEach((id,r)=>grid[r+1][c]=id);});return grid;}
async function modGenerateOrders(){const assets=[...modState.assignments].filter(([,k])=>k==="asset").map(([n])=>n),fair=[...modState.assignments].filter(([,k])=>k==="fairuse").map(([n])=>n),pd=[...modState.assignments].filter(([,k])=>k==="pd").map(([n])=>n);const total=assets.length+fair.length+pd.length;if(!total)return alert("There are no Per Asset, Fair Use, or Public Domain tabs to create yet. Sources marked N/A or left unassigned are skipped.");if(!confirm(`Create output tabs for ${assets.length} Per Asset source(s), ${fair.length} Fair Use source(s), and ${pd.length} Public Domain source(s)? Sources marked N/A or left unassigned will be skipped.`))return;const button=mod$("modGenerateOrders");button.disabled=true;try{const made=[];if(assets.length){const blocks=assets.map(modVendorBlock);const name=modUniqueTabName("MO Asset");await modWriteTab(name,modAssetMatrix(blocks),true);made.push(`${name} (Per Asset)`);}if(fair.length){const rows=modState.sourceRows.filter(r=>fair.includes(r.vendor));const values=[["Element / Original File Name","Source","Show In","Show Out","Description","Tracker Row"],...rows.map(r=>[r.filename,r.vendor,r.timeIn,r.timeOut,r.description,r.rowNumber])];const name=modUniqueTabName("MO Fair Use");await modWriteTab(name,values);made.push(`${name} (Fair Use)`);}if(pd.length){const rows=modState.sourceRows.filter(r=>pd.includes(r.vendor)),values=[modState.sourceHeaders,...rows.map(r=>r.values)];const name=modUniqueTabName("MO Public Domain");await modWriteTab(name,values);made.push(`${name} (Public Domain)`);}mod$("modOrderStatus").textContent="Created: "+made.join(" · ");}catch(e){mod$("modOrderStatus").textContent="Could not create all output tabs: "+modErr(e);}finally{button.disabled=false;}}
function modExtractArc(row){const text=(row.arc+" "+row.filename);const m=text.match(/[A-Z0-9]{4,5}ARC\d{4,}/i);return m?m[0].toUpperCase():"";}
function modFindScreener(arc){if(!arc)return null;return modState.screeners.find(f=>f.name.toUpperCase().includes(arc.toUpperCase()))||null;}
function modExtractAeTapeNumber(filename){const m=String(filename||"").match(/(?:^|[_-])(\d{4,})(?=[_-]CRS(?:[_\-.]|$))/i);return m?m[1]:"";}
function modApplyAeTapeTitle(tapeId,title){const normalized=String(tapeId||"").trim();if(!normalized)return;const value=String(title||"").trim();modState.aeTapeTitles.set(normalized,value);for(const shot of modState.shots){if(String(shot.aeTapeNumber||"").trim()!==normalized)continue;shot.aeProgramTitle=value;if(shot.inputProgram)shot.inputProgram.value=value;if(shot.confirmed)shot.confirmed=false;}modUpdateAeOrderState();}
function modBuildAeTapePage(){const root=mod$("modAeTapeList"),groups=new Map();for(const file of modState.screeners){const id=modExtractAeTapeNumber(file.name);if(!id)continue;if(!groups.has(id))groups.set(id,[]);groups.get(id).push(file);}const entries=[...groups.entries()].sort((a,b)=>a[0].localeCompare(b[0],undefined,{numeric:true}));root.replaceChildren();if(!entries.length){root.innerHTML='<div class="mod-empty">No A+E tape IDs were recognized in the uploaded elements folder. You can enter Tape # and Program/Tape Title manually on each shot.</div>';mod$("modAeTapeStatus").textContent="No matching A+E Tape IDs detected. The shot review remains available for manual entry.";mod$("modAeStartReview").disabled=false;return;}const table=document.createElement("table");table.className="mod-table mod-ae-tape-table";const head=document.createElement("tr");["A+E Tape #","Uploaded files","Matching ARC numbers","Program/Tape Title"].forEach(label=>{const th=document.createElement("th");th.textContent=label;head.append(th);});const thead=document.createElement("thead");thead.append(head);const tbody=document.createElement("tbody");for(const [id,files] of entries){const shots=modState.shots.filter(s=>String(s.aeTapeNumber||"").trim()===id),arcs=[...new Set(shots.map(s=>s.arc).filter(Boolean))],tr=document.createElement("tr"),tape=document.createElement("td"),fileCount=document.createElement("td"),arcCell=document.createElement("td"),titleCell=document.createElement("td"),input=document.createElement("input");tape.textContent=id;fileCount.textContent=String(files.length);arcCell.textContent=arcs.length?arcs.join(", "):"No EDL rows matched";input.type="text";input.value=modState.aeTapeTitles.get(id)||"";input.placeholder="Enter program or tape title";input.setAttribute("aria-label",`Program/Tape Title for A+E Tape ${id}`);input.addEventListener("input",()=>{modApplyAeTapeTitle(id,input.value);modUpdateAeTapeStatus(entries);});titleCell.append(input);tr.append(tape,fileCount,arcCell,titleCell);tbody.append(tr);}table.append(thead,tbody);root.append(table);modUpdateAeTapeStatus(entries);}
function modUpdateAeTapeStatus(entries){const ids=entries||[...new Map(modState.screeners.map(f=>[modExtractAeTapeNumber(f.name),f]).filter(([id])=>id))];const required=ids.filter(([id])=>modState.shots.some(s=>String(s.aeTapeNumber||"").trim()===id));const missing=required.filter(([id])=>!String(modState.aeTapeTitles.get(id)||"").trim());const matched=required.length;mod$("modAeTapeStatus").textContent=`Detected ${ids.length} unique A+E Tape ID${ids.length===1?"":"s"}; ${matched} ID${matched===1?"":"s"} match this order's EDL shots. ${missing.length?`Add a title for each matched ID before continuing (${missing.map(([id])=>id).join(", ")}).`:`Matched tape titles are ready. IDs without matching EDL shots are optional.`}`;mod$("modAeStartReview").disabled=missing.length>0;}
function modAeShotIssues(shot){const issues=[];if(!shot)return ["shot is unavailable"];if(!modLooksLikeTc(shot.ocrIn))issues.push("confirm TC In");if(!modLooksLikeTc(shot.ocrOut))issues.push("confirm TC Out");if(!String(shot.aeTapeNumber||"").trim())issues.push("enter A+E Tape #");if(!String(shot.aeProgramTitle||"").trim())issues.push("enter Program/Tape Title");if(!String(shot.aeDescription||"").trim())issues.push("enter Clip/Shot Description");const fps=modState.screenerMetadata.get(shot.screener?modScreenerKey(shot.screener):"")?.fps||modState.mediaMetadata?.fps||24;if(!(modDurationSeconds(shot.sourceDuration,fps)>0))issues.push("Source Duration is missing or invalid");if(!shot.processedOnce||!shot.framesReady)issues.push("wait for this shot's screenshots to finish loading");if(!shot.screener||shot.sourceInAvailable!==true)issues.push("Source In image is unavailable");return issues;}
function modUpdateAeReviewControls(){const controls=mod$("modAeReviewControls"),previous=mod$("modAePrevShot"),count=mod$("modAeReviewCount"),validation=mod$("modAeValidation"),tapePage=mod$("modAeTapePage"),container=mod$("modShots");if(!controls)return;const active=modState.aeReviewMode&&modIsAEMediaResources(modState.activeSource);controls.classList.toggle("hidden",!active);if(!active)return;const index=modState.aeReviewIndex,shot=modState.shots[index],confirmed=modState.shots.filter(s=>s.confirmed).length;if(previous)previous.disabled=index<=0;count.textContent=`Shot ${Math.min(index+1,modState.shots.length)} of ${modState.shots.length} · ${confirmed} confirmed`;const button=shot?.confirmButton;if(button){const issues=modAeShotIssues(shot);button.disabled=issues.length>0||(shot.confirmed&&index===modState.shots.length-1);button.textContent=shot.confirmed?(index===modState.shots.length-1?"Last Shot Confirmed":"Continue to Next Shot"):(index===modState.shots.length-1?"Confirm Last Shot":"Confirm Shot & Continue");if(validation)validation.textContent=issues.length?`To confirm: ${issues.join(" · ")}`:"Ready to confirm.";}else if(validation)validation.textContent="";if(tapePage&&container&&!tapePage.classList.contains("hidden"))controls.classList.add("hidden");}
async function modShowAeShot(index){if(!modState.aeReviewMode||!modState.shots.length)return;const nextIndex=Math.max(0,Math.min(modState.shots.length-1,index)),previous=modState.shots[modState.aeReviewIndex];if(previous&&previous!==modState.shots[nextIndex]){previous.nearViewport=false;previous.keepFrames=false;modState.releaseShotFrames?.(previous);}modState.aeReviewIndex=nextIndex;modState.shots.forEach((shot,i)=>shot.card.classList.toggle("mod-ae-active",i===nextIndex));const current=modState.shots[nextIndex];current.nearViewport=true;current.keepFrames=true;mod$("modShots").classList.remove("hidden");mod$("modAeTapePage").classList.add("hidden");modUpdateAeReviewControls();mod$("modFrameStatus").textContent=`Preparing shot ${nextIndex+1} of ${modState.shots.length}…`;await modState.ensureShotFrames?.(current);modUpdateAeReviewControls();}
async function modConfirmAeShot(shot){const issues=modAeShotIssues(shot);if(issues.length){mod$("modFrameStatus").textContent=`This shot needs attention before confirmation: ${issues.join("; ")}.`;return;}shot.confirmed=true;modUpdateAeOrderState();if(modState.aeReviewIndex<modState.shots.length-1)await modShowAeShot(modState.aeReviewIndex+1);else{mod$("modFrameStatus").textContent=`Review complete: all ${modState.shots.length} shots are confirmed. Download the A+E Order Form when the order details are ready.`;modUpdateAeReviewControls();}}
function modLoadLocalFiles(){const cut=mod$("modCutFile").files[0];modState.cut=cut||null;modState.screeners=[...mod$("modScreenerFiles").files].filter(f=>f.type.startsWith("video/")||/\.(mov|mp4|mxf|m4v|avi|mpg|mpeg|mkv|webm|mts|m2ts)$/i.test(f.name));}
function modVideoLoad(video,file){
  const key=`${file.name}:${file.size}:${file.lastModified}`;
  if(video.dataset.fileKey===key&&video.readyState>=1)return Promise.resolve(video.duration);
  return new Promise((resolve,reject)=>{
    video.pause();
    if(video.dataset.objectUrl)URL.revokeObjectURL(video.dataset.objectUrl);
    video.removeAttribute("src");video.load();
    video.onloadedmetadata=()=>{video.dataset.fileKey=key;delete video.dataset.lastTargetTime;delete video.dataset.lastMediaTime;resolve(video.duration);};
    video.onerror=()=>reject(Error(`Could not read video “${file.name}” in this browser.`));
    video.dataset.objectUrl=URL.createObjectURL(file);video.src=video.dataset.objectUrl;
  });
}
function modNextDecodedFrame(video,targetTime,fps){
  if(typeof video.requestVideoFrameCallback!=="function")return new Promise(resolve=>setTimeout(resolve,180));
  return new Promise((resolve,reject)=>{
    let done=false,callbackId=null;
    const previousTarget=Number(video.dataset.lastTargetTime),previousMedia=Number(video.dataset.lastMediaTime),hasPrevious=Number.isFinite(previousTarget)&&Number.isFinite(previousMedia),expectedDelta=targetTime-previousTarget;
    const tolerance=Math.max(.10,2.5/(Number(fps)||24));
    const cleanup=()=>{clearTimeout(timer);if(callbackId!==null&&typeof video.cancelVideoFrameCallback==="function")video.cancelVideoFrameCallback(callbackId);};
    const finish=(error,mediaTime)=>{if(done)return;done=true;cleanup();if(error){reject(error);return;}video.dataset.lastTargetTime=String(targetTime);if(Number.isFinite(mediaTime))video.dataset.lastMediaTime=String(mediaTime);resolve(mediaTime);};
    const timer=setTimeout(()=>finish(Error(`No new frame appeared at ${Number(targetTime).toFixed(2)}s; refusing to reuse the previous screenshot.`)),3000);
    const onFrame=(_now,metadata)=>{callbackId=null;const mediaTime=Number(metadata?.mediaTime);if(!hasPrevious||!Number.isFinite(mediaTime)||Math.abs((mediaTime-previousMedia)-expectedDelta)<=tolerance){finish(null,mediaTime);return;}callbackId=video.requestVideoFrameCallback(onFrame);};
    callbackId=video.requestVideoFrameCallback(onFrame);
  });
}
async function modSeek(video,time,fps=24){
  const t=Number(time);
  if(!Number.isFinite(t)||t<0)throw Error("The requested frame is earlier than the cut timeline start. Check the start TC.");
  if(t>video.duration+.05)throw Error(`The requested frame (${t.toFixed(2)}s) is beyond this video’s duration (${video.duration.toFixed(2)}s). Check the cut timeline start TC.`);
  if(Math.abs(video.currentTime-t)>.002){
    const framePromise=modNextDecodedFrame(video,t,fps).then(value=>({value}),error=>({error}));
    await new Promise((resolve,reject)=>{
      const timer=setTimeout(()=>{cleanup();reject(Error("Timed out while seeking to the requested timecode."));},10000);
      const cleanup=()=>{clearTimeout(timer);video.removeEventListener("seeked",done);video.removeEventListener("error",bad);};
      const done=()=>{cleanup();resolve();};const bad=()=>{cleanup();reject(Error("Could not seek to the requested timecode."));};
      video.addEventListener("seeked",done,{once:true});video.addEventListener("error",bad,{once:true});video.currentTime=t;
    });
    const frameResult=await framePromise;if(frameResult.error)throw frameResult.error;
  }else{video.dataset.lastTargetTime=String(t);video.dataset.lastMediaTime=String(video.currentTime);}
}
function modTcSeconds(tc,fps=24){
  const p=String(tc||"").replaceAll(";",":").split(":").map(Number);
  if(p.length<3||p.some(Number.isNaN))throw Error(`Invalid timecode “${tc}”. Use HH:MM:SS:FF.`);
  return p.slice(0,3).reduce((n,x)=>n*60+x,0)+(p[3]||0)/fps;
}
async function modCapture(video,file,tc,canvas,offset=0,ocrEngine=null,readOcr=false){
  if(!file)return "";await modVideoLoad(video,file);
  let fps,seconds;
  if(video===mod$("modVideoCut")){
    if(!modState.mediaMetadata)throw Error("Reference cut metadata has not been detected.");
    fps=modState.mediaMetadata.fps;seconds=modCutSeconds(tc)+offset/fps;
  }else{
    let metadata=modState.screenerMetadata.get(modScreenerKey(file));
    if(!metadata)throw Error(`Could not read embedded start timecode and frame rate for screener ${file.name}.`);
    fps=metadata.fps;let frames=modParseTcFrames(tc,fps)-modParseTcFrames(metadata.startTimecode,fps);
    seconds=frames/fps-offset;
    if(seconds<0||seconds>video.duration+.05){
      try{const visualStart=await modVisualStartTc(video,file,fps);if(visualStart){const visualFrames=modParseTcFrames(tc,fps)-modParseTcFrames(visualStart,fps),visualSeconds=visualFrames/fps-offset;if(visualSeconds>=0&&visualSeconds<=video.duration+.05){metadata={...metadata,startTimecode:visualStart,source:metadata.source+" + first-frame BITC OCR"};modState.screenerMetadata.set(modScreenerKey(file),metadata);seconds=visualSeconds;} }}catch(e){console.warn("Could not align screener from first-frame BITC",e);}
    }
    if(seconds<0||seconds>video.duration+.05)throw Error(`Source TC ${tc} maps to ${seconds.toFixed(2)}s in a ${video.duration.toFixed(2)}s screener (detected start ${metadata.startTimecode}). Check the Source In/Out column mapping or the screener’s start timecode.`);
  }
  await modSeek(video,seconds,fps);
  const sourceWidth=video.videoWidth||640,sourceHeight=video.videoHeight||360;
  const scale=Math.min(1,960/sourceWidth),previewWidth=Math.max(1,Math.round(sourceWidth*scale)),previewHeight=Math.max(1,Math.round(sourceHeight*scale));
  let ocr="",full=null;
  try{
    const target=canvas.getContext("2d");canvas.width=previewWidth;canvas.height=previewHeight;
    if(readOcr){full=document.createElement("canvas");full.width=sourceWidth;full.height=sourceHeight;full.getContext("2d").drawImage(video,0,0,sourceWidth,sourceHeight);ocr=await modReadBurntInTc(ocrEngine,full);target.drawImage(full,0,0,previewWidth,previewHeight);}
    else target.drawImage(video,0,0,previewWidth,previewHeight);
    canvas.dataset.captured="true";canvas.hidden=false;delete canvas.dataset.error;
    const placeholder=canvas.parentElement?.querySelector(".mod-frame-placeholder");if(placeholder)placeholder.hidden=true;
  }finally{if(full){full.width=1;full.height=1;}}
  return ocr;
}
function modCutSeconds(tc){
  const metadata=modState.mediaMetadata;
  if(!metadata)throw Error("Reference cut metadata has not been detected.");
  const frames=modParseTcFrames(tc,metadata.fps)-modParseTcFrames(metadata.startTimecode,metadata.fps);
  if(frames<0)throw Error(`EDL time ${tc} is before the reference file start ${metadata.startTimecode}.`);
  return frames/metadata.fps;
}
async function modVisualStartTc(video,file,fps){
  await modVideoLoad(video,file);
  if(video.readyState<2)await new Promise((resolve,reject)=>{const timer=setTimeout(()=>{video.removeEventListener("loadeddata",ok);reject(Error("Timed out decoding the first video frame for timecode alignment."));},10000);const ok=()=>{clearTimeout(timer);resolve();};video.addEventListener("loadeddata",ok,{once:true});});
  await modSeek(video,0,fps);
  const canvas=document.createElement("canvas");canvas.width=video.videoWidth||640;canvas.height=video.videoHeight||360;canvas.getContext("2d").drawImage(video,0,0,canvas.width,canvas.height);
  let paddle=null;try{paddle=await modEnsurePaddle();}catch(e){console.warn("PaddleOCR unavailable for start-timecode alignment",e);}
  const tc=await modReadBurntInTc(paddle,canvas);
  return modLooksLikeTc(tc)?tc:"";
}
async function modAlignReferenceCutStart(){
  const meta=modState.mediaMetadata;if(!meta||meta.visualStartChecked||!/^00:00:00[:;]00$/.test(meta.startTimecode))return;
  meta.visualStartChecked=true;mod$("modFrameStatus").textContent="Reference cut metadata starts at zero; checking its first visible frame for the show timecode…";
  try{const tc=await modVisualStartTc(mod$("modVideoCut"),modState.cut,meta.fps);if(tc&&!/^00:00:00[:;]00$/.test(tc)){meta.startTimecode=tc;meta.source+=" + first-frame BITC OCR";mod$("modTimelineStart").value=tc;mod$("modMetadataStatus").textContent=`Metadata timecode was reset to zero. Read first-frame BITC ${tc}; cut frames will use this show-timecode start.`;}}catch(e){console.warn("Could not align reference cut from first-frame BITC",e);}
}
function modWaitForMedia(video,eventName,timeoutMs=15000){
  return new Promise((resolve,reject)=>{
    if(eventName==="loadedmetadata"&&video.readyState>=1)return resolve();
    let timer;
    const cleanup=()=>{clearTimeout(timer);video.removeEventListener(eventName,ok);video.removeEventListener("error",bad);};
    const ok=()=>{cleanup();resolve();};
    const bad=()=>{cleanup();reject(Error(`The browser could not load this cut (${video.error?.message||"unsupported or unreadable video format"}).`));};
    timer=setTimeout(()=>{cleanup();reject(Error("Timed out while preparing the cut player."));},timeoutMs);
    video.addEventListener(eventName,ok,{once:true});video.addEventListener("error",bad,{once:true});
  });
}
function modAttachCutRange(video,row,status){
  video.dataset.rangeStart=String(modCutSeconds(row.timeIn));
  video.dataset.rangeEnd=String(modCutSeconds(row.timeOut));
  video.dataset.rangeFps=String(modState.mediaMetadata.fps);
  video.tabIndex=0;
  if(Number(video.dataset.rangeEnd)<=Number(video.dataset.rangeStart))throw Error(`Cut OUT ${row.timeOut} must be later than IN ${row.timeIn}.`);
  video.addEventListener("timeupdate",()=>{
    const start=Number(video.dataset.rangeStart),end=Number(video.dataset.rangeEnd),epsilon=1/(Number(video.dataset.rangeFps)||24);
    if(video.currentTime<start-epsilon){video.pause();video.currentTime=start;}
    else if(video.currentTime>=end){video.pause();video.currentTime=end;status.textContent=`Reached Cut OUT ${row.timeOut}. Use “Play cut range” to replay.`;}
  });
  video.addEventListener("play",()=>{
    const start=Number(video.dataset.rangeStart),end=Number(video.dataset.rangeEnd);
    if(video.currentTime>=end){video.currentTime=start;status.textContent=`Playing ${row.timeIn} to ${row.timeOut} with picture and audio.`;}
  });
}
async function modPlayCutRange(video,row,status,button,placeholder){
  try{
    status.className="mod-player-status";if(button){button.disabled=true;button.textContent="Loading reference cut section…";}
    if(!modState.cut||!modState.cutObjectUrl)throw Error("Choose the reference cut first.");
    const start=modCutSeconds(row.timeIn),end=modCutSeconds(row.timeOut),fps=modState.mediaMetadata.fps;
    if(end<=start)throw Error(`Cut OUT ${row.timeOut} must be later than IN ${row.timeIn}.`);
    if(modState.activeRangeVideo&&modState.activeRangeVideo!==video){modState.activeRangeVideo.pause();modState.activeRangeVideo.removeAttribute("src");delete modState.activeRangeVideo.dataset.cutUrl;modState.activeRangeVideo.load();}
    modState.activeRangeVideo=video;
    if(video.dataset.cutUrl!==modState.cutObjectUrl){
      video.pause();video.src=modState.cutObjectUrl;video.dataset.cutUrl=modState.cutObjectUrl;video.load();
      status.textContent="Opening the local reference cut…";await modWaitForMedia(video,"loadedmetadata");
    }
    if(end>video.duration+.05)throw Error(`Cut OUT ${row.timeOut} is beyond the reference cut duration (${video.duration.toFixed(2)} seconds). Check the detected start timecode.`);
    video.muted=false;video.volume=1;
    if(Math.abs(video.currentTime-start)>.002){const seeked=modWaitForMedia(video,"seeked");video.currentTime=start;await seeked;}
    video.hidden=false;if(placeholder)placeholder.hidden=true;video.focus({preventScroll:true});
    if(button){button.disabled=false;button.textContent="Play Reference Cut Section";}
    status.textContent=`Starting ${row.timeIn} to ${row.timeOut} with picture and audio…`;
    try{await video.play();status.textContent=`Playing ${row.timeIn} to ${row.timeOut} with picture and audio.`;}
    catch(e){status.textContent=`Ready at ${row.timeIn}. Press Play on the video controls to start picture and audio.`;}
  }catch(e){status.textContent=modErr(e);status.className="mod-player-status error";if(button){button.disabled=false;button.textContent="Play Reference Cut Section";}}
}
function modStepCutFrame(event){
  if(event.key!=="ArrowLeft"&&event.key!=="ArrowRight")return;
  const video=modState.activeRangeVideo,wrap=video?.closest(".mod-cut-player-wrap");
  if(!video||video.hidden||!wrap||(!wrap.contains(document.activeElement)&&!wrap.contains(event.target)))return;
  if(event.target?.matches?.("input,textarea,select,[contenteditable='true']"))return;
  const fps=Number(video.dataset.rangeFps),start=Number(video.dataset.rangeStart),end=Number(video.dataset.rangeEnd);
  if(!Number.isFinite(fps)||fps<=0||!Number.isFinite(start)||!Number.isFinite(end))return;
  event.preventDefault();event.stopPropagation();
  video.pause();
  const frame=1/fps,max=Math.max(start,end-frame),delta=event.key==="ArrowRight"?frame:-frame;
  video.currentTime=Math.max(start,Math.min(max,video.currentTime+delta));
  const status=wrap.querySelector("[data-player-status]");
  if(status)status.textContent=`Paused · stepped one frame ${delta>0?"forward":"back"}. Use the arrow keys again to continue frame by frame.`;
}
function modSetFramePlaceholder(canvas,message){canvas.width=1;canvas.height=1;canvas.hidden=true;delete canvas.dataset.captured;const placeholder=canvas.parentElement?.querySelector(".mod-frame-placeholder");if(placeholder){placeholder.textContent=message;placeholder.hidden=false;}}
function modMarkFrameUnavailable(canvas,label){const ctx=canvas.getContext("2d");canvas.width=640;canvas.height=360;ctx.fillStyle="#151b2a";ctx.fillRect(0,0,canvas.width,canvas.height);ctx.fillStyle="#d5d9e3";ctx.font="600 20px Arial, sans-serif";ctx.textAlign="center";ctx.textBaseline="middle";ctx.fillText(label,canvas.width/2,canvas.height/2,canvas.width-40);canvas.hidden=false;canvas.dataset.captured="";canvas.dataset.error="true";const placeholder=canvas.parentElement?.querySelector(".mod-frame-placeholder");if(placeholder)placeholder.hidden=true;}
function modOpenFrame(canvas){
  try{const w=window.open();if(w){w.document.write(`<title>Frame review</title><img src="${canvas.toDataURL("image/png")}" style="max-width:100%;height:auto;background:#111">`);w.document.close();}}catch(e){console.warn("Could not open full-size frame",e);}
}
function modEsc(s){return String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));}
async function modEnsurePaddle(){if(!mod$("modRunOcr").checked)return null;if(modState.paddleOcr)return modState.paddleOcr;if(!modState.paddleLoadPromise)modState.paddleLoadPromise=(async()=>{const {PaddleOCR}=await import("https://esm.sh/@paddleocr/paddleocr-js");modState.paddleOcr=await PaddleOCR.create({lang:"en",ocrVersion:"PP-OCRv5",ortOptions:{backend:"wasm",wasmPaths:"https://cdn.jsdelivr.net/npm/onnxruntime-web/dist/"}});return modState.paddleOcr;})().catch(e=>{modState.paddleLoadPromise=null;throw e;});return modState.paddleLoadPromise;}
async function modEnsureTesseract(){if(!window.Tesseract)return null;if(modState.worker)return modState.worker;modState.worker=await Tesseract.createWorker("eng",1,{logger:()=>{}});await modState.worker.setParameters({tessedit_char_whitelist:"0123456789:;.",tessedit_pageseg_mode:"7",preserve_interword_spaces:"1"});return modState.worker;}
function modOcrCrop(source,region,mode="raw"){const x=Math.round(source.width*region.x),y=Math.round(source.height*region.y),w=Math.round(source.width*region.w),h=Math.round(source.height*region.h),scale=4,crop=document.createElement("canvas");crop.width=w*scale;crop.height=h*scale;const ctx=crop.getContext("2d",{willReadFrequently:true});ctx.imageSmoothingEnabled=true;ctx.imageSmoothingQuality="high";ctx.drawImage(source,x,y,w,h,0,0,crop.width,crop.height);if(mode!=="raw"){const pixels=ctx.getImageData(0,0,crop.width,crop.height);for(let i=0;i<pixels.data.length;i+=4){const gray=.299*pixels.data[i]+.587*pixels.data[i+1]+.114*pixels.data[i+2],white=gray>145,v=mode==="invert"?(white?0:255):(white?255:0);pixels.data[i]=pixels.data[i+1]=pixels.data[i+2]=v;pixels.data[i+3]=255;}ctx.putImageData(pixels,0,0);}return crop;}
async function modReadBurntInTc(paddle,canvas){const regions=[{x:.28,y:0,w:.68,h:.14},{x:.08,y:.02,w:.84,h:.20},{x:.08,y:.68,w:.84,h:.27}];if(paddle){for(const region of regions)for(const mode of ["raw","threshold"]){try{const [result]=await paddle.predict(modOcrCrop(canvas,region,mode));const tc=modCleanOcr((result?.items||[]).map(x=>x.text||"").join(" "));if(tc)return tc;}catch(e){console.warn("PaddleOCR crop failed",e);}}}try{const worker=await modEnsureTesseract();if(worker)for(const region of regions)for(const mode of ["threshold","invert","raw"]){const result=await worker.recognize(modOcrCrop(canvas,region,mode));const tc=modCleanOcr(result.data.text);if(tc)return tc;}}catch(e){console.warn("Tesseract OCR failed",e);}return "";}
async function modBuildFrames(){
  modLoadLocalFiles();
  if(modState.metadataPromise)await modState.metadataPromise;
  if(modState.screenerMetadataPromise)await modState.screenerMetadataPromise;
  if(!modState.cut)return alert("Choose the reference cut first.");
  if(!modState.mediaMetadata)return alert("Wait for the reference cut's start timecode and frame rate to be detected before building the shot review.");
  if(!modState.screeners.length)return alert("Choose the raw screener folder.");
  const isAeVendor=modIsAEMediaResources(modState.activeSource);mod$("modExportAeOrder").classList.toggle("hidden",!isAeVendor);mod$("modAeOrderSetup").classList.toggle("hidden",!isAeVendor);mod$("modWriteSecond").classList.toggle("hidden",isAeVendor);mod$("modAeTapePage").classList.add("hidden");mod$("modAeReviewControls").classList.add("hidden");if(isAeVendor&&!String(mod$("modAeProject").value||"").trim())mod$("modAeProject").value=modState.title||"";if(isAeVendor)modBuildAeDurationColumn();
  await modAlignReferenceCutStart();
  const fps=modState.mediaMetadata.fps;
  const rows=[...modVendorRows()].sort((a,b)=>{
    let ai,bi;try{ai=modParseTcFrames(a.timeIn,fps);}catch{ai=Number.POSITIVE_INFINITY;}
    try{bi=modParseTcFrames(b.timeIn,fps);}catch{bi=Number.POSITIVE_INFINITY;}
    if(ai!==bi)return ai-bi;
    let ao,bo;try{ao=modParseTcFrames(a.timeOut,fps);}catch{ao=Number.POSITIVE_INFINITY;}
    try{bo=modParseTcFrames(b.timeOut,fps);}catch{bo=Number.POSITIVE_INFINITY;}
    return ao-bo||(a.rowNumber||0)-(b.rowNumber||0);
  });
  if(!rows.length)return;
  const badCutRows=rows.filter(r=>!modLooksLikeTc(r.timeIn)||!modLooksLikeTc(r.timeOut));
  if(badCutRows.length){const examples=badCutRows.slice(0,3).map(r=>`row ${r.rowNumber}: In “${r.timeIn||"blank"}”, Out “${r.timeOut||"blank"}”`).join("; ");mod$("modFrameStatus").textContent=`Shot review stopped before capture: ${badCutRows.length} row(s) do not have valid Cut/Timeline In and Out timecodes (${examples}). Check the timecode column samples above.`;return;}
  modState.frameObserver?.disconnect();modState.frameCleanupObserver?.disconnect();
  const container=mod$("modShots");container.replaceChildren();container.classList.remove("hidden");container.classList.toggle("mod-ae-click-review",isAeVendor);modState.shots=[];modState.aeReviewMode=false;modState.aeReviewIndex=0;mod$("modWriteSecond").disabled=true;mod$("modProgressBar").style.width="0%";
  mod$("modFrameStatus").textContent=isAeVendor?`Preparing ${rows.length} A+E shots in show time order for one-by-one review.`:`Building ${rows.length} shot cards in show Cut In / Cut Out order. Screenshots will load as you scroll.`;
  const fragment=document.createDocumentFragment();
  const shots=rows.map((row,index)=>{
    const arc=modExtractArc(row),file=modFindScreener(arc),detectedTape=isAeVendor?modExtractAeTapeNumber(file?.name||row.filename):"",card=document.createElement("article");card.className="mod-shot";
    const shot={...row,arc,screener:file,ocrIn:"",ocrOut:"",aeTapeNumber:detectedTape,aeProgramTitle:detectedTape?(modState.aeTapeTitles.get(detectedTape)||""):"",aeDescription:"",confirmed:false,sourceInAvailable:null,card,processing:false,queued:false,keepFrames:false,framesReady:false,processedOnce:false,ocrDone:!mod$("modRunOcr").checked,nearViewport:false,cleanupNear:false};
    const badge=file?`<span class="mod-note">${modEsc(file.name)}</span>`:`<span class="mod-badge">Screener not matched</span>`;
    card.innerHTML=`<div class="mod-shot-head"><strong class="mod-shot-title">Shot ${index+1} · ${modEsc(row.filename||arc||"Unidentified clip")}</strong><div class="mod-toolbar"><span class="mod-badge">${modEsc(row.vendor)}</span>${badge}</div></div><div class="mod-frame-grid">${[["Cut In · 3 frames in",row.timeIn],["Cut Out · 3 frames in",row.timeOut],["Screener Source In",row.srcIn],["Screener Source Out",row.srcOut]].map((x,i)=>`<div class="mod-frame"><canvas data-frame="${i}" hidden></canvas><div class="mod-frame-placeholder">Waiting for this shot to enter view</div><label>${x[0]} · ${modEsc(x[1])}</label></div>`).join("")}</div><div class="mod-shot-fields"><div class="field"><label>Burnt-In TC In suggestion</label><input data-ocr="in" placeholder="OCR suggestion or enter manually"></div><div class="field"><label>Burnt-In TC Out suggestion</label><input data-ocr="out" placeholder="OCR suggestion or enter manually"></div></div>${isAeVendor?`<div class="mod-shot-fields mod-ae-shot-fields"><div class="field"><label>A+E Tape #</label><input data-ae="tape" value="${modEsc(shot.aeTapeNumber)}" placeholder="Enter A+E tape number"></div><div class="field"><label>Program/Tape Title</label><input data-ae="program" value="${modEsc(shot.aeProgramTitle)}" placeholder="Enter program or tape title"></div><div class="field"><label>Clip/Shot Description</label><input data-ae="description" placeholder="Enter clip or shot description"></div><div class="field"><label>Source Duration (EDL · exact)</label><input data-ae="duration" value="${modEsc(row.sourceDuration||"")}" readonly></div></div>`:""}<div class="mod-note mod-help">Cut: ${modEsc(row.timeIn)} – ${modEsc(row.timeOut)} · Source: ${modEsc(row.srcIn)} – ${modEsc(row.srcOut)} · Review is ordered by show Cut In, then Cut Out. Click a frame to enlarge.</div>`;
    shot.canvases=[...card.querySelectorAll("canvas")];shot.inputIn=card.querySelector('[data-ocr="in"]');shot.inputOut=card.querySelector('[data-ocr="out"]');
    shot.inputIn.addEventListener("input",()=>{shot.ocrIn=shot.inputIn.value;shot.confirmed=false;modUpdateAeOrderState();});shot.inputOut.addEventListener("input",()=>{shot.ocrOut=shot.inputOut.value;shot.confirmed=false;modUpdateAeOrderState();});
    if(isAeVendor){shot.inputTape=card.querySelector('[data-ae="tape"]');shot.inputProgram=card.querySelector('[data-ae="program"]');shot.inputDescription=card.querySelector('[data-ae="description"]');shot.inputDuration=card.querySelector('[data-ae="duration"]');shot.inputTape.addEventListener("input",()=>{shot.aeTapeNumber=shot.inputTape.value;const mappedTitle=modState.aeTapeTitles.get(String(shot.aeTapeNumber||"").trim())||"";shot.aeProgramTitle=mappedTitle;shot.inputProgram.value=mappedTitle;shot.confirmed=false;modUpdateAeOrderState();});shot.inputProgram.addEventListener("input",()=>{shot.aeProgramTitle=shot.inputProgram.value;shot.confirmed=false;modUpdateAeOrderState();});shot.inputDescription.addEventListener("input",()=>{shot.aeDescription=shot.inputDescription.value;shot.confirmed=false;modUpdateAeOrderState();});const actions=document.createElement("div");actions.className="mod-actions";shot.confirmButton=document.createElement("button");shot.confirmButton.type="button";shot.confirmButton.className="button primary";shot.confirmButton.addEventListener("click",()=>modConfirmAeShot(shot));actions.append(shot.confirmButton);card.append(actions);}
    shot.canvases.forEach(c=>c.addEventListener("click",()=>modOpenFrame(c)));fragment.append(card);return shot;
  });
  container.replaceChildren(fragment);modState.shots=shots;mod$("modBuildFrames").disabled=true;
  let completed=shots.filter(s=>s.processedOnce).length,running=false;const queue=[];let worker=null,ocrMode=mod$("modRunOcr").checked?"PaddleOCR":"off";
  const update=()=>{mod$("modProgressBar").style.width=Math.round(completed/shots.length*100)+"%";if(isAeVendor){mod$("modFrameStatus").textContent=`Screenshots prepared: ${completed} of ${shots.length}. The A+E review shows one shot at a time. OCR: ${ocrMode}.`;if(completed===shots.length){mod$("modBuildFrames").disabled=false;mod$("modFrameStatus").textContent=`All ${shots.length} shots have been prepared. Confirm each shot in the click-through review before downloading the A+E form.`;}}else{mod$("modFrameStatus").textContent=`Screenshots prepared: ${completed} of ${shots.length}. Scroll through the list; images load near the viewport and offscreen previews are released to keep the page responsive. OCR: ${ocrMode}.`;if(completed===shots.length){mod$("modWriteSecond").disabled=false;mod$("modBuildFrames").disabled=false;mod$("modFrameStatus").textContent=`All ${shots.length} shots have been prepared in show time order. You can review and write the Per Second order tab.`;}}};
  const captureShot=async shot=>{
    shot.processing=true;shot.canvases.forEach(c=>{const ph=c.parentElement?.querySelector(".mod-frame-placeholder");if(ph&&!c.dataset.captured)ph.textContent="Loading frame…";});
    const errors=[],wantsOcr=mod$("modRunOcr").checked&&shot.screener&&!shot.ocrDone;
    if(wantsOcr&&!worker){mod$("modFrameStatus").textContent="Loading OCR for the visible source frames…";try{worker=await modEnsurePaddle();}catch(e){ocrMode="Tesseract fallback";console.warn("PaddleOCR could not start; using Tesseract",e);}}
    const cutRequests=[[shot.timeIn,shot.canvases[0],3,"Cut In"],[shot.timeOut,shot.canvases[1],-3,"Cut Out"]];
    for(const [tc,canvas,frames,label] of cutRequests){try{await modCapture(mod$("modVideoCut"),modState.cut,tc,canvas,frames);}catch(e){errors.push(`${label}: ${modErr(e)}`);modMarkFrameUnavailable(canvas,`${label} frame unavailable`);}}
    if(shot.screener){
      const sourceRequests=[[shot.srcIn,shot.canvases[2],"in","Source In"],[shot.srcOut,shot.canvases[3],"out","Source Out"]];
      for(const [tc,canvas,which,label] of sourceRequests){if(!modLooksLikeTc(tc)){errors.push(`${label}: “${tc||"blank"}” is not a timecode; check the tracker column sample.`);modMarkFrameUnavailable(canvas,`${label} timecode unavailable`);if(which==="in")shot.sourceInAvailable=false;continue;}try{const ocr=await modCapture(mod$("modVideoSource"),shot.screener,tc,canvas,0,worker,wantsOcr);if(which==="in")shot.sourceInAvailable=true;if(wantsOcr){if(which==="in"&&!shot.ocrIn){shot.ocrIn=ocr;shot.inputIn.value=ocr;}else if(which==="out"&&!shot.ocrOut){shot.ocrOut=ocr;shot.inputOut.value=ocr;}}}catch(e){errors.push(`${label}: ${modErr(e)}`);modMarkFrameUnavailable(canvas,`${label} frame unavailable`);if(which==="in")shot.sourceInAvailable=false;}}
    }else{shot.sourceInAvailable=false;errors.push("Source In/Out: no matching screener file");modMarkFrameUnavailable(shot.canvases[2],"No screener matched");modMarkFrameUnavailable(shot.canvases[3],"No screener matched");}
    if(wantsOcr)shot.ocrDone=true;
    if(modLooksLikeTc(shot.srcIn)&&modLooksLikeTc(shot.srcOut)&&shot.screener){try{const sourceFps=modState.screenerMetadata.get(modScreenerKey(shot.screener))?.fps||fps;if(modParseTcFrames(shot.srcOut,sourceFps)<modParseTcFrames(shot.srcIn,sourceFps))errors.push("Source Out precedes Source In; verify the source-timecode columns.");}catch{}}
    shot.error=errors.join(" · ");if(errors.length){if(!shot.warningNode){shot.warningNode=document.createElement("p");shot.warningNode.className="mod-status error mod-shot-warning";shot.card.append(shot.warningNode);}shot.warningNode.textContent="Review this row: "+shot.error;}else if(shot.warningNode){shot.warningNode.remove();shot.warningNode=null;}
    shot.processing=false;shot.framesReady=true;if(!shot.processedOnce){shot.processedOnce=true;completed++;}if(!shot.nearViewport&&!shot.cleanupNear&&!shot.keepFrames){shot.canvases.forEach(c=>modSetFramePlaceholder(c,"Preview released to save memory · scroll back to reload"));shot.framesReady=false;}update();modUpdateAeOrderState();
  };
  const settleShot=shot=>{const waiters=shot.frameWaiters||[];shot.frameWaiters=[];waiters.forEach(resolve=>resolve());};
  const runQueue=async()=>{if(running)return;running=true;while(queue.length){const shot=queue.shift();shot.queued=false;if(shot.processing||shot.framesReady){settleShot(shot);continue;}try{await captureShot(shot);}catch(e){shot.processing=false;shot.error=modErr(e);shot.canvases.forEach(c=>modMarkFrameUnavailable(c,"Frame unavailable"));shot.framesReady=true;if(!shot.processedOnce){shot.processedOnce=true;completed++;}update();modUpdateAeOrderState();}settleShot(shot);}running=false;};
  const enqueue=shot=>{if(shot.framesReady)return Promise.resolve();return new Promise(resolve=>{shot.frameWaiters=shot.frameWaiters||[];shot.frameWaiters.push(resolve);if(shot.processing||shot.queued)return;shot.queued=true;queue.push(shot);runQueue();});};
  modState.frameObserver=!isAeVendor&&typeof IntersectionObserver==="function"?new IntersectionObserver(entries=>{for(const entry of entries){const shot=shots.find(s=>s.card===entry.target);if(!shot)continue;shot.nearViewport=entry.isIntersecting;if(entry.isIntersecting)enqueue(shot);else if(!shot.cleanupNear&&shot.processedOnce&&!shot.processing&&!shot.keepFrames&&shot.framesReady){shot.canvases.forEach(c=>modSetFramePlaceholder(c,"Preview released to save memory · scroll back to reload"));shot.framesReady=false;}}},{rootMargin:"1200px 0px"}):null;
  modState.frameCleanupObserver=!isAeVendor&&typeof IntersectionObserver==="function"?new IntersectionObserver(entries=>{for(const entry of entries){const shot=shots.find(s=>s.card===entry.target);if(!shot)continue;const wasNear=shot.cleanupNear;shot.cleanupNear=entry.isIntersecting;if(!entry.isIntersecting&&wasNear&&shot.processedOnce&&!shot.processing&&!shot.keepFrames&&shot.framesReady){shot.canvases.forEach(c=>modSetFramePlaceholder(c,"Preview released to save memory · scroll back to reload"));shot.framesReady=false;}}},{rootMargin:"2200px 0px"}):null;
  if(isAeVendor){}else if(modState.frameObserver){for(const shot of shots){modState.frameObserver.observe(shot.card);modState.frameCleanupObserver?.observe(shot.card);}}else{for(const shot of shots)enqueue(shot);}
  if(!mod$("modRunOcr").checked)for(const shot of shots)shot.ocrDone=true;
  modState.ensureShotFrames=async shot=>{shot.keepFrames=true;await enqueue(shot);};modState.releaseShotFrames=shot=>{shot.keepFrames=false;if(!shot.nearViewport&&!shot.cleanupNear&&shot.framesReady){shot.canvases.forEach(c=>modSetFramePlaceholder(c,"Preview released to save memory · scroll back to reload"));shot.framesReady=false;}};
  if(isAeVendor){modState.aeReviewMode=true;modState.aeReviewIndex=0;container.classList.add("hidden");mod$("modAeTapePage").classList.remove("hidden");modBuildAeTapePage();modUpdateAeOrderState();mod$("modFrameStatus").textContent=`Detected A+E tape IDs from ${modState.screeners.length} uploaded element file(s). Enter titles for IDs matched to this order, then start the one-by-one review.`;}
  else{modUpdateAeOrderState();mod$("modFrameStatus").textContent=`All ${shots.length} shot cards are in show time order. Scroll through the list to load the images; the reviewed order can be written after every shot has been prepared.`;if(!modState.frameObserver)update();}
}
function modCleanOcr(text){const normalized=String(text||"").toUpperCase().replace(/[OQD]/g,"0").replace(/[IL|]/g,"1").replace(/[.,]/g,":").replace(/\s+/g,"");const m=normalized.match(/\d{2}[:;]\d{2}[:;]\d{2}[:;]\d{2}/g);return m?m[0].replaceAll(";",":"):"";}
async function modEnsureExcelJs(){if(window.ExcelJS?.Workbook)return window.ExcelJS;if(!modState.excelJsPromise)modState.excelJsPromise=new Promise((resolve,reject)=>{const script=document.createElement("script");script.src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js";script.onload=()=>window.ExcelJS?.Workbook?resolve(window.ExcelJS):reject(Error("ExcelJS did not initialize."));script.onerror=()=>reject(Error("Could not load the Excel workbook generator."));document.head.append(script);}).catch(e=>{modState.excelJsPromise=null;throw e;});return modState.excelJsPromise;}
function modImageJpeg(canvas){const maxWidth=640,scale=Math.min(1,maxWidth/(canvas.width||1)),temp=document.createElement("canvas");temp.width=Math.max(1,Math.round((canvas.width||640)*scale));temp.height=Math.max(1,Math.round((canvas.height||360)*scale));temp.getContext("2d").drawImage(canvas,0,0,temp.width,temp.height);const data=temp.toDataURL("image/jpeg",.88);temp.width=1;temp.height=1;return data;}
function modAeSortedShots(){const fps=modState.mediaMetadata?.fps||24;return [...modState.shots].filter(s=>s.confirmed).sort((a,b)=>String(a.aeTapeNumber).localeCompare(String(b.aeTapeNumber),undefined,{numeric:true,sensitivity:"base"})||String(a.aeProgramTitle).localeCompare(String(b.aeProgramTitle),undefined,{numeric:true,sensitivity:"base"})||(modTcFramesSafe(a.ocrIn,fps)-modTcFramesSafe(b.ocrIn,fps))||(a.rowNumber||0)-(b.rowNumber||0));}
function modTcFramesSafe(tc,fps){try{return modParseTcFrames(tc,fps);}catch{return Number.MAX_SAFE_INTEGER;}}
async function modExportAeOrderForm(){
  if(!modIsAEMediaResources(modState.activeSource))return;
  modUpdateAeOrderState();if(mod$("modExportAeOrder").disabled)return;
  modState.exportingAe=true;const button=mod$("modExportAeOrder"),status=mod$("modAeOrderStatus");button.disabled=true;
  try{
    status.textContent="Loading the Excel order-form generator…";const ExcelJS=await modEnsureExcelJs();const workbook=new ExcelJS.Workbook();workbook.creator="Lovely Runner";workbook.created=new Date();workbook.subject="A+E Media Resources per-second order";workbook.title="A+E Media Resources Order Form";
    const ws=workbook.addWorksheet("Hi-Res Clip Order Form",{properties:{defaultRowHeight:20},views:[{state:"frozen",xSplit:1,ySplit:12}]});
    ws.columns=[{key:"spacer",width:3},{key:"tape",width:15},{key:"program",width:43},{key:"in",width:15},{key:"out",width:15},{key:"description",width:48},{key:"duration",width:14},{key:"rights",width:22},{key:"image",width:34}];
    ws.mergeCells("B2:H2");ws.getCell("B2").value="Hi-Res Clip Order Template";ws.getCell("B2").font={bold:true,size:15};ws.getCell("B2").alignment={vertical:"middle"};ws.getRow(2).height=26;
    ws.getCell("B3").value="Client:";ws.getCell("C3").value="A+E Factual Studios";ws.getCell("B4").value="Project:";ws.getCell("C4").value=String(mod$("modAeProject").value||"").trim();ws.getCell("C4").alignment={wrapText:true,vertical:"top"};ws.getRow(4).height=40;ws.getCell("B5").value="Date:";ws.getCell("C5").value=new Date();ws.getCell("C5").numFmt="mm/dd/yyyy";
    ws.getCell("C7").value="Only A+E owned and public domain clips are eligible for use.";ws.getCell("C8").value="Shot sheets are available upon request.";ws.getCell("C9").value="List all assets used, including timecode IN and OUT points. Frames are not required.";ws.getCell("C7").font={italic:true,color:{argb:"FF555555"}};ws.getCell("C8").font={italic:true,color:{argb:"FF555555"}};ws.getCell("C9").font={italic:true,color:{argb:"FF555555"}};
    ws.getCell("D11").value="Hour: Min: Sec";ws.getCell("F11").value="Total Time Used:";ws.getCell("F11").font={bold:true};ws.getCell("B12").value="A&E Tape#";ws.getCell("C12").value="Program/Tape Title";ws.getCell("D12").value="Source TC In";ws.getCell("E12").value="Source TC Out";ws.getCell("F12").value="Clip / Shot Description";ws.getCell("G12").value="Duration";ws.getCell("H12").value="Rights Assessment";ws.getCell("I12").value="Image";
    for(let c=2;c<=9;c++){const cell=ws.getRow(12).getCell(c);cell.font={bold:true,color:{argb:"FFFFFFFF"}};cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF315A88"}};cell.alignment={wrapText:true,vertical:"middle"};cell.border={bottom:{style:"thin",color:{argb:"FF18324F"}}};}ws.getRow(12).height=32;
    const shots=modAeSortedShots(),fps=modState.mediaMetadata?.fps||24;let row=13,lastDetail=12,totalDuration=0,groupKey="";
    for(let i=0;i<shots.length;i++){
      const shot=shots[i],key=String(shot.aeTapeNumber).trim()+"\u0000"+String(shot.aeProgramTitle).trim();
      if(key!==groupKey){groupKey=key;ws.getCell(`B${row}`).value=shot.aeTapeNumber;ws.getCell(`C${row}`).value=shot.aeProgramTitle;ws.getCell(`D${row}`).value="In";ws.getCell(`E${row}`).value="Out";ws.getCell(`F${row}`).value="Description";ws.getCell(`G${row}`).value="Source Duration (EDL)";ws.getCell(`I${row}`).value="Image";for(let c=2;c<=9;c++){const cell=ws.getCell(row,c);cell.font={bold:true};cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFE7F0FA"}};cell.alignment={wrapText:true,vertical:"middle"};}ws.getRow(row).height=26;row++;}
      if(!shot.canvases[2].dataset.captured){status.textContent=`Preparing Source In image ${i+1} of ${shots.length}…`;await modState.ensureShotFrames(shot);}
      if(!shot.canvases[2].dataset.captured||shot.canvases[2].dataset.error==="true")throw Error(`Could not include the Source In image for tracker row ${shot.rowNumber}.`);
      const sourceFps=modState.screenerMetadata.get(modScreenerKey(shot.screener))?.fps||fps,inValue=modTcExcelDay(shot.ocrIn,sourceFps),outValue=modTcExcelDay(shot.ocrOut,sourceFps),durationSeconds=modDurationSeconds(shot.sourceDuration,sourceFps),durationText=String(shot.sourceDuration??"").trim();if(inValue==null||outValue==null||durationSeconds==null||!durationText)throw Error(`A timecode or Source Duration became invalid on tracker row ${shot.rowNumber}.`);
      ws.getCell(`B${row}`).value=String(shot.aeTapeNumber).trim();ws.getCell(`C${row}`).value=String(shot.aeProgramTitle).trim();ws.getCell(`D${row}`).value=inValue;ws.getCell(`E${row}`).value=outValue;ws.getCell(`F${row}`).value=String(shot.aeDescription).trim();ws.getCell(`G${row}`).value=durationText;ws.getCell(`H${row}`).value="";ws.getCell(`D${row}`).numFmt="h:mm:ss";ws.getCell(`E${row}`).numFmt="h:mm:ss";ws.getCell(`G${row}`).numFmt="@";
      for(let c=2;c<=9;c++){const cell=ws.getCell(row,c);cell.alignment={wrapText:true,vertical:"top"};cell.border={bottom:{style:"hair",color:{argb:"FFD7DEE8"}}};}
      ws.getRow(row).height=104;const imageData=modImageJpeg(shot.canvases[2]),imageId=workbook.addImage({base64:imageData,extension:"jpeg"});ws.addImage(imageId,`I${row}:I${row}`);totalDuration+=durationSeconds/86400;lastDetail=row;row++;
      if(!shot.nearViewport&&!shot.cleanupNear)modState.releaseShotFrames?.(shot);
    }
    ws.getCell("G11").value=totalDuration;ws.getCell("G11").numFmt="[h]:mm:ss";ws.getCell("G11").font={bold:true};
    status.textContent="Creating the downloadable order form…";const buffer=await workbook.xlsx.writeBuffer();const blob=new Blob([buffer],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),url=URL.createObjectURL(blob),a=document.createElement("a"),slug=String(mod$("modAeProject").value||"A+E Order").replace(/[^A-Za-z0-9_-]+/g,"-").replace(/^-|-$/g,"");a.href=url;a.download=`A+E Media Resources Order - ${slug||"Order"}.xlsx`;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);status.textContent=`Downloaded the A+E order form with ${shots.length} shot(s) and Source In images.`;
  }catch(e){status.textContent="Could not create the A+E order form: "+modErr(e);}finally{modState.exportingAe=false;modUpdateAeOrderState();}
}
async function modWriteSecond(){if(!modState.shots.length)return;const vendor=modState.activeSource;if(!confirm(`Write ${modState.shots.length} reviewed per-second rows for ${vendor} to a new tracker tab?`))return;const headers=["Vendor","Original File Name","ARC Number","Description","Source In","Source Out","Cut / Timeline In","Cut / Timeline Out","Burnt-In TC In (Review)","Burnt-In TC Out (Review)","Screener File","Tracker Row","Review Notes"];const values=[headers,...modState.shots.map(s=>[s.vendor,s.filename,s.arc,s.description,s.srcIn,s.srcOut,s.timeIn,s.timeOut,s.ocrIn,s.ocrOut,s.screener?.name||"",s.rowNumber,""])];try{const name=modUniqueTabName("MO Sec "+vendor.replace(/[^A-Za-z0-9 ]/g,"").slice(0,35));await modWriteTab(name,values);mod$("modFrameStatus").textContent=`Created “${name}”. Burnt-In TC fields contain OCR suggestions for human review.`;}catch(e){mod$("modFrameStatus").textContent="Could not write the per-second tab: "+modErr(e);}}
mod$("modCutFile").addEventListener("change",()=>{const file=mod$("modCutFile").files[0];if(modState.activeRangeVideo){modState.activeRangeVideo.pause();modState.activeRangeVideo.removeAttribute("src");delete modState.activeRangeVideo.dataset.cutUrl;modState.activeRangeVideo.load();modState.activeRangeVideo=null;}if(modState.cutObjectUrl)URL.revokeObjectURL(modState.cutObjectUrl);modState.cut=file||null;modState.cutObjectUrl=file?URL.createObjectURL(file):"";modState.mediaMetadata=null;document.querySelector(".mod-cut-options")?.classList.add("hidden");if(file)modState.metadataPromise=modDetectMetadata(file).catch(e=>{console.warn("Cut metadata detection failed",e);return null;});else{mod$("modMetadataStatus").textContent="Choose a reference cut to detect its embedded start timecode and frame rate.";}});mod$("modScreenerFiles").addEventListener("change",()=>{modLoadLocalFiles();const selected=[...modState.screeners],previous=modState.screenerMetadataPromise||Promise.resolve();modState.screenerMetadataPromise=previous.catch(()=>null).then(()=>modDetectScreenerMetadata(selected)).catch(e=>{console.warn("Screener metadata scan failed",e);return null;});});mod$("modConnect").addEventListener("click",modConnect);mod$("modDisconnect").addEventListener("click",modDisconnect);mod$("modRatesTab").addEventListener("change",()=>modLoadRateTab().catch(e=>modStatus("Could not read vendor rates: "+modErr(e),"error")));mod$("modSourceTab").addEventListener("change",async()=>{try{const x=await modReadTab(mod$("modSourceTab").value);modBuildMap(x[0]||[]);}catch(e){mod$("modLoadStatus").textContent=modErr(e);}});mod$("modLoadRows").addEventListener("click",modLoadRows);mod$("modGenerateOrders").addEventListener("click",modGenerateOrders);mod$("modBuildFrames").addEventListener("click",modBuildFrames);mod$("modWriteSecond").addEventListener("click",modWriteSecond);mod$("modExportAeOrder").addEventListener("click",modExportAeOrderForm);mod$("modAeProject").addEventListener("input",modUpdateAeOrderState);mod$("modAeDurationColumn").addEventListener("change",modUpdateAeDurationValues);
mod$("modAeStartReview").addEventListener("click",()=>{if(mod$("modAeStartReview").disabled)return;mod$("modAeTapePage").classList.add("hidden");mod$("modShots").classList.remove("hidden");mod$("modAeReviewControls").classList.remove("hidden");modShowAeShot(modState.aeReviewIndex);});mod$("modAePrevShot").addEventListener("click",()=>modShowAeShot(modState.aeReviewIndex-1));mod$("modAeEditTapeTitles").addEventListener("click",()=>{const shot=modState.shots[modState.aeReviewIndex];if(shot){shot.nearViewport=false;shot.keepFrames=false;modState.releaseShotFrames?.(shot);}mod$("modShots").classList.add("hidden");mod$("modAeTapePage").classList.remove("hidden");modUpdateAeTapeStatus();modUpdateAeReviewControls();});
modInitGoogle();

window.addEventListener("pagehide",()=>{modState.frameObserver?.disconnect();modState.frameCleanupObserver?.disconnect();if(modState.cutObjectUrl)URL.revokeObjectURL(modState.cutObjectUrl);});

document.addEventListener("keydown",modStepCutFrame,true);
