/* الباحث الإداري عن المواقع في سوريا - منطق التطبيق */
(function(){
"use strict";

/* ---------- Master data ---------- */
const MASTER = MASTER_LOCATIONS.map(r=>({
  gEn:r[0],gAr:r[1],gP:r[2],
  dEn:r[3],dAr:r[4],dP:r[5],
  sEn:r[6],sAr:r[7],sP:r[8],
  cEn:r[9],cAr:r[10],cP:r[11],
  type:r[12],lon:r[13],lat:r[14]
}));

/* Coarse grid used only to shortlist candidates before an exact haversine
   comparison — never trusted alone, so it cannot mis-rank a nearest point. */
const CELL = 1.0;
const grid = new Map();
MASTER.forEach((p,idx)=>{
  const key = Math.floor(p.lat/CELL)+"_"+Math.floor(p.lon/CELL);
  if(!grid.has(key)) grid.set(key,[]);
  grid.get(key).push(idx);
});

const toRad = d=>d*Math.PI/180;

function haversine(lat1,lon1,lat2,lon2){
  const R=6371;
  const dLat=toRad(lat2-lat1), dLon=toRad(lon2-lon1);
  const a=Math.sin(dLat/2)**2 + Math.cos(toRad(lat1))*Math.cos(toRad(lat2))*Math.sin(dLon/2)**2;
  return 2*R*Math.asin(Math.sqrt(a));
}

function findNearest(lat,lon){
  const cLat=Math.floor(lat/CELL), cLon=Math.floor(lon/CELL);
  let best=null,bestDist=Infinity;

  // Expand the search box ring by ring. After each ring, check whether any
  // point outside the searched box could still possibly be closer than the
  // best found so far before stopping. The minimum distance to an unsearched
  // cell is `ring*CELL` degrees; converting that to km with cos(lat) (the
  // shorter of the lat/lon axes) keeps the bound conservative so this never
  // stops early and picks the wrong nearest point in sparsely covered areas.
  for(let ring=0; ring<50; ring++){
    for(let dLat=-ring; dLat<=ring; dLat++){
      for(let dLon=-ring; dLon<=ring; dLon++){
        if(Math.max(Math.abs(dLat),Math.abs(dLon))!==ring) continue;
        const bucket=grid.get((cLat+dLat)+"_"+(cLon+dLon));
        if(!bucket) continue;
        for(const idx of bucket){
          const p=MASTER[idx];
          const d=haversine(lat,lon,p.lat,p.lon);
          if(d<bestDist){bestDist=d;best=p;}
        }
      }
    }
    if(best!==null){
      const safeRadiusKm = ring * CELL * 111 * Math.cos(toRad(lat));
      if(bestDist <= safeRadiusKm) break;
    }
  }
  return {point:best,distKm:bestDist};
}

/* ---------- Inside-Syria check (flags likely data-entry errors) ---------- */
function pointInRing(lat,lon,ring){
  let inside=false;
  for(let i=0,j=ring.length-1;i<ring.length;j=i++){
    const xi=ring[i][0], yi=ring[i][1], xj=ring[j][0], yj=ring[j][1];
    const intersect = ((yi>lat)!==(yj>lat)) && (lon < (xj-xi)*(lat-yi)/(yj-yi)+xi);
    if(intersect) inside=!inside;
  }
  return inside;
}
function pointInGeometry(lat,lon,geometry){
  const polys = geometry.type==="Polygon" ? [geometry.coordinates] : geometry.coordinates;
  return polys.some(rings=>{
    if(!pointInRing(lat,lon,rings[0])) return false;
    for(let k=1;k<rings.length;k++){ if(pointInRing(lat,lon,rings[k])) return false; }
    return true;
  });
}
function isInsideSyria(lat,lon){
  return BOUNDARIES.governorate.features.some(f=>pointInGeometry(lat,lon,f.geometry));
}

/* ---------- State ---------- */
let originalColumns = [];
let matchedData = [];
let filteredData = [];
let workbookSheets = {};
let mapLevel = "governorate";
let filters = {gP:"",dP:"",sP:""};
let currentPage = 1, pageSize = 25, searchTerm = "";
let map=null, tileLayer=null, boundaryLayer=null, markerLayer=null;

/* ---------- Theme ---------- */
const root = document.documentElement;
function applyTheme(t){
  root.setAttribute("data-theme", t);
  document.getElementById("themeIconSun").hidden = (t!=="light");
  document.getElementById("themeIconMoon").hidden = (t==="light");
  localStorage.setItem("sal_admin_lookup_theme", t);
  if(tileLayer && map){
    map.removeLayer(tileLayer);
    tileLayer = buildTileLayer();
    tileLayer.addTo(map);
  }
}
document.getElementById("themeToggle").addEventListener("click", ()=>{
  const cur = root.getAttribute("data-theme");
  applyTheme(cur==="dark" ? "light" : "dark");
});
applyTheme(localStorage.getItem("sal_admin_lookup_theme") || "dark");

/* ---------- Upload ---------- */
const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
document.getElementById("chooseFileBtn").addEventListener("click", ()=>fileInput.click());
dropzone.addEventListener("click", ()=>fileInput.click());
["dragenter","dragover"].forEach(evt=>dropzone.addEventListener(evt,e=>{e.preventDefault();dropzone.classList.add("drag");}));
["dragleave","drop"].forEach(evt=>dropzone.addEventListener(evt,e=>{e.preventDefault();dropzone.classList.remove("drag");}));
dropzone.addEventListener("drop", e=>{
  if(e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener("change", e=>{
  if(e.target.files.length) handleFile(e.target.files[0]);
});

document.getElementById("demoBtn").addEventListener("click", loadDemoData);
document.getElementById("resetBtn").addEventListener("click", ()=>location.reload());

function handleFile(file){
  const reader = new FileReader();
  reader.onload = e=>{
    const data = new Uint8Array(e.target.result);
    const wb = XLSX.read(data, {type:"array", cellDates:false});
    workbookSheets = {};
    wb.SheetNames.forEach(name=>{ workbookSheets[name] = wb.Sheets[name]; });
    const sheetSelect = document.getElementById("sheetSelect");
    const sheetField = document.getElementById("sheetField");
    sheetSelect.innerHTML = "";
    if(wb.SheetNames.length > 1){
      sheetField.hidden = false;
      wb.SheetNames.forEach(name=>{
        const opt = document.createElement("option");
        opt.value = name; opt.textContent = name;
        sheetSelect.appendChild(opt);
      });
    } else {
      sheetField.hidden = true;
    }
    loadSheet(wb.SheetNames[0]);
    document.getElementById("resetBtn").hidden = false;
  };
  reader.readAsArrayBuffer(file);
}

document.getElementById("sheetSelect").addEventListener("change", e=>loadSheet(e.target.value));

let rawRows = [];
function loadSheet(sheetName){
  const ws = workbookSheets[sheetName];
  rawRows = XLSX.utils.sheet_to_json(ws, {defval:""});
  if(rawRows.length===0){ alert("لا توجد بيانات في هذه الورقة"); return; }
  originalColumns = Object.keys(rawRows[0]);
  populateMappingUI();
  document.getElementById("mappingSection").hidden = false;
  document.getElementById("mappingSection").scrollIntoView({behavior:"smooth", block:"start"});
}

function guessColumn(keywords){
  let bestCol = "", bestScore = -1;
  originalColumns.forEach(col=>{
    const low = col.toString().trim().toLowerCase();
    let score = -1;
    keywords.forEach((kw,i)=>{
      if(low===kw) score = Math.max(score, 100-i);
      else if(low.includes(kw)) score = Math.max(score, 50-i);
    });
    if(score>bestScore){ bestScore=score; bestCol=col; }
  });
  return bestScore>=0 ? bestCol : originalColumns[0];
}

function populateMappingUI(){
  const latSelect = document.getElementById("latSelect");
  const lonSelect = document.getElementById("lonSelect");
  latSelect.innerHTML = ""; lonSelect.innerHTML = "";
  originalColumns.forEach(col=>{
    const o1 = document.createElement("option"); o1.value=col; o1.textContent=col; latSelect.appendChild(o1);
    const o2 = document.createElement("option"); o2.value=col; o2.textContent=col; lonSelect.appendChild(o2);
  });
  latSelect.value = guessColumn(["latitude","lat","y","خط العرض","العرض","لاتيتيود"]);
  lonSelect.value = guessColumn(["longitude","long","lon","lng","x","خط الطول","الطول","لونجيتيود"]);
  renderPreview();
}
document.getElementById("latSelect").addEventListener("change", renderPreview);
document.getElementById("lonSelect").addEventListener("change", renderPreview);

function renderPreview(){
  const latCol = document.getElementById("latSelect").value;
  const lonCol = document.getElementById("lonSelect").value;
  const table = document.getElementById("previewTable");
  const cols = originalColumns;
  let html = "<thead><tr>" + cols.map(c=>{
    const hl = (c===latCol||c===lonCol) ? ' style="color:var(--accent)"' : "";
    return `<th${hl}>${escapeHtml(c)}</th>`;
  }).join("") + "</tr></thead><tbody>";
  rawRows.slice(0,5).forEach(row=>{
    html += "<tr>" + cols.map(c=>`<td>${escapeHtml(row[c])}</td>`).join("") + "</tr>";
  });
  html += "</tbody>";
  table.innerHTML = html;
}

function escapeHtml(v){
  if(v===null||v===undefined) return "";
  return String(v).replace(/[&<>"']/g, m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]));
}

function parseCoord(v){
  if(typeof v === "number") return v;
  if(v===null||v===undefined) return NaN;
  const s = String(v).trim().replace(",", ".");
  const n = parseFloat(s);
  return n;
}

/* ---------- Processing / matching ---------- */
document.getElementById("processBtn").addEventListener("click", startProcessing);

function startProcessing(){
  const latCol = document.getElementById("latSelect").value;
  const lonCol = document.getElementById("lonSelect").value;
  document.getElementById("mappingSection").hidden = true;
  document.getElementById("progressSection").hidden = false;
  document.getElementById("progressSection").scrollIntoView({behavior:"smooth", block:"start"});

  matchedData = [];
  const total = rawRows.length;
  const CHUNK = 300;
  let i = 0;
  const bar = document.getElementById("progressBar");
  const label = document.getElementById("progressLabel");

  function step(){
    const end = Math.min(i+CHUNK, total);
    for(; i<end; i++){
      const row = rawRows[i];
      const lat = parseCoord(row[latCol]);
      const lon = parseCoord(row[lonCol]);
      const valid = isFinite(lat) && isFinite(lon) && Math.abs(lat)<=90 && Math.abs(lon)<=180 && !(lat===0&&lon===0);
      let rec = Object.assign({}, row, {
        adm_id:i, adm_lat:lat, adm_lon:lon, adm_valid:valid, adm_outsideSyria:false,
        adm_gEn:"",adm_gAr:"",adm_gP:"",adm_dEn:"",adm_dAr:"",adm_dP:"",
        adm_sEn:"",adm_sAr:"",adm_sP:"",adm_cEn:"",adm_cAr:"",adm_cP:"",
        adm_type:"",adm_distKm:null
      });
      if(valid){
        const {point,distKm} = findNearest(lat,lon);
        const insideSyria = isInsideSyria(lat,lon);
        rec.adm_outsideSyria = !insideSyria;
        if(point){
          // "nearest known point" is informative regardless of containment.
          rec.adm_cEn=point.cEn; rec.adm_cAr=point.cAr; rec.adm_cP=point.cP;
          rec.adm_type=point.type; rec.adm_distKm=Math.round(distKm*100)/100;
        }
        if(insideSyria && point){
          rec.adm_gEn=point.gEn; rec.adm_gAr=point.gAr; rec.adm_gP=point.gP;
          rec.adm_dEn=point.dEn; rec.adm_dAr=point.dAr; rec.adm_dP=point.dP;
          rec.adm_sEn=point.sEn; rec.adm_sAr=point.sAr; rec.adm_sP=point.sP;
        } else {
          rec.adm_gEn="Outside Syria"; rec.adm_gAr="خارج سوريا"; rec.adm_gP="OUTSIDE";
          rec.adm_dEn=""; rec.adm_dAr=""; rec.adm_dP="";
          rec.adm_sEn=""; rec.adm_sAr=""; rec.adm_sP="";
        }
      }
      matchedData.push(rec);
    }
    bar.style.width = Math.round((i/total)*100) + "%";
    label.textContent = i.toLocaleString("en-US") + " / " + total.toLocaleString("en-US");
    if(i<total){
      setTimeout(step, 0);
    } else {
      finishProcessing();
    }
  }
  step();
}

function finishProcessing(){
  document.getElementById("progressSection").hidden = true;
  const invalidCount = matchedData.filter(r=>!r.adm_valid).length;
  const outsideCount = matchedData.filter(r=>r.adm_valid && r.adm_outsideSyria).length;
  const alertEl = document.getElementById("invalidAlert");
  const messages = [];
  if(invalidCount>0) messages.push(`تم تجاهل ${invalidCount} صف بسبب عدم وجود إحداثيات صالحة (لن تظهر على الخريطة أو ضمن التصفية الإدارية).`);
  if(outsideCount>0) messages.push(`تم تصنيف ${outsideCount} موقع كـ "خارج سوريا" لأن إحداثياته تقع خارج حدود الجمهورية العربية السورية — قد يكون هذا خطأ في إدخال الإحداثيات، يمكنك عزلها عبر تصفية "خارج سوريا" في قائمة المحافظة.`);
  if(messages.length){
    alertEl.hidden = false;
    alertEl.textContent = "تنبيه: " + messages.join(" ");
  } else {
    alertEl.hidden = true;
  }
  buildFilterOptions();
  buildExportColumnLists();
  document.getElementById("filtersSection").hidden = false;
  document.getElementById("resultsSection").hidden = false;
  document.getElementById("statsSection").hidden = false;
  document.getElementById("exportSection").hidden = false;
  applyFilters();
  document.getElementById("filtersSection").scrollIntoView({behavior:"smooth", block:"start"});
}

/* ---------- Filters ---------- */
function uniqueBy(list, keyFn){
  const seen = new Set(); const out = [];
  list.forEach(item=>{
    const k = keyFn(item);
    if(k && !seen.has(k)){ seen.add(k); out.push(item); }
  });
  return out;
}

function buildFilterOptions(){
  const govFilter = document.getElementById("govFilter");
  const distFilter = document.getElementById("distFilter");
  const subdistFilter = document.getElementById("subdistFilter");

  function fillGov(){
    const valid = matchedData.filter(r=>r.adm_valid);
    const govs = uniqueBy(valid, r=>r.adm_gP).map(r=>({p:r.adm_gP,ar:r.adm_gAr,en:r.adm_gEn}))
      .sort((a,b)=>a.ar.localeCompare(b.ar,"ar"));
    govFilter.innerHTML = '<option value="">كل المحافظات</option>' +
      govs.map(g=>`<option value="${g.p}">${escapeHtml(g.ar)}</option>`).join("");
  }
  function fillDist(){
    let valid = matchedData.filter(r=>r.adm_valid);
    if(filters.gP) valid = valid.filter(r=>r.adm_gP===filters.gP);
    const dists = uniqueBy(valid, r=>r.adm_dP).map(r=>({p:r.adm_dP,ar:r.adm_dAr,en:r.adm_dEn}))
      .sort((a,b)=>a.ar.localeCompare(b.ar,"ar"));
    distFilter.innerHTML = '<option value="">كل المناطق</option>' +
      dists.map(d=>`<option value="${d.p}">${escapeHtml(d.ar)}</option>`).join("");
  }
  function fillSubdist(){
    let valid = matchedData.filter(r=>r.adm_valid);
    if(filters.gP) valid = valid.filter(r=>r.adm_gP===filters.gP);
    if(filters.dP) valid = valid.filter(r=>r.adm_dP===filters.dP);
    const subs = uniqueBy(valid, r=>r.adm_sP).map(r=>({p:r.adm_sP,ar:r.adm_sAr,en:r.adm_sEn}))
      .sort((a,b)=>a.ar.localeCompare(b.ar,"ar"));
    subdistFilter.innerHTML = '<option value="">كل النواحي</option>' +
      subs.map(s=>`<option value="${s.p}">${escapeHtml(s.ar)}</option>`).join("");
  }

  fillGov(); fillDist(); fillSubdist();

  govFilter.onchange = ()=>{
    filters.gP = govFilter.value; filters.dP=""; filters.sP="";
    fillDist(); fillSubdist();
    currentPage=1; applyFilters();
  };
  distFilter.onchange = ()=>{
    filters.dP = distFilter.value; filters.sP="";
    fillSubdist();
    currentPage=1; applyFilters();
  };
  subdistFilter.onchange = ()=>{
    filters.sP = subdistFilter.value;
    currentPage=1; applyFilters();
  };
  document.getElementById("clearFiltersBtn").onclick = ()=>{
    filters = {gP:"",dP:"",sP:""};
    govFilter.value=""; fillDist(); fillSubdist();
    currentPage=1; applyFilters();
  };
}

function applyFilters(){
  filteredData = matchedData.filter(r=>{
    if(!r.adm_valid) return false;
    if(filters.gP && r.adm_gP!==filters.gP) return false;
    if(filters.dP && r.adm_dP!==filters.dP) return false;
    if(filters.sP && r.adm_sP!==filters.sP) return false;
    return true;
  });
  document.getElementById("resultCount").textContent = filteredData.length.toLocaleString("en-US");
  currentPage = 1;
  renderTable();
  renderMap();
  renderStats();
  updateExportCount();
}

/* ---------- Table ---------- */
const tableHead = ()=>{
  return originalColumns.concat(["adm_gAr","adm_dAr","adm_sAr","adm_cAr"]);
};
const admLabels = {adm_gAr:"المحافظة", adm_dAr:"المنطقة", adm_sAr:"الناحية", adm_cAr:"القرية / الحي"};

function getDisplayRows(){
  let rows = filteredData;
  if(searchTerm){
    const q = searchTerm.toLowerCase();
    rows = rows.filter(r=>{
      return originalColumns.some(c=>String(r[c]).toLowerCase().includes(q)) ||
        r.adm_gAr.toLowerCase().includes(q) || r.adm_dAr.toLowerCase().includes(q) ||
        r.adm_sAr.toLowerCase().includes(q) || r.adm_cAr.toLowerCase().includes(q);
    });
  }
  return rows;
}

function renderTable(){
  const rows = getDisplayRows();
  const cols = tableHead();
  const table = document.getElementById("dataTable");
  let thead = "<thead><tr>" + cols.map(c=>`<th>${escapeHtml(admLabels[c]||c)}</th>`).join("") + "</tr></thead>";

  const totalPages = Math.max(1, Math.ceil(rows.length/pageSize));
  if(currentPage>totalPages) currentPage = totalPages;
  const startIdx = (currentPage-1)*pageSize;
  const pageRows = rows.slice(startIdx, startIdx+pageSize);

  let tbody = "<tbody>";
  pageRows.forEach(r=>{
    tbody += "<tr>" + cols.map(c=>`<td>${escapeHtml(r[c])}</td>`).join("") + "</tr>";
  });
  tbody += "</tbody>";
  table.innerHTML = thead + tbody;

  document.getElementById("pageIndicator").textContent = `${currentPage} / ${totalPages}`;
  const rangeStart = rows.length? startIdx+1 : 0;
  const rangeEnd = Math.min(startIdx+pageSize, rows.length);
  document.getElementById("tableRangeLabel").textContent = `عرض ${rangeStart.toLocaleString("en-US")}–${rangeEnd.toLocaleString("en-US")} من ${rows.length.toLocaleString("en-US")}`;
}

document.getElementById("tableSearch").addEventListener("input", e=>{
  searchTerm = e.target.value.trim();
  currentPage = 1;
  renderTable();
});
document.getElementById("pageSizeSelect").addEventListener("change", e=>{
  pageSize = parseInt(e.target.value,10);
  currentPage = 1;
  renderTable();
});
document.getElementById("prevPageBtn").addEventListener("click", ()=>{
  if(currentPage>1){ currentPage--; renderTable(); }
});
document.getElementById("nextPageBtn").addEventListener("click", ()=>{
  const rows = getDisplayRows();
  const totalPages = Math.max(1, Math.ceil(rows.length/pageSize));
  if(currentPage<totalPages){ currentPage++; renderTable(); }
});

/* ---------- Map ---------- */
function buildTileLayer(){
  const theme = root.getAttribute("data-theme");
  const url = theme==="light"
    ? "https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png"
    : "https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png";
  return L.tileLayer(url, {
    attribution:"&copy; OpenStreetMap &copy; CARTO",
    subdomains:"abcd", maxZoom:19
  });
}

function initMap(){
  if(map) return;
  map = L.map("mapEl", {zoomControl:true, attributionControl:true}).setView([35.0,38.3],7);
  tileLayer = buildTileLayer();
  tileLayer.addTo(map);
  boundaryLayer = L.geoJSON(null).addTo(map);
  markerLayer = L.layerGroup().addTo(map);
}

document.querySelectorAll(".pill").forEach(btn=>{
  btn.addEventListener("click", ()=>{
    document.querySelectorAll(".pill").forEach(b=>b.classList.remove("active"));
    btn.classList.add("active");
    mapLevel = btn.dataset.level;
    renderMap();
  });
});

function pcodeFieldForLevel(level){
  return level==="governorate" ? "adm_gP" : level==="district" ? "adm_dP" : "adm_sP";
}
function filterKeyForLevel(level){
  return level==="governorate" ? "gP" : level==="district" ? "dP" : "sP";
}

function renderMap(){
  initMap();
  boundaryLayer.clearLayers();
  markerLayer.clearLayers();

  const field = pcodeFieldForLevel(mapLevel);
  const relevantPcodes = new Set(filteredData.map(r=>r[field]).filter(Boolean));
  const source = BOUNDARIES[mapLevel];
  const isDark = root.getAttribute("data-theme")!=="light";
  const strokeColor = isDark ? "#3f8f8a" : "#2e6e69";

  const feats = source.features.filter(f=>relevantPcodes.has(f.properties.p));
  const fc = {type:"FeatureCollection", features: feats.length? feats : source.features};

  boundaryLayer.addData(fc);
  boundaryLayer.setStyle({color:strokeColor, weight:1.4, fillColor:strokeColor, fillOpacity: feats.length?0.12:0.04, opacity: feats.length?0.9:0.35});
  boundaryLayer.eachLayer(layer=>{
    const props = layer.feature.properties;
    if(mapLevel==="governorate"){
      layer.bindTooltip(props.nAr, {permanent:true, direction:"center", className:"gov-label", opacity:0.85});
    } else {
      layer.bindTooltip(props.nAr, {direction:"center", opacity:0.9});
    }
    layer.on("mouseover", ()=>layer.setStyle({weight:3, fillOpacity:0.22}));
    layer.on("mouseout", ()=>layer.setStyle({weight:1.4, fillOpacity: feats.length?0.12:0.04}));
    layer.on("click", ()=>{
      const key = filterKeyForLevel(mapLevel);
      filters[key] = props.p;
      if(mapLevel==="governorate"){ filters.dP=""; filters.sP=""; document.getElementById("govFilter").value=props.p; }
      if(mapLevel==="district"){ filters.sP=""; document.getElementById("distFilter").value=props.p; }
      if(mapLevel==="subdistrict"){ document.getElementById("subdistFilter").value=props.p; }
      buildFilterOptions();
      if(mapLevel==="governorate") document.getElementById("govFilter").value=props.p;
      if(mapLevel==="district") document.getElementById("distFilter").value=props.p;
      if(mapLevel==="subdistrict") document.getElementById("subdistFilter").value=props.p;
      applyFilters();
    });
  });

  filteredData.forEach(r=>{
    const marker = L.circleMarker([r.adm_lat, r.adm_lon], {
      radius:5, color:"#ffffff", weight:1, fillColor: isDark ? "#d4a72c" : "#a9790e", fillOpacity:0.95
    });
    const idCol = originalColumns[0];
    let rowsHtml = originalColumns.slice(0,5).map(c=>`<div class="pop-row"><span>${escapeHtml(c)}</span><b>${escapeHtml(r[c])}</b></div>`).join("");
    marker.bindPopup(
      `<div class="pop-title">${escapeHtml(r[idCol])}</div>${rowsHtml}
       <div class="pop-row"><span>المحافظة</span><b>${escapeHtml(r.adm_gAr)}</b></div>
       <div class="pop-row"><span>المنطقة</span><b>${escapeHtml(r.adm_dAr)}</b></div>
       <div class="pop-row"><span>الناحية</span><b>${escapeHtml(r.adm_sAr)}</b></div>
       <div class="pop-row"><span>القرية/الحي الأقرب</span><b>${escapeHtml(r.adm_cAr)}</b></div>`
    );
    marker.addTo(markerLayer);
  });

  try{
    if(feats.length){
      map.fitBounds(boundaryLayer.getBounds(), {padding:[24,24]});
    } else if(filteredData.length){
      const bounds = L.latLngBounds(filteredData.map(r=>[r.adm_lat,r.adm_lon]));
      map.fitBounds(bounds, {padding:[40,40]});
    }
  }catch(e){/* no-op */}

  setTimeout(()=>map.invalidateSize(), 200);
}

/* ---------- Stats ---------- */
function countBy(list, pField, arField){
  const map = new Map();
  list.forEach(r=>{
    const key = r[pField];
    if(!key) return;
    if(!map.has(key)) map.set(key, {name:r[arField], count:0});
    map.get(key).count++;
  });
  return Array.from(map.values()).sort((a,b)=>b.count-a.count);
}

function renderBars(containerId, data){
  const el = document.getElementById(containerId);
  const top = data.slice(0,8);
  const maxCount = data.length? Math.max(...data.map(d=>d.count)) : 1;
  el.innerHTML = top.map(d=>`
    <div class="bar-row${d.outside?" outside":""}">
      <div class="bar-labels"><span>${escapeHtml(d.name)}</span><b>${d.count.toLocaleString("en-US")}</b></div>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.max(4,Math.round(d.count/maxCount*100))}%"></div></div>
    </div>`).join("") + (data.length>8 ? `<div class="progress-label">+ ${data.length-8} أخرى</div>` : "");
}

function renderStats(){
  const valid = filteredData;
  document.getElementById("statTotal").textContent = valid.length.toLocaleString("en-US");

  const outsideRows = valid.filter(r=>r.adm_outsideSyria);
  const insideRows = valid.filter(r=>!r.adm_outsideSyria);

  const govCounts = countBy(insideRows,"adm_gP","adm_gAr");
  const distCounts = countBy(insideRows,"adm_dP","adm_dAr");
  const subdistCounts = countBy(insideRows,"adm_sP","adm_sAr");
  document.getElementById("statGov").textContent = govCounts.length;
  document.getElementById("statDist").textContent = distCounts.length;
  document.getElementById("statSubdist").textContent = subdistCounts.length;
  document.getElementById("statOutside").textContent = outsideRows.length.toLocaleString("en-US");

  const govBars = govCounts.slice();
  if(outsideRows.length>0) govBars.push({name:"خارج سوريا", count:outsideRows.length, outside:true});
  govBars.sort((a,b)=>b.count-a.count);

  renderBars("barsGov", govBars);
  renderBars("barsDist", distCounts);
  renderBars("barsSubdist", subdistCounts);
}

/* ---------- Export ---------- */
const ADM_FIELDS = [
  {key:"gov", ar:"المحافظة", en:"Governorate", get:(r,lang)=> lang==="ar"?r.adm_gAr:r.adm_gEn, default:true},
  {key:"dist", ar:"المنطقة", en:"District", get:(r,lang)=> lang==="ar"?r.adm_dAr:r.adm_dEn, default:true},
  {key:"subdist", ar:"الناحية", en:"Subdistrict", get:(r,lang)=> lang==="ar"?r.adm_sAr:r.adm_sEn, default:true},
  {key:"comm", ar:"القرية / الحي الأقرب", en:"Nearest Community/Neighborhood", get:(r,lang)=> lang==="ar"?r.adm_cAr:r.adm_cEn, default:true},
  {key:"type", ar:"نوع الموقع الأقرب", en:"Nearest Location Type", get:(r,lang)=> r.adm_type==="C" ? (lang==="ar"?"قرية":"Community") : (r.adm_type==="N" ? (lang==="ar"?"حي":"Neighborhood") : ""), default:false},
  {key:"govP", ar:"رمز المحافظة", en:"Governorate Pcode", get:(r)=>r.adm_gP, default:false},
  {key:"distP", ar:"رمز المنطقة", en:"District Pcode", get:(r)=>r.adm_dP, default:false},
  {key:"subdistP", ar:"رمز الناحية", en:"Subdistrict Pcode", get:(r)=>r.adm_sP, default:false},
  {key:"dist_km", ar:"المسافة لأقرب نقطة (كم)", en:"Distance to Nearest Point (km)", get:(r)=>r.adm_distKm, default:false},
];

function buildExportColumnLists(){
  const origList = document.getElementById("origColsList");
  origList.innerHTML = originalColumns.map(c=>`
    <label class="check-item"><input type="checkbox" class="orig-col-cb" value="${escapeHtml(c)}" checked>${escapeHtml(c)}</label>
  `).join("");
  const admList = document.getElementById("admColsList");
  admList.innerHTML = ADM_FIELDS.map(f=>`
    <label class="check-item"><input type="checkbox" class="adm-col-cb" value="${f.key}" ${f.default?"checked":""}>
      <span class="adm-col-label" data-ar="${escapeHtml(f.ar)}" data-en="${escapeHtml(f.en)}">${escapeHtml(f.ar)}</span>
    </label>
  `).join("");
  document.querySelectorAll('input[name="exportLang"]').forEach(r=>{
    r.addEventListener("change", updateExportLabelsLang);
  });
  document.querySelectorAll(".orig-col-cb,.adm-col-cb").forEach(cb=>cb.addEventListener("change", updateExportCount));
  updateExportLabelsLang();
}

function currentExportLang(){
  return document.querySelector('input[name="exportLang"]:checked').value;
}
function updateExportLabelsLang(){
  const lang = currentExportLang();
  document.querySelectorAll(".adm-col-label").forEach(el=>{
    el.textContent = lang==="ar" ? el.dataset.ar : el.dataset.en;
  });
  updateExportCount();
}

function updateExportCount(){
  document.getElementById("exportCountLabel").textContent =
    `سيتم تصدير ${filteredData.length.toLocaleString("en-US")} صف بحسب التصفية الحالية`;
}

document.getElementById("downloadBtn").addEventListener("click", ()=>{
  const lang = currentExportLang();
  const selOrig = Array.from(document.querySelectorAll(".orig-col-cb:checked")).map(cb=>cb.value);
  const selAdmKeys = Array.from(document.querySelectorAll(".adm-col-cb:checked")).map(cb=>cb.value);
  const admFields = ADM_FIELDS.filter(f=>selAdmKeys.includes(f.key));

  if(selOrig.length===0 && admFields.length===0){
    alert("الرجاء اختيار عمود واحد على الأقل للتصدير");
    return;
  }
  if(filteredData.length===0){
    alert("لا توجد بيانات مطابقة للتصفية الحالية للتصدير");
    return;
  }

  const exportRows = filteredData.map(r=>{
    const row = {};
    selOrig.forEach(c=>{ row[c] = r[c]; });
    admFields.forEach(f=>{ row[lang==="ar"?f.ar:f.en] = f.get(r, lang); });
    return row;
  });

  const ws = XLSX.utils.json_to_sheet(exportRows);
  const allCols = selOrig.concat(admFields.map(f=>lang==="ar"?f.ar:f.en));
  ws["!cols"] = allCols.map(()=>({wch:20}));
  const wb = XLSX.utils.book_new();
  if(lang==="ar") wb.Workbook = {Views:[{RTL:true}]};
  XLSX.utils.book_append_sheet(wb, ws, lang==="ar" ? "النتائج" : "Results");
  const fname = (lang==="ar" ? "نتائج_المواقع_الادارية" : "Administrative_Locations_Results") + ".xlsx";
  XLSX.writeFile(wb, fname);
});

/* ---------- Demo data ---------- */
function loadDemoData(){
  const sample = [];
  const step = Math.floor(MASTER.length/45);
  for(let i=0;i<45;i++){
    const base = MASTER[(i*step) % MASTER.length];
    const jitterLat = (Math.random()-0.5)*0.03;
    const jitterLon = (Math.random()-0.5)*0.03;
    sample.push({
      "اسم الموقع": "موقع تجريبي " + (i+1),
      "النوع": ["فرع بيع","مستودع","مكتب"][i%3],
      "خط العرض": Math.round((base.lat+jitterLat)*1e6)/1e6,
      "خط الطول": Math.round((base.lon+jitterLon)*1e6)/1e6
    });
  }
  rawRows = sample;
  originalColumns = Object.keys(sample[0]);
  workbookSheets = {};
  document.getElementById("sheetField").hidden = true;
  populateMappingUI();
  document.getElementById("mappingSection").hidden = false;
  document.getElementById("resetBtn").hidden = false;
  document.getElementById("mappingSection").scrollIntoView({behavior:"smooth", block:"start"});
}

})();
