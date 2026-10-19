const APP_PASSWORD = "cpm@2026";

const AUTH_KEY = "dashboard_auth";
const AUTH_TIME_KEY = "dashboard_auth_time";
const SESSION_TIMEOUT = 5 * 60 * 1000; // 5 phút

window.addEventListener("DOMContentLoaded", () => {
    const overlay = document.getElementById("passwordOverlay");
    const input = document.getElementById("passwordInput");
    const btn = document.getElementById("passwordBtn");
    const error = document.getElementById("passwordError");

    if (!overlay) return;

    function logout() {
        localStorage.removeItem(AUTH_KEY);
        localStorage.removeItem(AUTH_TIME_KEY);

        report = null;

        overlay.classList.remove("hidden");

        input.value = "";
        error.textContent = "";

        const fileName = document.getElementById("fileName");
        const summaryBody = document.getElementById("summaryBody");

        if (fileName) {
            fileName.textContent = "No file selected";
        }

        if (summaryBody) {
            summaryBody.innerHTML = `
                <tr>
                    <td colspan="5" class="empty">
                        Session expired. Please login again.
                    </td>
                </tr>
            `;
        }

        input.focus();
    }

    function startSessionTimer(remaining = SESSION_TIMEOUT) {
        setTimeout(() => {
            alert("Session expired. Please login again.");
            logout();
        }, remaining);
    }

    const auth = localStorage.getItem(AUTH_KEY);
    const loginTime = Number(localStorage.getItem(AUTH_TIME_KEY) || 0);

    if (
        auth === "ok" &&
        loginTime &&
        Date.now() - loginTime < SESSION_TIMEOUT
    ) {
        overlay.classList.add("hidden");

        startSessionTimer(
            SESSION_TIMEOUT - (Date.now() - loginTime)
        );

        return;
    }

    function login() {
        if (input.value === APP_PASSWORD) {

            localStorage.setItem(AUTH_KEY, "ok");
            localStorage.setItem(
                AUTH_TIME_KEY,
                Date.now()
            );

            overlay.classList.add("hidden");

            startSessionTimer();

        } else {
            error.textContent = "Incorrect password";
            input.value = "";
            input.focus();
        }
    }

    btn.addEventListener("click", login);

    input.addEventListener("keypress", (e) => {
        if (e.key === "Enter") {
            login();
        }
    });

    input.focus();
});
"use strict";
let report=null;
const $=id=>document.getElementById(id);
const fmt=(v,d=0)=>Number(v||0).toLocaleString(undefined,{minimumFractionDigits:d,maximumFractionDigits:d});
const colors=["#246bfe","#00a6a6","#7c3aed","#f59e0b","#ef476f","#64748b","#22c55e","#06b6d4","#a855f7"];

$("importBtn").onclick=()=>$("fileInput").click();
$("dropZone").onclick=e=>{if(e.target.id!=="downloadTemplate")$("fileInput").click()};
$("fileInput").onchange=e=>e.target.files[0]&&loadFile(e.target.files[0]);
["dragenter","dragover"].forEach(x=>$("dropZone").addEventListener(x,e=>{e.preventDefault();$("dropZone").classList.add("drag")}));
["dragleave","drop"].forEach(x=>$("dropZone").addEventListener(x,e=>{e.preventDefault();$("dropZone").classList.remove("drag")}));
$("dropZone").addEventListener("drop",e=>e.dataTransfer.files[0]&&loadFile(e.dataTransfer.files[0]));
$("viewFilter").onchange=render;
$("monthFilter").onchange=render;
$("resetBtn").onclick=()=>{$("viewFilter").value="ytd";$("monthFilter").selectedIndex=Math.max(0,report.reportingMonthIndex);render()};
$("exportBtn").onclick=exportCSV;
$("downloadTemplate").onclick=()=>{const a=document.createElement("a");a.href="Reporting_Template.xlsx";a.download="Reporting_Template.xlsx";a.click()};

async function loadFile(file){
  setStatus("Reading "+file.name+"...","");
  try{
    if(typeof XLSX==="undefined")throw new Error("Excel reader library could not load. Connect to the internet once and refresh the page.");
    const buffer=await file.arrayBuffer();
    const wb=XLSX.read(buffer,{type:"array",cellDates:true});
    const ws=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
    report=parseTemplate(rows);
    $("fileName").textContent=file.name;
    configureFilters();
    $("resetBtn").disabled=false;$("exportBtn").disabled=false;
    setStatus(`Loaded ${wb.SheetNames[0]} · ${report.year}`,"ok");
    render();
  }catch(err){report=null;setStatus(err.message,"error");showError(err.message)}
}
function clean(v){return String(v??"").trim().toLowerCase().replace(/\s+/g," ")}
function num(v){const n=typeof v==="number"?v:parseFloat(String(v??"").replace(/,/g,""));return Number.isFinite(n)?n:0}
function findRow(rows,name,start=0){const target=clean(name);for(let i=start;i<rows.length;i++)if(clean(rows[i][0])===target)return i;return -1}
function findSection(rows,name){const target=clean(name);for(let i=0;i<rows.length;i++)for(const c of rows[i])if(clean(c)===target)return i;return -1}
function parseTemplate(rows){
  const monthRow=findRow(rows,"Month");
  if(monthRow<0)throw new Error('Invalid template: row "Month" was not found.');
  const months=rows[monthRow].slice(1,13).map((x,i)=>String(x||`${i+1}`).replace(/^\d+\./,""));
  if(months.filter(Boolean).length!==12)throw new Error("Invalid template: 12 monthly columns are required.");
  const year=num((rows[monthRow-1]||[])[1])||new Date().getFullYear();
  const reportingRaw=(rows[0]||[])[1];
  let reportingMonthIndex=11;
  if(reportingRaw instanceof Date)reportingMonthIndex=reportingRaw.getMonth();
  else if(reportingRaw){const s=String(reportingRaw).toLowerCase();const idx=months.findIndex(m=>s.includes(String(m).toLowerCase().slice(0,3)));if(idx>=0)reportingMonthIndex=idx}
  const hc=findRow(rows,"Total Staff HCs"), hours=findRow(rows,"Total Working Hours log count"), perStaff=findRow(rows,"Working Hour/Staff");
  const whSec=findSection(rows,"Working Hours by Segment"), jrSec=findSection(rows,"Job Requests by Segments"), typeSec=findSection(rows,"Job Requests by Type");
  if([hc,hours,perStaff,whSec,jrSec,typeSec].some(x=>x<0))throw new Error("Invalid template: one or more required KPI sections were not found.");
  const series=r=>rows[r].slice(1,13).map(num);
  const section=(start,end,skipTotal=false)=>{const o={};for(let r=start+1;r<end;r++){const label=String(rows[r]?.[0]||"").trim();if(!label)continue;if(skipTotal&&clean(label)==="total jrs received")continue;o[label]=series(r)}return o};
  const nextAfterType=rows.length;
  return {year,months,reportingMonthIndex,hc:series(hc),hours:series(hours),perStaff:series(perStaff),jrs:series(findRow(rows,"Total JRs Received",jrSec+1)),hoursBySegment:section(whSec,jrSec),jrBySegment:section(jrSec,typeSec,true),jrByType:section(typeSec,nextAfterType,true)};
}
function configureFilters(){
  $("monthFilter").innerHTML=report.months.map((m,i)=>`<option value="${i}">${escapeHtml(m)} ${report.year}</option>`).join("");
  $("monthFilter").value=String(report.reportingMonthIndex);$("monthFilter").disabled=false;
}
function period(){const monthly=$("viewFilter").value==="month", idx=num($("monthFilter").value);return {monthly,idx,label:monthly?`${report.months[idx]} ${report.year}`:`YTD through ${report.months[idx]} ${report.year}`}}
function aggregate(series,p){return p.monthly?num(series[p.idx]):series.slice(0,p.idx+1).reduce((a,b)=>a+num(b),0)}
function snapshot(obj,p){return Object.fromEntries(Object.entries(obj).map(([k,v])=>[k,aggregate(v,p)]))}
function render(){if(!report)return;const p=period();
  $("kpiHC").textContent=fmt(report.hc[p.idx],2);$("kpiHours").textContent=fmt(aggregate(report.hours,p),2);$("kpiJR").textContent=fmt(aggregate(report.jrs,p));
  const ps=p.monthly?report.perStaff[p.idx]:(report.hc[p.idx]?aggregate(report.hours,p)/report.hc[p.idx]:0);$("kpiPerStaff").textContent=fmt(ps,2);
  ["kpiHCSub","kpiHoursSub","kpiPerStaffSub","kpiJRSub","hoursSegmentSubtitle","jrSegmentSubtitle","jrTypeSubtitle"].forEach(id=>$(id).textContent=p.label);
  renderTrend(p.idx);renderBars("hoursSegmentChart",snapshot(report.hoursBySegment,p),2);renderDonut(snapshot(report.jrBySegment,p));renderBars("jrTypeChart",snapshot(report.jrByType,p),0);renderTable(p);
}
function renderBars(id,data,dec){const entries=Object.entries(data).sort((a,b)=>b[1]-a[1]),max=Math.max(...entries.map(x=>x[1]),1);$(id).innerHTML=entries.map(([k,v])=>`<div class="bar-row"><div class="bar-label" title="${escapeHtml(k)}">${escapeHtml(k)}</div><div class="bar-track"><div class="bar-fill" style="width:${v/max*100}%"></div></div><div class="bar-value">${fmt(v,dec)}</div></div>`).join("")||'<div class="empty">No data</div>'}
function renderDonut(data){const entries=Object.entries(data).sort((a,b)=>b[1]-a[1]),total=entries.reduce((s,x)=>s+x[1],0);let a=0;const stops=entries.map(([,v],i)=>{const from=a;a+=total?v/total*100:0;return `${colors[i%colors.length]} ${from}% ${a}%`}).join(",");$("jrSegmentChart").innerHTML=`<div class="donut" style="background:conic-gradient(${stops||'#e8edf3 0 100%'})"><div class="donut-center"><strong>${fmt(total)}</strong><span>Total JRs</span></div></div><div class="donut-legend">${entries.map(([k,v],i)=>`<div class="legend-row"><i class="legend-swatch" style="background:${colors[i%colors.length]}"></i><span>${escapeHtml(k)}</span><b>${fmt(v)}</b></div>`).join("")}</div>`}
function renderTrend(last){const labels=report.months.slice(0,last+1),h=report.hours.slice(0,last+1),j=report.jrs.slice(0,last+1),W=1100,H=270,pad={l:45,r:25,t:18,b:35};const maxH=Math.max(...h,1),maxJ=Math.max(...j,1),x=i=>pad.l+(labels.length===1?(W-pad.l-pad.r)/2:i*(W-pad.l-pad.r)/(labels.length-1)),yh=v=>H-pad.b-v/maxH*(H-pad.t-pad.b),yj=v=>H-pad.b-v/maxJ*(H-pad.t-pad.b);const path=(arr,y)=>arr.map((v,i)=>(i?'L':'M')+x(i)+','+y(v)).join(' ');let grid='';for(let n=0;n<=4;n++){const y=pad.t+n*(H-pad.t-pad.b)/4;grid+=`<line class="gridline" x1="${pad.l}" y1="${y}" x2="${W-pad.r}" y2="${y}"/><text class="axistext" x="5" y="${y+4}">${fmt(maxH*(4-n)/4)}</text>`}const labs=labels.map((m,i)=>`<text class="axistext" x="${x(i)}" y="${H-10}" text-anchor="middle">${escapeHtml(m)}</text>`).join('');const pts=h.map((v,i)=>`<circle class="point-hours" cx="${x(i)}" cy="${yh(v)}" r="4"><title>${labels[i]}: ${fmt(v,2)} hours</title></circle>`).join('')+j.map((v,i)=>`<circle class="point-jobs" cx="${x(i)}" cy="${yj(v)}" r="4"><title>${labels[i]}: ${fmt(v)} JRs</title></circle>`).join('');$("trendChart").innerHTML=`<svg class="trend-svg" viewBox="0 0 ${W} ${H}" preserveAspectRatio="none"><defs><linearGradient id="areaGradient" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#246bfe" stop-opacity=".2"/><stop offset="1" stop-color="#246bfe" stop-opacity="0"/></linearGradient></defs>${grid}<path class="hour-area" d="${path(h,yh)} L${x(h.length-1)},${H-pad.b} L${x(0)},${H-pad.b} Z"/><path class="hour-line" d="${path(h,yh)}"/><path class="job-line" d="${path(j,yj)}"/>${pts}${labs}</svg>`}
function renderTable(p){const idxs=p.monthly?[p.idx]:Array.from({length:p.idx+1},(_,i)=>i);$("summaryBody").innerHTML=idxs.map(i=>`<tr><td>${escapeHtml(report.months[i])} ${report.year}</td><td>${fmt(report.hc[i],2)}</td><td>${fmt(report.hours[i],2)}</td><td>${fmt(report.perStaff[i],2)}</td><td>${fmt(report.jrs[i])}</td></tr>`).join('')}
function exportCSV(){if(!report)return;const p=period(),idxs=p.monthly?[p.idx]:Array.from({length:p.idx+1},(_,i)=>i);const rows=[["Month","Staff HCs","Working Hours","Hours / Staff","Job Requests"],...idxs.map(i=>[`${report.months[i]} ${report.year}`,report.hc[i],report.hours[i],report.perStaff[i],report.jrs[i]])];const csv='\ufeff'+rows.map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(',')).join('\r\n');const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));a.download=`dashboard_${p.label.replace(/\s+/g,'_')}.csv`;a.click();URL.revokeObjectURL(a.href)}
function setStatus(msg,cls){$("status").textContent=msg;$("status").className="status "+cls}
function showError(msg){["trendChart","hoursSegmentChart","jrSegmentChart","jrTypeChart"].forEach(id=>$(id).innerHTML=`<div class="error-box">${escapeHtml(msg)}</div>`)}
function escapeHtml(s){return String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[c]))}
