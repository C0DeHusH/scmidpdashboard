const state={boot:null,charts:[],model:null,requests:[],period:'ytd',theme:document.documentElement.dataset.theme||'dark',branchData:null,classFilter:'A',statusData:null,statusChart:null,management:null,managementEdits:{},showPointData:false,areaDetail:null,reorderCard:null,reorderClassFilter:'All',reorderBrandFilter:'All Brands',requestOutputBranch:'',requestBranchCount:0,requestBuilderModels:[]};
const isAdmin=Boolean(window.SCM_CONFIG?.isAdmin);
if(typeof Chart!=='undefined'){Chart.defaults.font.family='Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif';Chart.defaults.font.size=11;Chart.defaults.font.weight='500';}

const $=s=>document.querySelector(s); const $$=s=>[...document.querySelectorAll(s)];
const requestControllers=new Map();
const nativeFetch=window.fetch.bind(window);
let toastTimer=null;
let authRedirectPending=false;
function escapeHtml(value){return String(value??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]))}
function adminLoginUrl(){
  const next=window.location.pathname+window.location.search+window.location.hash;
  return `/login?next=${encodeURIComponent(next)}`;
}
function scheduleAdminRelogin(message='Your Admin session expired. Please sign in again.'){
  if(authRedirectPending)return;authRedirectPending=true;toast(message,false);
  window.setTimeout(()=>window.location.assign(adminLoginUrl()),700);
}
async function appFetch(url,options={}){
  const response=await nativeFetch(url,{credentials:'same-origin',cache:'no-store',...options});
  const path=typeof url==='string'?url:(url?.url||'');
  if(isAdmin&&path.startsWith('/admin/')&&(response.status===401||response.status===403)){
    let data={};try{data=await response.clone().json()}catch(_){ }
    if(data.reauth_required||String(data.error||'').toLowerCase().includes('admin'))scheduleAdminRelogin(data.error);
  }
  return response;
}
async function verifyAdminSession(){
  if(!isAdmin)return true;
  try{
    const r=await nativeFetch('/api/session',{credentials:'same-origin',cache:'no-store'});
    const d=await r.json();
    if(d.configuration_required){toast('Vercel Admin session setup is incomplete. Configure SCM_SECRET_KEY and redeploy.',false);return false}
    if(!r.ok||!d.is_admin){scheduleAdminRelogin('Admin access needs to be refreshed. Please sign in again.');return false}
    return true;
  }catch(_){return true}
}
async function fetchJson(url,options={},requestKey=''){
  let controller=null;
  if(requestKey){requestControllers.get(requestKey)?.abort();controller=new AbortController();requestControllers.set(requestKey,controller);options={...options,signal:controller.signal}}
  try{
    const response=await appFetch(url,options);let data={};
    try{data=await response.json()}catch(_){data={}}
    if(!response.ok)throw new Error(data.error||`Request failed (${response.status})`);
    return data;
  }catch(err){if(err?.name==='AbortError')return null;throw err}
  finally{if(requestKey&&requestControllers.get(requestKey)===controller)requestControllers.delete(requestKey)}
}
function toast(msg,ok=true){const t=$('#toast');if(!t)return;clearTimeout(toastTimer);t.textContent=msg;t.className=`fixed bottom-4 right-4 z-50 max-w-sm rounded-xl border px-4 py-3 text-sm shadow-2xl ${ok?'border-emerald-600 bg-emerald-950 text-emerald-200':'border-red-600 bg-red-950 text-red-200'}`;t.classList.remove('hidden');toastTimer=setTimeout(()=>t.classList.add('hidden'),3500)}
function responseFilename(response,fallbackName){const cd=response.headers.get('content-disposition')||'';const utf=cd.match(/filename\*=UTF-8''([^;]+)/i),plain=cd.match(/filename=\"?([^\";]+)\"?/i);const raw=utf?.[1]||plain?.[1];if(!raw)return fallbackName;try{return decodeURIComponent(raw.replaceAll('\"','').trim())}catch(_){return raw.replaceAll('\"','').trim()||fallbackName}}
async function downloadResponse(response,fallbackName){const blob=await response.blob(),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=responseFilename(response,fallbackName);document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000)}
function openClearDataModal(){if(!isAdmin)return;const m=$('#clearDataModal'),i=$('#clearDataConfirmation');if(!m)return;if(i)i.value='';m.classList.remove('hidden');m.classList.add('flex');setTimeout(()=>i?.focus(),30)}
function closeClearDataModal(){const m=$('#clearDataModal');if(!m)return;m.classList.add('hidden');m.classList.remove('flex')}
async function clearAllUserData(){const confirmation=String($('#clearDataConfirmation')?.value||'').trim();if(confirmation!=='CLEAR DATA'){toast('Type CLEAR DATA exactly to confirm.',false);return}const btn=$('#confirmClearDataBtn');if(btn){btn.disabled=true;btn.textContent='Clearing…'}try{const r=await appFetch('/admin/clear-data',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({confirmation})});const d=await r.json();if(!r.ok){toast(d.error||'Clear Data failed.',false);return}localStorage.removeItem('scm_weekly_schedule_collapsed');localStorage.removeItem('scm_allocation_import_collapsed');closeClearDataModal();toast(d.message||'All user data cleared.');setTimeout(()=>window.location.reload(),700)}finally{if(btn){btn.disabled=false;btn.textContent='Clear All Data'}}}
function roundHalfUp(v){const n=Number(v);if(!Number.isFinite(n))return null;return Math.sign(n)*Math.floor(Math.abs(n)+0.5)}
function fmt(v){const n=roundHalfUp(v);return n===null?'—':String(n)}
function ceilDoi(v){const n=Number(v);return Number.isFinite(n)?String(Math.ceil(Math.max(0,n))):'—'}
function php(v){const n=roundHalfUp(v);return n===null?'—':new Intl.NumberFormat('en-PH',{style:'currency',currency:'PHP',maximumFractionDigits:0}).format(n)}
function statusClass(s){const slug=String(s||'unknown').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-+|-+$/g,'');return 'status-'+(slug||'unknown')}
function miniCard(label,val,suffix=''){return `<div class="theme-surface rounded-xl border p-3"><div class="text-[10px] font-black uppercase tracking-wider text-slate-500">${escapeHtml(label)}</div><div class="mt-1 text-xl font-black">${escapeHtml(val)}${escapeHtml(suffix)}</div></div>`}
function metricCard(label,val,delta,unit,good='low'){const suffix=unit==='percent'?'%':' d';const d=Number(delta),flat=delta===null||!Number.isFinite(d)||roundHalfUp(d)===0;const favorable=good==='high'?d>0:good==='low'?d<0:null;const tone=delta===null?'text-slate-600':flat?'text-slate-500':favorable===true?'text-emerald-500':favorable===false?'text-red-500':'text-blue-400';const meaning=delta===null?'No comparison':flat?'• Flat vs previous':`${d<0?'▼':'▲'} ${fmt(Math.abs(d))} vs previous${good==='balanced'?'':favorable?' • Positive':' • Negative'}`;return `<div class="theme-surface rounded-xl border p-4"><div class="text-[10px] font-black uppercase tracking-wider text-slate-500">${escapeHtml(label)}</div><div class="mt-1 text-2xl font-black">${unit==='days'?ceilDoi(val):fmt(val)}${escapeHtml(suffix)}</div><div class="mt-1 text-[10px] font-bold ${tone}">${escapeHtml(meaning)}</div></div>`}
function palette(){return state.theme==='light'?{axis:'#244663',grid:'rgba(37,99,235,.14)',actual:'#1d4ed8',trend:'#0369a1',fill:'rgba(37,99,235,.10)'}:{axis:'#d5e5f7',grid:'rgba(96,165,250,.14)',actual:'#3b82f6',trend:'#93c5fd',fill:'rgba(59,130,246,.15)'}}
function monthOnly(label){const s=String(label||'');const m=s.match(/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i);return m?m[1].slice(0,3):s}
function shortCurrentPeriodLabel(label){
  const s=String(label||'').trim();
  const m=s.match(/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{1,2})/i);
  return m?`${m[1].slice(0,3)} ${Number(m[2])}`:s;
}
function longPeriodLabel(label){
  const s=String(label||'').trim();
  const m=s.match(/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{1,2}),?\s+(\d{4})$/i);
  if(!m)return s;
  const names={jan:'January',feb:'February',mar:'March',apr:'April',may:'May',jun:'June',jul:'July',aug:'August',sep:'September',oct:'October',nov:'November',dec:'December'};
  return `${names[m[1].slice(0,3).toLowerCase()]||m[1]} ${Number(m[2])}, ${m[3]}`;
}
function displayLabels(labels){
  if(state.period!=='ytd')return labels;
  return labels.map((label,i)=>i===labels.length-1?shortCurrentPeriodLabel(label):monthOnly(label));
}
const pointValueLabelPlugin={
  id:'pointValueLabels',
  afterDatasetsDraw(chart,args,opts){
    if(!opts?.enabled)return;
    const actualIndex=chart.data.datasets.findIndex(ds=>ds.label==='Actual');
    if(actualIndex<0)return;
    const meta=chart.getDatasetMeta(actualIndex),values=chart.data.datasets[actualIndex].data,ctx=chart.ctx;
    ctx.save();
    ctx.fillStyle=opts.color||'#e2e8f0';
    ctx.font='600 10px Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif';
    ctx.textAlign='center';
    ctx.textBaseline='bottom';
    meta.data.forEach((pt,i)=>{
      const v=values[i];
      if(!Number.isFinite(Number(v))||pt.skip)return;
      const y=Math.max(chart.chartArea.top+12,pt.y-8);
      ctx.fillText(fmt(v),pt.x,y);
    });
    ctx.restore();
  }
};
function syncPointDataButton(){const b=$('#pointDataBtn');if(!b)return;b.textContent=state.showPointData?'Hide Data':'Show Data';b.setAttribute('aria-pressed',state.showPointData?'true':'false');b.classList.toggle('border-blue-400/50',state.showPointData);b.classList.toggle('text-blue-400',state.showPointData);b.classList.toggle('border-slate-700',!state.showPointData);b.classList.toggle('text-slate-300',!state.showPointData)}
function setPointDataVisibility(show){state.showPointData=!!show;syncPointDataButton();state.charts.forEach(c=>{if(c.options.plugins.pointValueLabels)c.options.plugins.pointValueLabels.enabled=state.showPointData;c.update('none')})}
function makeCombinedChart(canvas,labels,values,trend,unit,good='low'){
  const p=palette(),actual=values.map(roundHalfUp),trueTrend=trend.map(v=>Number(v)),labs=displayLabels(labels);
  const numeric=[...actual,...trueTrend].filter(Number.isFinite);
  const min=numeric.length?Math.min(...numeric):0,max=numeric.length?Math.max(...numeric):1,range=Math.max(1,max-min);

  // Trend presentation rule (v1.6): keep the mathematically straight regression
  // direction line inside the SAME chart, with no markers and no curve-following.
  // Apply ONE constant vertical visual offset only, so the complete dashed line
  // remains straight and sits cleanly above the Actual series like the reference.
  const gap=Math.max(1,range*0.10);
  let requiredOffset=0;
  trueTrend.forEach((t,i)=>{
    const a=actual[i];
    if(Number.isFinite(t)&&Number.isFinite(a)) requiredOffset=Math.max(requiredOffset,(a+gap)-t);
  });
  const visualOffset=Math.max(0,requiredOffset);
  const displayTrend=trueTrend.map(t=>Number.isFinite(t)?t+visualOffset:null);

  const first=trueTrend.find(v=>Number.isFinite(v));
  const last=[...trueTrend].reverse().find(v=>Number.isFinite(v));
  const trendDirection=(!Number.isFinite(first)||!Number.isFinite(last)||Math.abs(last-first)<0.0001)?'flat':(last>first?'up':'down');
  const trendIcon=trendDirection==='up'?'↗':trendDirection==='down'?'↘':'→';
  const trendText=trendDirection==='up'?'Upward':trendDirection==='down'?'Downward':'Flat';
  const latestDelta=actual.length>1?Number(actual[actual.length-1])-Number(actual[actual.length-2]):0;
  const latestDirection=Math.abs(latestDelta)<0.0001?'flat':latestDelta>0?'up':'down';
  const latestIcon=latestDirection==='up'?'↑':latestDirection==='down'?'↓':'→';
  const latestFavorable=Math.abs(latestDelta)<0.0001||good==='balanced'?null:(good==='high'?latestDelta>0:latestDelta<0);
  const latestTone=latestFavorable===true?'positive':latestFavorable===false?'negative':'neutral';
  const latestMeaning=latestFavorable===true?'Positive':latestFavorable===false?'Negative':latestDirection==='flat'?'Flat':'Latest';
  const legend=canvas.closest('.kpi-visual')?.querySelector('[data-kpi-chart-legend]');
  if(legend){
    legend.innerHTML=`<span class="kpi-legend-item"><i class="kpi-legend-line trend"></i><b>Trend Direction</b><span class="kpi-direction-icon" title="${trendText} trend" aria-label="${trendText} trend">${trendIcon}</span></span>${actual.length>1?`<span class="kpi-legend-item kpi-latest-move ${latestTone}"><b>Latest</b><span class="kpi-latest-icon">${latestIcon}</span><span>${latestMeaning}</span></span>`:''}<span class="kpi-legend-item"><i class="kpi-legend-line actual"></i><b>Actual</b></span>`;
  }

  const c=new Chart(canvas,{type:'line',data:{labels:labs,datasets:[
    {label:'Actual',data:actual,borderColor:p.actual,backgroundColor:p.fill,pointBackgroundColor:p.actual,pointBorderColor:p.actual,pointRadius:state.period==='ytd'?3:3.5,pointHoverRadius:6,borderWidth:2.7,tension:0,fill:true,order:1},
    {label:'Trend Direction',data:displayTrend,borderColor:p.trend,backgroundColor:'transparent',pointRadius:0,pointHoverRadius:0,borderWidth:2.6,borderDash:[9,6],borderCapStyle:'round',borderJoinStyle:'miter',tension:0,fill:false,spanGaps:true,order:0,trueTrend:trueTrend}
  ]},plugins:[pointValueLabelPlugin],options:{responsive:true,maintainAspectRatio:false,resizeDelay:80,animation:{duration:520,easing:'easeOutQuart'},interaction:{mode:'index',intersect:false},layout:{padding:{left:2,right:4,top:10,bottom:0}},plugins:{pointValueLabels:{enabled:state.showPointData,color:p.axis},legend:{display:false},tooltip:{callbacks:{title:(items)=>{const i=items?.[0]?.dataIndex??-1;if(i<0)return '';if(state.period==='ytd'&&i===labels.length-1)return `Latest data · ${longPeriodLabel(labels[i])}`;return labs[i]||String(labels[i]||'');},label:(ctx)=>{const suffix=unit==='percent'?'%':' days';if(String(ctx.dataset.label||'')==='Trend Direction'){const raw=ctx.dataset.trueTrend?.[ctx.dataIndex];return `Trend Direction: ${fmt(raw)}${suffix}`;}return `Actual: ${fmt(ctx.parsed.y)}${suffix}`;}}}},scales:{x:{offset:false,grid:{display:false},ticks:{color:p.axis,autoSkip:labels.length>12,maxTicksLimit:state.period==='ytd'?12:10,maxRotation:0,minRotation:0,padding:4,font:{size:state.period==='ytd'?9:10}}},y:{grace:'20%',grid:{color:p.grid},ticks:{precision:0,color:p.axis,font:{size:10},callback:v=>unit==='percent'?fmt(v)+'%':fmt(v)}}}}});state.charts.push(c)
}
function renderKPIs(kpis){state.charts.forEach(c=>c.destroy());state.charts=[];const grid=$('#kpiGrid');grid.innerHTML='';const period=state.period,periodLabel=period==='ytd'?'YTD':'WEEKLY';const hasValues=Object.values(kpis||{}).some(k=>((k?.[period]?.values)||[]).length);if(!hasValues){grid.innerHTML='<div class="glass rounded-2xl border p-10 text-center"><div class="text-sm font-black text-slate-300">No KPI data loaded</div><div class="mt-2 text-xs text-slate-500">Import a new SCM workbook to populate KPI graphs.</div></div>';return;}Object.entries(kpis).forEach(([name,k])=>{const id=name.replace(/[^a-z0-9]/gi,'_'),d=k[period];const vals=(d.values||[]).map(Number).filter(Number.isFinite);const avg=vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:null;const high=vals.length?Math.max(...vals):null;const low=vals.length?Math.min(...vals):null;const suffix=k.meta.unit==='percent'?'%':' d';const stat=v=>v===null?'—':`${k.meta.unit==='days'?ceilDoi(v):fmt(v)}${suffix}`;const el=document.createElement('div');el.className='glass rounded-2xl border p-4 sm:p-5 overflow-hidden';el.innerHTML=`<div class="grid min-w-0 gap-4 xl:grid-cols-[250px_minmax(0,1fr)]"><div class="min-w-0 w-full">${metricCard(k.meta.label,d.latest,d.delta,k.meta.unit,k.meta.good)}<div class="theme-surface mt-3 w-full rounded-xl border p-3"><div class="mb-2 text-[10px] font-black tracking-widest text-blue-400">${periodLabel} SUMMARY</div><div class="grid grid-cols-1 gap-2"><div class="flex items-center justify-between gap-3 rounded-lg border border-slate-800 px-3 py-3"><div class="text-[9px] font-black uppercase tracking-wider text-slate-600">Average</div><div class="text-sm font-black">${stat(avg)}</div></div><div class="flex items-center justify-between gap-3 rounded-lg border border-slate-800 px-3 py-3"><div class="text-[9px] font-black uppercase tracking-wider text-slate-600">High</div><div class="text-sm font-black">${stat(high)}</div></div><div class="flex items-center justify-between gap-3 rounded-lg border border-slate-800 px-3 py-3"><div class="text-[9px] font-black uppercase tracking-wider text-slate-600">Low</div><div class="text-sm font-black">${stat(low)}</div></div></div></div></div><div class="kpi-visual min-w-0"><div class="mb-1 flex items-center justify-between"><div class="text-[10px] font-black tracking-widest text-blue-400">${periodLabel} · ACTUAL + TREND</div>${period==='ytd'&&d.labels?.length?`<span class="kpi-latest-date"><span>Latest data</span><strong>${escapeHtml(longPeriodLabel(d.labels[d.labels.length-1]))}</strong></span>`:`<span class="text-[10px] text-slate-500">Weekly view</span>`}</div><div class="kpi-chart-legend" data-kpi-chart-legend aria-label="KPI chart legend"></div><div class="kpi-chart min-w-0 w-full"><canvas id="c_${id}"></canvas></div></div></div>`;grid.appendChild(el);makeCombinedChart(el.querySelector(`#c_${id}`),d.labels,d.values,d.trend,k.meta.unit,k.meta.good)});animateCards(grid);}
function renderArea(d){const s=d.summary;$('#areaCards').innerHTML=[['Class A',s.class_rates.A],['Class B',s.class_rates.B],['Class C',s.class_rates.C],['Overall Avg',s.overall_avg]].map(x=>miniCard(x[0],fmt(x[1]),'%')).join('');$('#areaRanking').innerHTML=d.ranking.map((r,i)=>`<button type="button" data-area-detail="${escapeHtml(r.area)}" aria-pressed="${d.selected===r.area}" class="theme-row w-full text-left grid grid-cols-[32px_1fr_auto] items-center gap-2 rounded-xl border px-3 py-2 hover:border-blue-400 focus-visible:outline focus-visible:outline-blue-400 ${d.selected===r.area?'border-blue-400':''}"><div class="text-xs font-black ${i<3?'text-red-500':'text-slate-600'}">${i+1}</div><div><div class="text-xs font-bold">${escapeHtml(r.area)}</div><div class="mt-1 text-[10px] text-slate-500">A ${fmt(r.class_a)}% • B ${fmt(r.class_b)}% • C ${fmt(r.class_c)}%</div></div><div class="text-sm font-black ${i<3?'text-red-500':''}">${fmt(r.overall_avg)}% <span aria-hidden="true">›</span></div></button>`).join('');$('#branchClassAScope').textContent=d.selected==='Overall'?'All branches · sorted highest to lowest':`${d.selected} branches · sorted highest to lowest`;$('#branchClassARanking').innerHTML=(d.branch_class_a||[]).map((r,i)=>`<div class="theme-row grid grid-cols-[32px_1fr_auto] items-center gap-2 rounded-xl border px-3 py-2"><div class="text-xs font-black ${i<5?'text-red-500':'text-slate-600'}">${i+1}</div><div><div class="text-xs font-bold">${escapeHtml(r.branch)}</div><div class="mt-1 text-[10px] text-slate-500">${escapeHtml(r.area||'—')} • ${r.class_a_counts.stockout}/${r.class_a_counts.status_count} Class A status</div></div><div class="text-sm font-black ${i<5?'text-red-500':'text-blue-400'}">${fmt(r.class_a)}%</div></div>`).join('')||'<div class="py-8 text-center text-xs text-slate-600">No branch Class A data for this Area.</div>';animateCards($('#areaCards'));animateRefresh($('#areaRanking'));animateRefresh($('#branchClassARanking'));renderAreaDrilldown(d);$$('[data-area-detail]').forEach(b=>b.onclick=()=>selectArea(b.dataset.areaDetail,true))}
function adsText(n){return n===null||n===undefined||!Number.isFinite(Number(n))?'—':Number(n).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})}
function salesText(n){return n===null||n===undefined?'—':fmt(n)}
async function selectArea(area,scroll=false){$('#areaFilter').value=area;try{const d=await fetchJson('/api/area?area='+encodeURIComponent(area),{},'area');if(d){renderArea(d);if(scroll)$('#areaDrilldown').scrollIntoView({behavior:'smooth',block:'start'})}}catch(err){toast(err.message||'Unable to load Area data.',false)}}
function renderAreaDrilldown(d){
  state.areaDetail=d;const box=$('#areaDrilldown'),open=d.selected!=='Overall';box.classList.toggle('hidden',!open);if(!open)return;
  $('#areaDetailTitle').textContent=d.selected+' · Branch & Model Stock-out Detail';
  $('#areaDetailBranch').innerHTML='<option value="">All Branches</option>'+optionHtml((d.branch_class_a||[]).map(r=>r.branch).sort());
  $('#areaDetailClass').value='';$('#areaDetailStatus').value='';$('#areaDetailSearch').value='';
  $('#areaBranchRows').innerHTML=(d.branch_class_a||[]).map(r=>`<tr class="theme-row border-b border-slate-800"><td class="p-2"><button type="button" class="text-blue-400 font-bold hover:underline" data-detail-branch="${escapeHtml(r.branch)}">${escapeHtml(r.branch)}</button></td>${['A','B','C'].map(c=>`<td>${r.counts?.[c]?.status_count?fmt(r.class_rates[c])+'% ('+r.counts[c].stockout+'/'+r.counts[c].status_count+')':'—'}</td>`).join('')}<td>${fmt(r.overall_avg)}%</td></tr>`).join('');
  $$('[data-detail-branch]').forEach(b=>b.onclick=()=>{$('#areaDetailBranch').value=b.dataset.detailBranch;renderAreaModelDetails()});renderAreaModelDetails();
}
function renderAreaModelDetails(){
 const d=state.areaDetail;if(!d)return;const branch=$('#areaDetailBranch').value,cls=$('#areaDetailClass').value,status=$('#areaDetailStatus').value,q=$('#areaDetailSearch').value.trim().toLowerCase();
 const rows=(d.model_details||[]).filter(r=>(!branch||r.branch===branch)&&(!cls||r.class===cls)&&(!status||(status==='stockout'?r.is_stockout:status==='missing'?!r.stock_status:!!r.stock_status&&!r.is_stockout))&&(!q||`${r.model} ${r.brand}`.toLowerCase().includes(q)));
 const valid=rows.filter(r=>r.stock_status).length,outs=rows.filter(r=>r.is_stockout).length;
 $('#areaDetailCount').textContent=`${rows.length} model records · ${outs} stock-outs / ${valid} recorded statuses`;
 $('#areaDetailRows').innerHTML=rows.map(r=>`<tr class="theme-row border-b border-slate-800"><td class="p-2 font-bold">${escapeHtml(r.branch)}</td><td>${classPill(r.class)}</td><td>${r.rank===999999?'—':fmt(r.rank)}</td><td>${escapeHtml(r.brand)}</td><td class="font-bold">${escapeHtml(r.model)}</td><td><span class="status-pill ${statusClass(r.stock_status)}">${escapeHtml(r.stock_status||'Not recorded')}</span></td><td>${fmt(r.inventory)}</td><td>${ceilDoi(r.doi)}</td><td>${salesText(r.sales_3m)}</td></tr>`).join('')||'<tr><td colspan="9" class="p-5 text-center text-slate-500">No models match the selected filters.</td></tr>';
}
function renderClassModels(){const d=state.branchData;if(!d)return;const cls=state.classFilter,models=d.classes[cls]||[];$$('.classTab').forEach(b=>{const active=b.dataset.class===cls;b.classList.toggle('class-tab-active',active);b.classList.toggle('text-slate-400',!active)});$('#classSummary').textContent=`Class ${cls} • ${models.length} models • ranked by branch classification`;const head=`<div class="model-table-head"><div>Rank</div><div>Model</div><div>Brand</div><div>Stock Status</div><div>Inventory</div><div>Suggested</div><div>DoI</div></div>`;const rows=models.map(m=>`<div class="model-row"><div class="rank-cell"><span class="rank-chip">#${fmt(m.rank)}</span></div><div class="model-main"><div class="model-label-mobile">Model</div><div class="model-value">${escapeHtml(m.model||'—')}</div></div><div><div class="model-label-mobile">Brand</div><div class="text-xs font-semibold text-slate-500">${escapeHtml(m.brand||'—')}</div></div><div class="model-status"><div class="model-label-mobile">Stock Status</div><span class="status-pill ${statusClass(m.stock_status)}">${escapeHtml(m.stock_status||'—')}</span></div><div><div class="model-label-mobile">Inventory</div><div class="model-value">${fmt(m.inventory)}</div></div><div><div class="model-label-mobile">Suggested Transfer</div><div class="model-value">${fmt(m.suggested_transfer)}</div></div><div><div class="model-label-mobile">DoI</div><div class="model-value">${ceilDoi(m.doi)}</div></div></div>`).join('');$('#modelTable').innerHTML=head+(rows||'<div class="py-12 text-center text-xs text-slate-600">No models in this class.</div>')}
function renderBranch(d){state.branchData=d;$('#branchArea').textContent=d.area?`${d.area} • ${d.branch||''}`:'Area not assigned';const s=d.summary;$('#branchCards').innerHTML=[['Class A',s.class_rates.A],['Class B',s.class_rates.B],['Class C',s.class_rates.C],['Overall Avg',s.overall_avg]].map(x=>miniCard(x[0],fmt(x[1]),'%')).join('');renderClassModels();animateCards($('#branchCards'));animateRefresh($('#modelTable'))}
function optionHtml(values){return (values||[]).map(v=>{const safe=escapeHtml(v);return `<option value="${safe}">${safe}</option>`}).join('')}
function animateRefresh(el){if(!el)return;el.classList.remove('data-refresh');void el.offsetWidth;el.classList.add('data-refresh');setTimeout(()=>el.classList.remove('data-refresh'),380)}
function animateCards(container){if(!container)return;[...container.children].forEach((el,i)=>{el.classList.remove('card-cascade');el.style.animationDelay=`${Math.min(i,8)*35}ms`;void el.offsetWidth;el.classList.add('card-cascade')})}
function statusMixHtml(statuses){const entries=Object.entries(statuses||{});if(!entries.length)return '<span class="text-[10px] text-slate-600">No status data</span>';return entries.map(([name,count])=>`<span class="status-pill ${statusClass(name)}">${escapeHtml(name)} ${fmt(count)}</span>`).join(' ')}
function populateStatusFilters(filters){
  const map=[['#statusAreaFilter',filters.areas],['#statusBranchFilter',filters.branches],['#statusBrandFilter',filters.brands],['#statusModelFilter',filters.models],['#statusClassFilter',filters.classes],['#statusStockFilter',filters.statuses]];
  map.forEach(([sel,vals])=>{$(sel).innerHTML=optionHtml(vals)});
}
function statusGraphColor(status){const s=String(status||'').toLowerCase();if(s.includes('stockout')||s.includes('stock out'))return '#ef4444';if(s.includes('re-order')||s.includes('re order')||s.includes('reorder'))return '#f59e0b';if(s.includes('over'))return '#22c55e';if(s.includes('critical')||s.includes('low'))return '#f97316';return state.theme==='light'?'#64748b':'#94a3b8'}
function renderReorderModelCard(d){
  state.reorderCard=d||{rows:[],summary:{counts:{},avg_doi:{}}};
  const rows=state.reorderCard.rows||[],sum=state.reorderCard.summary||{};
  const src=state.reorderCard.source_sheet||'Reorder';
  $('#reorderCardSource').textContent=rows.length?`Latest imported ${src} source • Rank is imported directly and DoI follows the approved round-up rule.`:'No Reorder model data found in the latest imported workbook.';
  const brands=['All Brands',...new Set(rows.map(r=>r.brand||'Unspecified').filter(Boolean))].sort((a,b)=>a==='All Brands'?-1:b==='All Brands'?1:a.localeCompare(b));
  const brandSel=$('#reorderBrandFilter');if(brandSel){const current=state.reorderBrandFilter||'All Brands';brandSel.innerHTML=brands.map(b=>{const safe=escapeHtml(b);return `<option value="${safe}">${safe}</option>`}).join('');brandSel.value=brands.includes(current)?current:'All Brands';state.reorderBrandFilter=brandSel.value;}
  const selected=state.reorderClassFilter||'All',selectedBrand=state.reorderBrandFilter||'All Brands';
  const brandRows=rows.filter(r=>selectedBrand==='All Brands'||(r.brand||'Unspecified')===selectedBrand);
  const scopedCounts={A:0,B:0,C:0},scopedAvg={A:0,B:0,C:0};['A','B','C'].forEach(cls=>{const clsRows=brandRows.filter(r=>r.class===cls);scopedCounts[cls]=clsRows.length;scopedAvg[cls]=clsRows.length?Math.ceil(clsRows.reduce((a,r)=>a+Number(r.doi||0),0)/clsRows.length):0;});
  $('#reorderSummaryCards').innerHTML=[['Class A',scopedCounts.A,`Avg DoI ${ceilDoi(scopedAvg.A)} d`],['Class B',scopedCounts.B,`Avg DoI ${ceilDoi(scopedAvg.B)} d`],['Class C',scopedCounts.C,`Avg DoI ${ceilDoi(scopedAvg.C)} d`],['ABC Models',brandRows.length,selectedBrand]].map(x=>`<div class="theme-surface rounded-xl border p-3"><div class="text-[10px] font-black uppercase tracking-wider text-slate-500">${x[0]}</div><div class="mt-1 text-xl font-black">${fmt(x[1])}</div><div class="mt-1 text-[9px] font-bold text-slate-600">${escapeHtml(x[2])}</div></div>`).join('');
  const visible=brandRows.filter(r=>selected==='All'||r.class===selected);
  $('#reorderModelRows').innerHTML=visible.map(r=>`<tr class="theme-row border-t border-slate-800/70"><td class="px-3 py-3">${classPill(r.class)}</td><td class="px-3 py-3 font-black">${Number(r.rank)>=999999?'—':fmt(r.rank)}</td><td class="px-3 py-3 font-bold text-slate-500">${escapeHtml(r.brand||'—')}</td><td class="px-3 py-3 font-black">${escapeHtml(r.model||'—')}</td><td class="px-3 py-3 text-right font-black">${ceilDoi(r.doi)} d</td><td class="px-3 py-3"><span class="status-pill ${statusClass(r.stock_status)}">${escapeHtml(r.stock_status||'—')}</span></td></tr>`).join('')||'<tr><td colspan="6" class="py-10 text-center text-xs text-slate-600">No models for the selected Class / Brand filter.</td></tr>';
  $$('.reorderClassFilter').forEach(b=>{const active=b.dataset.reorderClass===selected;b.classList.toggle('border-blue-400',active);b.classList.toggle('text-blue-400',active);b.classList.toggle('border-slate-700',!active);b.classList.toggle('text-slate-400',!active)});
  animateCards($('#reorderSummaryCards'));
}

function renderStatusChart(d){
  const canvas=$('#statusSummaryChart');if(!canvas||!d)return;if(state.statusChart){state.statusChart.destroy();state.statusChart=null}
  const rows=d.area_status||[];const wrap=canvas.parentElement;let empty=wrap?.querySelector('.status-empty-message');if(!rows.length){canvas.style.display='none';if(!empty&&wrap){empty=document.createElement('div');empty.className='status-empty-message flex h-full min-h-[240px] items-center justify-center text-center text-xs text-slate-500';empty.innerHTML='<div><div class="font-black text-slate-300">No status data loaded</div><div class="mt-2">Import a new SCM workbook to populate this graph.</div></div>';wrap.appendChild(empty)}$('#statusChartScope').textContent='No status data loaded.';return}canvas.style.display='block';if(empty)empty.remove();
  const statuses=[...new Set(rows.flatMap(r=>Object.keys(r.statuses||{})))];
  const p=palette();
  const areaScope=$('#statusAreaFilter')?.value||'Overall',brandScope=$('#statusBrandFilter')?.value||'All Brands',classScope=$('#statusClassFilter')?.value||'All Classes',modelScope=$('#statusModelFilter')?.value||'All Models',statusScope=$('#statusStockFilter')?.value||'All Statuses';
  const scopeBits=[areaScope!=='Overall'?areaScope:null,brandScope!=='All Brands'?brandScope:null,classScope!=='All Classes'?`Class ${classScope}`:null,modelScope!=='All Models'?modelScope:null,statusScope!=='All Statuses'?statusScope:null].filter(Boolean);
  $('#statusChartScope').textContent=rows.length?`By Area • ${scopeBits.length?scopeBits.join(' • '):'all Brand / Model statuses'} • inherits the summary filters above`:'No status data matches the selected Brand & Model Summary filters.';
  state.statusChart=new Chart(canvas,{type:'bar',data:{labels:rows.map(r=>r.area),datasets:statuses.map(st=>({label:st,data:rows.map(r=>Number(r.statuses?.[st]||0)),backgroundColor:statusGraphColor(st),borderWidth:0,borderRadius:5,barThickness:20}))},options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,animation:{duration:340,easing:'easeOutQuart'},interaction:{mode:'index',intersect:false},plugins:{legend:{display:true,position:'top',align:'end',labels:{color:p.axis,boxWidth:10,boxHeight:10,font:{size:9,weight:'bold'}}},tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmt(c.parsed.x)}`}}},scales:{x:{stacked:true,beginAtZero:true,grid:{color:p.grid},ticks:{color:p.axis,precision:0,font:{size:9}}},y:{stacked:true,grid:{display:false},ticks:{color:p.axis,font:{size:9},autoSkip:false}}}}});
}
function classPill(cls){const c=String(cls||'—').toUpperCase();const tone=c==='A'?'border-blue-400/40 bg-blue-500/10 text-blue-400':c==='B'?'border-blue-400/35 bg-blue-400/10 text-blue-300':c==='C'?'border-blue-300/30 bg-blue-300/10 text-blue-200':'border-slate-700 text-slate-500';return `<span class="inline-flex rounded-full border px-2 py-1 text-[9px] font-black ${tone}">Class ${escapeHtml(c)}</span>`}
function renderStatusSummary(d){
  state.statusData=d;const s=d.summary||{};
  $('#statusCards').innerHTML=[['Areas',s.areas],['Brands',s.brands],['Models',s.models],['Branches',s.branches],['Inventory',s.inventory],['Stock-Out Rate',s.stockout_rate,'%']].map(x=>miniCard(x[0],fmt(x[1]),x[2]||'')).join('');
  animateCards($('#statusCards'));animateRefresh(document.querySelector('.status-chart-panel'));
  const am=d.area_models||[];
  $('#modelStatusScope').textContent=`${am.length} Area × Brand × Model × Class combinations • use the filters above to isolate Area, Class, Brand, Model, Branch, or Stock Status`;
  $('#modelStatusSummary').innerHTML=am.map(r=>`<tr><td class="font-black">${escapeHtml(r.area||'—')}</td><td>${classPill(r.class)}</td><td class="font-bold">${escapeHtml(r.brand||'—')}</td><td class="font-black">${escapeHtml(r.model||'—')}</td><td>${fmt(r.branches)}</td><td class="font-black">${fmt(r.inventory)}</td><td>${fmt(r.suggested_transfer)}</td><td>${ceilDoi(r.avg_doi)} d</td><td><div class="font-black ${Number(r.stockout_rate)>0?'text-red-500':'text-emerald-500'}">${fmt(r.stockout_rate)}%</div><div class="text-[9px] text-slate-600">${fmt(r.stockout_count)} / ${fmt(r.records)}</div></td><td><div class="flex max-w-[380px] flex-wrap gap-1">${statusMixHtml(r.statuses)}</div></td></tr>`).join('')||'<tr><td colspan="10" class="py-12 text-center text-xs text-slate-600">No Area / Model status data matches the selected filters.</td></tr>';
  renderStatusChart(d);animateRefresh($('#modelStatusSummary'));
}
async function loadStatusSummary(){
  const qs=new URLSearchParams({area:$('#statusAreaFilter').value||'Overall',branch:$('#statusBranchFilter').value||'All Branches',brand:$('#statusBrandFilter').value||'All Brands',model:$('#statusModelFilter').value||'All Models',class:$('#statusClassFilter').value||'All Classes',status:$('#statusStockFilter').value||'All Statuses'});
  try{const d=await fetchJson('/api/status-summary?'+qs.toString(),{},'status-summary');if(d)renderStatusSummary(d)}catch(err){toast(err.message||'Unable to load status summary.',false)}
}
function datalistOptions(values){return (values||[]).map(v=>`<option value="${escapeHtml(v)}"></option>`).join('')}
// One editable combobox: type to narrow choices, or open the full list with the arrow.
const modelCombos=new Map();
function initModelCombo(id,listId,getOptions,onType,onChoose){
 const input=$('#'+id),list=$('#'+listId),wrap=input.closest('.model-combo'),toggle=wrap.querySelector('button');let shown=[],active=-1,showAll=false;
 function close(){list.hidden=true;input.setAttribute('aria-expanded','false');input.removeAttribute('aria-activedescendant');active=-1}
 function draw(){if(list.hidden)return;const q=input.value.trim().toLowerCase();shown=getOptions().filter(o=>showAll||o.value.toLowerCase().includes(q)||o.label.toLowerCase().includes(q));active=-1;input.removeAttribute('aria-activedescendant');list.replaceChildren();if(!shown.length){const empty=document.createElement('div');empty.className='model-combo-empty';empty.textContent='No matching models';list.append(empty)}shown.forEach((o,i)=>{const el=document.createElement('div');el.id=listId+'-option-'+i;el.setAttribute('role','option');el.setAttribute('aria-selected','false');el.className='model-combo-option';el.textContent=o.label;el.addEventListener('mousedown',e=>e.preventDefault());el.addEventListener('click',()=>choose(i));list.append(el)})}
 function open(all=false){showAll=all;list.hidden=false;input.setAttribute('aria-expanded','true');draw()}
 function choose(i){const o=shown[i];if(!o)return;input.value=o.value;close();onChoose(o);input.focus();close()}
 function move(i){if(!shown.length)return;active=(i+shown.length)%shown.length;[...list.children].forEach((el,j)=>{el.setAttribute('aria-selected',String(j===active))});const el=list.children[active];input.setAttribute('aria-activedescendant',el.id);el.scrollIntoView({block:'nearest'})}
 input.addEventListener('input',()=>{open(false);onType()});input.addEventListener('focus',()=>open(true));toggle.addEventListener('mousedown',e=>e.preventDefault());toggle.addEventListener('click',()=>{if(!list.hidden){close();return}input.focus();open(true)});
 input.addEventListener('keydown',e=>{if(e.key==='Escape'||e.key==='Tab'){close();return}if(e.key==='ArrowDown'||e.key==='ArrowUp'){e.preventDefault();if(list.hidden)open(true);move(e.key==='ArrowDown'?active+1:active<0?shown.length-1:active-1)}else if(e.key==='Enter'&&!list.hidden){e.preventDefault();if(active>=0)choose(active);else if(shown.length===1)choose(0);else close()}});input.addEventListener('blur',()=>{close()});document.addEventListener('pointerdown',e=>{if(!wrap.contains(e.target))close()});
 modelCombos.set(id,{refresh:draw,close});
}
function refreshSimDropdown(){modelCombos.get('simModel')?.refresh()}
async function loadModel(){state.model=null;calcNewDoi();const branch=$('#simBranch').value,brand=$('#simBrand')?.value||'All Brands',model=$('#simModel').value.trim();if(!branch||!model){state.model=null;$('#modelSnapshot').innerHTML='';calcNewDoi();return}const exact=(state.requestBuilderModels||[]).find(x=>String(x).toLowerCase()===model.toLowerCase());if(!exact){state.model=null;$('#modelSnapshot').innerHTML='<div class="col-span-2 rounded-xl border border-slate-800 px-3 py-4 text-xs text-slate-500">Type a model or choose one from the dropdown suggestions.</div>';calcNewDoi();return}if(exact!==model)$('#simModel').value=exact;try{const d=await fetchJson(`/api/model?branch=${encodeURIComponent(branch)}&brand=${encodeURIComponent(brand)}&model=${encodeURIComponent(exact)}`,{},'sim-model');if(!d||$('#simBranch').value!==branch||($('#simBrand')?.value||'All Brands')!==brand||$('#simModel').value!==exact)return;state.model=d;renderModel()}catch(err){toast(err.message||'Unable to load model.',false)}}
function renderModel(){const m=state.model;if(!m)return;$('#modelSnapshot').innerHTML=miniCard('Inventory',fmt(m.inventory))+miniCard('3 Months Sales (Qty)',salesText(m.sales_3m))+miniCard('Avg. Daily Sale (Qty)',adsText(m.avg_daily_sale))+miniCard('Class',m.class||'—')+miniCard('Stock Status',m.stock_status||'—')+miniCard('Suggested Transfer',fmt(m.suggested_transfer))+miniCard('Current DoI',ceilDoi(m.doi),' d');calcNewDoi()}
function calcNewDoi(){const m=state.model,q=Number($('#requestQty').value||0);if(!m){$('#newDoi').textContent='—';return}$('#newDoi').textContent=m.avg_daily_sale>0?`${ceilDoi((m.inventory+q)/m.avg_daily_sale)} days`:'N/A (no ADS)'}
async function populateModels(resetBrand=false){state.model=null;state.requestBuilderModels=[];refreshSimDropdown();$('#modelSnapshot').innerHTML='';calcNewDoi();const branch=$('#simBranch').value,brandEl=$('#simBrand'),modelEl=$('#simModel'),previousModel=String(modelEl?.value||'').trim();let brand=resetBrand?'All Brands':String(brandEl?.value||'All Brands');try{const d=await fetchJson(`/api/branch-models?branch=${encodeURIComponent(branch)}&brand=${encodeURIComponent(brand)}`,{},'sim-models');if(!d)return;const brands=['All Brands',...(d.brands||[])];if(brandEl){brandEl.innerHTML=optionHtml(brands);if(!brands.includes(brand))brand='All Brands';brandEl.value=brand}state.requestBuilderModels=d.models||[];const exactPrevious=state.requestBuilderModels.find(x=>String(x).toLowerCase()===previousModel.toLowerCase());if(modelEl)modelEl.value=exactPrevious||'';refreshSimDropdown();state.model=null;if(exactPrevious)await loadModel();else{$('#modelSnapshot').innerHTML='';calcNewDoi()}}catch(err){toast(err.message||'Unable to load branch models.',false)}}
function renderRequest(){
  const tb=$('#requestTable'),empty=$('#requestEmpty'),selector=$('#requestOutputBranch'),hint=$('#requestBranchHint');
  empty.classList.toggle('hidden',state.requests.length>0);
  const branches=[...new Set(state.requests.map(x=>String(x.branch||'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  if(selector){
    const previous=String(selector.value||state.requestOutputBranch||'');
    let desired='';
    if(branches.length===1)desired=branches[0];
    else if(branches.length>1&&state.requestBranchCount>1&&branches.includes(previous))desired=previous;
    selector.innerHTML=branches.length?`${branches.length>1?'<option value="">Select branch to output…</option>':''}${branches.map(branch=>`<option value="${escapeHtml(branch)}">${escapeHtml(branch)} (${state.requests.filter(x=>x.branch===branch).length} line${state.requests.filter(x=>x.branch===branch).length===1?'':'s'})</option>`).join('')}`:'<option value="">Add request lines first</option>';
    selector.value=desired;selector.disabled=!branches.length;state.requestOutputBranch=desired;
  }
  if(hint)hint.classList.toggle('hidden',branches.length<=1);
  state.requestBranchCount=branches.length;
  tb.innerHTML=state.requests.map((x,i)=>`<tr class="border-b border-slate-900"><td class="px-2 py-3 font-black">${escapeHtml(x.branch||'—')}</td><td class="font-bold">${escapeHtml(x.model)}</td><td>${escapeHtml(x.class)}</td><td>${salesText(x.sales_3m)}</td><td>${adsText(x.avg_daily_sale)}</td><td>${fmt(x.inventory)}</td><td>${fmt(x.requested_qty)}</td><td class="font-bold ${statusClass(x.stock_status)}">${escapeHtml(x.stock_status)}</td><td>${ceilDoi(x.doi)}</td><td class="font-bold text-blue-400">${typeof x.new_doi==='number'?ceilDoi(x.new_doi):escapeHtml(x.new_doi)}</td><td class="max-w-[220px] truncate pr-2">${escapeHtml(x.remarks||'—')}</td><td><button class="removeReq text-red-500" data-i="${i}">Remove</button></td></tr>`).join('');
  $$('.removeReq').forEach(b=>b.onclick=()=>{state.requests.splice(Number(b.dataset.i),1);renderRequest()});
}
function selectedRequestOutput(){
  if(!state.requests.length){toast('Add at least one request line.',false);return null}
  const branches=[...new Set(state.requests.map(x=>String(x.branch||'').trim()).filter(Boolean))];
  let branch=branches.length===1?branches[0]:String($('#requestOutputBranch')?.value||'').trim();
  if(!branch){toast('Multiple branches are queued. Select the branch you want to output.',false);$('#requestOutputBranch')?.focus();return null}
  const rows=state.requests.filter(x=>x.branch===branch);
  if(!rows.length){toast('No request lines were found for the selected branch.',false);return null}
  state.requestOutputBranch=branch;
  return {branch,items:rows.map(x=>({model:x.model,brand:x.brand,requested_qty:x.requested_qty,remarks:x.remarks}))};
}
function managementOutputPayload(){return {brand:$('#managementBrandFilter').value||'All Brands',class:$('#managementClassFilter').value||'All Classes',status:$('#managementStatusFilter').value||'All Statuses',model:managementModelValue(),orders:state.managementEdits,keys:(state.management?.renderedRows||[]).map(r=>r.key)}}
async function responseErrorMessage(response,fallback){let data={};try{data=await response.clone().json()}catch(_){ }return data.error||data.detail||fallback}
async function openDirectPrint(url,payload,title='Print Preview'){
  const win=window.open('','_blank');
  if(!win){toast('Print preview was blocked by the browser. Allow pop-ups for this dashboard and try again.',false);return}
  try{
    win.document.open();win.document.write(`<title>${escapeHtml(title)}</title><body style="font-family:Arial,sans-serif;padding:24px;color:#0f172a">Preparing print preview…</body>`);win.document.close();
    const r=await appFetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
    if(!r.ok){const msg=await responseErrorMessage(r,'Unable to prepare print preview.');win.close();toast(msg,false);return}
    const html=await r.text();win.document.open();win.document.write(html);win.document.close();
  }catch(err){try{win.close()}catch(_){ }toast(err?.message||'Unable to prepare print preview.',false)}
}
function syncThemeButton(){const light=state.theme==='light';$('#themeLabel').textContent=light?'Dark Mode':'Light Mode';const icon=$('#themeIcon');if(icon)icon.innerHTML=light?'<svg viewBox="0 0 24 24"><path d="M21 12.8A8.5 8.5 0 1 1 11.2 3 7 7 0 0 0 21 12.8Z"/></svg>':'<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41"/></svg>'}
function setTheme(theme){state.theme=theme;document.documentElement.dataset.theme=theme;localStorage.setItem('scm-theme',theme);syncThemeButton();if(state.boot)renderKPIs(state.boot.kpis);if(state.statusData)renderStatusChart(state.statusData)}
function applySidebarPreference(){
  const collapsed=localStorage.getItem('scm-sidebar-collapsed')==='1';document.body.classList.toggle('sidebar-collapsed',collapsed);document.body.classList.toggle('sidebar-expanded',!collapsed);const b=$('#sidebarToggle');if(b)b.title=collapsed?'Expand sidebar':'Collapse sidebar';
}
function toggleSidebar(){
  if(window.innerWidth<1024){document.body.classList.toggle('sidebar-mobile-open');return}
  const collapsed=!document.body.classList.contains('sidebar-collapsed');localStorage.setItem('scm-sidebar-collapsed',collapsed?'1':'0');applySidebarPreference();setTimeout(()=>window.dispatchEvent(new Event('resize')),230)
}

function initWorkspaceChrome(){
  applySidebarPreference();
  $('#sidebarToggle').onclick=toggleSidebar;$('#mobileSidebarBtn').onclick=()=>document.body.classList.add('sidebar-mobile-open');$('#sidebarBackdrop').onclick=()=>document.body.classList.remove('sidebar-mobile-open');
  window.addEventListener('resize',()=>{if(window.innerWidth>=1024)document.body.classList.remove('sidebar-mobile-open')});
}

function managementModelValue(){return String($('#managementModelFilter')?.value||'').trim()||'All Models'}
function populateManagementFilters(filters,selected){
  const config=[['#managementBrandFilter',filters.brands,selected?.brand],['#managementClassFilter',filters.classes,selected?.class],['#managementStatusFilter',filters.statuses,selected?.status]];
  config.forEach(([id,values,current])=>{const el=$(id);if(!el)return;const keep=el.value||current;el.innerHTML=optionHtml(values||[]);el.value=(values||[]).includes(keep)?keep:(values||[])[0]||''});
  refreshManagementDropdown();
}
function refreshManagementDropdown(){modelCombos.get('managementModelFilter')?.refresh()}
const managementNumericFields=new Set(['unit_cost','inventory','po_balance','quantity']);
function managementEditFor(row){return {unit_cost:Number(row.unit_cost||0),inventory:Number(row.inventory||0),po_balance:Number(row.po_balance||0),quantity:Number(row.allocation||0),remarks:String(row.remarks||'')}}
function managementRowsWithEdits(d){return (d.rows||[]).map(r=>{const edit=state.managementEdits[r.key];const row={...r};if(edit&&typeof edit==='object'){['unit_cost','inventory','po_balance'].forEach(f=>{if(Object.prototype.hasOwnProperty.call(edit,f))row[f]=edit[f]})}const allocation=Math.max(0,Number(edit?.quantity??r.allocation??0)),remarks=String(edit?.remarks??r.remarks??'');row.allocation=allocation;row.remarks=remarks;row.total_amount=Number(row.unit_cost||0)*allocation;const ads=Number(row.avg_daily_sale||0);row.new_doi=ads>0?Math.ceil((Number(row.inventory||0)+Number(row.po_balance||0)+allocation)/ads):0;return row})}
function renderManagementCards(rows){const totalAllocation=rows.reduce((a,r)=>a+Number(r.allocation||0),0),grand=rows.reduce((a,r)=>a+Number(r.total_amount||0),0),inv=rows.reduce((a,r)=>a+Number(r.inventory||0),0),po=rows.reduce((a,r)=>a+Number(r.po_balance||0),0);$('#managementCards').innerHTML=[['Models',rows.length,''],['Inv.',inv,''],['PO Balance',po,''],['Order Quantity',totalAllocation,''],['Grand Total',php(grand),'']].map(x=>miniCard(x[0],typeof x[1]==='string'?x[1]:fmt(x[1]),x[2]||'')).join('');$('#managementTotalAllocation').textContent=fmt(totalAllocation);$('#managementGrandTotal').textContent=php(grand)}
function refreshManagementComputed(key){const row=state.management?.renderedRows?.find(x=>x.key===key);if(!row)return;row.total_amount=Number(row.unit_cost||0)*Number(row.allocation||0);row.new_doi=Number(row.avg_daily_sale||0)>0?Math.ceil((Number(row.inventory||0)+Number(row.po_balance||0)+Number(row.allocation||0))/Number(row.avg_daily_sale)):0;const enc=encodeURIComponent(key);$$(`[data-management-total="${enc}"]`).forEach(el=>el.textContent=php(row.total_amount));$$(`[data-management-doi="${enc}"]`).forEach(el=>el.textContent=fmt(row.new_doi));renderManagementCards(state.management.renderedRows)}
function updateManagementField(key,field,value){const row=state.management?.renderedRows?.find(x=>x.key===key);if(!row)return;let v=value;if(managementNumericFields.has(field))v=Math.max(0,Number(value||0));if(field==='quantity')row.allocation=v;else row[field]=v;state.managementEdits[key]=managementEditFor(row);refreshManagementComputed(key)}
function bindManagementInputs(){if(!isAdmin)return;$$('.managementFieldInput').forEach(inp=>{const handler=()=>{const key=decodeURIComponent(inp.dataset.managementKey||''),field=inp.dataset.managementField;updateManagementField(key,field,inp.value)};inp.oninput=handler;inp.onchange=handler})}
function managementClassSelect(key,value){return `<select data-management-key="${key}" data-management-field="class" class="managementFieldInput management-edit font-black"><option value="A" ${value==='A'?'selected':''}>Class A</option><option value="B" ${value==='B'?'selected':''}>Class B</option><option value="C" ${value==='C'?'selected':''}>Class C</option><option value="" ${!value?'selected':''}>Unclassified</option></select>`}
function managementInput(key,field,value,type='text',extra=''){return `<input data-management-key="${key}" data-management-field="${field}" type="${type}" value="${escapeHtml(value)}" class="managementFieldInput management-edit ${type==='number'?'num':''}" ${type==='number'?'min="0"':''} ${extra}>`}
function renderManagement(d){
  state.management=d;populateManagementFilters(d.filters||{},d.selected||{});$('#managementSourceTitle').textContent=d.title||'Management Order Planning';
  const rows=managementRowsWithEdits(d);state.management.renderedRows=rows;const body=$('#managementTableBody'),mobile=$('#managementMobileCards');$('#managementEmpty').classList.toggle('hidden',rows.length>0);
  const locked=(label,value,align='left')=>`<div class="management-readonly ${align==='right'?'justify-end text-right':''}"><span>${value??'—'}</span></div>`;
  body.innerHTML=rows.map((r,idx)=>{const key=encodeURIComponent(r.key),qty=Number(r.allocation||0);return `<tr class="theme-row align-top"><td class="ref-cell px-1 py-2 text-center"><span class="management-row-number">${idx+1}</span></td><td class="ref-cell px-2 py-2"><div class="management-identity"><div class="flex items-center gap-2"><span class="management-brand">${escapeHtml(r.brand||'—')}</span></div><div class="management-model mt-1">${escapeHtml(r.model||'—')}</div></div></td><td class="ref-cell px-1 py-2 text-center">${classPill(r.class)}</td><td class="ref-cell px-1 py-2 text-center font-bold">${salesText(r.sales_3m)}</td><td class="ref-cell px-1 py-2 text-center font-bold">${adsText(r.avg_daily_sale)}</td><td class="ref-cell px-1 py-2">${isAdmin?managementInput(key,'unit_cost',Number(r.unit_cost||0),'number','step="0.01"'):`<div class="text-right font-semibold">${php(r.unit_cost)}</div>`}</td><td class="ref-cell px-1 py-2">${isAdmin?managementInput(key,'inventory',Number(r.inventory||0),'number','step="1"'):`<div class="text-center font-bold">${fmt(r.inventory)}</div>`}</td><td class="ref-cell px-1 py-2">${locked('',ceilDoi(r.doi),'right')}</td><td class="management-status ref-cell px-1 py-2"><div class="management-readonly justify-center"><span class="status-pill ${statusClass(r.stock_status)}">${escapeHtml(r.stock_status||'—')}</span></div></td><td class="ref-cell px-1 py-2">${isAdmin?managementInput(key,'po_balance',Number(r.po_balance||0),'number','step="1"'):`<div class="text-center font-bold">${fmt(r.po_balance)}</div>`}</td><td class="input-cell px-1 py-2">${isAdmin?managementInput(key,'quantity',qty,'number','step="1"').replace('management-edit','management-edit management-order-input'):`<div class="text-center font-black">${fmt(qty)}</div>`}</td><td class="input-cell px-2 py-2">${isAdmin?managementInput(key,'remarks',r.remarks||'','text','placeholder="Management note"'):`<span class="break-words text-slate-500">${escapeHtml(r.remarks||'—')}</span>`}</td><td class="result-cell px-1 py-2 text-center"><div class="management-result management-computed" data-management-doi="${key}">${ceilDoi(r.new_doi)}</div></td><td class="result-cell px-1 py-2 text-right"><div class="management-result management-computed" data-management-total="${key}">${php(r.total_amount)}</div></td></tr>`}).join('');
  mobile.innerHTML=rows.map(r=>{const key=encodeURIComponent(r.key),qty=Number(r.allocation||0);const editField=(label,field,value,type='text',extra='')=>`<label class="text-[9px] font-black uppercase tracking-wider text-slate-500">${label}${isAdmin?managementInput(key,field,value,type,extra):`<div class="mt-1 text-xs font-black text-[var(--text)]">${field==='unit_cost'?php(value):fmt(value)}</div>`}</label>`;return `<article class="theme-row rounded-xl border p-4"><div class="flex flex-wrap items-start justify-between gap-2"><div><div class="flex items-center gap-2">${classPill(r.class)}<span class="management-brand">${escapeHtml(r.brand||'—')}</span></div><div class="management-model mt-1">${escapeHtml(r.model||'—')}</div><div class="mt-1 text-[10px] text-slate-500">3Mo Sales Qty Total: <strong>${salesText(r.sales_3m)}</strong></div><div class="mt-1 text-[10px] text-slate-500">Avg. Daily Sale (Qty): <strong>${adsText(r.avg_daily_sale)}</strong></div></div><span class="status-pill ${statusClass(r.stock_status)}">${escapeHtml(r.stock_status||'—')}</span></div><div class="mt-3 grid grid-cols-2 gap-2 sm:grid-cols-3">${editField('Unit Cost','unit_cost',Number(r.unit_cost||0),'number','step="0.01"')}${editField('Inv.','inventory',Number(r.inventory||0),'number','step="1"')}<div><div class="text-[9px] font-black uppercase tracking-wider text-slate-500">DoI</div>${locked('',ceilDoi(r.doi),'right')}</div>${editField('PO Balance','po_balance',Number(r.po_balance||0),'number','step="1"')}<label class="text-[9px] font-black uppercase tracking-wider text-slate-500">Order Quantity${isAdmin?managementInput(key,'quantity',qty,'number','step="1"').replace('management-edit','management-edit management-order-input'):`<div class="mt-1 text-xs font-black">${fmt(qty)}</div>`}</label><div><div class="text-[9px] font-black uppercase tracking-wider text-slate-500">New DoI</div><div class="management-result mt-1 text-right management-computed" data-management-doi="${key}">${ceilDoi(r.new_doi)}</div></div></div><div class="mt-3 grid gap-2 sm:grid-cols-[1fr_180px]"><label class="text-[9px] font-black uppercase tracking-wider text-slate-500">Remarks${isAdmin?managementInput(key,'remarks',r.remarks||'','text','placeholder="Management note"'):`<div class="mt-1 text-xs text-slate-400">${escapeHtml(r.remarks||'—')}</div>`}</label><div><div class="text-[9px] font-black uppercase tracking-wider text-slate-500">Total Amount</div><div class="management-result mt-1 text-right management-computed" data-management-total="${key}">${php(r.total_amount)}</div></div></div></article>`}).join('');
  renderManagementCards(rows);bindManagementInputs();
}
async function loadManagement(){
  const q=new URLSearchParams({brand:$('#managementBrandFilter')?.value||'All Brands',class:$('#managementClassFilter')?.value||'All Classes',status:$('#managementStatusFilter')?.value||'All Statuses',model:managementModelValue()});try{const d=await fetchJson('/api/management?'+q.toString(),{},'management');if(d)renderManagement(d)}catch(err){toast(err.message||'Unable to load Management data.',false)}
}
async function init(){
  document.body.classList.add('app-busy');
  try{
    if(!(await verifyAdminSession()))return;
    const d=await fetchJson('/api/bootstrap',{cache:'no-store'},'bootstrap');if(!d)throw new Error('Dashboard bootstrap was cancelled.');state.boot=d;
    $('#lastRefresh').textContent=`${d.data_source||'Local Data'} • Refreshed ${new Date(d.generated_at).toLocaleString()}`;
    const dataChip=$('#workspaceDataChip');if(dataChip){const ready=!!d.has_data;dataChip.classList.toggle('status-ready',ready);dataChip.classList.toggle('status-empty',!ready);dataChip.innerHTML=`<span class="workspace-chip-dot"></span><strong>${escapeHtml(ready?(d.data_source||'Data Ready'):'No Data Loaded')}</strong>`;}
    $('#areaFilter').innerHTML=optionHtml(d.areas);
    $('#branchFilter').innerHTML=optionHtml(d.branches);
    $('#simBranch').innerHTML=optionHtml(d.branches);
    populateStatusFilters(d.status_summary.filters);
    renderKPIs(d.kpis);renderReorderModelCard(d.reorder_card);renderArea(d.area);renderBranch(d.branch);renderStatusSummary(d.status_summary);
    $('#branchFilter').value=d.branch.branch;$('#simBranch').value=d.branch.branch;await populateModels(true);
    await loadManagement();
  }finally{
    requestAnimationFrame(()=>document.body.classList.remove('app-busy'));
  }
}
const workspaceViews={
  dashboard:{title:'Executive Dashboard',subtitle:'Network inventory health, stock-out exposure, replenishment priorities, and distribution execution in one decision workspace.'},
  management:{title:'Management Order Plan',subtitle:'Review model-level inventory, PO balance, order quantity, projected DoI, and financial exposure with management-ready controls.'},
  simulator:{title:'Branch Request Simulator',subtitle:'Model branch requests before execution and immediately see inventory, stock-status, suggested-transfer, and new DoI impact.'}
};
function syncWorkspaceHero(tab){const m=workspaceViews[tab]||workspaceViews.dashboard;const t=$('#workspaceHeroTitle'),sub=$('#workspaceHeroSubtitle');if(t)t.textContent=m.title;if(sub)sub.textContent=m.subtitle;document.title=`${m.title} · SCM IDP Control Tower`;}
$$('.mainTab').forEach(btn=>btn.onclick=()=>{
  $$('.mainTab').forEach(b=>b.classList.remove('tab-active'));btn.classList.add('tab-active');
  const tab=btn.dataset.tab,panels={dashboard:$('#dashboardTab'),management:$('#managementTab'),simulator:$('#simulatorTab')};
  syncWorkspaceHero(tab);try{const u=new URL(window.location.href);if(tab==='dashboard')u.searchParams.delete('tab');else u.searchParams.set('tab',tab);history.replaceState({},'',u.pathname+(u.search||''));}catch(_){}
  Object.entries(panels).forEach(([name,panel])=>panel.classList.toggle('hidden',name!==tab));
  const active=panels[tab];if(active){active.classList.remove('view-enter');void active.offsetWidth;active.classList.add('view-enter');setTimeout(()=>active.classList.remove('view-enter'),460)}
  document.body.classList.remove('sidebar-mobile-open');window.scrollTo({top:0,behavior:'smooth'});
});
$$('.classTab').forEach(btn=>btn.onclick=()=>{state.classFilter=btn.dataset.class;renderClassModels();animateRefresh($('#modelTable'))});
$$('.reorderClassFilter').forEach(btn=>btn.onclick=()=>{state.reorderClassFilter=btn.dataset.reorderClass||'All';renderReorderModelCard(state.reorderCard)});if($('#reorderBrandFilter'))$('#reorderBrandFilter').onchange=e=>{state.reorderBrandFilter=e.target.value||'All Brands';renderReorderModelCard(state.reorderCard)};
$('#themeBtn').onclick=()=>setTheme(state.theme==='dark'?'light':'dark');syncThemeButton();syncPointDataButton();$('#pointDataBtn').onclick=()=>setPointDataVisibility(!state.showPointData);
initWorkspaceChrome();
const requestedWorkspaceTab=new URLSearchParams(window.location.search).get('tab');if(['management','simulator'].includes(requestedWorkspaceTab)){const requestedBtn=$(`.mainTab[data-tab="${requestedWorkspaceTab}"]`);if(requestedBtn)requestedBtn.click();}else{syncWorkspaceHero('dashboard');if(requestedWorkspaceTab==='delivery'){const u=new URL(location.href);u.searchParams.delete('tab');history.replaceState({},'',u.pathname+u.search)}}
$('#periodFilter').onchange=e=>{state.period=e.target.value;renderKPIs(state.boot.kpis)};
$('#areaFilter').onchange=e=>selectArea(e.target.value);
$('#areaDetailClose').onclick=()=>selectArea('Overall');
['#areaDetailBranch','#areaDetailClass','#areaDetailStatus'].forEach(id=>$(id).onchange=renderAreaModelDetails);$('#areaDetailSearch').oninput=renderAreaModelDetails;
$('#branchFilter').onchange=async e=>{try{const d=await fetchJson('/api/branch?branch='+encodeURIComponent(e.target.value),{},'branch');if(d)renderBranch(d)}catch(err){toast(err.message||'Unable to load Branch data.',false)}};
$('#simBranch').onchange=()=>populateModels(true);$('#simBrand').onchange=()=>populateModels(false);
initModelCombo('simModel','simModelDropdown',()=>state.requestBuilderModels.map(m=>({value:m,label:m})),loadModel,loadModel);$('#requestQty').oninput=calcNewDoi;
['#statusAreaFilter','#statusBranchFilter','#statusBrandFilter','#statusModelFilter','#statusClassFilter','#statusStockFilter'].forEach(sel=>$(sel).onchange=loadStatusSummary);
$('#statusResetBtn').onclick=()=>{$('#statusAreaFilter').value='Overall';$('#statusBranchFilter').value='All Branches';$('#statusBrandFilter').value='All Brands';$('#statusModelFilter').value='All Models';$('#statusClassFilter').value='All Classes';$('#statusStockFilter').value='All Statuses';loadStatusSummary()};
['#managementClassFilter','#managementStatusFilter'].forEach(sel=>$(sel).onchange=loadManagement);$('#managementBrandFilter').onchange=()=>{$('#managementModelFilter').value='';loadManagement()};let managementModelTimer=null;
initModelCombo('managementModelFilter','managementModelDropdown',()=>[{value:'',label:'All Models'},...(state.management?.filters?.model_options||[]).map(r=>({value:r.model,label:`${r.model} · ${r.brand} · Class ${r.class||'—'} · 3M Sales: ${salesText(r.sales_3m)} · ADS: ${adsText(r.avg_daily_sale)}`,brand:r.brand}))],()=>{clearTimeout(managementModelTimer);managementModelTimer=setTimeout(loadManagement,220)},chosen=>{clearTimeout(managementModelTimer);if(chosen.brand)$('#managementBrandFilter').value=chosen.brand;loadManagement()});
$('#managementResetBtn').onclick=()=>{$('#managementBrandFilter').value='All Brands';$('#managementClassFilter').value='All Classes';$('#managementStatusFilter').value='All Statuses';$('#managementModelFilter').value='';loadManagement()};
if(isAdmin){
    if($('#saveManagementBtn'))$('#saveManagementBtn').onclick=async()=>{if(!Object.keys(state.managementEdits).length){toast('No Management planning changes to save.');return}const r=await appFetch('/admin/management/allocations',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({allocations:state.managementEdits})});const d=await r.json();if(!r.ok){toast(d.error||'Management order save failed.',false);return}toast((d.message||'Order Plan saved.')+' Ready to Download Excel or Print Plan.');state.managementEdits={};await loadManagement()};
  if($('#downloadManagementBtn'))$('#downloadManagementBtn').onclick=async()=>{const payload=managementOutputPayload();const r=await appFetch('/admin/export/management',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});if(!r.ok){toast(await responseErrorMessage(r,'Management Excel export failed.'),false);return}await downloadResponse(r,'Management_Order_Plan.xlsx')};
  const clearDataModal=$('#clearDataModal');if($('#clearAllDataBtn'))$('#clearAllDataBtn').onclick=openClearDataModal;if($('#cancelClearDataBtn'))$('#cancelClearDataBtn').onclick=closeClearDataModal;if($('#confirmClearDataBtn'))$('#confirmClearDataBtn').onclick=clearAllUserData;if(clearDataModal)clearDataModal.onclick=e=>{if(e.target===clearDataModal)closeClearDataModal()};if($('#clearDataConfirmation'))$('#clearDataConfirmation').onkeydown=e=>{if(e.key==='Enter')clearAllUserData()};
  $('#pptBtn').onclick=async()=>{$('#adminActionsMenu')?.classList.add('hidden');const btn=$('#pptBtn'),label=btn?.querySelector('.sidebar-label'),oldText=label?.textContent||'Export Deck';const area=$('#areaFilter').value,branch=$('#branchFilter').value,reorderBrand=state.reorderBrandFilter||'All Brands';if(btn)btn.disabled=true;if(label)label.textContent='Generating Deck…';toast('Generating the executive PowerPoint deck…');try{const r=await appFetch(`/admin/export/pptx?area=${encodeURIComponent(area)}&branch=${encodeURIComponent(branch)}&reorder_brand=${encodeURIComponent(reorderBrand)}`);if(!r.ok){let e={};try{e=await r.json()}catch(_){ }toast(e.error||'PowerPoint export failed.',false);return}await downloadResponse(r,'SCM_Executive_Review.pptx');toast('PowerPoint deck generated successfully.')}catch(err){toast('PowerPoint export failed. '+(err?.message||''),false)}finally{if(btn)btn.disabled=false;if(label)label.textContent=oldText}};
  $('#addRequest').onclick=()=>{const m=state.model,q=Number($('#requestQty').value||0);if(!m||q<=0){toast('Enter a request quantity greater than zero.',false);return}const nd=m.avg_daily_sale>0?(m.inventory+q)/m.avg_daily_sale:'N/A';state.requests.push({...m,requested_qty:q,new_doi:nd,remarks:$('#remarks').value.trim()});renderRequest();toast(`Request line added for ${m.branch}.`);$('#remarks').value=''};
  if($('#requestOutputBranch'))$('#requestOutputBranch').onchange=e=>{state.requestOutputBranch=e.target.value||''};
  $('#downloadRequest').onclick=async()=>{const payload=selectedRequestOutput();if(!payload)return;const r=await appFetch('/admin/export/request',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});if(!r.ok){toast(await responseErrorMessage(r,'Branch Request Excel export failed.'),false);return}await downloadResponse(r,`Branch_Request_${payload.branch.replaceAll(' ','_')}.xlsx`)};
}
init().catch(e=>{console.error(e);toast('Dashboard failed to initialize: '+e.message,false)});
