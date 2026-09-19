(function(){
  const modal=document.getElementById('columnGuideModal');
  function setModal(open){if(!modal)return;modal.classList.toggle('open',open);modal.setAttribute('aria-hidden',open?'false':'true');document.body.classList.toggle('modal-open',open)}
  document.querySelectorAll('[data-open-column-guide]').forEach(btn=>btn.addEventListener('click',()=>setModal(true)));
  document.querySelectorAll('[data-close-column-guide]').forEach(btn=>btn.addEventListener('click',()=>setModal(false)));
  document.addEventListener('keydown',e=>{if(e.key==='Escape'){setModal(false);document.body.classList.remove('sidebar-mobile-open')}});

  const form=document.querySelector('[data-auto-filter]');
  if(form){
    const status=document.getElementById('filterStatus');
    const areaFilter=document.getElementById('areaFilter');
    const branchFilter=document.getElementById('branchFilter');
    let timer=null,submitting=false;
    const submitFilter=()=>{if(submitting)return;submitting=true;if(status){status.classList.add('is-loading');status.innerHTML='<span class="filter-spinner"></span>Updating view…'}form.requestSubmit()};
    form.querySelectorAll('.auto-filter-control').forEach(control=>control.addEventListener('change',()=>{if(control===areaFilter&&branchFilter)branchFilter.value='';submitFilter()}));
    form.querySelectorAll('.auto-filter-text').forEach(input=>{
      input.addEventListener('input',()=>{clearTimeout(timer);if(status){status.classList.remove('is-loading');status.innerHTML='<span class="status-dot"></span>Waiting for typing…'}timer=setTimeout(submitFilter,420)});
      input.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();clearTimeout(timer);submitFilter()}});
    });
  }
})();

window._agingCharts=[];
window._agingChartData=null;
window.renderAgingCharts=function(data){
  if(typeof Chart==='undefined')return;
  window._agingChartData=data;
  (window._agingCharts||[]).forEach(c=>{try{c.destroy()}catch(_){}});window._agingCharts=[];
  const light=document.documentElement.dataset.theme==='light';
  const axis=light?'#5d7898':'#8da2ba', gridColor=light?'rgba(100,116,139,.14)':'rgba(148,163,184,.075)', tooltipBg=light?'rgba(255,255,255,.98)':'rgba(2,6,23,.96)', tooltipText=light?'#0b1f3a':'#f8fafc', tooltipBody=light?'#496887':'#cbd5e1', chartBorder=light?'#eef5fc':'#08111f';
  Chart.defaults.color=axis;Chart.defaults.borderColor=gridColor;Chart.defaults.font.family='Segoe UI Variable, Segoe UI, Arial, sans-serif';Chart.defaults.font.size=11;
  const grid={color:gridColor,drawBorder:false};
  const tooltip={backgroundColor:tooltipBg,titleColor:tooltipText,bodyColor:tooltipBody,borderColor:'rgba(59,130,246,.24)',borderWidth:1,padding:12,cornerRadius:10,displayColors:false};
  const base={responsive:true,maintainAspectRatio:false,animation:{duration:540,easing:'easeOutQuart'},plugins:{legend:{display:false},tooltip},scales:{x:{grid:{display:false},ticks:{maxRotation:0,minRotation:0,color:axis}},y:{beginAtZero:true,grid,ticks:{precision:0,color:axis}}}};
  const add=c=>{window._agingCharts.push(c);return c};

  const age=document.getElementById('ageChart');
  if(age)add(new Chart(age,{type:'bar',data:{labels:data.buckets,datasets:[{data:data.bucketValues,backgroundColor:['rgba(59,130,246,.88)','rgba(56,189,248,.86)','rgba(245,158,11,.88)','rgba(239,68,68,.88)'],hoverBackgroundColor:['#3b82f6','#38bdf8','#f59e0b','#ef4444'],borderRadius:9,borderSkipped:false,maxBarThickness:72}]},options:{...base,plugins:{...base.plugins,tooltip:{...tooltip,callbacks:{label:ctx=>`${Number(ctx.raw||0).toLocaleString()} units`}}}}}));

  const area=document.getElementById('areaChart');
  if(area)add(new Chart(area,{type:'bar',data:{labels:data.areas,datasets:[{data:data.areaValues,backgroundColor:'rgba(245,158,11,.84)',hoverBackgroundColor:'#f59e0b',borderRadius:7,borderSkipped:false,maxBarThickness:28}]},options:{...base,indexAxis:'y',plugins:{...base.plugins,tooltip:{...tooltip,callbacks:{label:ctx=>{const pct=(data.areaPctValues||[])[ctx.dataIndex]||0;return `${Number(ctx.raw||0).toLocaleString()} units aged 91+ · ${Number(pct).toFixed(1)}% of area stock`}}}},scales:{x:{beginAtZero:true,grid,ticks:{precision:0,color:axis}},y:{grid:{display:false},ticks:{color:axis}}}}}));

  const brand=document.getElementById('brandChart');
  if(brand)add(new Chart(brand,{type:'doughnut',data:{labels:data.brands,datasets:[{data:data.brandValues,backgroundColor:['#2563eb','#38bdf8','#22c55e','#f59e0b','#fb7185','#8b5cf6','#14b8a6','#94a3b8'],borderColor:chartBorder,borderWidth:4,hoverOffset:6}]},options:{responsive:true,maintainAspectRatio:false,cutout:'70%',animation:{duration:540,easing:'easeOutQuart'},plugins:{legend:{position:'right',labels:{boxWidth:8,boxHeight:8,usePointStyle:true,pointStyle:'circle',padding:16,color:axis}},tooltip:{...tooltip,callbacks:{label:ctx=>`${ctx.label}: ${Number(ctx.raw||0).toLocaleString()} units`}}}}}));

  const model=document.getElementById('modelChart');
  if(model)add(new Chart(model,{type:'bar',data:{labels:data.models,datasets:[{data:data.modelValues,backgroundColor:'rgba(239,68,68,.78)',hoverBackgroundColor:'#ef4444',borderRadius:7,borderSkipped:false,maxBarThickness:25}]},options:{...base,indexAxis:'y',plugins:{...base.plugins,tooltip:{...tooltip,callbacks:{label:ctx=>`${Number(ctx.raw||0).toLocaleString()} units aged 91+ days`}}},scales:{x:{beginAtZero:true,grid,ticks:{precision:0,color:axis}},y:{grid:{display:false},ticks:{autoSkip:false,color:axis,callback:function(value){const s=this.getLabelForValue(value);return s.length>38?s.slice(0,38)+'…':s}}}}}}));
};
window.addEventListener('scm-theme-change',()=>{if(window._agingChartData)window.renderAgingCharts(window._agingChartData)});
