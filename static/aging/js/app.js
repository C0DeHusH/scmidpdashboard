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

  const unitForm=document.querySelector('[data-unit-filter]');
  if(unitForm){
    const search=document.getElementById('unitTraceSearch');
    const age=document.getElementById('unitAgeFilter');
    const sort=document.getElementById('unitSortFilter');
    const bandButtons=[...document.querySelectorAll('[data-unit-band]')];
    let unitTimer=null,unitSubmitting=false;
    const submitUnit=()=>{
      if(unitSubmitting)return;
      unitSubmitting=true;
      unitForm.classList.add('is-updating');
      unitForm.requestSubmit();
    };
    age?.addEventListener('change',submitUnit);
    sort?.addEventListener('change',submitUnit);
    search?.addEventListener('input',()=>{
      clearTimeout(unitTimer);
      unitTimer=setTimeout(submitUnit,420);
    });
    search?.addEventListener('keydown',event=>{
      if(event.key==='Enter'){event.preventDefault();clearTimeout(unitTimer);submitUnit()}
      if(event.key==='Escape'){search.value='';clearTimeout(unitTimer);submitUnit()}
    });
    bandButtons.forEach(button=>button.addEventListener('click',()=>{
      if(age)age.value=button.dataset.unitBand||'all';
      submitUnit();
    }));
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
  Chart.defaults.color=axis;Chart.defaults.borderColor=gridColor;Chart.defaults.font.family='Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif';Chart.defaults.font.size=11;Chart.defaults.font.weight='500';
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

(function initModelSummarySorting(){
  const table=document.querySelector('[data-sortable-model-summary]');
  if(!table)return;
  const tbody=table.tBodies[0];
  if(!tbody)return;

  const buttons=[...table.querySelectorAll('.table-sort-button[data-sort-index]')];
  const search=document.getElementById('modelTableSearch');
  const sortField=document.getElementById('modelSortField');
  const descBtn=document.getElementById('modelSortDesc');
  const ascBtn=document.getElementById('modelSortAsc');
  const visibleCount=document.getElementById('modelVisibleCount');
  const sortStatus=document.getElementById('modelSortStatus');
  const noResults=document.getElementById('modelNoResults');
  const presetButtons=[...document.querySelectorAll('[data-model-sort-preset]')];

  const config={
    model:{index:1,type:'text',label:'Standard Description'},
    risk:{index:2,type:'number',label:'Exposure Level'},
    qty:{index:3,type:'number',label:'Total Units'},
    avgAge:{index:4,type:'number',label:'Average Age'},
    aged90:{index:5,type:'number',label:'91+ Units'},
    agedPct:{index:6,type:'number',label:'91+ Percentage'},
    agedValue:{index:7,type:'number',label:'Aged Value'},
    oldest:{index:8,type:'number',label:'Oldest Unit'},
    branches:{index:9,type:'number',label:'Branches'},
    areas:{index:10,type:'number',label:'Areas'}
  };
  const presetMap={risk:['aged90','descending'],capital:['agedValue','descending'],oldest:['oldest','descending']};
  const rows=[...tbody.rows].filter(row=>row.hasAttribute('data-original-index'));
  let activeKey='aged90';
  let direction='descending';

  function sortableValue(row,index,type){
    const cell=row.cells[index];
    const raw=cell?.dataset.sortValue ?? cell?.textContent ?? '';
    if(type==='number'){
      const number=Number(String(raw).replace(/[^0-9.+-]/g,''));
      return Number.isFinite(number)?number:0;
    }
    return String(raw).trim().toLocaleLowerCase();
  }

  function keyForButton(btn){
    return btn.dataset.sortKey || Object.keys(config).find(key=>config[key].index===Number(btn.dataset.sortIndex)) || 'model';
  }

  function directionText(){return direction==='descending'?'Highest → Lowest':'Lowest → Highest'}

  function updateHeaderState(){
    const active=config[activeKey];
    buttons.forEach(btn=>{
      const key=keyForButton(btn);
      const th=btn.closest('th');
      const selected=key===activeKey;
      th?.setAttribute('aria-sort',selected?direction:'none');
      const label=config[key]?.label || (btn.querySelector('span')?.textContent||'column').trim();
      const nextDirection=selected?(direction==='descending'?'ascending':'descending'):(config[key]?.type==='number'?'descending':'ascending');
      const nextLabel=nextDirection==='descending'?'highest to lowest':'lowest to highest';
      btn.setAttribute('aria-label',`${label}: ${selected?`sorted ${directionText().toLowerCase()}`:'not sorted'}. Activate to sort ${nextLabel}.`);
      btn.title=`Sort ${label} ${nextLabel}`;
    });
    if(sortField)sortField.value=activeKey;
    if(descBtn){descBtn.classList.toggle('active',direction==='descending');descBtn.setAttribute('aria-pressed',String(direction==='descending'))}
    if(ascBtn){ascBtn.classList.toggle('active',direction==='ascending');ascBtn.setAttribute('aria-pressed',String(direction==='ascending'))}
    if(sortStatus)sortStatus.textContent=`${active.label} · ${directionText()}`;
  }

  function updatePresetState(){
    presetButtons.forEach(btn=>{
      const preset=presetMap[btn.dataset.modelSortPreset];
      btn.classList.toggle('active',Boolean(preset&&preset[0]===activeKey&&preset[1]===direction));
    });
  }

  function applySearchAndRanks(){
    const term=(search?.value||'').trim().toLocaleLowerCase();
    let shown=0;
    rows.forEach(row=>{
      const name=(row.dataset.modelName||'').toLocaleLowerCase();
      const visible=!term||name.includes(term);
      row.hidden=!visible;
      if(visible){
        shown+=1;
        const rank=row.querySelector('.model-rank');
        if(rank)rank.textContent=String(shown);
      }
    });
    if(visibleCount)visibleCount.textContent=`${shown} ${shown===1?'model':'models'} shown`;
    if(noResults)noResults.classList.toggle('hidden',shown!==0);
    table.classList.toggle('has-filtered-rows',Boolean(term));
  }

  function sortRows(key=activeKey,nextDirection=direction,{announce=true}={}){
    if(!config[key])return;
    activeKey=key;
    direction=nextDirection==='ascending'?'ascending':'descending';
    const {index,type}=config[activeKey];
    table.classList.add('is-sorting');
    rows.sort((a,b)=>{
      const av=sortableValue(a,index,type),bv=sortableValue(b,index,type);
      const cmp=type==='number'?(av-bv):av.localeCompare(bv,undefined,{numeric:true,sensitivity:'base'});
      if(cmp===0)return Number(a.dataset.originalIndex||0)-Number(b.dataset.originalIndex||0);
      return direction==='ascending'?cmp:-cmp;
    });
    const fragment=document.createDocumentFragment();
    rows.forEach(row=>fragment.appendChild(row));
    tbody.appendChild(fragment);
    updateHeaderState();
    updatePresetState();
    applySearchAndRanks();
    window.setTimeout(()=>table.classList.remove('is-sorting'),120);
    if(announce&&window.showAgingToast){window.showAgingToast(`Models sorted by ${config[activeKey].label} · ${directionText()}`,'success')}
  }

  buttons.forEach(btn=>{
    btn.addEventListener('click',()=>{
      const key=keyForButton(btn);
      const next=activeKey===key?(direction==='descending'?'ascending':'descending'):(config[key].type==='number'?'descending':'ascending');
      sortRows(key,next);
    });
  });

  sortField?.addEventListener('change',()=>{
    const key=sortField.value;
    const next=config[key]?.type==='text'?'ascending':direction;
    sortRows(key,next);
  });
  descBtn?.addEventListener('click',()=>sortRows(activeKey,'descending'));
  ascBtn?.addEventListener('click',()=>sortRows(activeKey,'ascending'));
  presetButtons.forEach(btn=>btn.addEventListener('click',()=>{
    const preset=presetMap[btn.dataset.modelSortPreset];
    if(preset)sortRows(preset[0],preset[1]);
  }));
  search?.addEventListener('input',applySearchAndRanks);
  search?.addEventListener('keydown',event=>{if(event.key==='Escape'){search.value='';applySearchAndRanks();search.blur()}});

  sortRows('aged90','descending',{announce:false});
})();
