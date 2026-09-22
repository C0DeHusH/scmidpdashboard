(function(){
  const $=(selector)=>document.querySelector(selector);
  const $$=(selector)=>[...document.querySelectorAll(selector)];
  let modal=null;
  let globalController=null;
  let unitController=null;
  let globalTimer=null;
  let unitTimer=null;

  function setModal(open){
    modal=$('#columnGuideModal');
    if(!modal)return;
    modal.classList.toggle('open',open);
    modal.setAttribute('aria-hidden',open?'false':'true');
    document.body.classList.toggle('modal-open',open);
  }

  window.bindAgingColumnGuideButtons=function bindAgingColumnGuideButtons(){
    $$('[data-open-column-guide]').forEach(btn=>{
      if(btn.dataset.columnGuideBound==='1')return;
      btn.dataset.columnGuideBound='1';
      btn.addEventListener('click',()=>setModal(true));
    });
    $$('[data-close-column-guide]').forEach(btn=>{
      if(btn.dataset.columnGuideBound==='1')return;
      btn.dataset.columnGuideBound='1';
      btn.addEventListener('click',()=>setModal(false));
    });
  };

  function setRegionBusy(region,busy,label='Updating…'){
    if(!region)return;
    region.classList.toggle('is-partial-refreshing',busy);
    region.setAttribute('aria-busy',busy?'true':'false');
    if(busy){
      region.style.minHeight=`${Math.max(region.offsetHeight,80)}px`;
      region.dataset.loadingLabel=label;
    }else{
      region.style.minHeight='';
      delete region.dataset.loadingLabel;
    }
  }

  async function fetchPartial(url,controller){
    const response=await fetch(url,{credentials:'same-origin',cache:'no-store',headers:{'X-SCM-Partial':'1'},signal:controller.signal});
    let data={};
    try{data=await response.json()}catch(_){ }
    if(!response.ok)throw new Error(data.error||`Unable to refresh view (${response.status}).`);
    return data;
  }

  function updateHistory(viewUrl){
    if(!viewUrl)return;
    try{
      const next=new URL(viewUrl,window.location.origin);
      history.replaceState({},'',next.pathname+next.search+window.location.hash);
    }catch(_){ }
  }

  function setSelectOptions(select,items,value,{valueKey=null,labelKey=null,emptyLabel=null}={}){
    if(!select)return;
    const options=[];
    if(emptyLabel!==null)options.push(new Option(emptyLabel,''));
    (items||[]).forEach(item=>{
      const optionValue=valueKey?String(item?.[valueKey]??''):String(item??'');
      const optionLabel=labelKey?String(item?.[labelKey]??optionValue):String(item??'');
      options.push(new Option(optionLabel,optionValue));
    });
    select.replaceChildren(...options);
    select.value=value||'';
  }

  function syncGlobalControls(data){
    const form=$('#filterForm');
    if(!form)return;
    const filters=data.filters||{};
    const options=data.options||{};
    const area=form.elements.namedItem('area');
    const branch=form.elements.namedItem('branch');
    const brand=form.elements.namedItem('brand');
    setSelectOptions(area,options.areas,filters.area,{emptyLabel:'All Areas'});
    setSelectOptions(branch,options.branches,filters.branch,{valueKey:'branch_key',labelKey:'branch_name',emptyLabel:filters.area?`All Branches in ${filters.area}`:'All Branches'});
    setSelectOptions(brand,options.brands,filters.brand,{emptyLabel:'All Brands'});
    const asOf=form.elements.namedItem('as_of'); if(asOf)asOf.value=data.as_of||'';
    const basis=form.elements.namedItem('basis'); if(basis)basis.value=data.basis||'branch';
    const std=form.elements.namedItem('std'); if(std)std.value=filters.std||'';
    const q=form.elements.namedItem('q'); if(q)q.value=filters.q||'';
    const xlsx=$('#agingExportXlsx'); if(xlsx)xlsx.href=data.export_xlsx||xlsx.href;
    const csv=$('#agingExportCsv'); if(csv)csv.href=data.export_csv||csv.href;
    const exportCount=$('#agingExportRowCount');
    if(exportCount){
      const count=Number(data.row_count||0);
      exportCount.innerHTML=`<strong>${count.toLocaleString()}</strong>&nbsp; filtered rows`;
    }
    const hero=$('#workspaceHeroMeta');
    if(hero){
      hero.innerHTML=`<span class="workspace-chip"><span class="workspace-chip-dot"></span>Local Control Tower</span><span class="workspace-chip">As of · <strong>${escapeHtml(data.as_of||'—')}</strong></span><span class="workspace-chip">Basis · <strong>${data.basis==='company'?'Company':'Branch'}</strong></span><span class="workspace-chip">Access · <strong>${escapeHtml(document.body.dataset.role||'')}</strong></span>`;
    }
  }

  function escapeHtml(value){return String(value??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]))}

  function globalParams({reset=false}={}){
    const params=reset?new URLSearchParams():new URLSearchParams(new FormData($('#filterForm')));
    if(!reset){
      const unitForm=$('#unitTraceForm');
      if(unitForm){
        ['unit_age','unit_sort','unit_q'].forEach(name=>{
          const field=unitForm.elements.namedItem(name);
          if(field&&String(field.value||'')!=='')params.set(name,field.value);
        });
      }
    }
    return params;
  }

  function buildLiveExportUrl(anchor){
    const endpoint=anchor?.dataset?.agingExportEndpoint;
    if(!endpoint)return anchor?.href||'#';
    const isUnit=anchor.dataset.agingExportMode==='unit';
    // Main Aging export follows only the global Aging filters (Area / Branch /
    // Brand / Model / unit key). Unit Trace-only search/band/sort must never
    // silently empty the main Aging workbook.
    const params=isUnit?globalParams():new URLSearchParams(new FormData($('#filterForm')));
    if(isUnit)params.set('detail','1');
    const query=params.toString();
    return `${endpoint}${query?`?${query}`:''}`;
  }

  document.addEventListener('click',event=>{
    const anchor=event.target.closest?.('[data-aging-export-endpoint]');
    if(!anchor)return;
    event.preventDefault();
    const url=buildLiveExportUrl(anchor);
    anchor.href=url;
    window.location.assign(url);
  });

  window.refreshAgingView=async function refreshAgingView({reset=false}={}){
    const summaryRegion=$('#agingSummaryRegion');
    const resultsRegion=$('#agingResultsRegion');
    const status=$('#filterStatus');
    if(!summaryRegion||!resultsRegion)return;
    globalController?.abort();
    unitController?.abort();
    globalController=new AbortController();
    const controller=globalController;
    const params=globalParams({reset});
    const filterAnchor=$('#filterForm');
    const anchorTop=filterAnchor?.getBoundingClientRect().top ?? null;
    const focusId=document.activeElement?.id||'';
    const query=params.toString();
    const url=`/aging/partial${query?`?${query}`:''}`;
    setRegionBusy(summaryRegion,true,'Updating summary…');
    setRegionBusy(resultsRegion,true,'Updating cards and tables…');
    if(status){status.classList.add('is-loading');status.innerHTML='<span class="filter-spinner"></span>Updating cards…'}
    try{
      const data=await fetchPartial(url,controller);
      if(controller.signal.aborted)return;
      summaryRegion.innerHTML=data.summary_html||'';
      resultsRegion.innerHTML=data.results_html||'';
      syncGlobalControls(data);
      updateHistory(data.view_url);
      if(Number.isFinite(anchorTop)&&filterAnchor){
        const delta=filterAnchor.getBoundingClientRect().top-anchorTop;
        if(Math.abs(delta)>1)window.scrollBy({top:delta,left:0,behavior:'auto'});
      }
      if(focusId){const focusTarget=document.getElementById(focusId);try{focusTarget?.focus({preventScroll:true})}catch(_){focusTarget?.focus()}}
      window.bindAgingColumnGuideButtons?.();
      window.initModelSummarySorting?.();
      window.initAgingUnitFilters?.();
      if(Number(data.row_count||0)>0)window.renderAgingCharts?.(data.chart_data||{});
      else{
        (window._agingCharts||[]).forEach(chart=>{try{chart.destroy()}catch(_){ }});
        window._agingCharts=[];
      }
      if(status){status.classList.remove('is-loading');status.innerHTML='<span class="status-dot"></span>Updated without page refresh'}
    }catch(error){
      if(error?.name==='AbortError')return;
      if(status){status.classList.remove('is-loading');status.innerHTML='<span class="status-dot"></span>Auto update on'}
      window.showAgingToast?.(error?.message||'Unable to refresh Aging view.','error');
    }finally{
      if(globalController===controller)globalController=null;
      setRegionBusy(summaryRegion,false);
      setRegionBusy(resultsRegion,false);
    }
  };

  window.initAgingGlobalFilters=function initAgingGlobalFilters(){
    const form=$('[data-auto-filter]');
    if(!form||form.dataset.partialBound==='1')return;
    form.dataset.partialBound='1';
    const status=$('#filterStatus');
    const areaFilter=$('#areaFilter');
    const branchFilter=$('#branchFilter');
    const queue=()=>window.refreshAgingView?.();
    form.addEventListener('submit',event=>{event.preventDefault();clearTimeout(globalTimer);queue()});
    form.querySelectorAll('.auto-filter-control').forEach(control=>control.addEventListener('change',()=>{
      if(control===areaFilter&&branchFilter)branchFilter.value='';
      clearTimeout(globalTimer);queue();
    }));
    form.querySelectorAll('.auto-filter-text').forEach(input=>{
      input.addEventListener('input',()=>{
        globalController?.abort();
        clearTimeout(globalTimer);
        if(status){status.classList.remove('is-loading');status.innerHTML='<span class="status-dot"></span>Waiting for typing…'}
        globalTimer=setTimeout(queue,360);
      });
      input.addEventListener('keydown',event=>{
        if(event.key==='Enter'){event.preventDefault();clearTimeout(globalTimer);queue()}
      });
    });
    $('#agingResetFilters')?.addEventListener('click',()=>{clearTimeout(globalTimer);window.refreshAgingView?.({reset:true})});
  };

  function unitParams(form){return new URLSearchParams(new FormData(form))}

  window.refreshAgingUnitTrace=async function refreshAgingUnitTrace({focusId='',reset=false}={}){
    const section=$('#unitTrace');
    const form=$('#unitTraceForm');
    if(!section||!form)return;
    unitController?.abort();
    unitController=new AbortController();
    const controller=unitController;
    const params=unitParams(form);
    if(reset){params.set('unit_age','all');params.set('unit_sort','oldest');params.delete('unit_q')}
    const beforeTop=section.getBoundingClientRect().top;
    setRegionBusy(section,true,'Updating unit table…');
    try{
      const data=await fetchPartial(`/aging/partial/units?${params.toString()}`,controller);
      if(controller.signal.aborted)return;
      const holder=document.createElement('template');holder.innerHTML=String(data.html||'').trim();
      const next=holder.content.firstElementChild;
      if(!next||next.id!=='unitTrace')throw new Error('Unit Trace update returned an invalid component.');
      section.className=next.className;
      section.innerHTML=next.innerHTML;
      updateHistory(data.view_url);
      window.bindAgingColumnGuideButtons?.();
      window.initAgingUnitFilters?.();
      const afterTop=section.getBoundingClientRect().top;
      const delta=afterTop-beforeTop;
      if(Math.abs(delta)>1)window.scrollBy({top:delta,left:0,behavior:'auto'});
      const target=focusId?document.getElementById(focusId):null;
      if(target){try{target.focus({preventScroll:true})}catch(_){target.focus()}}
    }catch(error){
      if(error?.name==='AbortError')return;
      window.showAgingToast?.(error?.message||'Unable to refresh Unit Trace.','error');
    }finally{
      if(unitController===controller)unitController=null;
      setRegionBusy(section,false);
    }
  };

  window.initAgingUnitFilters=function initAgingUnitFilters(){
    const form=$('[data-unit-filter]');
    if(!form||form.dataset.partialBound==='1')return;
    form.dataset.partialBound='1';
    const search=$('#unitTraceSearch');
    const age=$('#unitAgeFilter');
    const sort=$('#unitSortFilter');
    const submit=(focusId='')=>window.refreshAgingUnitTrace?.({focusId});
    form.addEventListener('submit',event=>{event.preventDefault();clearTimeout(unitTimer);submit(document.activeElement?.id||'')});
    age?.addEventListener('change',()=>submit(age.id));
    sort?.addEventListener('change',()=>submit(sort.id));
    search?.addEventListener('input',()=>{unitController?.abort();clearTimeout(unitTimer);unitTimer=setTimeout(()=>submit(search.id),360)});
    search?.addEventListener('keydown',event=>{
      if(event.key==='Enter'){event.preventDefault();clearTimeout(unitTimer);submit(search.id)}
      if(event.key==='Escape'){event.preventDefault();search.value='';clearTimeout(unitTimer);submit(search.id)}
    });
    $$('[data-unit-band]').forEach(button=>{
      if(button.dataset.partialBound==='1')return;
      button.dataset.partialBound='1';
      button.addEventListener('click',()=>{
        if(age)age.value=button.dataset.unitBand||'all';
        submit(age?.id||'');
      });
    });
    $('.unit-filter-reset')?.addEventListener('click',event=>{
      event.preventDefault();
      if(age)age.value='all';if(sort)sort.value='oldest';if(search)search.value='';
      window.refreshAgingUnitTrace?.({focusId:age?.id||'',reset:true});
    });
  };

  document.addEventListener('keydown',event=>{
    if(event.key==='Escape'){
      setModal(false);
      document.body.classList.remove('sidebar-mobile-open');
    }
  });

  window.bindAgingColumnGuideButtons();
  window.initAgingGlobalFilters();
  window.initAgingUnitFilters();
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

window.initModelSummarySorting=function initModelSummarySorting(){
  const table=document.querySelector('[data-sortable-model-summary]');
  if(!table||table.dataset.sortBound==='1')return;
  table.dataset.sortBound='1';
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
};

window.initModelSummarySorting?.();
