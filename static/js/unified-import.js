(function(){
  const $=s=>document.querySelector(s);
  const modal=$('#unifiedImportModal');
  if(!modal)return;
  const input=$('#unifiedImportFile'),drop=$('#unifiedImportDropzone'),submit=$('#unifiedImportSubmit');
  const fileName=$('#unifiedImportFileName'),fileMeta=$('#unifiedImportFileMeta');
  const progress=$('#unifiedImportProgressBar'),pct=$('#unifiedImportProgressPct'),progressLabel=$('#unifiedImportProgressLabel'),stage=$('#unifiedImportStage'),result=$('#unifiedImportResult');
  let file=null,busy=false,phaseTimer=null;

  const adminLoginUrl=()=>{const next=window.location.pathname+window.location.search+window.location.hash;return `/login?next=${encodeURIComponent(next)}`};
  const redirectToLogin=message=>{
    setProgress(100,'Admin session needs to be refreshed',message||'Please sign in again to continue.');
    result.hidden=false;result.className='unified-import-result is-error';result.innerHTML=`<b>Admin sign-in required.</b><span>${escapeHtml(message||'Please sign in again to continue.')}</span>`;
    window.setTimeout(()=>window.location.assign(adminLoginUrl()),700);
  };

  const parseXhrResponse=xhr=>{
    const raw=String(xhr.responseText||'').trim();
    if(!raw)return {};
    try{return JSON.parse(raw)}catch(_err){
      const text=raw.replace(/<[^>]*>/g,' ').replace(/\s+/g,' ').trim();
      return {error:text?`Server response: ${text.slice(0,220)}`:`Server returned HTTP ${xhr.status||'error'} without a readable message.`};
    }
  };
  const errorDetail=data=>{
    const parts=[];
    if(data?.error)parts.push(String(data.error));
    if(data?.stage)parts.push(`Stage: ${data.stage}`);
    if(data?.reference)parts.push(`Reference: ${data.reference}`);
    return parts.join(' · ')||'The previous data remains active. No changes were committed.';
  };

  const setProgress=(n,label,detail)=>{
    const value=Math.max(0,Math.min(100,Number(n)||0));
    progress.style.width=value+'%';pct.textContent=Math.round(value)+'%';
    if(label)progressLabel.textContent=label;if(detail)stage.textContent=detail;
  };
  const reset=()=>{
    if(phaseTimer)clearInterval(phaseTimer);phaseTimer=null;file=null;busy=false;
    input.value='';submit.disabled=true;submit.textContent='Import All Data';
    fileName.textContent='Choose consolidated workbook';
    fileMeta.textContent='Required sheets: Raw · KPI_YTD_Input · KPI_WEEKLY_Input · Aging · Reorder/Management';
    result.hidden=true;result.innerHTML='';drop.classList.remove('has-file','is-dragging','is-error');
    setProgress(0,'Ready for one-step refresh','Select the consolidated .xlsx file. Aging no longer needs a separate import.');
  };
  const open=()=>{reset();modal.classList.add('is-open');modal.setAttribute('aria-hidden','false');document.body.classList.add('modal-open');setTimeout(()=>drop.focus?.(),80)};
  const close=()=>{if(busy)return;modal.classList.remove('is-open');modal.setAttribute('aria-hidden','true');document.body.classList.remove('modal-open')};
  const setFile=f=>{
    if(!f)return;
    if(!/\.xlsx$/i.test(f.name)){drop.classList.add('is-error');file=null;submit.disabled=true;fileName.textContent='Unsupported file';fileMeta.textContent='Please select the consolidated .xlsx workbook.';return}
    file=f;drop.classList.remove('is-error');drop.classList.add('has-file');submit.disabled=false;
    const mb=(f.size/1024/1024).toFixed(1);fileName.textContent=f.name;fileMeta.textContent=`${mb} MB · Ready to refresh all analytical modules`;
    setProgress(4,'Workbook selected','Click Import All Data to validate and refresh SCM + Aging together.');
  };
  document.addEventListener('click',e=>{const el=e.target.closest?.('[data-open-unified-import]');if(!el)return;e.preventDefault();open()});
  document.querySelectorAll('[data-close-unified-import]').forEach(el=>el.addEventListener('click',close));
  input.addEventListener('change',()=>setFile(input.files?.[0]));
  drop.addEventListener('keydown',e=>{if(!busy&&(e.key==='Enter'||e.key===' ')){e.preventDefault();input.click()}});
  ['dragenter','dragover'].forEach(name=>drop.addEventListener(name,e=>{e.preventDefault();if(!busy)drop.classList.add('is-dragging')}));
  ['dragleave','drop'].forEach(name=>drop.addEventListener(name,e=>{e.preventDefault();drop.classList.remove('is-dragging')}));
  drop.addEventListener('drop',e=>{if(!busy)setFile(e.dataTransfer?.files?.[0])});
  document.addEventListener('keydown',e=>{if(e.key==='Escape'&&modal.classList.contains('is-open'))close()});

  submit.addEventListener('click',()=>{
    if(!file||busy)return;
    busy=true;submit.disabled=true;submit.textContent='Importing…';result.hidden=true;drop.classList.remove('is-error');
    setProgress(10,'Uploading workbook','Sending one consolidated source to the control tower…');
    const fd=new FormData();fd.append('file',file);
    const xhr=new XMLHttpRequest();xhr.open('POST','/admin/import',true);xhr.withCredentials=true;xhr.timeout=210000;
    xhr.upload.onprogress=e=>{if(e.lengthComputable){const uploadPct=Math.min(34,10+(e.loaded/e.total)*24);setProgress(uploadPct,'Uploading workbook','Secure upload in progress…')}};
    xhr.upload.onload=()=>{
      let synthetic=36;setProgress(synthetic,'Validating workbook','Checking workbook structure before any live data is changed…');
      phaseTimer=setInterval(()=>{synthetic=Math.min(88,synthetic+4);let label='Staging control-tower data',detail='Parsing inventory, KPI and management intelligence safely in memory…';if(synthetic>=60){label='Refreshing Motorcycle Aging';detail='Mapping Aging branches to the Raw Branch/Area master and recalculating age exposure…'}if(synthetic>=78){label='Committing unified refresh';detail='Creating rollback snapshots and activating the validated dashboard revision…'}setProgress(synthetic,label,detail)},420);
    };
    xhr.onload=()=>{
      if(phaseTimer)clearInterval(phaseTimer);phaseTimer=null;
      const data=parseXhrResponse(xhr);
      if((xhr.status===401||xhr.status===403)&&(data.reauth_required||String(data.error||'').toLowerCase().includes('admin'))){busy=false;redirectToLogin(data.error);return}
      if(xhr.status<200||xhr.status>=300){
        busy=false;submit.disabled=false;submit.textContent='Try Again';drop.classList.add('is-error');
        const detail=errorDetail(data);
        setProgress(100,'Import could not be completed',detail);
        result.hidden=false;result.className='unified-import-result is-error';
        result.innerHTML=`<b>Refresh stopped safely.</b><span>${escapeHtml(detail)}</span>`;return
      }
      const a=data.modules?.aging||{};const k=data.modules?.kpi||{};const warnings=Array.isArray(data.warnings)?data.warnings:[];const persistence=data.persistence||{};
      const persistenceNote=persistence.mode==='runtime-fallback'?' · Runtime mode':(persistence.mode==='durable-cloud'?' · Cloud-synced':'');
      const kpiNote=` · KPI YTD ${Number(k.ytd_periods||0)} / Weekly ${Number(k.weekly_periods||0)}`;
      setProgress(100,'All modules refreshed',`Executive, KPI, Management and Motorcycle Aging are aligned.${kpiNote}${persistenceNote} ${data.reference?`Reference: ${data.reference}`:''}`);
      const warningNote=warnings.length?`<small class="result-warning">${escapeHtml(warnings[0])}</small>`:'';
      result.hidden=false;result.className='unified-import-result is-success';result.innerHTML=`<div><b>Unified refresh complete</b><span>${Number(a.rows||0).toLocaleString()} aging units · ${Number(a.branches||0).toLocaleString()} branches · ${Number(a.areas||0).toLocaleString()} areas · KPI YTD ${Number(k.ytd_periods||0)} / Weekly ${Number(k.weekly_periods||0)} · As of ${escapeHtml(a.as_of_date||'—')}${persistenceNote}${data.reference?` · ${escapeHtml(data.reference)}`:''}</span>${warningNote}</div><span class="result-check">✓</span>`;
      submit.textContent='Refreshing View…';
      window.setTimeout(()=>{const u=new URL(window.location.href);u.searchParams.delete('open_import');window.location.replace(u.pathname+u.search+u.hash)},850);
    };
    xhr.onerror=()=>{if(phaseTimer)clearInterval(phaseTimer);phaseTimer=null;busy=false;submit.disabled=false;submit.textContent='Try Again';drop.classList.add('is-error');const detail='Connection interrupted before the server returned a result. The previous data remains active.';setProgress(100,'Connection interrupted',detail);result.hidden=false;result.className='unified-import-result is-error';result.innerHTML=`<b>Refresh stopped safely.</b><span>${escapeHtml(detail)}</span>`;};
    xhr.ontimeout=()=>{if(phaseTimer)clearInterval(phaseTimer);phaseTimer=null;busy=false;submit.disabled=false;submit.textContent='Try Again';drop.classList.add('is-error');const detail='The import exceeded the browser wait limit. The previous data remains active. Retry once; if it persists, check /health storage status and the deployment logs.';setProgress(100,'Import timed out',detail);result.hidden=false;result.className='unified-import-result is-error';result.innerHTML=`<b>Refresh stopped safely.</b><span>${escapeHtml(detail)}</span>`;};
    xhr.send(fd);
  });

  function escapeHtml(value){return String(value??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]))}
  if(new URLSearchParams(window.location.search).get('open_import')==='1')window.setTimeout(open,180);
})();
