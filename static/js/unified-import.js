(function(){
  const $=s=>document.querySelector(s);
  const modal=$('#unifiedImportModal');
  if(!modal)return;
  const input=$('#unifiedImportFile'),drop=$('#unifiedImportDropzone'),submit=$('#unifiedImportSubmit');
  const fileName=$('#unifiedImportFileName'),fileMeta=$('#unifiedImportFileMeta');
  const progress=$('#unifiedImportProgressBar'),pct=$('#unifiedImportProgressPct'),progressLabel=$('#unifiedImportProgressLabel'),stage=$('#unifiedImportStage'),result=$('#unifiedImportResult');
  let file=null,busy=false,phaseTimer=null;

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
  document.querySelectorAll('[data-open-unified-import]').forEach(el=>el.addEventListener('click',e=>{e.preventDefault();open()}));
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
    setProgress(10,'Uploading workbook','Sending one consolidated source to the local control tower…');
    const fd=new FormData();fd.append('file',file);
    const xhr=new XMLHttpRequest();xhr.open('POST','/admin/import',true);xhr.responseType='json';
    xhr.upload.onprogress=e=>{if(e.lengthComputable){const uploadPct=Math.min(34,10+(e.loaded/e.total)*24);setProgress(uploadPct,'Uploading workbook','Secure local upload in progress…')}};
    xhr.upload.onload=()=>{
      let synthetic=36;setProgress(synthetic,'Validating workbook','Checking Raw, KPI, Reorder/Management and Aging sheets…');
      phaseTimer=setInterval(()=>{synthetic=Math.min(88,synthetic+4);let label='Refreshing analytics',detail='Updating inventory, KPI and management intelligence…';if(synthetic>=60){label='Refreshing Motorcycle Aging';detail='Mapping Aging branches to the Raw Branch/Area master and recalculating age exposure…'}if(synthetic>=78){label='Finalizing unified refresh';detail='Synchronizing filters, caches and local saved source…'}setProgress(synthetic,label,detail)},360);
    };
    xhr.onload=()=>{
      if(phaseTimer)clearInterval(phaseTimer);phaseTimer=null;
      const data=xhr.response||{};
      if(xhr.status<200||xhr.status>=300){busy=false;submit.disabled=false;submit.textContent='Try Again';drop.classList.add('is-error');setProgress(100,'Import could not be completed',data.error||'The previous data remains active.');result.hidden=false;result.className='unified-import-result is-error';result.innerHTML=`<b>Refresh stopped safely.</b><span>${escapeHtml(data.error||'Import failed.')}</span>`;return}
      const a=data.modules?.aging||{};setProgress(100,'All modules refreshed','Executive, Management and Motorcycle Aging are now aligned to the same workbook.');
      result.hidden=false;result.className='unified-import-result is-success';result.innerHTML=`<div><b>Unified refresh complete</b><span>${Number(a.rows||0).toLocaleString()} aging units · ${Number(a.branches||0).toLocaleString()} branches · ${Number(a.areas||0).toLocaleString()} areas · As of ${escapeHtml(a.as_of_date||'—')}</span></div><span class="result-check">✓</span>`;
      submit.textContent='Refreshing View…';
      window.setTimeout(()=>{const u=new URL(window.location.href);u.searchParams.delete('open_import');window.location.replace(u.pathname+u.search+u.hash)},850);
    };
    xhr.onerror=()=>{if(phaseTimer)clearInterval(phaseTimer);phaseTimer=null;busy=false;submit.disabled=false;submit.textContent='Try Again';drop.classList.add('is-error');setProgress(100,'Connection interrupted','No data was intentionally cleared. Please retry the import.');};
    xhr.send(fd);
  });

  function escapeHtml(value){return String(value??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]))}
  if(new URLSearchParams(window.location.search).get('open_import')==='1')window.setTimeout(open,180);
})();
