(function(){
  const host=document.getElementById('dataOpsHost');
  const trigger=document.getElementById('dataOpsTrigger');
  const panel=document.getElementById('dataOpsPanel');
  if(!host||!trigger||!panel)return;

  let closeTimer=null;
  let pinned=false;
  const desktop=()=>window.matchMedia('(min-width:1024px) and (hover:hover)').matches;

  function setOpen(open,{pin=pinned}={}){
    if(closeTimer){clearTimeout(closeTimer);closeTimer=null;}
    pinned=!!pin;
    host.classList.toggle('is-open',open);
    trigger.setAttribute('aria-expanded',open?'true':'false');
    panel.setAttribute('aria-hidden',open?'false':'true');
  }

  function openHover(){
    if(desktop())setOpen(true,{pin:false});
  }
  function scheduleClose(){
    if(!desktop()||pinned)return;
    if(closeTimer)clearTimeout(closeTimer);
    closeTimer=setTimeout(()=>setOpen(false,{pin:false}),170);
  }

  host.addEventListener('pointerenter',openHover);
  host.addEventListener('pointerleave',scheduleClose);

  trigger.addEventListener('click',e=>{
    e.preventDefault();
    e.stopPropagation();
    const next=!host.classList.contains('is-open') || !pinned;
    setOpen(next,{pin:next});
  });

  host.addEventListener('focusin',()=>setOpen(true,{pin:pinned}));
  host.addEventListener('focusout',e=>{
    if(!host.contains(e.relatedTarget) && !pinned)setOpen(false,{pin:false});
  });

  panel.addEventListener('click',e=>{
    if(e.target.closest('button,a')){
      pinned=false;
      // Let modal/export/reset handlers run first, then close the accordion.
      requestAnimationFrame(()=>setOpen(false,{pin:false}));
    }
  });

  document.addEventListener('click',e=>{
    if(!host.contains(e.target))setOpen(false,{pin:false});
  });
  document.addEventListener('keydown',e=>{
    if(e.key==='Escape'&&host.classList.contains('is-open')){
      setOpen(false,{pin:false});
      trigger.focus();
    }
  });

  window.addEventListener('resize',()=>{
    if(!desktop())pinned=false;
  },{passive:true});
})();
