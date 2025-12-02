(() => {
    // Prevent duplicate button injection
    if (window.__capThumbBtnV2p2) {
      console.log('[CapThumb] Floating button already injected (v2.2).');
      return;
    }
  
    // ---------------- Config ----------------
    const WARMUP_MS = 1000; // wait after warming the original PDF via Shares API
  
    // ---------------- Small utils ----------------
    const wait = (ms) => new Promise(r => setTimeout(r, ms));
    const waitFor = async (getter, { timeout = 12000, interval = 120 } = {}) => {
      const end = performance.now() + timeout;
      while (performance.now() < end) {
        const v = getter();
        if (v) return v;
        await wait(interval);
      }
      return null;
    };
  
    // ---------------- Core helpers ----------------
    function isSharePoint(url){ try{ return new URL(url).hostname.includes('sharepoint.com'); }catch{ return false; } }
    function toBase64Url(str){ const b = new TextEncoder().encode(str); let s=''; for(const x of b) s+=String.fromCharCode(x); return btoa(s).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/g,''); }
  
    function* shareLinkThumbCandidates(shareUrl){
      const origin = new URL(shareUrl).origin;
      const token = 'u!' + toBase64Url(shareUrl);
      // We’ll retry/blank-check each one
    //   yield `${origin}/_api/v2.0/shares/${token}/driveItem/thumbnails/0/c640x640/content`;
    //   yield `${origin}/_api/v2.0/shares/${token}/driveItem/thumbnails/0/c1024x1024/content`;
      yield `${origin}/_api/v2.0/shares/${token}/driveItem/thumbnails/0/small/content`;
    //   yield `${origin}/_api/v2.0/shares/${token}/driveItem/thumbnails/0/medium/content`;
    //   yield `${origin}/_api/v2.0/shares/${token}/driveItem/thumbnails/0/large/content`;
    //   yield `${origin}/_api/v2.0/shares/${token}/driveItem/thumbnails/0/smallSquare/content`;
    //   yield `${origin}/_api/v2.0/shares/${token}/driveItem/thumbnails/0/mediumSquare/content`;
    //   yield `${origin}/_api/v2.0/shares/${token}/driveItem/thumbnails/0/largeSquare/content`;
    //   yield `${origin}/_api/v2.0/shares/${token}/driveItem/thumbnails/0/c256x256/content`;
    }
  
    function appendCacheBust(url, n){ const u = new URL(url); u.searchParams.set('cb', `${Date.now()}_${n}`); return u.toString(); }
  
    function imageLooksBlank(img){
      const w=img.naturalWidth, h=img.naturalHeight;
      const c=document.createElement('canvas'); c.width=Math.min(256,w||1); c.height=Math.min(256,h||1);
      const ctx=c.getContext('2d');
      const s=Math.min(c.width/w,c.height/h); const dw=w*s, dh=h*s, dx=(c.width-dw)/2, dy=(c.height-dh)/2;
      ctx.drawImage(img,dx,dy,dw,dh);
      const data=ctx.getImageData(0,0,c.width,c.height).data;
      let sum=0,sum2=0,n=0;
      for(let i=0;i<data.length;i+=64){ const r=data[i],g=data[i+1],b=data[i+2]; const y=0.2126*r+0.7152*g+0.0722*b; sum+=y; sum2+=y*y; n++; }
      if(!n) return true;
      const mean=sum/n; const std=Math.sqrt(Math.max(0,(sum2/n)-(mean*mean)));
      const blank = mean>245 && std<4;
      console.log(`[ImgCheck] mean=${mean.toFixed(1)} std=${std.toFixed(1)} → ${blank?'blank-ish':'ok'}`);
      return blank;
    }
  
    // Unchanged: no settle wait here
    async function loadImageOnce(url){
      console.log(`[ImgLoad] <img> src: ${url}`);
      const img=new Image(); img.crossOrigin='anonymous'; img.decoding='async'; img.referrerPolicy='no-referrer';
      await new Promise((res,rej)=>{ img.onload=res; img.onerror=rej; img.src=url; });
      console.log(`[ImgLoad] naturalSize=${img.naturalWidth}x${img.naturalHeight}`);
      return img;
    }
  
    async function loadThumbnailWithRetries(baseUrl,{attempts=1,delayMs=2000}={}){
      for(let i=1;i<=attempts;i++){
        try {
            const url = appendCacheBust(baseUrl,i);
            const img = await loadImageOnce(url);
            if(!imageLooksBlank(img)) return img;
            console.warn(`[Retry] Thumbnail looked blank (attempt ${i}/${attempts}). Waiting ${delayMs}ms…`);
          } catch(e) {
            if (e && e.message && /403/.test(e.message)) {
              console.warn(`[Retry] Got 403 Forbidden at ${baseUrl}. Not retrying further.`);
              return null; // bail out immediately
            }
            console.warn(`[Retry] Image load failed (attempt ${i}/${attempts}).`, e);
          }
          await wait(delayMs);
          
      }
      return null;
    }
  
    async function loadAnyCandidateWithRetries(shareUrl){
      for(const base of shareLinkThumbCandidates(shareUrl)){
        console.log(`[Try] Candidate base URL: ${base}`);
        const img = await loadThumbnailWithRetries(base);
        if(img) return img;
        console.warn('[Try] Candidate failed after retries, moving on…');
      }
      return null;
    }
  
    // NEW: warm up original PDF via Shares API (priming SharePoint) before we fetch thumbnail
    async function warmUpSharePointPdf(shareUrl, warmMs = WARMUP_MS){
      try {
        const origin = new URL(shareUrl).origin;
        const token = 'u!' + toBase64Url(shareUrl);
        const metaUrl = `${origin}/_api/v2.0/shares/${token}/driveItem?$select=name,@microsoft.graph.downloadUrl,webUrl`;
        console.log('[Warmup] Fetching driveItem metadata…', metaUrl);
        const res = await fetch(metaUrl, { credentials: 'include' });
        if (!res.ok) {
          console.warn('[Warmup] driveItem metadata request failed:', res.status, res.statusText);
        } else {
          const item = await res.json().catch(()=>null);
          const dl = item && item['@microsoft.graph.downloadUrl'];
          const webUrl = item && item.webUrl;
          // Prefer downloading a small byte range; this usually primes the server-side render/cache
          const warmTarget = dl || webUrl || shareUrl;
          console.log('[Warmup] Priming PDF via:', dl ? 'downloadUrl' : (webUrl ? 'webUrl' : 'shareUrl'));
          try {
            const warmRes = await fetch(warmTarget, {
              credentials: 'include',
              headers: { 'Range': 'bytes=0-32767' }, // small range to avoid big transfer
              cache: 'no-store',
            });
            console.log('[Warmup] Prime fetch status:', warmRes.status, warmRes.statusText);
            // We do not need the body; just ensure request hits the backend
            try { warmRes.body?.cancel?.(); } catch {}
          } catch (e) {
            console.warn('[Warmup] Prime fetch error:', e);
          }
        }
      } catch (e) {
        console.warn('[Warmup] Unexpected error:', e);
      }
      if (warmMs > 0) {
        console.log(`[Warmup] Waiting ${warmMs}ms for SharePoint to finish rendering…`);
        await wait(warmMs);
      }
    }
  
    async function imageTo180ThumbBlob(img){
      const sw=img.naturalWidth, sh=img.naturalHeight, target=180;
      const c=document.createElement('canvas'); c.width=target; c.height=target; const ctx=c.getContext('2d');
      const scale=Math.max(target/sw,target/sh); const dw=sw*scale, dh=sh*scale, dx=(target-dw)/2, dy=(target-dh)/2;
      console.log(`[Thumb] cover scale=${scale.toFixed(2)} draw=${dw.toFixed(1)}x${dh.toFixed(1)} at (${dx.toFixed(1)},${dy.toFixed(1)})`);
      ctx.drawImage(img,dx,dy,dw,dh);
      return await new Promise(res=>c.toBlob(res,'image/jpeg',0.9));
    }
  
    async function realDblClick(el){
      const mk=(t,o={})=>new MouseEvent(t,{bubbles:true,cancelable:true,view:window,...o});
      el.dispatchEvent(mk('mouseover')); el.dispatchEvent(mk('mousemove'));
      el.dispatchEvent(mk('mousedown')); el.dispatchEvent(mk('mouseup')); el.dispatchEvent(mk('click'));
      await wait(40);
      el.dispatchEvent(mk('mousedown')); el.dispatchEvent(mk('mouseup')); el.dispatchEvent(mk('click')); el.dispatchEvent(mk('dblclick'));
    }
  
    const isVisible = (el)=>!!(el && el.offsetParent!==null);
    const getZ = (el)=>{ const z = +getComputedStyle(el).zIndex; return Number.isFinite(z)?z:0; };
    function findActiveDialog(){
      const q = '#commonEditorCalloutId, .ms-Callout, [role="dialog"], .ms-Dialog-main, .fui-DialogActions, .ms-Dialog-actions';
      const candidates = Array.from(document.querySelectorAll(q)).filter(isVisible);
      if(!candidates.length) return null;
      const best = candidates.reduce((a,b)=> (getZ(a)>=getZ(b)?a:b));
      console.log('[Dialog] Active container picked:', best, 'z=', getZ(best));
      return best;
    }
  
    function findUploadishButton(scope){
      const roots = [];
      if(scope) roots.push(scope);
      roots.push(...document.querySelectorAll('.fui-DialogActions, .ms-Dialog-actions, [role="dialog"], .ms-Dialog-main'));
      roots.push(document.body);
  
      const isUploadish = (btn)=>{
        const t=(btn.textContent||'').toLowerCase();
        const title=(btn.getAttribute('title')||'').toLowerCase();
        const aria=(btn.getAttribute('aria-label')||'').toLowerCase();
        const auto=(btn.getAttribute('data-automationid')||'').toLowerCase();
        return /upload|add|save/.test(t) || /upload|add|save/.test(title) || /upload|add|save/.test(aria) || /upload|primary/.test(auto);
      };
  
      let btn = document.querySelector('.fui-DialogActions button.ms-Button--primary, .ms-Dialog-actions button.ms-Button--primary');
      if(btn && isVisible(btn) && !btn.disabled) return btn;
  
      for(const root of roots){
        const cand = Array.from(root.querySelectorAll('button, [role="button"]'))
          .find(b => isUploadish(b) && isVisible(b) && !(b.disabled || b.getAttribute('aria-disabled')==='true'));
        if(cand) return cand;
      }
      return null;
    }
  
    async function clickPrimaryAction(){
      const scope = findActiveDialog();
      const btn = await waitFor(()=>findUploadishButton(scope), { timeout: 12000, interval: 120 });
      if(!btn){ console.warn('[Upload] Primary action button not found.'); return false; }
      try{ btn.scrollIntoView({block:'center'}); }catch{}
      const clickable = btn.closest('.ms-Button') || btn;
      console.log('[Upload] Clicking primary action…', clickable);
      const fire = (t)=>clickable.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,view:window,buttons:1}));
      const firePtr=(t)=>clickable.dispatchEvent(new PointerEvent(t,{bubbles:true,cancelable:true,pointerId:1,pointerType:'mouse',buttons:1}));
      firePtr('pointerover'); fire('mouseover');
      firePtr('pointerenter'); fire('mouseenter');
      firePtr('pointerdown');  fire('mousedown');
      firePtr('pointerup');    fire('mouseup');
      clickable.click();
      console.log('[Upload] Click dispatched.');
      return true;
    }
  
    async function pressEnterOnDialog(){
      const dlg = findActiveDialog();
      if(!dlg) return false;
      const ev = new KeyboardEvent('keydown',{bubbles:true,cancelable:true,key:'Enter',code:'Enter',which:13,keyCode:13});
      console.log('[Upload] Pressing Enter on active dialog as fallback.');
      return dlg.dispatchEvent(ev);
    }
  
    function buildThumbFilename(row){
      const pf = row.querySelector('[data-automationid="field-PosterFileName"]');
      const first = row.querySelector('[data-automationid="field-FirstName"]')?.textContent?.trim() || '';
      const last  = row.querySelector('[data-automationid="field-LastName"]')?.textContent?.trim() || '';
      const clean = s => (s||'').trim().replace(/\s+/g,'_').replace(/[^\w.-]/g,'');
      if(pf && pf.textContent.trim()){
        const base = pf.textContent.trim().replace(/\.(jpe?g|png|webp|gif|pdf)$/i,'');
        return `${clean(base)}_thumb.jpeg`;
      }
      if(first||last) return `Capstone_${clean(first)}_${clean(last)}_thumb.jpeg`;
      return 'Capstone_thumbnail.jpeg';
    }
  
    // ---------------- New Item form guard ----------------
    function isNewItemFormOpen() {
      const header = document.querySelector('#reactClientFormHeader');
      if (!header) return false;
      const txt = (header.textContent || '').trim().toLowerCase();
      const open = /(^|\b)new item\b/.test(txt);
      if (open) console.warn('[NewItem] Detected "New item" form open.');
      return open;
    }
  
    async function tryCloseNewItemForm() {
      if (!isNewItemFormOpen()) return true;
  
      // Look for common close/cancel controls
      const selectors = [
        'button[aria-label="Close"]',
        'button[title="Close"]',
        '.ms-Panel-closeButton',
        '.ms-Dialog-button--close',
        'button[aria-label*="Close"]',
        'button[title*="Close"]',
        'button[aria-label="Cancel"]',
        'button[title="Cancel"]',
        'button:has(> span, > div):not([disabled])' // very broad fallback
      ];
      let closed = false;
  
      for (const sel of selectors) {
        const btn = document.querySelector(sel);
        if (btn && isVisible(btn)) {
          try { btn.scrollIntoView({ block: 'center' }); } catch {}
          console.log('[NewItem] Clicking close/cancel button:', sel, btn);
          btn.click();
          closed = true;
          break;
        }
      }
  
      if (!closed) {
        // Fallback: send Escape key to active dialog/panel
        const ev = new KeyboardEvent('keydown',{bubbles:true,cancelable:true,key:'Escape',code:'Escape',which:27,keyCode:27});
        console.log('[NewItem] No close button found; sending Escape.');
        document.dispatchEvent(ev);
        closed = true; // assume UI handles it
      }
  
      await wait(500); // allow UI to close
      const stillOpen = isNewItemFormOpen();
      console.log(`[NewItem] Close attempt finished. Still open? ${stillOpen}`);
      return !stillOpen;
    }
  
    // ---------------- Per-row processor (with retry on New Item) ----------------
    async function processRowOnce(row) {
      console.log('\n[Row] Processing selected row…');
      if (isNewItemFormOpen()) {
        const ok = await tryCloseNewItemForm();
        if (!ok) { console.warn('[Row] Could not close New item form at start.'); return false; }
      }
  
      const thumbCell = row.querySelector('[data-automationid="field-PosterThumbnail"]');
      const posterHref = row.querySelector('[data-automationid="field-Poster"] a[href]')?.getAttribute('href');
  
      if (!thumbCell) { console.warn('[Row] No PosterThumbnail cell.'); return true; }
      if (!posterHref) { console.warn('[Row] No Poster link; cannot make thumbnail.'); return true; }
      if (!isSharePoint(posterHref)) { console.warn('[Row] Poster is not SharePoint; skipping.'); return true; }
  
      console.log(`[Row] SharePoint poster link: ${posterHref}`);
  
      try { thumbCell.scrollIntoView({block:'center'}); } catch {}
      thumbCell.focus?.();
      await realDblClick(thumbCell);
      console.log('[Row] Double-clicked PosterThumbnail cell.');
  
      // If New Item pops right after double-click, close & fail this attempt so we can retry
      await wait(150);
      if (isNewItemFormOpen()) {
        const ok = await tryCloseNewItemForm();
        return false; // signal: retry row
      }
  
      // Find editor/callout and file input
      let callout = await waitFor(()=>document.querySelector('#commonEditorCalloutId, .ms-Callout, [role="dialog"], .ms-Dialog-main'), { timeout: 6000, interval: 80 });
      if (!callout) {
        console.warn('[Row] Editor container not found.');
        return false; // try again
      }
      const fileInput = document.querySelector('input[type="file"][accept*="image"], input[type="file"][name="thumbnailFiles"]');
      if (!fileInput) {
        console.warn('[Row] Thumbnail file input not found.');
        return false; // try again
      }
  
      // NEW: Warm up the original PDF before we hit thumbnails
      await warmUpSharePointPdf(posterHref, WARMUP_MS);
  
      console.log('[Row] Loading PDF thumbnail (with blank detection/retries)…');
      const img = await loadAnyCandidateWithRetries(posterHref);
      if (!img) { console.warn('[Row] Unable to obtain a non-blank thumbnail.'); return false; }
  
      const thumbBlob = await imageTo180ThumbBlob(img);
      if (!thumbBlob) { console.warn('[Row] Failed to create thumbnail blob.'); return false; }
  
      const filename = buildThumbFilename(row);
      const dt = new DataTransfer(); dt.items.add(new File([thumbBlob], filename, { type:'image/jpeg' }));
      fileInput.files = dt.files;
      console.log(`[Row] Assigned File to input. Filename=${filename}, Size=${dt.files[0].size} bytes`);
  
      fileInput.dispatchEvent(new Event('input', {bubbles:true}));
      fileInput.dispatchEvent(new Event('change',{bubbles:true}));
      fileInput.dispatchEvent(new CustomEvent('DataModifiedEvent',{bubbles:true,cancelable:true,detail:{field:'PosterThumbnail',filename}}));
      console.log('[Row] Change events dispatched; locating primary action…');
      await wait(150);
  
      // New Item could also open after change—guard again
      if (isNewItemFormOpen()) {
        const ok = await tryCloseNewItemForm();
        return false; // retry this row cleanly
      }
  
      callout = findActiveDialog();
      const clicked = await clickPrimaryAction();
      if (!clicked) {
        const pressed = await pressEnterOnDialog();
        if (!pressed) {
          console.warn('[Row] Could not activate primary action automatically.');
          return false; // retry
        }
      }
  
      await wait(300);
      console.log('[Row] Upload triggered.');
      return true;
    }
  
    async function processRowWithRetry(row, maxAttempts = 1) {
      for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        console.log(`[Row] Attempt ${attempt}/${maxAttempts}`);
        const ok = await processRowOnce(row);
        if (ok) return true;
        console.warn('[Row] Attempt did not complete; retrying after cleanup…');
        // extra cleanup: try to close any straggler New Item or dialogs
        await tryCloseNewItemForm();
        await wait(600);
      }
      console.error('[Row] Failed after max attempts; moving on.');
      return false;
    }
  
    async function runAllSelected(){
      const rows = Array.from(document.querySelectorAll('[role="row"][aria-selected="true"]'));
      if (!rows.length) { console.warn('No selected rows found.'); return; }
  
      for (const row of rows) {
        try {
          await processRowWithRetry(row, 1);
          await wait(700); // let dialogs settle/close before next
        } catch (e) {
          console.error('Error processing row:', e, row);
        }
      }
    }
  
    // ---------------- Floating button UI ----------------
    const btn = document.createElement('button');
    btn.id = 'cap-thumb-btn-v2p2';
    btn.textContent = 'Make Thumbnails';
    btn.title = 'Generate & upload thumbnails for selected rows (Alt+Shift+T)';
    Object.assign(btn.style, {
      position: 'fixed',
      right: '20px',
      bottom: '24px',
      zIndex: 2147483647,
      padding: '10px 14px',
      background: '#1f6feb',
      color: '#fff',
      border: 'none',
      borderRadius: '10px',
      boxShadow: '0 4px 12px rgba(0,0,0,.2)',
      font: '600 13px/1.2 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif',
      cursor: 'pointer'
    });
  
    btn.addEventListener('click', async () => {
      if (btn.dataset.busy === '1') return;
      btn.dataset.busy = '1';
      const old = btn.textContent;
      btn.textContent = 'Working…';
      btn.style.opacity = '0.8';
      try {
        await runAllSelected();
      } finally {
        btn.dataset.busy = '0';
        btn.textContent = old;
        btn.style.opacity = '1';
      }
    });
  
    // Right-click to remove the button
    btn.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      btn.remove();
      window.__capThumbBtnV2p2 = null;
      console.log('[CapThumb] Floating button removed.');
    });
  
    document.body.appendChild(btn);
    window.__capThumbBtnV2p2 = btn;
  
    // Keyboard shortcut: Alt+Shift+T
    window.addEventListener('keydown', async (e) => {
      if (e.altKey && e.shiftKey && (e.key === 'T' || e.key === 't')) {
        e.preventDefault();
        btn.click();
      }
    });
  
    console.log('[CapThumb] Floating button injected (v2.2). Warms original PDF via Shares API, then waits 5s.');
  })();
  
