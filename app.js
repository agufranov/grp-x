(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const E = { file:$('file'), dataset:$('dataset'), work:$('workspace'), select:$('select'), svg:$('diagram'), cep:$('cepstrum'), chartTitle:$('chartTitle'), chartMeta:$('chartMeta'), chartEmpty:$('chartEmpty'), rows:$('rows'), summary:$('summary'), toast:$('toast') };
  let M;
  const txt = v => String(v ?? '').trim();
  const num = v => v == null || txt(v) === '' ? NaN : (typeof v === 'number' ? v : Number(txt(v).replace(',', '.')));
  const isNum = v => Number.isFinite(num(v));
  const fmt = n => Number(n).toLocaleString('ru-RU', { maximumFractionDigits: 2 });
  const esc = v => txt(v).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

  function parse(wb, file) {
    const sheet = wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header:1, defval:null, raw:true });
    const hi = rows.findIndex(r => txt(r[0]).toLowerCase() === 'порт' && txt(r[1]).toLowerCase().includes('глуб'));
    if (hi < 0) throw Error('Не найдены столбцы «Порт» и «Глубина».');
    const header = rows[hi], nktIndex = header.findIndex(v => txt(v).toLowerCase() === 'нкт');
    let nkt = NaN;
    for (let c=nktIndex+1; c<=Math.min(nktIndex+3, header.length-1); c++) if (isNum(header[c])) { nkt=num(header[c]); break; }
    if (!isNum(nkt)) throw Error('Не удалось определить положение конца НКТ.');
    const points=[];
    for (let r=hi+1; r<rows.length; r++) {
      if (isNum(rows[r][0]) && isNum(rows[r][1])) points.push({port:num(rows[r][0]), depth:num(rows[r][1])});
      else if (points.length) break;
    }
    if (!points.length) throw Error('Не найдены точки перфорации.');
    const ports=new Set(points.map(p=>p.port)), tests=[];
    for (let r=hi+points.length+1; r<rows.length; r++) {
      const id=rows[r][0], name=rows[r][1];
      if (!isNum(id) || typeof name !== 'string' || !txt(name)) continue;
      const clean=txt(name), match=clean.match(/^\s*(\d+)/) || clean.match(/(\d+)\s*$/), candidate=match ? +match[1] : null;
      tests.push({id:num(id), name:clean, ballPort:ports.has(candidate) ? candidate : null});
    }
    if (!tests.length) throw Error('Не найдена таблица испытаний.');
    const cepstrumSheets=new Map();
    wb.SheetNames.forEach(name => {
      const normalized=txt(name).toLowerCase().replace(/с$/, 'c');
      if (/^\d+c$/.test(normalized)) cepstrumSheets.set(normalized, name);
    });
    return {mode:'xlsx',file:file.name, size:file.size, sheet, nkt, points, tests, wb, cepstrumSheets, cepstrumCache:new Map()};
  }

  async function load(file) {
    if (!file || !/\.xlsx?$/i.test(file.name)) return error('Выберите файл XLSX.');
    try {
      const wb=XLSX.read(await file.arrayBuffer(), {type:'array', dense:true});
      M=parse(wb,file); E.dataset.value=''; render();
    } catch (e) { error(e.message || 'Не удалось прочитать файл.'); }
    finally { E.file.value=''; }
  }

  function render() {
    E.select.innerHTML=M.tests.map((t,i)=>`<option value="${i}">${t.id} · ${esc(t.name)}</option>`).join('');
    E.select.disabled=false;
    E.rows.innerHTML=M.tests.map(t=>{const p=M.points.find(x=>x.port===t.ballPort);return `<tr><td>${t.id}</td><td>${esc(t.name)}</td><td>${p?'№ '+p.port:'—'}</td><td>${p?fmt(p.depth)+' м':'—'}</td></tr>`}).join('');
    E.summary.textContent=`· ${M.points.length} перфораций, ${M.tests.length} испытаний`;
    drawWell();
  }

  async function drawWell() {
    const test=M.tests[+E.select.value||0], ball=M.points.find(p=>p.port===test.ballPort);
    const deepest=Math.max(...M.points.map(p=>p.depth),M.nkt), lo=0, hi=Math.ceil(deepest/500)*500;
    const w=Math.max(320,E.svg.clientWidth), h=300, y=150, L=62, R=w-20, x=d=>L+d/hi*(R-L), nx=x(M.nkt), bx=ball?x(ball.depth):R+1;
    let s='<defs><linearGradient id="pipe" x1="0" x2="0" y1="0" y2="1"><stop stop-color="#dce8e3"/><stop offset=".5" stop-color="#51665e"/><stop offset="1" stop-color="#94a69f"/></linearGradient><pattern id="hatch" width="8" height="8" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><line y2="8" stroke="#fff" stroke-opacity=".08" stroke-width="3"/></pattern></defs>';
    s+=`<rect x="${L}" y="136" width="${R-L}" height="28" rx="14" fill="url(#pipe)"/><rect x="${bx}" y="118" width="${Math.max(0,R-bx)}" height="64" fill="#10191d" opacity=".68"/><rect x="${bx}" y="118" width="${Math.max(0,R-bx)}" height="64" fill="url(#hatch)"/><line x1="${L}" y1="${y}" x2="${nx}" y2="${y}" stroke="#688df2" stroke-width="7"/><line x1="${nx}" y1="125" x2="${nx}" y2="175" stroke="#86a3f5" stroke-width="4"/><circle cx="${nx}" cy="${y}" r="7" fill="#b9c9f7"/><text x="${nx}" y="103" text-anchor="middle" fill="#b9c9f7" font-size="13" font-weight="700">КОНЕЦ НКТ</text><text x="${nx}" y="119" text-anchor="middle" fill="#91a29c" font-size="11">${fmt(M.nkt)} м</text>`;
    M.points.forEach((p,i)=>{const px=x(p.depth),muted=ball&&p.depth>ball.depth,up=i%2===0,ly=up?55:245,lend=up?95:205,col=muted?'#68756f':'#c5e77b',tc=muted?'#7d8984':'#edf3ee';s+=`<line x1="${px}" y1="127" x2="${px}" y2="173" stroke="${col}" stroke-width="3"/><circle cx="${px}" cy="${y}" r="8" fill="${col}" stroke="#29343a" stroke-width="3"/><line x1="${px}" y1="${up?124:176}" x2="${px}" y2="${lend}" stroke="#fff" stroke-opacity=".14"/><text x="${px}" y="${ly}" text-anchor="middle" fill="${tc}" font-size="13" font-weight="700">№ ${p.port}</text><text x="${px}" y="${ly+15}" text-anchor="middle" fill="#91a29c" font-size="10">${fmt(p.depth)} м</text>`;});
    if(ball){const side=bx+150<=R?1:-1,lineEnd=bx+side*125,textX=bx+side*8,anchor=side===1?'start':'end';s+=`<circle cx="${bx}" cy="${y}" r="16" fill="#ef8058" stroke="#ffd0bf" stroke-width="3"/><circle cx="${bx-4}" cy="${y-4}" r="4" fill="#ffc2aa"/><path d="M ${bx} ${y-16}V108H${lineEnd}" fill="none" stroke="#ef8058" stroke-width="2"/><text x="${textX}" y="100" text-anchor="${anchor}" fill="#ffb296" font-size="12" font-weight="800">ШАР - ПОРТ ${ball.port}</text>`;}
    E.svg.setAttribute('viewBox',`0 0 ${w} ${h}`); E.svg.style.minWidth='0'; E.svg.innerHTML=s; await drawCepstrum(test);
  }

  async function cepstrumFor(test) {
    if(M.cepstrumCache.has(test.id)) return M.cepstrumCache.get(test.id);
    if(M.mode==='json'){
      const response=await fetch(`${M.experimentBase}/${test.id}.json`);if(!response.ok)throw Error(`Не удалось загрузить испытание ${test.id}`);
      const payload=await response.json(),value={name:payload.cepstrumSheet,data:(payload.cepstrum||[]).map(p=>({x:p[0],y:p[1]})),source:payload.source||[]};M.cepstrumCache.set(test.id,value);return value;
    }
    const name=M.cepstrumSheets.get(`${test.id}c`); if(!name)return null;
    const rows=XLSX.utils.sheet_to_json(M.wb.Sheets[name],{header:1,defval:null,raw:true}),data=[];
    for(const r of rows)if(isNum(r[0])&&isNum(r[1]))data.push({x:num(r[0]),y:num(r[1])});
    const value={name,data}; M.cepstrumCache.set(test.id,value); return value;
  }
  function thin(data,limit=1600){if(data.length<=limit)return data;const out=[data[0]],buckets=Math.floor((limit-2)/2),step=(data.length-2)/buckets;for(let b=0;b<buckets;b++){const from=1+Math.floor(b*step),to=Math.min(data.length-1,1+Math.floor((b+1)*step));let low=data[from],high=data[from];for(let i=from+1;i<to;i++){if(data[i].y<low.y)low=data[i];if(data[i].y>high.y)high=data[i]}out.push(...(low.x<high.x?[low,high]:[high,low]));}out.push(data.at(-1));return out;}
  async function drawCepstrum(test){
    let cep;try{cep=await cepstrumFor(test);}catch(e){error(e.message);cep=null;} E.chartTitle.textContent=`${test.id} · ${test.name}`;
    if(!cep||!cep.data.length){E.chartMeta.textContent='нет листа';E.chartEmpty.classList.remove('hidden');E.chartEmpty.parentElement.classList.add('empty-chart');return;}
    E.chartEmpty.classList.add('hidden');E.chartEmpty.parentElement.classList.remove('empty-chart');
    const data=thin(cep.data),xmin=cep.data[0].x,xmax=cep.data.at(-1).x,ys=cep.data.map(p=>p.y),sorted=[...ys].sort((a,b)=>a-b),quantile=p=>sorted[Math.floor((sorted.length-1)*p)];
    let ymin=quantile(.01),ymax=quantile(.99);if(ymax===ymin){ymin=Math.min(...ys);ymax=Math.max(...ys)||ymin+1;}const pad=(ymax-ymin)*.08;ymin-=pad;ymax+=pad;
    E.chartMeta.textContent=`лист «${cep.name}» · ${cep.data.length.toLocaleString('ru-RU')} точек · Y 1–99%`;
    const w=Math.max(320,E.cep.clientWidth),h=280,L=62,R=w-20,T=17,B=h-34,x=v=>L+(v-xmin)/(xmax-xmin)*(R-L),y=v=>B-(Math.max(ymin,Math.min(ymax,v))-ymin)/(ymax-ymin)*(B-T);let s='';
    for(let i=0;i<=5;i++){const v=xmin+(xmax-xmin)*i/5,px=x(v);s+=`<line x1="${px}" y1="${T}" x2="${px}" y2="${B}" stroke="#17231f" stroke-opacity=".07"/><text x="${px}" y="${h-12}" text-anchor="middle" fill="#7c8781" font-size="9">${fmt(v)}</text>`;}
    for(let i=0;i<=4;i++){const v=ymin+(ymax-ymin)*i/4,py=y(v);s+=`<line x1="${L}" y1="${py}" x2="${R}" y2="${py}" stroke="#17231f" stroke-opacity=".07"/><text x="${L-9}" y="${py+3}" text-anchor="end" fill="#7c8781" font-size="9">${fmt(v)}</text>`;}
    const path=data.map((p,i)=>`${i?'L':'M'}${x(p.x).toFixed(2)},${y(p.y).toFixed(2)}`).join(' ');s+=`<path d="${path}" fill="none" stroke="#1e6c5b" stroke-width="1.3" vector-effect="non-scaling-stroke"/>`;
    E.cep.setAttribute('viewBox',`0 0 ${w} ${h}`);E.cep.innerHTML=s;
  }

  async function loadDataset(item){
    if(!item)return;
    try{
      const [commonResponse,indexResponse]=await Promise.all([fetch(item.common),fetch(item.experiments)]);
      if(!commonResponse.ok||!indexResponse.ok)throw Error('Не удалось загрузить готовый набор данных.');
      const common=await commonResponse.json(),tests=await indexResponse.json();
      M={mode:'json',file:common.sourceFile||common.name,sheet:'Конструкция',nkt:common.tubingDepth,points:common.perforations,tests,experimentBase:item.experiments.replace(/\/index\.json$/,''),cepstrumCache:new Map()};
      render();
    }catch(e){error(e.message||'Не удалось загрузить готовый набор данных.');}
  }
  async function initCatalog(){
    try{
      const response=await fetch('data/catalog.json');if(!response.ok)return;
      const catalog=await response.json();
      E.dataset.innerHTML='<option value="">Выберите скважину…</option>'+catalog.map((x,i)=>`<option value="${i}">${esc(x.name)} · ${x.count} испытаний</option>`).join('')+'<option value="upload">＋ Загрузить XLSX…</option>';
      E.dataset.onchange=()=>{if(E.dataset.value==='upload'){E.dataset.value='';E.file.click();}else if(E.dataset.value!=='')loadDataset(catalog[+E.dataset.value]);};
    }catch(_){/* При file:// остаётся обычная загрузка XLSX. */}
  }

  let timer;function error(m){E.toast.textContent=m;E.toast.classList.remove('hidden');clearTimeout(timer);timer=setTimeout(()=>E.toast.classList.add('hidden'),5000);}
  E.file.onchange=e=>load(e.target.files[0]);E.select.onchange=drawWell;
  E.dataset.onchange=()=>{if(E.dataset.value==='upload'){E.dataset.value='';E.file.click();}};
  window.addEventListener('keydown',event=>{
    if(!M||E.select.disabled||(event.key!=='ArrowUp'&&event.key!=='ArrowDown')||event.altKey||event.ctrlKey||event.metaKey)return;
    event.preventDefault();const current=+E.select.value||0,next=Math.max(0,Math.min(M.tests.length-1,current+(event.key==='ArrowUp'?-1:1)));
    if(next!==current){E.select.value=String(next);drawWell();}
  });
  let resizeTimer;window.addEventListener('resize',()=>{clearTimeout(resizeTimer);resizeTimer=setTimeout(()=>M&&drawWell(),100);});
  initCatalog();
})();
