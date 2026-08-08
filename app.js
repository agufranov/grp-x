(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const E = { file:$('file'), dataset:$('dataset'), work:$('workspace'), wellHead:$('wellHead'), stickySentinel:$('stickySentinel'), miniWell:$('miniWell'), select:$('select'), svg:$('diagram'), cep:$('cepstrum'), computed:$('computedCepstrum'), signal:$('signalGraph'), chartMeta:$('chartMeta'), chartEmpty:$('chartEmpty'), computedMeta:$('computedMeta'), computedEmpty:$('computedEmpty'), signalMeta:$('signalMeta'), signalEmpty:$('signalEmpty'), simPort:$('simPort'), simNoise:$('simNoise'), simVelocity:$('simVelocity'), simAttenuation:$('simAttenuation'), simPulse:$('simPulse'), simSignal:$('simSignal'), simSignalMeta:$('simSignalMeta'), simSignalEmpty:$('simSignalEmpty'), simCepstrum:$('simCepstrum'), simCepstrumMeta:$('simCepstrumMeta'), simCepstrumEmpty:$('simCepstrumEmpty'), rows:$('rows'), summary:$('summary'), toast:$('toast') };
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
    M.points.forEach((p,i)=>{const px=x(p.depth),muted=ball&&p.depth>ball.depth,lower=i%2!==0,ly=lower?82:47,lend=lower?105:70,col=muted?'#68756f':'#c5e77b',tc=muted?'#7d8984':'#edf3ee';s+=`<line x1="${px}" y1="127" x2="${px}" y2="173" stroke="${col}" stroke-width="3"/><circle cx="${px}" cy="${y}" r="8" fill="${col}" stroke="#29343a" stroke-width="3"/><line x1="${px}" y1="124" x2="${px}" y2="${lend}" stroke="#fff" stroke-opacity=".14"/><text x="${px}" y="${ly}" text-anchor="middle" fill="${tc}" font-size="13" font-weight="700">№ ${p.port}</text><text x="${px}" y="${ly+15}" text-anchor="middle" fill="#91a29c" font-size="10">${fmt(p.depth)} м</text>`;});
    if(ball){const side=bx+150<=R?1:-1,lineEnd=bx+side*125,textX=bx+side*8,anchor=side===1?'start':'end';s+=`<circle cx="${bx}" cy="${y}" r="16" fill="#ef8058" stroke="#ffd0bf" stroke-width="3"/><circle cx="${bx-4}" cy="${y-4}" r="4" fill="#ffc2aa"/><path d="M ${bx} ${y-16}V108H${lineEnd}" fill="none" stroke="#ef8058" stroke-width="2"/><text x="${textX}" y="100" text-anchor="${anchor}" fill="#ffb296" font-size="12" font-weight="800">ШАР - ПОРТ ${ball.port}</text>`;}
    E.svg.setAttribute('viewBox',`0 0 ${w} ${h}`); E.svg.style.minWidth='0'; E.svg.innerHTML=s;drawMiniWell(ball);drawSimulation(test,ball);scheduleStickyState();await drawCepstrum(test);
  }

  function drawMiniWell(ball){const deepest=Math.max(...M.points.map(p=>p.depth),M.nkt),maxDepth=Math.ceil(deepest/500)*500,w=Math.max(240,E.miniWell.clientWidth||E.wellHead.clientWidth),h=40,y=20,L=62,R=w-20,x=d=>L+d/maxDepth*(R-L),nktX=x(M.nkt),ballX=ball?x(ball.depth):R+1;let markup=`<rect x="${L}" y="16" width="${R-L}" height="8" rx="4" fill="#82908b"/><rect x="${ballX}" y="11" width="${Math.max(0,R-ballX)}" height="18" rx="2" fill="#121d21" opacity=".78"/><line x1="${L}" y1="${y}" x2="${nktX}" y2="${y}" stroke="#7597ef" stroke-width="3"/><line x1="${nktX}" y1="12" x2="${nktX}" y2="28" stroke="#9fb5f1" stroke-width="2"/>`;for(const p of M.points){const px=x(p.depth),muted=ball&&p.depth>ball.depth,color=muted?'#66736e':'#c5e77b';markup+=`<line x1="${px}" y1="10" x2="${px}" y2="30" stroke="${color}" stroke-width="2"/><circle cx="${px}" cy="${y}" r="3.5" fill="${color}" stroke="#29343a" stroke-width="1.5"/>`;}if(ball)markup+=`<circle cx="${ballX}" cy="${y}" r="7" fill="#ef8058" stroke="#ffd0bf" stroke-width="2"/>`;E.miniWell.setAttribute('viewBox',`0 0 ${w} ${h}`);E.miniWell.innerHTML=markup;}

  async function cepstrumFor(test) {
    if(M.cepstrumCache.has(test.id)) return M.cepstrumCache.get(test.id);
    if(M.mode==='json'){
      const response=await fetch(`${M.experimentBase}/${test.id}.json`);if(!response.ok)throw Error(`Не удалось загрузить испытание ${test.id}`);
      const payload=await response.json(),value={name:payload.cepstrumSheet,data:(payload.cepstrum||[]).map(p=>({x:p[0],y:p[1]})),source:payload.source||[]};M.cepstrumCache.set(test.id,value);return value;
    }
    const name=M.cepstrumSheets.get(`${test.id}c`),sourceName=String(test.id),data=[],source=[];
    if(name)for(const r of XLSX.utils.sheet_to_json(M.wb.Sheets[name],{header:1,defval:null,raw:true}))if(isNum(r[0])&&isNum(r[1]))data.push({x:num(r[0]),y:num(r[1])});
    if(M.wb.Sheets[sourceName])for(const r of XLSX.utils.sheet_to_json(M.wb.Sheets[sourceName],{header:1,defval:null,raw:true}))if(isNum(r[0])&&isNum(r[1]))source.push([num(r[0]),num(r[1])]);
    if(!data.length&&!source.length)return null;const value={name,data,source};M.cepstrumCache.set(test.id,value);return value;
  }
  function thin(data,limit=1800){if(data.length<=limit)return data;const out=[data[0]],buckets=Math.floor((limit-2)/2),step=(data.length-2)/buckets;for(let b=0;b<buckets;b++){const from=1+Math.floor(b*step),to=Math.min(data.length-1,1+Math.floor((b+1)*step));let low=data[from],high=data[from];for(let i=from+1;i<to;i++){if(data[i].y<low.y)low=data[i];if(data[i].y>high.y)high=data[i]}out.push(...(low.x<high.x?[low,high]:[high,low]));}out.push(data.at(-1));return out;}
  function fft(re,im,inverse=false){const n=re.length;for(let i=1,j=0;i<n;i++){let bit=n>>1;for(;j&bit;bit>>=1)j^=bit;j^=bit;if(i<j){[re[i],re[j]]=[re[j],re[i]];[im[i],im[j]]=[im[j],im[i]];}}for(let len=2;len<=n;len<<=1){const angle=(inverse?2:-2)*Math.PI/len,baseR=Math.cos(angle),baseI=Math.sin(angle);for(let i=0;i<n;i+=len){let wr=1,wi=0;for(let j=0;j<len/2;j++){const a=i+j,b=a+len/2,tr=re[b]*wr-im[b]*wi,ti=re[b]*wi+im[b]*wr;re[b]=re[a]-tr;im[b]=im[a]-ti;re[a]+=tr;im[a]+=ti;const next=wr*baseR-wi*baseI;wi=wr*baseI+wi*baseR;wr=next;}}}if(inverse)for(let i=0;i<n;i++){re[i]/=n;im[i]/=n;}}
  function calculateCepstrum(source){if(source.length<4)return[];const sorted=source.map(p=>({x:num(p[0]),y:num(p[1])})).filter(p=>isNum(p.x)&&isNum(p.y)).sort((a,b)=>a.x-b.x),diffs=[];for(let i=1;i<sorted.length;i++)if(sorted[i].x>sorted[i-1].x)diffs.push(sorted[i].x-sorted[i-1].x);diffs.sort((a,b)=>a-b);let step=diffs[Math.floor(diffs.length*.1)]||1,xmin=sorted[0].x,xmax=sorted.at(-1).x;step=Math.max(step,(xmax-xmin)/16383);const count=Math.floor((xmax-xmin)/step)+1,uniform=new Float64Array(count);let cursor=0;for(let i=0;i<count;i++){const x=xmin+i*step;while(cursor+1<sorted.length&&sorted[cursor+1].x<x)cursor++;const a=sorted[cursor],b=sorted[Math.min(cursor+1,sorted.length-1)],t=b.x===a.x?0:(x-a.x)/(b.x-a.x);uniform[i]=a.y+t*(b.y-a.y);}let n=1;while(n<count)n<<=1;const re=new Float64Array(n),im=new Float64Array(n),mean=uniform.reduce((sum,v)=>sum+v,0)/count;for(let i=0;i<count;i++)re[i]=uniform[i]-mean;fft(re,im);for(let i=0;i<n;i++){re[i]=Math.log(Math.max(1e-12,Math.hypot(re[i],im[i])));im[i]=0;}fft(re,im,true);const result=[];for(let i=0;i<Math.floor(n/2);i++)result.push({x:i*step,y:re[i]});return result;}
  function setGraphState(svg,empty,isEmpty,message){empty.textContent=message;empty.classList.toggle('hidden',!isEmpty);empty.parentElement.classList.toggle('empty-chart',isEmpty);if(isEmpty)svg.innerHTML='';}
  function drawPlot(svg,raw,{robust=false,color='#1e6c5b'}={}){if(!raw.length)return;const all=raw.map(p=>Array.isArray(p)?{x:num(p[0]),y:num(p[1])}:p).filter(p=>isNum(p.x)&&isNum(p.y)),data=thin(all),xmin=Math.min(...all.map(p=>p.x)),xmax=Math.max(...all.map(p=>p.x)),ys=all.map(p=>p.y),ordered=[...ys].sort((a,b)=>a-b),q=p=>ordered[Math.floor((ordered.length-1)*p)];let ymin=robust?q(.01):ordered[0],ymax=robust?q(.99):ordered.at(-1);if(ymax===ymin)ymax=ymin+1;const pad=(ymax-ymin)*.08;ymin-=pad;ymax+=pad;const w=Math.max(320,svg.clientWidth||svg.parentElement.clientWidth),h=280,L=62,R=w-20,T=17,B=h-34,x=v=>L+(v-xmin)/(xmax-xmin||1)*(R-L),y=v=>B-(Math.max(ymin,Math.min(ymax,v))-ymin)/(ymax-ymin)*(B-T);let markup='';for(let i=0;i<=5;i++){const v=xmin+(xmax-xmin)*i/5,px=x(v);markup+=`<line x1="${px}" y1="${T}" x2="${px}" y2="${B}" stroke="#17231f" stroke-opacity=".07"/><text x="${px}" y="${h-12}" text-anchor="middle" fill="#7c8781" font-size="9">${fmt(v)}</text>`;}for(let i=0;i<=4;i++){const v=ymin+(ymax-ymin)*i/4,py=y(v);markup+=`<line x1="${L}" y1="${py}" x2="${R}" y2="${py}" stroke="#17231f" stroke-opacity=".07"/><text x="${L-9}" y="${py+3}" text-anchor="end" fill="#7c8781" font-size="9">${fmt(v)}</text>`;}markup+=`<path d="${data.map((p,i)=>`${i?'L':'M'}${x(p.x).toFixed(2)},${y(p.y).toFixed(2)}`).join(' ')}" fill="none" stroke="${color}" stroke-width="1.3" vector-effect="non-scaling-stroke"/>`;svg.setAttribute('viewBox',`0 0 ${w} ${h}`);svg.innerHTML=markup;}
  async function drawCepstrum(test){
    let bundle;try{bundle=await cepstrumFor(test);}catch(e){error(e.message);bundle=null;}
    const stored=bundle?.data||[],source=bundle?.source||[],computed=calculateCepstrum(source);
    setGraphState(E.cep,E.chartEmpty,!stored.length,'Для этого испытания лист с кепстром не найден');E.chartMeta.textContent=stored.length?`лист «${bundle.name}» · ${stored.length.toLocaleString('ru-RU')} точек · Y 1–99%`:'нет листа';if(stored.length)drawPlot(E.cep,stored,{robust:true,color:'#1e6c5b'});
    setGraphState(E.computed,E.computedEmpty,!computed.length,'Недостаточно данных для расчёта кепстра');E.computedMeta.textContent=computed.length?`${computed.length.toLocaleString('ru-RU')} точек · IFFT(log|FFT|) · Y 1–99%`:'нет исходного сигнала';if(computed.length)drawPlot(E.computed,computed,{robust:true,color:'#586acb'});
    setGraphState(E.signal,E.signalEmpty,!source.length,'Для этого испытания исходный сигнал не найден');E.signalMeta.textContent=source.length?`${source.length.toLocaleString('ru-RU')} точек`:'нет листа';if(source.length)drawPlot(E.signal,source,{color:'#9b633c'});
  }

  function syncSimulationPorts(test,ball){const endDepth=ball?.depth??Math.max(...M.points.map(p=>p.depth)),available=M.points.filter(p=>p.depth<=endDepth).sort((a,b)=>a.port-b.port),key=`${M.file}:${test.id}:${ball?.port??'open'}`;if(E.simPort.dataset.key!==key){E.simPort.innerHTML=available.map(p=>`<option value="${p.port}">Порт ${p.port} · ${fmt(p.depth)} м</option>`).join('');const preferred=ball?.port??available.reduce((best,p)=>p.depth>best.depth?p:best,available[0]).port;E.simPort.value=String(preferred);E.simPort.dataset.key=key;}E.simPort.disabled=false;return available;}
  function seededRandom(seedText){let seed=2166136261;for(const ch of seedText){seed^=ch.charCodeAt(0);seed=Math.imul(seed,16777619);}return()=>{seed+=0x6D2B79F5;let t=seed;t=Math.imul(t^t>>>15,t|1);t^=t+Math.imul(t^t>>>7,t|61);return((t^t>>>14)>>>0)/4294967296;};}
  function addRicker(values,sampleRate,time,width,amplitude){const radius=Math.ceil(width*4*sampleRate),center=Math.round(time*sampleRate);for(let i=Math.max(0,center-radius);i<=Math.min(values.length-1,center+radius);i++){const z=Math.PI*((i/sampleRate)-time)/width,z2=z*z;values[i]+=amplitude*(1-2*z2)*Math.exp(-z2);}}
  function simulateSignal(test,ball){const velocity=+E.simVelocity.value,noise=+E.simNoise.value/100,attenuation=.02+(+E.simAttenuation.value/100)*.22,width=+E.simPulse.value/1000,endDepth=ball?.depth??Math.max(...M.points.map(p=>p.depth)),ports=syncSimulationPorts(test,ball),activePort=+E.simPort.value,sampleRate=512,duration=2*endDepth/velocity+.6,count=Math.ceil(duration*sampleRate),values=new Float64Array(count),reflectors=[];reflectors.push({depth:0,amplitude:1.25,time:.08,type:'устье'});for(const p of ports)reflectors.push({depth:p.depth,amplitude:p.port===activePort?-.9:.14,type:p.port===activePort?'активная перфорация':'муфта'});if(M.nkt<=endDepth)reflectors.push({depth:M.nkt,amplitude:-.52,type:'конец НКТ'});if(ball)reflectors.push({depth:ball.depth,amplitude:1.15,type:'шар'});for(const r of reflectors){const time=r.time??2*r.depth/velocity,pathKm=2*r.depth/1000,amplitude=r.amplitude*Math.exp(-attenuation*pathKm);addRicker(values,sampleRate,time,width,amplitude);}const random=seededRandom(`${M.file}:${test.id}:${activePort}`);let colored=0;for(let i=0;i<count;i++){const white=(random()+random()+random()+random()-2)*.5;colored=colored*.82+white*.18;values[i]+=noise*(white*.09+colored*.11);}const source=Array.from(values,(y,i)=>[i/sampleRate,y]);return{source,sampleRate,duration,reflectors:reflectors.length,activePort,endDepth};}
  function drawSimulation(test,ball){const simulation=simulateSignal(test,ball),cepstrum=calculateCepstrum(simulation.source);setGraphState(E.simSignal,E.simSignalEmpty,false,'');setGraphState(E.simCepstrum,E.simCepstrumEmpty,false,'');E.simSignalMeta.textContent=`${simulation.source.length.toLocaleString('ru-RU')} отсчётов · ${simulation.sampleRate} Гц · ${simulation.reflectors} отражателей`;E.simCepstrumMeta.textContent=`${cepstrum.length.toLocaleString('ru-RU')} точек · порт ${simulation.activePort} · Y 1–99%`;drawPlot(E.simSignal,simulation.source,{color:'#c25a45'});drawPlot(E.simCepstrum,cepstrum,{robust:true,color:'#7256b5'});}

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
  document.addEventListener('click',event=>{const head=event.target.closest('.chart-head');if(!head)return;const card=head.closest('.cepstrum-card');card.classList.toggle('collapsed');if(!card.classList.contains('collapsed')&&M)requestAnimationFrame(()=>drawWell());});
  const cursorGuide=$('cursorGuide');let guideFrame=0,guideX=-10;window.addEventListener('pointermove',event=>{if(event.pointerType==='touch')return;guideX=event.clientX;cursorGuide.classList.add('visible');if(!guideFrame)guideFrame=requestAnimationFrame(()=>{cursorGuide.style.transform=`translate3d(${guideX}px,0,0)`;guideFrame=0;});},{passive:true});document.documentElement.addEventListener('pointerleave',()=>cursorGuide.classList.remove('visible'));window.addEventListener('blur',()=>cursorGuide.classList.remove('visible'));
  const redrawSimulation=()=>{if(!M)return;const test=M.tests[+E.select.value||0],ball=M.points.find(p=>p.port===test.ballPort);drawSimulation(test,ball);};
  const simulationOutputs={simNoise:v=>`${v}%`,simVelocity:v=>`${v} м/с`,simAttenuation:v=>`${v}%`,simPulse:v=>`${v} мс`};
  [E.simNoise,E.simVelocity,E.simAttenuation,E.simPulse].forEach(input=>input.addEventListener('input',()=>{input.parentElement.querySelector('output').value=simulationOutputs[input.id](input.value);redrawSimulation();}));
  E.simPort.addEventListener('change',redrawSimulation);
  let stickyFrame=0;
  function updateStickyState(){
    stickyFrame=0;
    const headerStuck=E.stickySentinel.getBoundingClientRect().top<0;
    E.wellHead.classList.toggle('is-stuck',headerStuck);
    const headRect=E.wellHead.getBoundingClientRect(),diagramRect=E.svg.getBoundingClientRect();
    const wellStuck=headerStuck&&diagramRect.top+diagramRect.height/2<=headRect.bottom+20;
    E.wellHead.classList.toggle('is-well-stuck',wellStuck);
  }
  function scheduleStickyState(){if(!stickyFrame)stickyFrame=requestAnimationFrame(updateStickyState);}
  window.addEventListener('scroll',scheduleStickyState,{passive:true});
  scheduleStickyState();
  let resizeTimer;window.addEventListener('resize',()=>{clearTimeout(resizeTimer);resizeTimer=setTimeout(()=>{if(M)drawWell();scheduleStickyState();},100);});
  initCatalog();
})();
