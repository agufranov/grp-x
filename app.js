(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const E = { file:$('file'), wellButton:$('wellButton'), wellButtonName:$('wellButtonName'), wellChooser:$('wellChooser'), wellChooserClose:$('wellChooserClose'), wellOptions:$('wellOptions'), wellUpload:$('wellUpload'), testControl:$('testControl'), keyHint:$('keyHint'), work:$('workspace'), wellHead:$('wellHead'), stickySentinel:$('stickySentinel'), miniWell:$('miniWell'), select:$('select'), svg:$('diagram'), cep:$('cepstrum'), computed:$('computedCepstrum'), signal:$('signalGraph'), computedSignal:$('computedSignal'), chartMeta:$('chartMeta'), chartEmpty:$('chartEmpty'), computedMeta:$('computedMeta'), computedEmpty:$('computedEmpty'), signalMeta:$('signalMeta'), signalEmpty:$('signalEmpty'), computedSignalMeta:$('computedSignalMeta'), computedSignalEmpty:$('computedSignalEmpty'), signalFilterControls:$('signalFilterControls'), filterStart:$('filterStart'), filterEnd:$('filterEnd'), filterLow:$('filterLow'), filterHigh:$('filterHigh'), filterInterpolation:$('filterInterpolation'), filterStartOutput:$('filterStartOutput'), filterEndOutput:$('filterEndOutput'), filterLowOutput:$('filterLowOutput'), filterHighOutput:$('filterHighOutput'), filterAfterStop:$('filterAfterStop'), filterReset:$('filterReset'), simPort:$('simPort'), simNoise:$('simNoise'), simVelocity:$('simVelocity'), simAttenuation:$('simAttenuation'), simPulse:$('simPulse'), simSignal:$('simSignal'), simSignalMeta:$('simSignalMeta'), simSignalEmpty:$('simSignalEmpty'), simCepstrum:$('simCepstrum'), simCepstrumMeta:$('simCepstrumMeta'), simCepstrumEmpty:$('simCepstrumEmpty'), scaleTooltip:$('scaleTooltip'), rows:$('rows'), summary:$('summary'), toast:$('toast') };
  const WELL_GEOMETRY={height:155,top:20,bottom:175,axis:150,miniHeight:65,miniAxis:40,tickStart:125};
  const WELL_SCALE_DEFAULT=1,WELL_SCALE_MIN=.1,WELL_SCALE_MAX=32,WELL_PLOT_LEFT=62,WELL_PLOT_RIGHT=20;
  const CEPSTRUM_SCALE_DEFAULT=1,CEPSTRUM_SCALE_MIN=.1,CEPSTRUM_SCALE_MAX=32,CEPSTRUM_PLOT_LEFT=62,CEPSTRUM_PLOT_RIGHT=20;
  const INTERPOLATION=globalThis.GrpInterpolation,INTERPOLATION_DEFAULT=INTERPOLATION.DEFAULT_METHOD,INTERPOLATION_LABELS={pchip:'монотонный сплайн',linear:'линейная интерполяция'};
  const cepstrumScaleStates=new Map();
  const wellViewState={scale:WELL_SCALE_DEFAULT,offset:0,drag:null,frame:0};
  const signalAnalysisSettings=new Map();
  let currentSignalAnalysis=null,analysisFrame=0,drawRevision=0;
  let M,catalog=[],currentWellKey='';
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

  function setWellChooserLoading(active){E.wellChooser.classList.toggle('is-loading',active);E.wellChooser.setAttribute('aria-busy',String(active));E.wellChooser.querySelectorAll('button').forEach(button=>button.disabled=active);}
  function focusWellChooser(){if(E.wellChooser.hidden)return;const current=E.wellOptions.querySelector('[aria-current="true"]'),first=E.wellOptions.querySelector('button');(current||first||E.wellUpload).focus();}
  function openWellChooser(){E.wellChooser.hidden=false;document.body.classList.add('well-chooser-open');E.wellButton.setAttribute('aria-expanded','true');requestAnimationFrame(focusWellChooser);}
  function closeWellChooser(){if(!M||E.wellChooser.classList.contains('is-loading'))return;E.wellChooser.hidden=true;document.body.classList.remove('well-chooser-open');E.wellButton.setAttribute('aria-expanded','false');E.wellButton.focus({preventScroll:true});scheduleStickyState();}
  function completeWellSelection(label,key){currentWellKey=String(key);E.wellButtonName.textContent=label;E.wellButton.title='Выбрать другую скважину';E.wellChooserClose.hidden=false;E.wellOptions.querySelectorAll('[data-well-key]').forEach(button=>button.setAttribute('aria-current',String(button.dataset.wellKey===currentWellKey)));E.wellUpload.setAttribute('aria-current',String(currentWellKey==='upload'));setWellChooserLoading(false);closeWellChooser();}
  function renderWellOptions(){E.wellOptions.innerHTML=catalog.map((item,index)=>`<button class="well-choice" type="button" data-well-index="${index}" data-well-key="${esc(String(item.id??index))}" aria-current="${String(String(item.id??index)===currentWellKey)}"><span class="well-choice-icon" aria-hidden="true">№</span><span><strong>${esc(item.name)}</strong><small>${fmt(item.count)} испытаний</small></span><svg class="well-choice-arrow" viewBox="0 0 16 16" aria-hidden="true"><path d="M6 3l5 5-5 5"/></svg></button>`).join('');}

  async function load(file) {
    if (!file || !/\.xlsx?$/i.test(file.name)) return error('Выберите файл XLSX.');
    setWellChooserLoading(true);
    try {
      const wb=XLSX.read(await file.arrayBuffer(), {type:'array', dense:true});
      M=parse(wb,file);render();completeWellSelection(file.name.replace(/\.xlsx?$/i,''),'upload');
    } catch (e) { error(e.message || 'Не удалось прочитать файл.'); }
    finally { E.file.value='';setWellChooserLoading(false); }
  }

  function showWellDependentControls(){E.testControl.hidden=false;E.keyHint.hidden=false;}
  function render() {
    showWellDependentControls();
    E.select.innerHTML=M.tests.map((t,i)=>`<option value="${i}">${t.id} · ${esc(t.name)}</option>`).join('');
    E.select.disabled=false;
    E.rows.innerHTML=M.tests.map(t=>{const p=M.points.find(x=>x.port===t.ballPort);return `<tr><td>${t.id}</td><td>${esc(t.name)}</td><td>${p?'№ '+p.port:'—'}</td><td>${p?fmt(p.depth)+' м':'—'}</td></tr>`}).join('');
    E.summary.textContent=`· ${M.points.length} перфораций, ${M.tests.length} испытаний`;
    drawWell();
  }

  async function drawWell() {
    const revision=++drawRevision;
    const test=M.tests[+E.select.value||0], ball=M.points.find(p=>p.port===test.ballPort);
    drawWellDiagram(ball);drawMiniWell(ball);scheduleStickyState();const source=await drawCepstrum(test,revision);if(revision===drawRevision)drawSimulation(test,ball,source);
  }

  function wellHorizontalGeometry(width,minWidth,maxDepth){const w=Math.max(minWidth,width),L=WELL_PLOT_LEFT,R=w-WELL_PLOT_RIGHT,x=depth=>L+(depth/maxDepth-wellViewState.offset)*wellViewState.scale*(R-L);return{w,L,R,x,start:x(0),end:x(maxDepth)};}

  function drawWellDiagram(ball){
    const deepest=Math.max(...M.points.map(p=>p.depth),M.nkt),maxDepth=Math.ceil(deepest/500)*500,{w,L,R,x,start,end}=wellHorizontalGeometry(E.svg.clientWidth,320,maxDepth),h=WELL_GEOMETRY.height,y=WELL_GEOMETRY.axis,nx=x(M.nkt),bx=ball?x(ball.depth):end+1;
    let s=`<defs><clipPath id="wellPlotClip"><rect x="${L}" y="${WELL_GEOMETRY.top}" width="${R-L}" height="${h}"/></clipPath><linearGradient id="pipe" x1="0" x2="0" y1="0" y2="1"><stop stop-color="#dce8e3"/><stop offset=".5" stop-color="#51665e"/><stop offset="1" stop-color="#94a69f"/></linearGradient><pattern id="hatch" width="8" height="8" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><line y2="8" stroke="#fff" stroke-opacity=".08" stroke-width="3"/></pattern></defs><g clip-path="url(#wellPlotClip)">`;
    s+=`<rect x="${start}" y="136" width="${Math.max(0,end-start)}" height="28" rx="14" fill="url(#pipe)"/><rect x="${bx}" y="118" width="${Math.max(0,end-bx)}" height="64" fill="#10191d" opacity=".68"/><rect x="${bx}" y="118" width="${Math.max(0,end-bx)}" height="64" fill="url(#hatch)"/><line x1="${start}" y1="${y}" x2="${nx}" y2="${y}" stroke="#688df2" stroke-width="7"/><line x1="${nx}" y1="${WELL_GEOMETRY.tickStart}" x2="${nx}" y2="${WELL_GEOMETRY.bottom}" stroke="#86a3f5" stroke-width="4"/><circle cx="${nx}" cy="${y}" r="7" fill="#b9c9f7"/><text x="${nx}" y="103" text-anchor="middle" fill="#b9c9f7" font-size="13" font-weight="700">КОНЕЦ НКТ</text><text x="${nx}" y="119" text-anchor="middle" fill="#91a29c" font-size="11">${fmt(M.nkt)} м</text>`;
    M.points.forEach((p,i)=>{const px=x(p.depth),muted=ball&&p.depth>ball.depth,lower=i%2!==0,ly=lower?86:53,lend=lower?111:76,col=muted?'#68756f':'#c5e77b',tc=muted?'#7d8984':'#edf3ee';s+=`<line x1="${px}" y1="${WELL_GEOMETRY.tickStart}" x2="${px}" y2="${WELL_GEOMETRY.bottom}" stroke="${col}" stroke-width="3"/><circle cx="${px}" cy="${y}" r="8" fill="${col}" stroke="#29343a" stroke-width="3"/><line x1="${px}" y1="124" x2="${px}" y2="${lend}" stroke="#fff" stroke-opacity=".14"/><text x="${px}" y="${ly}" text-anchor="middle" fill="${tc}" font-size="13" font-weight="700">№ ${p.port}</text><text x="${px}" y="${ly+15}" text-anchor="middle" fill="#91a29c" font-size="10">${fmt(p.depth)} м</text>`;});
    if(ball)s+=`<circle cx="${bx}" cy="${y}" r="16" fill="#ef8058" stroke="#ffd0bf" stroke-width="3"/><circle cx="${bx-4}" cy="${y-4}" r="4" fill="#ffc2aa"/>`;
    E.svg.setAttribute('viewBox',`0 ${WELL_GEOMETRY.top} ${w} ${h}`);E.svg.style.minWidth='0';E.svg.innerHTML=s+'</g>';
  }

  function drawMiniWell(ball){
    const deepest=Math.max(...M.points.map(p=>p.depth),M.nkt),maxDepth=Math.ceil(deepest/500)*500,{w,L,R,x,start,end}=wellHorizontalGeometry(E.miniWell.clientWidth||E.wellHead.clientWidth,240,maxDepth),h=WELL_GEOMETRY.miniHeight,y=WELL_GEOMETRY.miniAxis,nktX=x(M.nkt),ballX=ball?x(ball.depth):end+1,shift=y-WELL_GEOMETRY.axis,sy=value=>value+shift;
    let markup=`<defs><clipPath id="miniWellPlotClip"><rect x="${L}" y="0" width="${R-L}" height="${h}"/></clipPath><linearGradient id="miniPipe" x1="0" x2="0" y1="0" y2="1"><stop stop-color="#dce8e3"/><stop offset=".5" stop-color="#51665e"/><stop offset="1" stop-color="#94a69f"/></linearGradient><pattern id="miniHatch" width="8" height="8" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><line y2="8" stroke="#fff" stroke-opacity=".08" stroke-width="3"/></pattern></defs><g clip-path="url(#miniWellPlotClip)">`;
    markup+=`<rect x="${start}" y="${sy(136)}" width="${Math.max(0,end-start)}" height="28" rx="14" fill="url(#miniPipe)"/><rect x="${ballX}" y="${sy(118)}" width="${Math.max(0,end-ballX)}" height="64" fill="#10191d" opacity=".68"/><rect x="${ballX}" y="${sy(118)}" width="${Math.max(0,end-ballX)}" height="64" fill="url(#miniHatch)"/><line x1="${start}" y1="${y}" x2="${nktX}" y2="${y}" stroke="#688df2" stroke-width="7"/><line x1="${nktX}" y1="${sy(WELL_GEOMETRY.tickStart)}" x2="${nktX}" y2="${h}" stroke="#86a3f5" stroke-width="4"/><circle cx="${nktX}" cy="${y}" r="7" fill="#b9c9f7"/>`;
    for(const p of M.points){const px=x(p.depth),muted=ball&&p.depth>ball.depth,color=muted?'#68756f':'#c5e77b';markup+=`<line x1="${px}" y1="${sy(WELL_GEOMETRY.tickStart)}" x2="${px}" y2="${h}" stroke="${color}" stroke-width="3"/><circle cx="${px}" cy="${y}" r="8" fill="${color}" stroke="#29343a" stroke-width="3"/>`;}
    if(ball)markup+=`<circle cx="${ballX}" cy="${y}" r="16" fill="#ef8058" stroke="#ffd0bf" stroke-width="3"/><circle cx="${ballX-4}" cy="${y-4}" r="4" fill="#ffc2aa"/>`;
    E.miniWell.setAttribute('viewBox',`0 0 ${w} ${h}`);E.miniWell.innerHTML=markup+'</g>';
  }

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
  function sortedSource(source){
    const sorted=source.map(p=>Array.isArray(p)?{x:num(p[0]),y:num(p[1])}:{x:num(p.x),y:num(p.y)}).filter(p=>isNum(p.x)&&isNum(p.y)).sort((a,b)=>a.x-b.x),unique=[];
    for(const point of sorted){const previous=unique.at(-1);if(previous&&previous.x===point.x)previous.y=point.y;else unique.push(point);}
    return unique;
  }
  function median(values){if(!values.length)return NaN;const ordered=[...values].sort((a,b)=>a-b);return ordered[Math.floor(ordered.length/2)];}
  function resampleSignal(source,start=-Infinity,end=Infinity,interpolation=INTERPOLATION_DEFAULT){
    const all=sortedSource(source),selected=all.filter(point=>point.x>=start&&point.x<=end);
    if(selected.length<4)return null;
    const diffs=[];for(let i=1;i<selected.length;i++)diffs.push(selected[i].x-selected[i-1].x);
    const xmin=selected[0].x,xmax=selected.at(-1).x,span=xmax-xmin;
    let step=median(diffs)||1;step=Math.max(step,span/16383);
    const count=Math.floor(span/step)+1;if(count<4)return null;
    const values=new Float64Array(count),interpolate=INTERPOLATION.createInterpolator(selected,interpolation);let cursor=0;
    for(let i=0;i<count;i++){const x=xmin+i*step;while(cursor+1<selected.length&&selected[cursor+1].x<x)cursor++;values[i]=interpolate(x,cursor);}
    return{xmin,xmax:xmin+(count-1)*step,step,values,sampleRate:1/step,nyquist:1/(2*step),points:selected,interpolation};
  }
  function linearTrend(values){
    const n=values.length,trend=new Float64Array(n);if(!n)return trend;
    let sumX=0,sumY=0,sumXX=0,sumXY=0;
    for(let i=0;i<n;i++){const x=n===1?0:i/(n-1);sumX+=x;sumY+=values[i];sumXX+=x*x;sumXY+=x*values[i];}
    const denominator=n*sumXX-sumX*sumX,slope=denominator?(n*sumXY-sumX*sumY)/denominator:0,intercept=(sumY-slope*sumX)/n;
    for(let i=0;i<n;i++)trend[i]=intercept+slope*(n===1?0:i/(n-1));return trend;
  }
  function zeroPhaseLowPass(values,cutoff,step){
    if(!(cutoff>0)||values.length<3)return Float64Array.from(values);
    const alpha=1-Math.exp(-2*Math.PI*cutoff*step),forward=new Float64Array(values.length),backward=new Float64Array(values.length);
    forward[0]=values[0];for(let i=1;i<values.length;i++)forward[i]=forward[i-1]+alpha*(values[i]-forward[i-1]);
    backward[values.length-1]=forward.at(-1);for(let i=values.length-2;i>=0;i--)backward[i]=backward[i+1]+alpha*(forward[i]-backward[i+1]);
    return backward;
  }
  function tukeyWindow(values,alpha=.12){
    const result=Float64Array.from(values),edge=alpha*(values.length-1)/2;if(!(edge>0))return result;
    for(let i=0;i<values.length;i++){const distance=Math.min(i,values.length-1-i),weight=distance>=edge?1:.5*(1-Math.cos(Math.PI*distance/edge));result[i]*=weight;}
    return result;
  }
  function cepstrumFromValues(values,step){
    if(values.length<4)return[];let n=1;while(n<values.length)n<<=1;
    const re=new Float64Array(n),im=new Float64Array(n),mean=values.reduce((sum,value)=>sum+value,0)/values.length;
    for(let i=0;i<values.length;i++)re[i]=values[i]-mean;fft(re,im);
    for(let i=0;i<n;i++){re[i]=Math.log(Math.max(1e-12,Math.hypot(re[i],im[i])));im[i]=0;}fft(re,im,true);
    const result=[];for(let i=0;i<Math.floor(n/2);i++)result.push({x:i*step,y:re[i]});return result;
  }
  function processSignal(source,{start=-Infinity,end=Infinity,low=0,high=Infinity,interpolation=INTERPOLATION_DEFAULT}={}){
    const grid=resampleSignal(source,start,end,interpolation);if(!grid)return null;
    const trend=linearTrend(grid.values),residual=new Float64Array(grid.values.length);
    for(let i=0;i<residual.length;i++)residual[i]=grid.values[i]-trend[i];
    let filtered=residual;
    if(high<grid.nyquist*.995)filtered=zeroPhaseLowPass(filtered,high,grid.step);
    if(low>0){const slow=zeroPhaseLowPass(filtered,low,grid.step),band=new Float64Array(filtered.length);for(let i=0;i<band.length;i++)band[i]=filtered[i]-slow[i];filtered=band;}
    const windowed=tukeyWindow(filtered),processed=Array.from(windowed,(value,i)=>[grid.xmin+i*grid.step,value]);
    return{grid,processed,windowed};
  }
  function analyzeSignal(source,settings={}){
    const processed=processSignal(source,settings);return processed?{grid:processed.grid,processed:processed.processed,cepstrum:cepstrumFromValues(processed.windowed,processed.grid.step)}:null;
  }
  function setGraphState(svg,empty,isEmpty,message){empty.textContent=message;empty.classList.toggle('hidden',!isEmpty);empty.parentElement.classList.toggle('empty-chart',isEmpty);if(isEmpty){svg.innerHTML='';const state=cepstrumScaleStates.get(svg);if(state){state.raw=[];state.options=null;}}}
  function wellObjectMarkerMarkup(svg,w,L,R,T,B,clipId){
    if(!M)return'';const test=M.tests[+E.select.value||0],ball=M.points.find(point=>point.port===test.ballPort),deepest=Math.max(M.nkt,...M.points.map(point=>point.depth)),maxDepth=Math.ceil(deepest/500)*500,wellGeometry=wellHorizontalGeometry(E.svg.clientWidth,320,maxDepth),wellRect=E.svg.getBoundingClientRect(),plotRect=svg.getBoundingClientRect(),screenProjection=wellRect.width>0&&plotRect.width>0;
    const project=depth=>{if(!screenProjection)return L+(depth/maxDepth-wellViewState.offset)*wellViewState.scale*(R-L);const screenX=wellRect.left+wellGeometry.x(depth)/wellGeometry.w*wellRect.width;return(screenX-plotRect.left)*w/plotRect.width;};
    let markup=`<g class="well-object-markers" clip-path="url(#${clipId})" pointer-events="none">`,nktX=project(M.nkt);
    markup+=`<g class="well-object-marker" data-kind="nkt" data-depth="${M.nkt}"><title>Конец НКТ · ${fmt(M.nkt)} м</title><line x1="${nktX}" y1="${T}" x2="${nktX}" y2="${B}" stroke="#688df2" stroke-opacity=".48" stroke-dasharray="3 3"/><text x="${nktX}" y="${T+9}" text-anchor="middle" fill="#526fc4" stroke="#fffdfa" stroke-width="2" paint-order="stroke" font-size="7" font-weight="800">НКТ</text></g>`;
    M.points.forEach((point,index)=>{const px=project(point.depth),muted=ball&&point.depth>ball.depth,color=muted?'#8a928e':'#76983f',opacity=muted?0.2:0.36,labelY=T+(index%2?18:9);markup+=`<g class="well-object-marker" data-kind="perforation" data-port="${point.port}" data-depth="${point.depth}"><title>Перфорация № ${point.port} · ${fmt(point.depth)} м</title><line x1="${px}" y1="${T}" x2="${px}" y2="${B}" stroke="${color}" stroke-opacity="${opacity}" stroke-dasharray="2 4"/><text x="${px}" y="${labelY}" text-anchor="middle" fill="${color}" stroke="#fffdfa" stroke-width="2" paint-order="stroke" font-size="7" font-weight="800">${point.port}</text></g>`;});
    if(ball){const ballX=project(ball.depth);markup+=`<g class="well-object-marker" data-kind="ball" data-port="${ball.port}" data-depth="${ball.depth}"><title>Шар · порт ${ball.port} · ${fmt(ball.depth)} м</title><line x1="${ballX}" y1="${T}" x2="${ballX}" y2="${B}" stroke="#ef8058" stroke-opacity=".62" stroke-width="1.3"/><circle cx="${ballX}" cy="${T+4}" r="3" fill="#ef8058" stroke="#fffdfa" stroke-width="1"/></g>`;}
    return markup+'</g>';
  }
  function drawPlot(svg,raw,{robust=false,color='#1e6c5b',xDomain=null,context=null,contextExclude=null,wellMarkers=false}={}){
    if(!raw.length)return;const scaleState=cepstrumScaleStates.get(svg),options={robust,color,xDomain,context,contextExclude,wellMarkers};if(scaleState){scaleState.raw=raw;scaleState.options=options;}
    const normalize=series=>(series||[]).map(p=>Array.isArray(p)?{x:num(p[0]),y:num(p[1])}:p).filter(p=>isNum(p.x)&&isNum(p.y)).sort((a,b)=>a.x-b.x),all=normalize(raw),contextAll=normalize(context);if(!all.length)return;
    const fixedXmin=num(xDomain?.[0]),fixedXmax=num(xDomain?.[1]),hasFixedXDomain=isNum(fixedXmin)&&isNum(fixedXmax)&&fixedXmax>fixedXmin,fullXmin=hasFixedXDomain?fixedXmin:all[0].x,fullXmax=hasFixedXDomain?fixedXmax:all.at(-1).x,scale=scaleState?.scale||CEPSTRUM_SCALE_DEFAULT,offset=scaleState?.offset||0,fullSpan=fullXmax-fullXmin||1,xmin=fullXmin+fullSpan*offset,xmax=xmin+fullSpan/scale;
    const visible=series=>{const lowerBound=value=>{let low=0,high=series.length;while(low<high){const middle=(low+high)>>1;if(series[middle].x<value)low=middle+1;else high=middle;}return low;},from=Math.max(0,lowerBound(xmin)-1),to=Math.min(series.length,lowerBound(xmax)+1);return thin(series.slice(from,Math.max(from+1,to)));},data=visible(all),contextData=contextAll.length?visible(contextAll):[];
    const scaleData=contextAll.length?all.concat(contextAll):all,ordered=scaleData.map(p=>p.y).sort((a,b)=>a-b),q=p=>ordered[Math.floor((ordered.length-1)*p)];let ymin=robust?q(.01):ordered[0],ymax=robust?q(.99):ordered.at(-1);if(ymax===ymin)ymax=ymin+1;const pad=(ymax-ymin)*.08;ymin-=pad;ymax+=pad;
    const compact=svg.classList.contains('signal-plot'),w=Math.max(320,svg.clientWidth||svg.parentElement.clientWidth),h=compact?104:280,L=CEPSTRUM_PLOT_LEFT,R=w-CEPSTRUM_PLOT_RIGHT,T=compact?8:17,B=h-(compact?24:34),x=value=>L+(value-xmin)/(xmax-xmin||1)*(R-L),y=value=>B-(Math.max(ymin,Math.min(ymax,value))-ymin)/(ymax-ymin)*(B-T),contextY=value=>B-(value-ymin)/(ymax-ymin)*(B-T),clipId=`${svg.id}PlotClip`,path=(series,projectY=y)=>series.map((point,i)=>`${i?'L':'M'}${x(point.x).toFixed(2)},${projectY(point.y).toFixed(2)}`).join(' '),xTicks=5,yTicks=compact?2:4,labelSize=compact?8:9;
    let markup=`<defs><clipPath id="${clipId}"><rect x="${L}" y="${T}" width="${R-L}" height="${B-T}"/></clipPath></defs>`;
    for(let i=0;i<=xTicks;i++){const value=xmin+(xmax-xmin)*i/xTicks,px=x(value);markup+=`<line x1="${px}" y1="${T}" x2="${px}" y2="${B}" stroke="#17231f" stroke-opacity=".07"/><text x="${px}" y="${h-(compact?7:12)}" text-anchor="middle" fill="#7c8781" font-size="${labelSize}">${fmt(value)}</text>`;}
    for(let i=0;i<=yTicks;i++){const value=ymin+(ymax-ymin)*i/yTicks,py=y(value);markup+=`<line x1="${L}" y1="${py}" x2="${R}" y2="${py}" stroke="#17231f" stroke-opacity=".07"/><text x="${L-9}" y="${py+3}" text-anchor="end" fill="#7c8781" font-size="${labelSize}">${fmt(value)}</text>`;}
    if(wellMarkers)markup+=wellObjectMarkerMarkup(svg,w,L,R,T,B,clipId);
    const excludeStart=num(contextExclude?.[0]),excludeEnd=num(contextExclude?.[1]),hasExclusion=isNum(excludeStart)&&isNum(excludeEnd)&&excludeEnd>excludeStart,contextSegments=hasExclusion?[contextData.filter(point=>point.x<=excludeStart),contextData.filter(point=>point.x>=excludeEnd)]:[contextData];
    for(const segment of contextSegments)if(segment.length>1)markup+=`<path class="plot-context" d="${path(segment,contextY)}" clip-path="url(#${clipId})" fill="none" stroke="#7f8984" stroke-opacity=".58" stroke-width="1.15" vector-effect="non-scaling-stroke"/>`;
    markup+=`<path class="plot-series" d="${path(data)}" clip-path="url(#${clipId})" fill="none" stroke="${color}" stroke-width="1.3" vector-effect="non-scaling-stroke"/>`;
    svg.setAttribute('viewBox',`0 0 ${w} ${h}`);svg.innerHTML=markup;
  }
  function updateCepstrumScaleHeader(svg,state){const card=svg.closest('.cepstrum-card'),scaleControl=card.querySelector('.chart-scale'),changed=state.scale!==CEPSTRUM_SCALE_DEFAULT||state.offset!==0;scaleControl.hidden=!changed;card.querySelector('.chart-scale-value').textContent=changed?`×${state.scale.toLocaleString('ru-RU',{maximumFractionDigits:2})}`:'';}
  function scheduleCepstrumScaleRedraw(svg,state){if(state.frame)return;state.frame=requestAnimationFrame(()=>{state.frame=0;if(state.raw.length&&state.options)drawPlot(svg,state.raw,state.options);});}
  function setCepstrumView(svg,state,{scale=state.scale,offset=state.offset}={}){if(!Number.isFinite(scale)||!Number.isFinite(offset))return;scale=Math.max(CEPSTRUM_SCALE_MIN,Math.min(CEPSTRUM_SCALE_MAX,scale));if(scale===state.scale&&offset===state.offset)return;state.scale=scale;state.offset=offset;updateCepstrumScaleHeader(svg,state);scheduleCepstrumScaleRedraw(svg,state);}
  function resetCepstrumView(svg,state){setCepstrumView(svg,state,{scale:CEPSTRUM_SCALE_DEFAULT,offset:0});}
  function horizontalPointerFraction(svg,event,left=CEPSTRUM_PLOT_LEFT,right=CEPSTRUM_PLOT_RIGHT){const matrix=svg.getScreenCTM();if(!matrix)return NaN;const point=svg.createSVGPoint();point.x=event.clientX;point.y=event.clientY;const x=point.matrixTransform(matrix.inverse()).x,width=svg.viewBox.baseVal.width;return(x-left)/(width-left-right);}
  function positionScaleTooltip(event){const x=Math.min(event.clientX+12,window.innerWidth-E.scaleTooltip.offsetWidth-4),y=Math.min(event.clientY+14,window.innerHeight-E.scaleTooltip.offsetHeight-4);E.scaleTooltip.style.transform=`translate3d(${Math.max(4,x)}px,${Math.max(4,y)}px,0)`;}
  function showScaleTooltip(event){if(event.pointerType==='touch')return;positionScaleTooltip(event);E.scaleTooltip.classList.add('visible');E.scaleTooltip.setAttribute('aria-hidden','false');}
  function hideScaleTooltip(){E.scaleTooltip.classList.remove('visible');E.scaleTooltip.setAttribute('aria-hidden','true');}
  function bindCepstrumScale(svg){const state={scale:CEPSTRUM_SCALE_DEFAULT,offset:0,raw:[],options:null,drag:null,frame:0};cepstrumScaleStates.set(svg,state);updateCepstrumScaleHeader(svg,state);svg.addEventListener('pointerenter',showScaleTooltip);svg.addEventListener('pointerdown',event=>{if(event.pointerType==='touch'||(event.button!==0&&event.button!==1))return;const pointerFraction=horizontalPointerFraction(svg,event),mode=event.button===0&&event.shiftKey?'scale':'pan';if(!Number.isFinite(pointerFraction)||(mode==='scale'&&!(pointerFraction>0)))return;event.preventDefault();showScaleTooltip(event);state.drag={pointerId:event.pointerId,mode,startFraction:pointerFraction,startScale:state.scale,startOffset:state.offset,anchorFraction:pointerFraction/state.scale};svg.setPointerCapture(event.pointerId);svg.classList.add(mode==='pan'?'is-panning':'is-scaling');document.documentElement.classList.add(mode==='pan'?'is-panning-cepstrum':'is-scaling-cepstrum');});svg.addEventListener('pointermove',event=>{showScaleTooltip(event);if(!state.drag||state.drag.pointerId!==event.pointerId)return;const pointerFraction=horizontalPointerFraction(svg,event);if(state.drag.mode==='pan')setCepstrumView(svg,state,{scale:state.drag.startScale,offset:state.drag.startOffset-(pointerFraction-state.drag.startFraction)/state.drag.startScale});else setCepstrumView(svg,state,{scale:pointerFraction/state.drag.anchorFraction,offset:state.drag.startOffset});});svg.addEventListener('pointerleave',()=>{if(!state.drag)hideScaleTooltip();});svg.addEventListener('contextmenu',event=>{event.preventDefault();resetCepstrumView(svg,state);});const finish=event=>{if(!state.drag||state.drag.pointerId!==event.pointerId)return;state.drag=null;svg.classList.remove('is-scaling','is-panning');document.documentElement.classList.remove('is-scaling-cepstrum','is-panning-cepstrum');if(svg.hasPointerCapture(event.pointerId))svg.releasePointerCapture(event.pointerId);if(!svg.matches(':hover'))hideScaleTooltip();};svg.addEventListener('pointerup',finish);svg.addEventListener('pointercancel',finish);svg.addEventListener('lostpointercapture',finish);}

  function scheduleWellViewRedraw(){if(wellViewState.frame)return;wellViewState.frame=requestAnimationFrame(()=>{wellViewState.frame=0;if(!M)return;const test=M.tests[+E.select.value||0],ball=M.points.find(p=>p.port===test.ballPort);drawWellDiagram(ball);drawMiniWell(ball);for(const [svg,state] of cepstrumScaleStates)if(state.raw.length&&state.options?.wellMarkers)drawPlot(svg,state.raw,state.options);scheduleStickyState();});}
  function setWellView({scale=wellViewState.scale,offset=wellViewState.offset}={}){if(!Number.isFinite(scale)||!Number.isFinite(offset))return;scale=Math.max(WELL_SCALE_MIN,Math.min(WELL_SCALE_MAX,scale));if(scale===wellViewState.scale&&offset===wellViewState.offset)return;wellViewState.scale=scale;wellViewState.offset=offset;scheduleWellViewRedraw();}
  function resetWellView(){setWellView({scale:WELL_SCALE_DEFAULT,offset:0});}
  function bindWellView(svg){svg.addEventListener('pointerenter',event=>{if(M)showScaleTooltip(event);});svg.addEventListener('pointerdown',event=>{if(!M||event.pointerType==='touch'||(event.button!==0&&event.button!==1))return;const pointerFraction=horizontalPointerFraction(svg,event,WELL_PLOT_LEFT,WELL_PLOT_RIGHT),mode=event.button===0&&event.shiftKey?'scale':'pan';if(!Number.isFinite(pointerFraction)||(mode==='scale'&&!(pointerFraction>0)))return;event.preventDefault();showScaleTooltip(event);wellViewState.drag={pointerId:event.pointerId,mode,startFraction:pointerFraction,startScale:wellViewState.scale,startOffset:wellViewState.offset,anchorFraction:pointerFraction/wellViewState.scale};svg.setPointerCapture(event.pointerId);svg.classList.add(mode==='pan'?'is-panning':'is-scaling');document.documentElement.classList.add(mode==='pan'?'is-panning-well':'is-scaling-well');});svg.addEventListener('pointermove',event=>{if(M)showScaleTooltip(event);const drag=wellViewState.drag;if(!drag||drag.pointerId!==event.pointerId)return;const pointerFraction=horizontalPointerFraction(svg,event,WELL_PLOT_LEFT,WELL_PLOT_RIGHT);if(drag.mode==='pan')setWellView({scale:drag.startScale,offset:drag.startOffset-(pointerFraction-drag.startFraction)/drag.startScale});else setWellView({scale:pointerFraction/drag.anchorFraction,offset:drag.startOffset});});svg.addEventListener('pointerleave',()=>{if(!wellViewState.drag)hideScaleTooltip();});svg.addEventListener('contextmenu',event=>{if(!M)return;event.preventDefault();resetWellView();});const finish=event=>{const drag=wellViewState.drag;if(!drag||drag.pointerId!==event.pointerId)return;wellViewState.drag=null;svg.classList.remove('is-scaling','is-panning');document.documentElement.classList.remove('is-scaling-well','is-panning-well');if(svg.hasPointerCapture(event.pointerId))svg.releasePointerCapture(event.pointerId);if(!svg.matches(':hover'))hideScaleTooltip();};svg.addEventListener('pointerup',finish);svg.addEventListener('pointercancel',finish);svg.addEventListener('lostpointercapture',finish);}
  function estimatePumpProfile(source){
    const grid=resampleSignal(source);if(!grid)return null;
    const smooth=zeroPhaseLowPass(grid.values,Math.min(.5,grid.nyquist*.25),grid.step),n=smooth.length,lag=Math.max(1,Math.round(.5/grid.step)),from=Math.max(lag,Math.floor(n*.08)),to=Math.min(n-lag-1,Math.floor(n*.75));
    let stopIndex=from,bestDrop=Infinity;for(let i=from;i<=to;i++){const change=smooth[i+lag]-smooth[i-lag];if(change<bestDrop){bestDrop=change;stopIndex=i;}}
    const preEnd=Math.max(lag,stopIndex-2*lag),preStart=Math.max(0,preEnd-Math.round(5/grid.step)),tailStart=Math.floor(n*.85),on=median(Array.from(smooth.slice(preStart,preEnd))),settled=median(Array.from(smooth.slice(tailStart))),range=Math.max(...smooth)-Math.min(...smooth),drop=on-settled;
    if(!(drop>range*.2))return null;
    const threshold=on-drop*.08,searchStart=Math.max(0,stopIndex-Math.round(5/grid.step));for(let i=searchStart;i<stopIndex;i++)if(smooth[i]<=threshold){stopIndex=i;break;}
    const minEnd=Math.min(n-1,stopIndex+Math.round(12/grid.step));let minIndex=stopIndex;for(let i=stopIndex+1;i<=minEnd;i++)if(smooth[i]<smooth[minIndex])minIndex=i;
    const reboundStart=Math.min(n-1,minIndex+Math.round(.5/grid.step)),reboundEnd=Math.min(n-1,minIndex+Math.round(14/grid.step));let reboundIndex=reboundStart;for(let i=reboundStart+1;i<=reboundEnd;i++)if(smooth[i]>smooth[reboundIndex])reboundIndex=i;
    return{start:grid.xmin,end:grid.xmax,stop:grid.xmin+stopIndex*grid.step,minimumTime:grid.xmin+minIndex*grid.step,reboundTime:grid.xmin+reboundIndex*grid.step,on,settled,minimum:smooth[minIndex],rebound:smooth[reboundIndex],step:grid.step};
  }
  function updateFilterOutputs(){
    E.filterStartOutput.value=`${fmt(E.filterStart.value)} с`;E.filterEndOutput.value=`${fmt(E.filterEnd.value)} с`;E.filterLowOutput.value=`${fmt(E.filterLow.value)} Гц`;E.filterHighOutput.value=`${fmt(E.filterHigh.value)} Гц`;
  }
  function syncFilterControls(){
    if(!currentSignalAnalysis)return;const {settings,bounds}=currentSignalAnalysis,timeStep=bounds.timeStep,freqStep=bounds.freqStep;
    for(const input of [E.filterStart,E.filterEnd]){input.min=String(bounds.xmin);input.max=String(bounds.xmax);input.step=String(timeStep);}
    for(const input of [E.filterLow,E.filterHigh]){input.min='0';input.max=String(bounds.nyquist);input.step=String(freqStep);}
    E.filterStart.value=String(settings.start);E.filterEnd.value=String(settings.end);E.filterLow.value=String(settings.low);E.filterHigh.value=String(settings.high);E.filterInterpolation.value=settings.interpolation;updateFilterOutputs();
  }
  function fullSignalContext(analysis){
    const {source,settings,bounds}=analysis,key=`${settings.low}:${settings.high}:${settings.interpolation}`;if(analysis.contextKey!==key){analysis.contextKey=key;analysis.context=processSignal(source,{start:bounds.xmin,end:bounds.xmax,low:settings.low,high:settings.high,interpolation:settings.interpolation})?.processed||[];}return analysis.context;
  }
  function renderSignalAnalysis(){
    analysisFrame=0;if(!currentSignalAnalysis)return;const {source,settings,bounds}=currentSignalAnalysis,result=analyzeSignal(source,settings),context=fullSignalContext(currentSignalAnalysis);
    if(!result){setGraphState(E.computed,E.computedEmpty,true,'Недостаточно данных в выбранном окне');E.computedMeta.textContent='недостаточно данных';setGraphState(E.computedSignal,E.computedSignalEmpty,true,'Недостаточно данных в выбранном окне');E.computedSignalMeta.textContent='недостаточно данных';return;}
    const computed=result.cepstrum,band=`${fmt(settings.low)}–${fmt(settings.high)} Гц`,windowText=`${fmt(result.grid.xmin)}–${fmt(result.grid.xmax)} с`,interpolation=INTERPOLATION_LABELS[settings.interpolation];
    setGraphState(E.computed,E.computedEmpty,!computed.length,'Недостаточно данных для расчёта кепстра');E.computedMeta.textContent=computed.length?`${computed.length.toLocaleString('ru-RU')} точек · ${windowText} · ${band} · ${interpolation}`:'нет исходного сигнала';if(computed.length)drawPlot(E.computed,computed,{robust:true,color:'#586acb',wellMarkers:true});
    setGraphState(E.computedSignal,E.computedSignalEmpty,false,'');E.computedSignalMeta.textContent=`${result.processed.length.toLocaleString('ru-RU')} отсчётов · ${windowText} · ${band} · ${interpolation}`;drawPlot(E.computedSignal,result.processed,{color:'#586acb',xDomain:[bounds.xmin,bounds.xmax],context,contextExclude:[result.grid.xmin,result.grid.xmax]});
    currentSignalAnalysis.result=result;
  }
  function scheduleSignalAnalysis(){if(!analysisFrame)analysisFrame=requestAnimationFrame(renderSignalAnalysis);}
  function configureSignalAnalysis(test,source){
    if(!source.length){currentSignalAnalysis=null;E.signalFilterControls.hidden=true;setGraphState(E.computed,E.computedEmpty,true,'Недостаточно данных для расчёта кепстра');E.computedMeta.textContent='нет исходного сигнала';setGraphState(E.computedSignal,E.computedSignalEmpty,true,'Недостаточно данных для расчёта кепстра');E.computedSignalMeta.textContent='нет исходного сигнала';setGraphState(E.signal,E.signalEmpty,true,'Для этого испытания исходный сигнал не найден');E.signalMeta.textContent='нет листа';return;}
    setGraphState(E.signal,E.signalEmpty,false,'');E.signalMeta.textContent=`${source.length.toLocaleString('ru-RU')} точек`;drawPlot(E.signal,source,{color:'#9b633c'});
    const grid=resampleSignal(source),profile=estimatePumpProfile(source);if(!grid){currentSignalAnalysis=null;E.signalFilterControls.hidden=true;setGraphState(E.computed,E.computedEmpty,true,'Недостаточно данных для расчёта кепстра');E.computedMeta.textContent='недостаточно данных';setGraphState(E.computedSignal,E.computedSignalEmpty,true,'Недостаточно данных для обработки сигнала');E.computedSignalMeta.textContent='недостаточно данных';return;}
    const key=`${M.file}:${test.id}`,duration=grid.xmax-grid.xmin,bounds={xmin:grid.xmin,xmax:grid.xmax,nyquist:grid.nyquist,timeStep:Math.max(grid.step,duration/1000),freqStep:Math.max(.01,grid.nyquist/200)};let settings=signalAnalysisSettings.get(key);
    if(!settings)settings={start:bounds.xmin,end:bounds.xmax,low:0,high:bounds.nyquist,interpolation:INTERPOLATION_DEFAULT};
    const minWindow=bounds.timeStep*8;settings.start=Math.max(bounds.xmin,Math.min(settings.start,bounds.xmax-minWindow));settings.end=Math.min(bounds.xmax,Math.max(settings.end,settings.start+minWindow));settings.low=Math.max(0,Math.min(settings.low,bounds.nyquist-bounds.freqStep));settings.high=Math.max(settings.low+bounds.freqStep,Math.min(settings.high,bounds.nyquist));if(!INTERPOLATION.SUPPORTED_METHODS.includes(settings.interpolation))settings.interpolation=INTERPOLATION_DEFAULT;signalAnalysisSettings.set(key,settings);
    currentSignalAnalysis={key,test,source,settings,bounds,profile,result:null};E.signalFilterControls.hidden=false;syncFilterControls();renderSignalAnalysis();
  }
  async function drawCepstrum(test,revision){
    let bundle;try{bundle=await cepstrumFor(test);}catch(e){error(e.message);bundle=null;}if(revision!==drawRevision)return null;
    const stored=bundle?.data||[],source=bundle?.source||[];
    setGraphState(E.cep,E.chartEmpty,!stored.length,'Для этого испытания лист с кепстром не найден');E.chartMeta.textContent=stored.length?`лист «${bundle.name}» · ${stored.length.toLocaleString('ru-RU')} точек · Y 1–99%`:'нет листа';if(stored.length)drawPlot(E.cep,stored,{robust:true,color:'#1e6c5b',wellMarkers:true});
    configureSignalAnalysis(test,source);return source;
  }

  function syncSimulationPorts(test,ball){const endDepth=ball?.depth??Math.max(...M.points.map(p=>p.depth)),available=M.points.filter(p=>p.depth<=endDepth).sort((a,b)=>a.port-b.port),key=`${M.file}:${test.id}:${ball?.port??'open'}`;if(E.simPort.dataset.key!==key){E.simPort.innerHTML=available.map(p=>`<option value="${p.port}">Порт ${p.port} · ${fmt(p.depth)} м</option>`).join('');const preferred=ball?.port??available.reduce((best,p)=>p.depth>best.depth?p:best,available[0]).port;E.simPort.value=String(preferred);E.simPort.dataset.key=key;}E.simPort.disabled=false;return available;}
  function seededRandom(seedText){let seed=2166136261;for(const ch of seedText){seed^=ch.charCodeAt(0);seed=Math.imul(seed,16777619);}return()=>{seed+=0x6D2B79F5;let t=seed;t=Math.imul(t^t>>>15,t|1);t^=t+Math.imul(t^t>>>7,t|61);return((t^t>>>14)>>>0)/4294967296;};}
  function addRicker(values,sampleRate,time,width,amplitude){const radius=Math.ceil(width*4*sampleRate),center=Math.round(time*sampleRate);for(let i=Math.max(0,center-radius);i<=Math.min(values.length-1,center+radius);i++){const z=Math.PI*((i/sampleRate)-time)/width,z2=z*z;values[i]+=amplitude*(1-2*z2)*Math.exp(-z2);}}
  function defaultPumpProfile(endDepth,velocity){const travel=2*endDepth/velocity,stop=12,minimumTime=stop+Math.max(2.5,travel*.45),reboundTime=minimumTime+Math.max(3,travel*.65);return{start:0,end:Math.max(48,reboundTime+24),stop,minimumTime,reboundTime,on:1,settled:.42,minimum:.23,rebound:.51,step:1/256};}
  function simulateSignal(test,ball,measuredSource=[]){
    const velocity=+E.simVelocity.value,noise=+E.simNoise.value/100,attenuation=.02+(+E.simAttenuation.value/100)*.22,width=+E.simPulse.value/1000,endDepth=ball?.depth??Math.max(...M.points.map(p=>p.depth)),ports=syncSimulationPorts(test,ball),activePort=+E.simPort.value,sampleRate=256,measuredProfile=estimatePumpProfile(measuredSource),profile=measuredProfile||defaultPumpProfile(endDepth,velocity),duration=Math.max(1,profile.end-profile.start),count=Math.ceil(duration*sampleRate)+1,values=new Float64Array(count),reflectors=[];
    const minimumTime=Math.max(profile.stop+.5,profile.minimumTime),reboundTime=Math.max(minimumTime+.5,profile.reboundTime),fastTau=Math.max(.7,(reboundTime-minimumTime)/2.4),slowTau=Math.max(8,(profile.end-reboundTime)/2.2),reboundExcess=Math.max(0,profile.rebound-profile.settled),slowAmplitude=reboundExcess*1.55,fastAmplitude=profile.settled+slowAmplitude-profile.minimum;
    for(let i=0;i<count;i++){
      const time=profile.start+i/sampleRate;
      if(time<=profile.stop)values[i]=profile.on;
      else if(time<minimumTime){const phase=(time-profile.stop)/(minimumTime-profile.stop),smooth=phase*phase*(3-2*phase);values[i]=profile.on+(profile.minimum-profile.on)*smooth;}
      else{const elapsed=time-minimumTime;values[i]=profile.settled+slowAmplitude*Math.exp(-elapsed/slowTau)-fastAmplitude*Math.exp(-elapsed/fastTau);}
    }
    reflectors.push({depth:0,amplitude:-1.25,delay:.08,type:'фронт остановки'});for(const point of ports)reflectors.push({depth:point.depth,amplitude:point.port===activePort?-.9:.14,type:point.port===activePort?'активная перфорация':'муфта'});if(M.nkt<=endDepth)reflectors.push({depth:M.nkt,amplitude:-.52,type:'конец НКТ'});if(ball)reflectors.push({depth:ball.depth,amplitude:1.15,type:'шар'});
    const pressureSpan=Math.max(1,Math.abs(profile.on-profile.minimum));for(const reflector of reflectors){const delay=reflector.delay??2*reflector.depth/velocity,pathKm=2*reflector.depth/1000,amplitude=reflector.amplitude*pressureSpan*.025*Math.exp(-attenuation*pathKm);addRicker(values,sampleRate,profile.stop-profile.start+delay,width,amplitude);}
    const random=seededRandom(`${M.file}:${test.id}:${activePort}`);let colored=0;for(let i=0;i<count;i++){const white=(random()+random()+random()+random()-2)*.5;colored=colored*.82+white*.18;const time=profile.start+i/sampleRate,pumpFactor=time<profile.stop?1:.45;values[i]+=noise*pressureSpan*.06*pumpFactor*(white*.65+colored*.35);}
    const source=Array.from(values,(value,i)=>[profile.start+i/sampleRate,value]);return{source,sampleRate,duration,reflectors:reflectors.length,activePort,endDepth,stopTime:profile.stop,endTime:profile.end,fitted:Boolean(measuredProfile)};
  }
  function drawSimulation(test,ball,measuredSource=[]){const simulation=simulateSignal(test,ball,measuredSource),analysis=analyzeSignal(simulation.source,{start:simulation.stopTime,end:simulation.endTime}),cepstrum=analysis?.cepstrum||[];setGraphState(E.simSignal,E.simSignalEmpty,false,'');setGraphState(E.simCepstrum,E.simCepstrumEmpty,!cepstrum.length,'Недостаточно данных для расчёта кепстра');E.simSignalMeta.textContent=`${simulation.source.length.toLocaleString('ru-RU')} отсчётов · остановка ${fmt(simulation.stopTime)} с · ${simulation.fitted?'профиль по данным':'типовой профиль'}`;E.simCepstrumMeta.textContent=`${cepstrum.length.toLocaleString('ru-RU')} точек · после остановки · порт ${simulation.activePort} · Y 1–99%`;drawPlot(E.simSignal,simulation.source,{color:'#c25a45'});if(cepstrum.length)drawPlot(E.simCepstrum,cepstrum,{robust:true,color:'#7256b5',wellMarkers:true});}

  async function loadDataset(item){
    if(!item)return;
    setWellChooserLoading(true);
    try{
      const [commonResponse,indexResponse]=await Promise.all([fetch(item.common),fetch(item.experiments)]);
      if(!commonResponse.ok||!indexResponse.ok)throw Error('Не удалось загрузить готовый набор данных.');
      const common=await commonResponse.json(),tests=await indexResponse.json();
      M={mode:'json',file:common.sourceFile||common.name,sheet:'Конструкция',nkt:common.tubingDepth,points:common.perforations,tests,experimentBase:item.experiments.replace(/\/index\.json$/,''),cepstrumCache:new Map()};
      render();completeWellSelection(item.name,item.id);
    }catch(e){error(e.message||'Не удалось загрузить готовый набор данных.');}
    finally{setWellChooserLoading(false);}
  }
  async function initCatalog(){
    try{
      const response=await fetch('data/catalog.json');if(!response.ok)throw Error('Каталог недоступен');
      catalog=await response.json();renderWellOptions();
    }catch(_){E.wellOptions.innerHTML='<p class="well-catalog-status">Встроенный каталог недоступен. Загрузите файл XLSX.</p>';}
    finally{if(!E.wellChooser.hidden)requestAnimationFrame(focusWellChooser);}
  }

  let timer;function error(m){E.toast.textContent=m;E.toast.classList.remove('hidden');clearTimeout(timer);timer=setTimeout(()=>E.toast.classList.add('hidden'),5000);}
  E.file.onchange=e=>load(e.target.files[0]);E.select.onchange=drawWell;
  E.wellButton.addEventListener('click',openWellChooser);
  E.wellChooserClose.addEventListener('click',closeWellChooser);
  E.wellChooser.addEventListener('click',event=>{if(event.target===E.wellChooser)closeWellChooser();});
  E.wellOptions.addEventListener('click',event=>{const button=event.target.closest('[data-well-index]');if(button)loadDataset(catalog[+button.dataset.wellIndex]);});
  E.wellUpload.addEventListener('click',()=>E.file.click());
  const setZoomModifier=active=>document.documentElement.classList.toggle('is-zoom-modifier',active);
  window.addEventListener('keydown',event=>{
    if(!E.wellChooser.hidden){
      if(event.key==='Escape'){event.preventDefault();closeWellChooser();}
      else if(event.key==='Tab'){
        const focusable=[...E.wellChooser.querySelectorAll('button:not([disabled])')].filter(element=>!element.hidden),first=focusable[0],last=focusable.at(-1);
        if(!first){event.preventDefault();return;}
        if(event.shiftKey&&(document.activeElement===first||!E.wellChooser.contains(document.activeElement))){event.preventDefault();last.focus();}
        else if(!event.shiftKey&&document.activeElement===last){event.preventDefault();first.focus();}
      }
      return;
    }
    if(event.key==='Shift')setZoomModifier(true);
    if(!M||E.select.disabled||(event.key!=='ArrowUp'&&event.key!=='ArrowDown')||event.altKey||event.ctrlKey||event.metaKey)return;
    event.preventDefault();const current=+E.select.value||0,next=Math.max(0,Math.min(M.tests.length-1,current+(event.key==='ArrowUp'?-1:1)));
    if(next!==current){E.select.value=String(next);drawWell();}
  });
  window.addEventListener('keyup',event=>{if(event.key==='Shift')setZoomModifier(false);});
  document.querySelectorAll('.cepstrum-scale-target').forEach(bindCepstrumScale);
  [E.svg,E.miniWell].forEach(bindWellView);
  document.addEventListener('click',event=>{const groupToggle=event.target.closest('.chart-group-toggle');if(groupToggle){const group=groupToggle.closest('.chart-group'),collapsed=group.classList.toggle('collapsed');groupToggle.setAttribute('aria-expanded',String(!collapsed));if(!collapsed&&M)requestAnimationFrame(()=>drawWell());return;}const reset=event.target.closest('.chart-scale-reset');if(reset){const svg=reset.closest('.cepstrum-card').querySelector('.cepstrum-scale-target'),state=cepstrumScaleStates.get(svg);resetCepstrumView(svg,state);return;}const head=event.target.closest('.chart-head');if(!head||event.target.closest('.chart-scale')||event.target.closest('.signal-filter-controls,.simulation-controls'))return;const card=head.closest('.cepstrum-card'),collapsed=card.classList.toggle('collapsed');head.querySelector('.chart-toggle').setAttribute('aria-expanded',String(!collapsed));if(!collapsed&&M)requestAnimationFrame(()=>drawWell());});
  const cursorGuide=$('cursorGuide');let guideFrame=0,guideX=-10;window.addEventListener('pointermove',event=>{if(event.pointerType==='touch')return;guideX=event.clientX;cursorGuide.classList.add('visible');if(!guideFrame)guideFrame=requestAnimationFrame(()=>{cursorGuide.style.transform=`translate3d(${guideX}px,0,0)`;guideFrame=0;});},{passive:true});document.documentElement.addEventListener('pointerleave',()=>{cursorGuide.classList.remove('visible');hideScaleTooltip();});window.addEventListener('blur',()=>{cursorGuide.classList.remove('visible');hideScaleTooltip();setZoomModifier(false);});
  function changeSignalFilter(input){
    if(!currentSignalAnalysis)return;const {settings,bounds}=currentSignalAnalysis,value=+input.value,minWindow=bounds.timeStep*8,minBand=bounds.freqStep;
    if(input===E.filterStart)settings.start=Math.max(bounds.xmin,Math.min(value,settings.end-minWindow));
    else if(input===E.filterEnd)settings.end=Math.min(bounds.xmax,Math.max(value,settings.start+minWindow));
    else if(input===E.filterLow)settings.low=Math.max(0,Math.min(value,settings.high-minBand));
    else if(input===E.filterHigh)settings.high=Math.min(bounds.nyquist,Math.max(value,settings.low+minBand));
    signalAnalysisSettings.set(currentSignalAnalysis.key,settings);syncFilterControls();scheduleSignalAnalysis();
  }
  [E.filterStart,E.filterEnd,E.filterLow,E.filterHigh].forEach(input=>input.addEventListener('input',()=>changeSignalFilter(input)));
  E.filterInterpolation.addEventListener('change',()=>{if(!currentSignalAnalysis)return;const {settings}=currentSignalAnalysis;settings.interpolation=INTERPOLATION.SUPPORTED_METHODS.includes(E.filterInterpolation.value)?E.filterInterpolation.value:INTERPOLATION_DEFAULT;signalAnalysisSettings.set(currentSignalAnalysis.key,settings);syncFilterControls();scheduleSignalAnalysis();});
  E.filterAfterStop.addEventListener('click',()=>{if(!currentSignalAnalysis)return;const {settings,bounds,profile}=currentSignalAnalysis;settings.start=Math.max(bounds.xmin,Math.min(profile?.stop??bounds.xmin,bounds.xmax-bounds.timeStep*8));settings.end=bounds.xmax;signalAnalysisSettings.set(currentSignalAnalysis.key,settings);syncFilterControls();scheduleSignalAnalysis();});
  E.filterReset.addEventListener('click',()=>{if(!currentSignalAnalysis)return;const {settings,bounds}=currentSignalAnalysis;Object.assign(settings,{start:bounds.xmin,end:bounds.xmax,low:0,high:bounds.nyquist,interpolation:INTERPOLATION_DEFAULT});signalAnalysisSettings.set(currentSignalAnalysis.key,settings);syncFilterControls();scheduleSignalAnalysis();});
  const redrawSimulation=()=>{if(!M)return;const test=M.tests[+E.select.value||0],ball=M.points.find(p=>p.port===test.ballPort),key=`${M.file}:${test.id}`,source=currentSignalAnalysis?.key===key?currentSignalAnalysis.source:[];drawSimulation(test,ball,source);};
  const simulationOutputs={simNoise:v=>`${v}%`,simVelocity:v=>`${v} м/с`,simAttenuation:v=>`${v}%`,simPulse:v=>`${v} мс`};
  [E.simNoise,E.simVelocity,E.simAttenuation,E.simPulse].forEach(input=>input.addEventListener('input',()=>{input.parentElement.querySelector('output').value=simulationOutputs[input.id](input.value);redrawSimulation();}));
  E.simPort.addEventListener('change',redrawSimulation);
  let stickyFrame=0;
  function updateStickyState(){
    stickyFrame=0;
    const headerStuck=E.stickySentinel.getBoundingClientRect().top<0;
    E.wellHead.classList.toggle('is-stuck',headerStuck);
    const headRect=E.wellHead.getBoundingClientRect(),diagramRect=E.svg.getBoundingClientRect();
    const diagramAxis=diagramRect.top+diagramRect.height*(WELL_GEOMETRY.axis-WELL_GEOMETRY.top)/WELL_GEOMETRY.height,stickyAxis=headRect.bottom+WELL_GEOMETRY.miniAxis;
    const wellStuck=headerStuck&&diagramAxis<=stickyAxis;
    E.wellHead.classList.toggle('is-well-stuck',wellStuck);
    cursorGuide.style.top=`${Math.max(0,Math.floor(wellStuck?headRect.bottom+WELL_GEOMETRY.miniHeight-2:diagramRect.bottom))}px`;
  }
  function scheduleStickyState(){if(!stickyFrame)stickyFrame=requestAnimationFrame(updateStickyState);}
  window.addEventListener('scroll',scheduleStickyState,{passive:true});
  scheduleStickyState();
  let resizeTimer;window.addEventListener('resize',()=>{clearTimeout(resizeTimer);resizeTimer=setTimeout(()=>{if(M)drawWell();scheduleStickyState();},100);});
  initCatalog();requestAnimationFrame(focusWellChooser);
})();
