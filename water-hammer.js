(() => {
  'use strict';

  const $ = id => document.getElementById(id);
  const E = {
    open: $('hammerHelp'), modal: $('hammerLab'), panel: document.querySelector('#hammerLab .hammer-panel'), close: $('hammerClose'),
    pipe: $('hammerPipe'), graph: $('hammerGraph'), progress: $('hammerProgress'), play: $('hammerPlay'), replay: $('hammerReplay'), trend: $('hammerTrend'),
    step: $('hammerStep'), time: $('hammerTime'), verdict: $('hammerVerdict'), verdictText: $('hammerVerdictText'), graphHint: $('hammerGraphHint'),
    causeTitle: $('hammerCauseTitle'), causeText: $('hammerCauseText'), waveTitle: $('hammerWaveTitle'), waveText: $('hammerWaveText')
  };
  if (!E.open || !E.modal) return;

  const MODES = {
    pump: {
      event: 'Насос остановлен', wave: 'волна разрежения', sign: '−', color: '#5479b6', soft: '#dce6f5', origin: 'left',
      lead: 'Насос больше не добавляет напор, а жидкость по инерции продолжает двигаться. От устья вниз идёт фронт пониженного давления.',
      verdictBefore: 'Высокое давление накачки', verdictBeforeText: 'Насос добавляет напор, который покрывает давление пласта и потери на движение жидкости.',
      verdictEvent: 'Резкое падение — это ожидаемо', verdictEventText: 'Из показания исчезает давление, которое тратилось на проталкивание жидкости. Волна разрежения усиливает первый ход вниз.',
      verdictEcho: 'Эхо сидит на падающем уровне', verdictEchoText: 'Отражения от неоднородностей возвращаются малыми колебаниями. Они не обязаны превратить весь график в пик вверх.',
      causeTitle: 'Пропал напор насоса', causeText: 'Жидкость по инерции ещё уходит в скважину.', waveTitle: 'Пошёл «минусовой» фронт', waveText: 'За ним давление ниже, чем перед ним.'
    },
    valve: {
      event: 'Задвижка закрыта', wave: 'волна сжатия', sign: '+', color: '#d46849', soft: '#f4ded7', origin: 'right',
      lead: 'Задвижка резко тормозит поток. Жидкость набегает на закрытую границу, сжимается, и от неё идёт фронт повышенного давления.',
      verdictBefore: 'Установившийся поток', verdictBeforeText: 'До закрытия жидкость движется, а давление почти постоянно.',
      verdictEvent: 'Вот знакомый пик вверх', verdictEventText: 'У закрытой задвижки импульс жидкости превращается в сжатие. Это другой граничный сценарий того же гидроудара.',
      verdictEcho: 'Пик затухает отражениями', verdictEchoText: 'Волна бегает между границами, меняя амплитуду и иногда знак.',
      causeTitle: 'Поток упёрся в задвижку', causeText: 'Скорость у закрытой границы резко стала нулевой.', waveTitle: 'Пошёл «плюсовой» фронт', waveText: 'За ним давление выше, чем перед ним.'
    }
  };
  let mode = 'pump', frame = 0, playing = false, startedAt = 0, startProgress = 0, returnFocus = null;
  const reducedMotion = matchMedia('(prefers-reduced-motion: reduce)').matches;
  const clamp = (value, low, high) => Math.max(low, Math.min(high, value));
  const smoothstep = value => { const x = clamp(value, 0, 1); return x * x * (3 - 2 * x); };
  const graphX = time => 58 + (time + 1) / 7 * 820;
  const pressure = (time, selectedMode = mode) => {
    if (time < 0) return { trend: selectedMode === 'pump' ? 8.15 : 4.8, wave: 0 };
    if (selectedMode === 'pump') {
      const trend = 3.2 + 4.95 * Math.exp(-time / .58);
      const firstDrop = -.62 * Math.exp(-Math.pow((time - .13) / .105, 2));
      const ringing = .34 * Math.exp(-time / 2.25) * Math.sin(2 * Math.PI * (time - .12) / 1.08);
      return { trend, wave: firstDrop + ringing };
    }
    const spike = 2.65 * Math.exp(-Math.pow((time - .14) / .12, 2));
    const ringing = 1.15 * Math.exp(-time / 2.2) * Math.sin(2 * Math.PI * (time + .02) / 1.08);
    return { trend: 4.8, wave: spike + ringing };
  };
  const pathFor = (accessor, yMap) => {
    let path = '';
    for (let i = 0; i <= 280; i++) {
      const time = -1 + i / 40, value = accessor(time);
      path += `${i ? 'L' : 'M'}${graphX(time).toFixed(2)} ${yMap(value).toFixed(2)}`;
    }
    return path;
  };

  function phaseInfo(time) {
    const config = MODES[mode];
    if (time < 0) return { stage: 'before', label: mode === 'pump' ? 'Насос работает' : 'Поток идёт через задвижку', front: null, direction: config.origin === 'left' ? 1 : -1 };
    if (time < 2.2) return { stage: 'out', label: `${config.event} · ${config.wave} идёт ${config.origin === 'left' ? 'вниз' : 'к насосу'}`, front: time / 2.2, direction: config.origin === 'left' ? 1 : -1 };
    if (time < 4.4) return { stage: 'back', label: 'Отражение возвращается к датчику', front: 1 - (time - 2.2) / 2.2, direction: config.origin === 'left' ? -1 : 1 };
    return { stage: 'settle', label: 'Колебания затухают', front: null, direction: 0 };
  }

  function renderPipe(time) {
    const config = MODES[mode], phase = phaseInfo(time), left = 82, right = 818, width = right - left, incidentFromLeft = config.origin === 'left';
    let changedStart = left, changedWidth = 0, frontX = null;
    if (phase.stage === 'out') {
      frontX = incidentFromLeft ? left + phase.front * width : right - phase.front * width;
      changedStart = incidentFromLeft ? left : frontX;
      changedWidth = incidentFromLeft ? frontX - left : right - frontX;
    } else if (phase.stage === 'back' || phase.stage === 'settle') {
      changedWidth = width;
      if (phase.stage === 'back') frontX = incidentFromLeft ? left + phase.front * width : right - phase.front * width;
    }
    const flowOpacity = time < 0 ? 1 : Math.max(.08, 1 - time / 1.6), arrowDirection = mode === 'pump' ? 1 : 1;
    let arrows = '';
    for (let x = 190; x <= 700; x += 128) {
      const shift = ((Math.max(time, 0) * 45) % 44) * arrowDirection;
      arrows += `<path d="M${x + shift} 107h34m-8-7 8 7-8 7" opacity="${flowOpacity.toFixed(2)}"/>`;
    }
    const front = frontX == null ? '' : `<g class="hammer-svg-front"><rect x="${frontX - 13}" y="70" width="26" height="75" rx="13" fill="${config.color}" opacity=".18"/><line x1="${frontX}" y1="65" x2="${frontX}" y2="151" stroke="${config.color}" stroke-width="4"/><path d="M${frontX - phase.direction * 12} 59l${phase.direction * 12} 7-12 7" fill="none" stroke="${config.color}" stroke-width="3"/><text x="${clamp(frontX, 155, 745)}" y="48" text-anchor="middle" fill="${config.color}" font-size="12" font-weight="800">${config.sign} ${config.wave.toUpperCase()}</text></g>`;
    const originLabel = mode === 'pump' ? 'НАСОС · ДАТЧИК' : 'НАСОС · ДАТЧИК', endLabel = mode === 'pump' ? 'ПЕРФОРАЦИИ · ПЛАСТ' : 'ЗАКРЫТАЯ ЗАДВИЖКА';
    const pumpState = mode === 'pump' && time >= 0 ? 'ВЫКЛ' : 'РАБОТАЕТ';
    const endState = mode === 'valve' && time >= 0 ? 'ЗАКРЫТА' : mode === 'valve' ? 'ОТКРЫТА' : 'ПРИНИМАЕТ ПОТОК';
    E.pipe.innerHTML = `<defs><linearGradient id="hammerPipeBody" x1="0" x2="0" y1="0" y2="1"><stop stop-color="#edf1ed"/><stop offset=".48" stop-color="#b9c6c0"/><stop offset=".52" stop-color="#8c9b94"/><stop offset="1" stop-color="#d8dfdb"/></linearGradient><clipPath id="hammerWaterClip"><rect x="${left}" y="83" width="${width}" height="48" rx="24"/></clipPath></defs>
      <text x="${left}" y="23" fill="#66736d" font-size="10" font-weight="800" letter-spacing="1.1">${originLabel}</text><text x="${right}" y="23" text-anchor="end" fill="#66736d" font-size="10" font-weight="800" letter-spacing="1.1">${endLabel}</text>
      <rect x="${left - 7}" y="76" width="${width + 14}" height="62" rx="31" fill="url(#hammerPipeBody)"/><rect x="${left}" y="83" width="${width}" height="48" rx="24" fill="#6f9f97"/>
      <rect x="${changedStart}" y="83" width="${Math.max(0, changedWidth)}" height="48" fill="${config.color}" clip-path="url(#hammerWaterClip)" opacity="${phase.stage === 'settle' ? (mode === 'pump' ? .52 : Math.max(0, .52 * (1 - (time - 4.4) / 1.6))).toFixed(2) : .88}"/>
      <g fill="none" stroke="#fff" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" clip-path="url(#hammerWaterClip)">${arrows}</g>${front}
      <g transform="translate(42 107)"><circle r="28" fill="#263d34"/><circle r="18" fill="#eef2ee"/><path d="M0-13C8-8 8-2 2 0C8 3 7 10 0 13C-4 6-4 2 0 0C-7-2-8-8 0-13z" fill="${pumpState === 'ВЫКЛ' ? '#89968f' : '#91bd58'}"/><circle r="3" fill="#263d34"/></g>
      <g transform="translate(858 107)">${mode === 'valve' ? `<path d="M-14-15L14 15M14-15L-14 15" stroke="${time >= 0 ? '#d46849' : '#667d73'}" stroke-width="7" stroke-linecap="round"/>` : '<path d="M-14-18v36m12-29v22m12-16v10" stroke="#91bd58" stroke-width="5" stroke-linecap="round"/>'}</g>
      <text x="42" y="153" text-anchor="middle" fill="#6d7973" font-size="9" font-weight="800">${pumpState}</text><text x="858" y="153" text-anchor="middle" fill="#6d7973" font-size="9" font-weight="800">${endState}</text>
      <text x="${left}" y="185" fill="#7b8580" font-size="10">устье</text><text x="${right}" y="185" text-anchor="end" fill="#7b8580" font-size="10">глубина →</text>`;
  }

  function renderGraph(time) {
    const isolate = E.trend.checked, config = MODES[mode];
    const yMin = isolate ? -1.6 : 1.8, yMax = isolate ? 1.6 : 9.0, y = value => 218 - (value - yMin) / (yMax - yMin) * 174;
    const total = t => { const p = pressure(t); return isolate ? p.wave : p.trend + p.wave; };
    const mainPath = pathFor(total, y), trendPath = isolate ? '' : pathFor(t => pressure(t).trend, y), current = pressure(time), currentValue = isolate ? current.wave : current.trend + current.wave, cx = graphX(time), cy = y(currentValue);
    E.graph.innerHTML = `<line x1="${graphX(0)}" y1="36" x2="${graphX(0)}" y2="218" stroke="#d6a755" stroke-width="1.5" stroke-dasharray="4 4"/><text x="${graphX(0) + 6}" y="31" fill="#9a742c" font-size="9" font-weight="800">СОБЫТИЕ</text>
      ${trendPath ? `<path d="${trendPath}" fill="none" stroke="#b5ada0" stroke-width="2" stroke-dasharray="6 5"/><text x="760" y="${y(pressure(5).trend) - 9}" fill="#91897d" font-size="9">общий уровень</text>` : `<line x1="58" y1="${y(0)}" x2="878" y2="${y(0)}" stroke="#aab5af" stroke-width="1.5" stroke-dasharray="5 4"/>`}
      <path d="${mainPath}" fill="none" stroke="${config.color}" stroke-width="3.4" stroke-linecap="round" stroke-linejoin="round"/><line x1="${cx}" y1="28" x2="${cx}" y2="218" stroke="#273d34" stroke-width="1" opacity=".35"/><circle cx="${cx}" cy="${cy}" r="6" fill="#fff" stroke="${config.color}" stroke-width="3"/><text x="${isolate ? 18 : 16}" y="29" fill="#66736d" font-size="9" font-weight="800">${isolate ? 'Δp' : 'p'}</text><text x="878" y="252" text-anchor="end" fill="#7b8580" font-size="9">время, с</text>`;
  }

  function render() {
    const progress = +E.progress.value / 1000, time = -1 + progress * 7, phase = phaseInfo(time), config = MODES[mode];
    E.step.textContent = phase.label;
    E.time.value = `t = ${time < 0 ? '−' + Math.abs(time).toFixed(1) : '+' + time.toFixed(1)} с`;
    if (phase.stage === 'before') { E.verdict.textContent = config.verdictBefore; E.verdictText.textContent = config.verdictBeforeText; }
    else if (phase.stage === 'out') { E.verdict.textContent = config.verdictEvent; E.verdictText.textContent = config.verdictEventText; }
    else { E.verdict.textContent = config.verdictEcho; E.verdictText.textContent = config.verdictEchoText; }
    renderPipe(time); renderGraph(time);
  }

  function setPlaying(active) {
    playing = active;
    E.play.classList.toggle('is-playing', active);
    E.play.querySelector('span').textContent = active ? 'Пауза' : 'Пуск';
    E.play.setAttribute('aria-label', active ? 'Приостановить анимацию' : 'Запустить анимацию');
    E.play.querySelector('path').setAttribute('d', active ? 'M5 4h3v10H5zm5 0h3v10h-3z' : 'M6 4l8 5-8 5z');
    if (frame) cancelAnimationFrame(frame);
    frame = 0;
    if (active) { startProgress = +E.progress.value; startedAt = performance.now(); frame = requestAnimationFrame(tick); }
  }
  function tick(now) {
    if (!playing) return;
    const next = startProgress + (now - startedAt) / 7.2;
    E.progress.value = String(Math.min(1000, next)); render();
    if (next >= 1000) setPlaying(false); else frame = requestAnimationFrame(tick);
  }
  function restart(autoplay = true) { E.progress.value = '0'; render(); setPlaying(autoplay && !reducedMotion); }
  function selectMode(nextMode) {
    if (!MODES[nextMode]) return;
    mode = nextMode;
    E.modal.querySelectorAll('[data-hammer-mode]').forEach(button => button.setAttribute('aria-pressed', String(button.dataset.hammerMode === mode)));
    const config = MODES[mode];
    E.causeTitle.textContent = config.causeTitle; E.causeText.textContent = config.causeText; E.waveTitle.textContent = config.waveTitle; E.waveText.textContent = config.waveText;
    E.graphHint.textContent = mode === 'pump' ? 'Общее давление: большой спад + малая рябь' : 'Общее давление: положительный импульс + отражения';
    E.trend.parentElement.querySelector('span').textContent = mode === 'pump' ? 'Убрать общий спад' : 'Показать только волну';
    restart(true);
  }
  function open() {
    returnFocus = document.activeElement;
    E.modal.hidden = false; document.body.classList.add('water-hammer-open');
    E.open.setAttribute('aria-expanded', 'true'); E.close.focus({ preventScroll: true }); restart(true);
  }
  function close() {
    if (E.modal.hidden) return;
    setPlaying(false); E.modal.hidden = true; document.body.classList.remove('water-hammer-open'); E.open.setAttribute('aria-expanded', 'false');
    if (returnFocus && typeof returnFocus.focus === 'function') returnFocus.focus({ preventScroll: true });
  }

  E.open.setAttribute('aria-expanded', 'false');
  E.open.addEventListener('click', open); E.close.addEventListener('click', close); E.modal.addEventListener('click', event => { if (event.target === E.modal) close(); });
  E.modal.querySelectorAll('[data-hammer-mode]').forEach(button => button.addEventListener('click', () => selectMode(button.dataset.hammerMode)));
  E.play.addEventListener('click', () => { if (!playing && +E.progress.value >= 1000) E.progress.value = '0'; setPlaying(!playing); });
  E.replay.addEventListener('click', () => restart(true)); E.progress.addEventListener('input', () => { setPlaying(false); render(); }); E.trend.addEventListener('change', render);
  window.addEventListener('keydown', event => {
    if (E.modal.hidden) return;
    event.stopImmediatePropagation();
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key === ' ') { event.preventDefault(); E.play.click(); return; }
    if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') { event.preventDefault(); setPlaying(false); E.progress.value = String(clamp(+E.progress.value + (event.key === 'ArrowRight' ? 25 : -25), 0, 1000)); render(); return; }
    if (event.key === 'Tab') {
      const focusable = [...E.modal.querySelectorAll('button,input')].filter(element => !element.disabled && !element.hidden), first = focusable[0], last = focusable.at(-1);
      if (event.shiftKey && (document.activeElement === first || !E.modal.contains(document.activeElement))) { event.preventDefault(); last.focus(); }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
    }
  }, true);
  render();
})();
