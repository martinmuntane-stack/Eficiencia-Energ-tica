/* Pestaña "Electricidad-Facturación": consumo eléctrico facturado de SET 01 y SET 05.
 * Lee los consumos de window.LINEA_BASE (js/linea-base-facturacion.js), pasajeros de DATOS_CONSUMOS
 * y grados-día de DATOS_EJEMPLO. Mismo estilo que el tablero de electricidad (app.js). */
(function () {
  'use strict';
  if (!window.LINEA_BASE) return;

  const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  const MES_CORTO = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const TOTAL = '__total';
  const COLORES = { set01: '--s-set01', set05: '--s-set03' };

  const $ = s => document.querySelector(s);
  const nf0 = new Intl.NumberFormat('es-AR', { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat('es-AR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const nf2 = new Intl.NumberFormat('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const mwh = kwh => nf1.format(kwh / 1000);
  const pct = x => nf1.format(x * 100) + ' %';
  const ok = v => v != null && isFinite(v);
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

  /* ---------- Datos ---------- */
  const SETS = window.LINEA_BASE.sets.map(s => {
    const porAnio = {};
    s.filas.forEach(f => {
      (porAnio[f.anio] = porAnio[f.anio] || Array(12).fill(null))[MESES.indexOf(f.mes)] = f.consumo;
    });
    return { id: s.id, nombre: s.nombre, edificios: s.edificios, color: COLORES[s.id] || '--s-resto', porAnio };
  });
  const ANIOS = [...new Set(SETS.flatMap(s => Object.keys(s.porAnio).map(Number)))].sort((a, b) => a - b);

  const pasajeros = anio => {
    const p = window.DATOS_CONSUMOS && window.DATOS_CONSUMOS.combustible.pasajeros[anio];
    return p || Array(12).fill(null);
  };
  const gradosDia = (anio, mes) => {
    const g = (window.DATOS_EJEMPLO ? window.DATOS_EJEMPLO.gradosDias : []).find(x => +x.anio === anio && x.mes === MESES[mes]);
    const hdd = window.HDD13 && window.HDD13[anio] ? window.HDD13[anio][mes] : null;
    return { cdd: g ? g.cdd : null, hdd: ok(hdd) ? hdd : (g && g.hdd != null ? g.hdd : null) };
  };

  const estado = { anio: ANIOS[ANIOS.length - 1], mes: 0, sel: TOTAL };
  const charts = {};
  const serieSet = (s, a) => s.porAnio[a] || Array(12).fill(null);
  const serieTotal = a => Array.from({ length: 12 }, (_, i) => {
    const vs = SETS.map(s => serieSet(s, a)[i]).filter(ok);
    return vs.length ? vs.reduce((x, y) => x + y, 0) : null;
  });
  const serieDe = (sel, a) => sel === TOTAL ? serieTotal(a) : serieSet(SETS.find(s => s.id === sel), a);
  const mesesCon = a => serieTotal(a).map((v, i) => ok(v) ? i + 1 : null).filter(Boolean);
  const suma = (arr, meses) => meses.reduce((t, m) => t + (arr[m - 1] || 0), 0);
  const nombreDe = sel => sel === TOTAL ? 'Total SET 01 + SET 05' : SETS.find(s => s.id === sel).nombre;

  function tokens() {
    const cs = getComputedStyle(document.documentElement), t = n => cs.getPropertyValue(n).trim();
    return { ink: t('--ink'), ink2: t('--ink-2'), muted: t('--muted'), grid: t('--grid'), axis: t('--axis'), surface: t('--surface'), accent: t('--accent'), font: t('--font'), mono: t('--mono'), color: t };
  }
  const colorDe = (sel, T) => sel === TOTAL ? T.accent : T.color(SETS.find(s => s.id === sel).color);
  function conAlfa(hex, a) {
    const h = hex.replace('#', '');
    if (h.length !== 6) return hex;
    const n = parseInt(h, 16);
    return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${a})`;
  }
  function opcionesBase(T) {
    return {
      responsive: true, maintainAspectRatio: false,
      animation: window.matchMedia('(prefers-reduced-motion: reduce)').matches ? false : { duration: 250 },
      plugins: { legend: { display: false }, tooltip: {
        backgroundColor: T.surface, titleColor: T.ink, bodyColor: T.ink2, borderColor: T.axis, borderWidth: 1, padding: 10, cornerRadius: 8, boxPadding: 4, usePointStyle: true,
        titleFont: { family: T.font, weight: '600', size: 12 }, bodyFont: { family: T.mono, size: 12 } } },
    };
  }
  function ejes(T) {
    return {
      x: { grid: { display: false }, border: { color: T.axis }, ticks: { color: T.ink2, font: { family: T.font, size: 12 } } },
      y: { beginAtZero: true, border: { display: false }, grid: { color: T.grid, drawTicks: false },
        ticks: { color: T.muted, font: { family: T.mono, size: 11 }, padding: 6, callback: v => nf0.format(v) },
        title: { display: true, text: 'MWh', color: T.muted, font: { family: T.mono, size: 11 } } },
    };
  }
  const kpi = (lab, val, unidad, pie) => `<div class="kpi"><div class="k-lab">${lab}</div><div class="k-val">${val}<small>${unidad}</small></div><div class="k-pie">${pie}</div></div>`;
  const delta = (actual, previo, et) => (!ok(actual) || !previo) ? '' :
    `<span class="delta ${actual > previo ? 'sube' : 'baja'}">${actual > previo ? '▲ +' : '▼ '}${nf1.format((actual / previo - 1) * 100)} %</span> ${et}`;

  /* ---------- Render ---------- */
  function render() {
    const T = tokens(), A = estado.anio, meses = mesesCon(A);
    if (estado.mes && !meses.includes(estado.mes)) estado.mes = 0;
    const filtro = estado.mes ? [estado.mes] : meses;
    renderSelectores(meses);
    $('#fuente').textContent = 'SET 01 (Edificios 5 y 6) y SET 05 (Edificio 4) · ' + (meses.length ? MESES[meses[0] - 1] + ' a ' + MESES[meses[meses.length - 1] - 1] + ' ' + A : 'sin meses') + ' · fuente: facturas';

    const tot = serieTotal(A), totP = suma(tot, filtro);
    const pie1 = estado.mes
      ? (meses.indexOf(estado.mes) > 0 ? delta(totP, tot[meses[meses.indexOf(estado.mes) - 1] - 1], 'vs ' + MESES[meses[meses.indexOf(estado.mes) - 1] - 1].toLowerCase()) : MESES[estado.mes - 1] + ' ' + A)
      : 'Promedio ' + mwh(totP / (filtro.length || 1)) + ' MWh/mes';
    const px = pasajeros(A), mp = filtro.filter(m => ok(px[m - 1]) && ok(tot[m - 1]));
    const pax = mp.reduce((t, m) => t + px[m - 1], 0);
    const may = SETS.map(s => ({ s, v: suma(serieSet(s, A), filtro) })).sort((a, b) => b.v - a.v)[0];
    const prev = serieTotal(A - 1), mPrev = filtro.filter(m => ok(prev[m - 1]));
    const dPrev = mPrev.length === filtro.length && filtro.length ? delta(totP, suma(prev, mPrev), 'vs mismo período ' + (A - 1)) : '';

    $('#kpis').innerHTML = [
      kpi('Consumo facturado SET 01 + SET 05', mwh(totP), 'MWh', pie1),
      kpi('Energía por pasajero', pax ? nf2.format(suma(tot, mp) / pax) : '—', 'kWh/pax', pax ? nf0.format(pax) + ' pasajeros' : 'Sin datos de pasajeros'),
      kpi('Mayor subestación', may ? mwh(may.v) : '—', 'MWh', may && totP ? may.s.nombre + ' · ' + pct(may.v / totP) + ' del total' : ''),
      kpi('Variación interanual', dPrev ? dPrev.split('</span>')[0] + '</span>' : '—', '', dPrev ? dPrev.split('</span>')[1].trim() : (A === ANIOS[0] ? 'año base, sin año anterior' : 'sin meses comparables')),
    ].join('');

    // Mensual apilado
    $('#leyenda-mensual').innerHTML = SETS.map(s => `<li><i style="background:${T.color(s.color)}"></i>${s.nombre}</li>`).join('');
    charts.mensual && charts.mensual.destroy();
    let opt = opcionesBase(T); opt.scales = ejes(T); opt.scales.x.stacked = true; opt.scales.y.stacked = true;
    opt.interaction = { mode: 'index', intersect: false };
    opt.plugins.tooltip.itemSort = (a, b) => b.datasetIndex - a.datasetIndex;
    opt.plugins.tooltip.callbacks = {
      title: it => MESES[meses[it[0].dataIndex] - 1] + ' ' + A,
      label: it => ' ' + it.dataset.label + ': ' + nf1.format(it.raw) + ' MWh',
      labelColor: it => ({ backgroundColor: T.color(SETS[it.datasetIndex].color), borderColor: 'transparent' }),
      footer: it => 'Total: ' + mwh(tot[meses[it[0].dataIndex] - 1]) + ' MWh',
    };
    opt.onClick = (e, els, ch) => {
      const pts = ch.getElementsAtEventForMode(e, 'index', { intersect: false }, false);
      if (!pts.length) return;
      const m = meses[pts[0].index]; estado.mes = estado.mes === m ? 0 : m; setTimeout(render, 0);
    };
    opt.onHover = (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; };
    charts.mensual = new Chart($('#ch-mensual'), { type: 'bar', data: {
      labels: meses.map(m => MES_CORTO[m - 1]),
      datasets: SETS.map((s, i) => {
        const base = T.color(s.color);
        return { label: s.nombre, data: meses.map(m => (serieSet(s, A)[m - 1] || 0) / 1000),
          backgroundColor: meses.map(m => estado.mes && m !== estado.mes ? conAlfa(base, 0.3) : base),
          borderColor: T.surface, borderWidth: { top: i === SETS.length - 1 ? 0 : 2 }, borderSkipped: false,
          borderRadius: i === SETS.length - 1 ? { topLeft: 4, topRight: 4 } : 0, maxBarThickness: 56, stack: 'e' };
      }) }, options: opt });

    // Detalle
    const sel = estado.sel, ser = serieDe(sel, A), col = colorDe(sel, T);
    $('#t-detalle').textContent = nombreDe(sel);
    const set = SETS.find(s => s.id === sel);
    $('#n-detalle').textContent = set ? 'Facturación de ' + set.nombre + ' — ' + set.edificios + '. Elegí otra subestación en el filtro o en la tabla.' : 'Suma de las facturas de SET 01 y SET 05. Elegí una subestación en el filtro o en la tabla.';
    const totAnio = suma(ser, meses), prom = meses.length ? totAnio / meses.length : null;
    const campos = [];
    if (set) campos.push(['Edificios', set.edificios]);
    campos.push(['Total ' + A, mwh(totAnio) + ' MWh']);
    if (sel !== TOTAL) campos.push(['Participación', suma(tot, meses) ? pct(totAnio / suma(tot, meses)) : '—']);
    campos.push(['Promedio mensual', prom != null ? mwh(prom) + ' MWh' : '—']);
    if (estado.mes) {
      const v = ser[estado.mes - 1], i = meses.indexOf(estado.mes), p = i > 0 ? ser[meses[i - 1] - 1] : null;
      let txt = ok(v) ? mwh(v) + ' MWh' : '—';
      if (ok(v) && p) txt += ` <span class="delta ${v > p ? 'sube' : 'baja'}">${v > p ? '+' : ''}${nf1.format((v / p - 1) * 100)} %</span>`;
      campos.push([MESES[estado.mes - 1], txt]);
    }
    $('#ficha').innerHTML = campos.map(([k, v]) => `<div><dt>${k}</dt><dd>${v}</dd></div>`).join('');
    charts.detalle && charts.detalle.destroy();
    opt = opcionesBase(T); opt.scales = ejes(T); opt.interaction = { mode: 'index', intersect: false };
    opt.plugins.tooltip.callbacks = {
      title: it => MESES[meses[it[0].dataIndex] - 1] + ' ' + A,
      label: it => ' ' + (it.raw == null ? 'sin dato' : nf1.format(it.raw) + ' MWh'),
      labelColor: () => ({ backgroundColor: col, borderColor: 'transparent' }),
    };
    charts.detalle = new Chart($('#ch-detalle'), { type: 'line', data: {
      labels: meses.map(m => MES_CORTO[m - 1]),
      datasets: [{ data: meses.map(m => ok(ser[m - 1]) ? ser[m - 1] / 1000 : null), borderColor: col,
        backgroundColor: conAlfa(col.startsWith('#') ? col : '#2a78d6', 0.12), fill: 'origin', borderWidth: 2, tension: 0.25,
        pointRadius: meses.map(m => m === estado.mes ? 6 : 3), pointHoverRadius: 6, pointBackgroundColor: col, pointBorderColor: T.surface, pointBorderWidth: 2, spanGaps: false }] }, options: opt });

    renderMatriz(meses, tot, T);
  }

  function renderSelectores(meses) {
    $('#f-anio').innerHTML = ANIOS.map(a => `<option value="${a}">${a}</option>`).join('');
    $('#f-anio').value = estado.anio;
    $('#f-mes').innerHTML = '<option value="0">Todo el período</option>' + meses.map(m => `<option value="${m}">${MESES[m - 1]}</option>`).join('');
    $('#f-mes').value = String(estado.mes);
    $('#f-set').innerHTML = `<option value="${TOTAL}">Total SET 01 + SET 05</option>` + SETS.map(s => `<option value="${s.id}">${esc(s.nombre)} — ${esc(s.edificios)}</option>`).join('');
    $('#f-set').value = estado.sel;
  }

  function renderMatriz(meses, tot, T) {
    const A = estado.anio, totP = suma(tot, meses), colSel = m => m === estado.mes ? ' col-sel' : '';
    let h = '<thead><tr><th class="izq">Subestación</th><th class="izq">Edificios</th>' +
      meses.map(m => `<th class="mes${colSel(m)}" data-mes="${m}" title="Filtrar ${MESES[m - 1]}">${MES_CORTO[m - 1]}</th>`).join('') + '<th>Total</th><th>% total</th></tr></thead><tbody>';
    const celdas = (ser, heat) => {
      const max = Math.max(0, ...meses.map(m => ser[m - 1] || 0));
      return meses.map(m => {
        const v = ser[m - 1];
        if (!ok(v)) return `<td class="num${colSel(m)}">—</td>`;
        const p = heat && max ? Math.round(8 + 92 * (v / max)) : 0;
        const bg = p ? ` style="background-color: color-mix(in srgb, var(--heat-1) ${p}%, transparent)"` : '';
        return `<td class="num${colSel(m)}"${bg} title="${nf0.format(v)} kWh">${mwh(v)}</td>`;
      }).join('');
    };
    const fila = (id, nombre, color, carga, ser, clases) => {
      const t = suma(ser, meses);
      return `<tr class="${clases}${id === estado.sel ? ' sel' : ''}" data-id="${esc(id)}"><td class="izq"><span class="medidor-nom"><i style="background:${color}"></i>${esc(nombre)}</span></td>` +
        `<td class="izq carga" title="${esc(carga)}">${esc(carga)}</td>${celdas(ser, !clases.includes('total'))}<td class="num">${mwh(t)}</td><td class="num">${totP ? pct(t / totP) : '—'}</td></tr>`;
    };
    SETS.forEach(s => { h += fila(s.id, s.nombre, T.color(s.color), s.edificios, serieSet(s, A), 'fila'); });
    h += fila(TOTAL, 'Total SET 01 + SET 05', T.accent, 'Suma de facturas', tot, 'fila total');

    const px = pasajeros(A), ctx = [];
    ctx.push(['Pasajeros', m => ok(px[m - 1]) ? nf0.format(px[m - 1]) : '—', nf0.format(meses.reduce((t, m) => t + (px[m - 1] || 0), 0))]);
    ctx.push(['kWh por pasajero', m => ok(px[m - 1]) && ok(tot[m - 1]) ? nf2.format(tot[m - 1] / px[m - 1]) : '—', (() => {
      const mp = meses.filter(m => ok(px[m - 1])), p = mp.reduce((t, m) => t + px[m - 1], 0);
      return p ? nf2.format(suma(tot, mp) / p) : '—';
    })()]);
    ctx.push(['Grados-día refrig. (base 18 °C)', m => { const g = gradosDia(A, m - 1); return ok(g.cdd) ? nf1.format(g.cdd) : '—'; }, '']);
    ctx.push(['Grados-día calef. (base 13 °C)', m => { const g = gradosDia(A, m - 1); return ok(g.hdd) ? nf1.format(g.hdd) : '—'; }, '']);
    h += `<tr class="grupo"><td colspan="${meses.length + 4}">Contexto</td></tr>`;
    ctx.forEach(([nom, f, t]) => { h += `<tr><td class="izq" colspan="2">${nom}</td>` + meses.map(m => `<td class="num${colSel(m)}">${f(m)}</td>`).join('') + `<td class="num">${t}</td><td></td></tr>`; });
    $('#matriz').innerHTML = h + '</tbody>';
  }

  /* ---------- Eventos ---------- */
  $('#f-anio').addEventListener('change', e => { estado.anio = +e.target.value; estado.mes = 0; render(); });
  $('#f-mes').addEventListener('change', e => { estado.mes = +e.target.value; render(); });
  $('#f-set').addEventListener('change', e => { estado.sel = e.target.value; render(); });
  $('#matriz').addEventListener('click', e => {
    const th = e.target.closest('th[data-mes]');
    if (th) { const m = +th.dataset.mes; estado.mes = estado.mes === m ? 0 : m; render(); return; }
    const tr = e.target.closest('tr[data-id]');
    if (tr) { estado.sel = tr.dataset.id; render(); }
  });
  window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', render);
  new MutationObserver(render).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });
  Chart.defaults.font.family = tokens().font;
  render();
})();
