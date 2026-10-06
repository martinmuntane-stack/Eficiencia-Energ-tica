/* Tablero de consumo de agua del Aeroparque.
 * Misma mecánica que app.js: Excel leído en el navegador, datos en localStorage, gráficos con Chart.js. */
(function () {
  'use strict';

  const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  const MES_CORTO = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const CLAVE = 'aep-agua-v1';
  const SECTORES = [
    { id: 'Terminal', color: '--a-1' },
    { id: 'Climatización', color: '--a-2' },
    { id: 'Riego', color: '--a-3' },
    { id: 'Gastronomía', color: '--a-4' },
    { id: 'Hangares', color: '--a-5' },
    { id: 'Otros', color: '--a-6' },
  ];
  const SECTOR = Object.fromEntries(SECTORES.map(s => [s.id, s]));
  const TOTAL = '__total';

  const $ = s => document.querySelector(s);
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase();
  const nf0 = new Intl.NumberFormat('es-AR', { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat('es-AR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const css = v => getComputedStyle(document.documentElement).getPropertyValue(v).trim();

  /* ---------- Datos de ejemplo (ilustrativos, determinísticos) ---------- */
  function ejemplo() {
    const puntos = [
      ['Sanitarios Terminal Norte', 'Terminal', 1450, 0.25],
      ['Sanitarios Terminal Sur', 'Terminal', 1280, 0.25],
      ['Torres de enfriamiento', 'Climatización', 900, 1.6],
      ['Riego parquizado', 'Riego', 700, 1.2],
      ['Patio gastronómico', 'Gastronomía', 1100, 0.3],
      ['Hangares', 'Hangares', 520, 0.1],
      ['Edificio administrativo', 'Otros', 340, 0.15],
    ];
    const consumo = [];
    puntos.forEach(([punto, sector, base, k], pi) => {
      for (let m = 1; m <= 12; m++) {
        const ruido = 1 + 0.06 * Math.sin(m * 2.3 + pi * 1.7);
        const verano = (Math.cos(((m - 1) / 12) * 2 * Math.PI)) * k * 0.5 + 1;
        consumo.push({ mes: m, punto, sector, m3: Math.round(base * verano * ruido) });
      }
    });
    return { anios: { 2025: { consumo, fuente: 'datos de ejemplo (ilustrativos)' } }, ejemplo: true };
  }

  function leerStore() {
    try {
      const s = JSON.parse(localStorage.getItem(CLAVE));
      if (s && s.anios && Object.keys(s.anios).length) return s;
    } catch (e) { /* sin storage */ }
    return ejemplo();
  }
  function guardar() {
    try {
      if (estado.store.ejemplo) localStorage.removeItem(CLAVE);
      else localStorage.setItem(CLAVE, JSON.stringify(estado.store));
    } catch (e) { /* ignorar */ }
  }

  /* ---------- Excel ---------- */
  function mesAIndice(v) {
    if (v instanceof Date) return v.getMonth() + 1;
    if (typeof v === 'number') return v >= 1 && v <= 12 ? Math.round(v) : null;
    const n = norm(v);
    if (/^\d{1,2}$/.test(n) && +n >= 1 && +n <= 12) return +n;
    if (n.startsWith('set')) return 9;
    const i = MESES.findIndex(m => n.startsWith(norm(m).slice(0, 3)));
    return i >= 0 ? i + 1 : null;
  }
  function aNumero(v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    if (v == null || v === '') return null;
    let s = String(v).trim().replace(/\s/g, '');
    if (/,\d{1,3}$/.test(s) || /\.\d{3},/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
    const n = Number(s);
    return isFinite(n) ? n : null;
  }
  function sectorDe(txt) {
    const n = norm(txt);
    const hit = SECTORES.find(s => norm(s.id) === n);
    if (hit) return hit.id;
    if (/sanit|terminal|bano|baño/.test(n)) return 'Terminal';
    if (/torre|climat|enfri|hvac|chiller/.test(n)) return 'Climatización';
    if (/riego|parque|jardin/.test(n)) return 'Riego';
    if (/gastro|cocina|comida|bar|restaur/.test(n)) return 'Gastronomía';
    if (/hangar/.test(n)) return 'Hangares';
    return 'Otros';
  }

  function leerLibro(wb, anioDef) {
    for (const nombre of wb.SheetNames) {
      const filas = XLSX.utils.sheet_to_json(wb.Sheets[nombre], { header: 1, raw: true, defval: null });
      for (let h = 0; h < Math.min(filas.length, 15); h++) {
        const cab = filas[h].map(norm);
        const iMes = cab.findIndex(c => c.startsWith('mes'));
        const iPun = cab.findIndex(c => c.startsWith('punto') || c.startsWith('medidor'));
        const iVal = cab.findIndex(c => c.includes('m3') || c.includes('consumo'));
        if (iMes < 0 || iPun < 0 || iVal < 0) continue;
        const iAnio = cab.findIndex(c => c === 'ano' || c === 'anio' || c.startsWith('ano') );
        const iSec = cab.findIndex(c => c.startsWith('sector'));
        const porAnio = {};
        for (const f of filas.slice(h + 1)) {
          const mes = mesAIndice(f[iMes]), val = aNumero(f[iVal]), punto = f[iPun];
          if (!mes || val == null || !punto) continue;
          const anio = (iAnio >= 0 && aNumero(f[iAnio])) || anioDef;
          (porAnio[anio] = porAnio[anio] || []).push({ mes, punto: String(punto).trim(), sector: sectorDe(iSec >= 0 ? f[iSec] : punto), m3: val });
        }
        if (Object.keys(porAnio).length) return porAnio;
      }
    }
    throw new Error('No encontré una tabla con columnas Mes, Punto y Consumo_m3.');
  }

  /* ---------- Estado y cálculos ---------- */
  const estado = { store: leerStore(), anio: null, mes: 0, punto: TOTAL };
  const filasAnio = () => (estado.store.anios[estado.anio] || { consumo: [] }).consumo;
  const puntos = () => {
    const m = new Map();
    filasAnio().forEach(r => m.set(r.punto, r.sector));
    return [...m].map(([punto, sector]) => ({ punto, sector })).sort((a, b) => a.punto.localeCompare(b.punto, 'es'));
  };
  const suma = (rows, mes) => rows.filter(r => !mes || r.mes === mes).reduce((a, r) => a + r.m3, 0);
  const serieMensual = rows => Array.from({ length: 12 }, (_, i) => suma(rows, i + 1));
  const filasPunto = () => estado.punto === TOTAL ? filasAnio() : filasAnio().filter(r => r.punto === estado.punto);
  const mesesConDatos = () => new Set(filasAnio().map(r => r.mes));

  /* ---------- Render ---------- */
  const charts = {};
  function chart(id, cfg) {
    if (charts[id]) charts[id].destroy();
    charts[id] = new Chart($('#' + id), cfg);
  }
  function opciones(extra) {
    const tick = css('--muted'), grid = css('--grid');
    return Object.assign({
      responsive: true, maintainAspectRatio: false, animation: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${nf0.format(c.parsed.y ?? c.parsed.x)} m³` } } },
      scales: {
        x: { stacked: true, ticks: { color: tick }, grid: { display: false } },
        y: { stacked: true, ticks: { color: tick, callback: v => nf0.format(v) }, grid: { color: grid }, beginAtZero: true },
      },
    }, extra || {});
  }

  function kpi(lab, val, unidad, pie) {
    return `<div class="kpi"><div class="k-lab">${lab}</div><div class="k-val">${val}<small>${unidad || ''}</small></div><div class="k-pie">${pie || ''}</div></div>`;
  }

  function render() {
    const mesSel = estado.mes;
    const rows = filasPunto();
    const todos = filasAnio();
    const periodo = mesSel ? MESES[mesSel - 1] : 'año ' + estado.anio;
    const total = suma(rows, mesSel);
    const presentes = [...mesesConDatos()].sort((a, b) => a - b);
    const nMeses = mesSel ? 1 : Math.max(1, presentes.length);

    // Variación vs. mes anterior (sólo con un mes elegido)
    let delta = '';
    if (mesSel > 1 && mesesConDatos().has(mesSel - 1)) {
      const ant = suma(rows, mesSel - 1);
      if (ant) {
        const d = (total - ant) / ant;
        delta = `<span class="delta ${d > 0 ? 'sube' : 'baja'}">${d > 0 ? '▲' : '▼'} ${nf1.format(Math.abs(d) * 100)} %</span> vs. ${MESES[mesSel - 2].toLowerCase()}`;
      }
    }
    const ranking = puntos().map(p => ({ ...p, v: suma(todos.filter(r => r.punto === p.punto), mesSel) })).sort((a, b) => b.v - a.v);
    const totalGlobal = suma(todos, mesSel);
    const top = ranking[0];
    const pico = serieMensual(rows).reduce((b, v, i) => v > b.v ? { v, i } : b, { v: -1, i: 0 });

    $('#kpis').innerHTML =
      kpi('Consumo del período', nf0.format(total), 'm³', (estado.punto === TOTAL ? 'Total Aeroparque' : estado.punto) + ' · ' + periodo) +
      kpi(mesSel ? 'Variación mensual' : 'Promedio mensual', mesSel ? (delta ? delta.split(' vs.')[0] : '—') : nf0.format(total / nMeses), mesSel ? '' : 'm³/mes', mesSel ? delta.split('</span>')[1] || 'sin mes anterior' : nMeses + ' meses con datos') +
      kpi('Mayor consumidor', top ? nf0.format(top.v) : '—', 'm³', top ? `${top.punto} (${nf1.format(totalGlobal ? top.v / totalGlobal * 100 : 0)} % del total)` : '') +
      kpi('Mes pico', pico.v >= 0 ? MES_CORTO[pico.i] : '—', '', pico.v >= 0 ? nf0.format(pico.v) + ' m³' : '');

    // Mensual apilado por sector
    const sectoresPresentes = SECTORES.filter(s => todos.some(r => r.sector === s.id));
    $('#leyenda').innerHTML = sectoresPresentes.map(s => `<li><i style="background:var(${s.color})"></i>${s.id}</li>`).join('');
    const hoy = mesesConDatos();
    chart('ch-mensual', {
      type: 'bar',
      data: {
        labels: MES_CORTO,
        datasets: sectoresPresentes.map(s => ({
          label: s.id,
          backgroundColor: css(s.color),
          data: MES_CORTO.map((_, i) => hoy.has(i + 1) ? suma(todos.filter(r => r.sector === s.id), i + 1) : null),
        })),
      },
      options: opciones({ onClick: (_, el) => { if (el.length) { estado.mes = el[0].index + 1 === estado.mes ? 0 : el[0].index + 1; sync(); } } }),
    });

    // Ranking
    $('#t-ranking').textContent = 'Ranking de puntos · ' + periodo;
    chart('ch-ranking', {
      type: 'bar',
      data: { labels: ranking.map(r => r.punto), datasets: [{ label: 'Consumo', backgroundColor: ranking.map(r => css(SECTOR[r.sector].color)), data: ranking.map(r => r.v) }] },
      options: opciones({
        indexAxis: 'y',
        onClick: (_, el) => { if (el.length) { estado.punto = ranking[el[0].index].punto; sync(); } },
        scales: {
          x: { ticks: { color: css('--muted'), callback: v => nf0.format(v) }, grid: { color: css('--grid') }, beginAtZero: true },
          y: { ticks: { color: css('--ink-2') }, grid: { display: false } },
        },
      }),
    });

    // Detalle
    $('#t-detalle').textContent = estado.punto === TOTAL ? 'Total Aeroparque' : estado.punto;
    const pt = puntos().find(p => p.punto === estado.punto);
    $('#n-detalle').textContent = pt ? 'Sector: ' + pt.sector : 'Suma de todos los puntos de medición';
    chart('ch-detalle', {
      type: 'line',
      data: { labels: MES_CORTO, datasets: [{ label: 'Consumo', borderColor: css('--accent'), backgroundColor: css('--accent'), pointRadius: 3, tension: 0.25, spanGaps: false,
        data: serieMensual(rows).map((v, i) => hoy.has(i + 1) ? v : null) }] },
      options: opciones({ scales: {
        x: { ticks: { color: css('--muted') }, grid: { display: false } },
        y: { ticks: { color: css('--muted'), callback: v => nf0.format(v) }, grid: { color: css('--grid') }, beginAtZero: true },
      } }),
    });

    // Matriz
    const ps = puntos();
    let h = '<thead><tr><th class="izq">Punto</th>' + MES_CORTO.map((m, i) => `<th class="mes${mesSel === i + 1 ? ' col-sel' : ''}" data-mes="${i + 1}">${m}</th>`).join('') + '<th>Total</th></tr></thead><tbody>';
    ps.forEach(p => {
      const ser = serieMensual(todos.filter(r => r.punto === p.punto));
      const max = Math.max(...ser, 1);
      h += `<tr class="fila${estado.punto === p.punto ? ' sel' : ''}" data-punto="${encodeURIComponent(p.punto)}"><td class="izq"><span class="medidor-nom"><i style="background:var(${SECTOR[p.sector].color})"></i>${p.punto}</span></td>` +
        ser.map((v, i) => hoy.has(i + 1)
          ? `<td class="num${mesSel === i + 1 ? ' col-sel' : ''}" style="background-color:color-mix(in srgb, var(--heat-1) ${Math.round(v / max * 100)}%, var(--heat-0))">${nf0.format(v)}</td>`
          : `<td class="num">—</td>`).join('') +
        `<td class="num">${nf0.format(ser.reduce((a, b) => a + b, 0))}</td></tr>`;
    });
    const tot = serieMensual(todos);
    h += `<tr class="total fila${estado.punto === TOTAL ? ' sel' : ''}" data-punto="${TOTAL}"><td class="izq">Total Aeroparque</td>` +
      tot.map((v, i) => `<td class="num">${hoy.has(i + 1) ? nf0.format(v) : '—'}</td>`).join('') + `<td class="num">${nf0.format(tot.reduce((a, b) => a + b, 0))}</td></tr></tbody>`;
    $('#matriz').innerHTML = h;
  }

  function llenarFiltros() {
    const anios = Object.keys(estado.store.anios).sort();
    if (!anios.includes(String(estado.anio))) estado.anio = anios[anios.length - 1];
    $('#f-anio').innerHTML = anios.map(a => `<option ${a == estado.anio ? 'selected' : ''}>${a}</option>`).join('');
    $('#f-mes').innerHTML = '<option value="0">Todo el año</option>' + MESES.map((m, i) => `<option value="${i + 1}" ${estado.mes === i + 1 ? 'selected' : ''}>${m}</option>`).join('');
    $('#f-punto').innerHTML = `<option value="${TOTAL}">Total Aeroparque</option>` +
      puntos().map(p => `<option value="${encodeURIComponent(p.punto)}" ${estado.punto === p.punto ? 'selected' : ''}>${p.punto}</option>`).join('');
    $('#fuente').textContent = 'Fuente: ' + (estado.store.anios[estado.anio].fuente || '—');
  }
  function sync() { llenarFiltros(); render(); }

  /* ---------- Eventos ---------- */
  $('#f-anio').onchange = e => { estado.anio = e.target.value; estado.punto = TOTAL; sync(); };
  $('#f-mes').onchange = e => { estado.mes = +e.target.value; sync(); };
  $('#f-punto').onchange = e => { estado.punto = e.target.value === TOTAL ? TOTAL : decodeURIComponent(e.target.value); sync(); };
  $('#matriz').addEventListener('click', e => {
    const th = e.target.closest('th[data-mes]');
    if (th) { estado.mes = +th.dataset.mes === estado.mes ? 0 : +th.dataset.mes; return sync(); }
    const tr = e.target.closest('tr[data-punto]');
    if (tr) { estado.punto = tr.dataset.punto === TOTAL ? TOTAL : decodeURIComponent(tr.dataset.punto); sync(); }
  });

  const panel = $('#panel-carga'), btn = $('#btn-cargar'), msg = $('#msg-carga');
  $('#anio-carga').value = new Date().getFullYear();
  btn.onclick = () => { panel.hidden = !panel.hidden; btn.setAttribute('aria-expanded', String(!panel.hidden)); };

  function cargar(file) {
    if (!file) return;
    const rd = new FileReader();
    rd.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type: 'array', cellDates: true });
        const porAnio = leerLibro(wb, +$('#anio-carga').value || new Date().getFullYear());
        if (estado.store.ejemplo) estado.store = { anios: {} };
        Object.entries(porAnio).forEach(([a, consumo]) => { estado.store.anios[a] = { consumo, fuente: file.name }; });
        estado.anio = Object.keys(porAnio).sort().pop(); estado.mes = 0; estado.punto = TOTAL;
        guardar(); sync();
        msg.className = 'msg ok'; msg.textContent = 'Listo: ' + Object.keys(porAnio).join(', ') + '.';
      } catch (err) { msg.className = 'msg error'; msg.textContent = err.message; }
    };
    rd.readAsArrayBuffer(file);
  }
  $('#archivo').onchange = e => { cargar(e.target.files[0]); e.target.value = ''; };
  const drop = $('#drop');
  ['dragover', 'dragenter'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('sobre'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('sobre'); }));
  drop.addEventListener('drop', e => cargar(e.dataTransfer.files[0]));
  $('#btn-ejemplo').onclick = () => { estado.store = ejemplo(); estado.anio = null; estado.mes = 0; estado.punto = TOTAL; guardar(); sync(); msg.textContent = ''; };

  window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', render);
  sync();
})();
