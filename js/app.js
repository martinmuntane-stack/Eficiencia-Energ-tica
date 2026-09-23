/* Tablero de consumo eléctrico del Aeroparque.
 * Lee el Excel en el navegador (SheetJS), guarda los datos en localStorage
 * y dibuja los gráficos con Chart.js. No necesita servidor. */
(function () {
  'use strict';

  const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  const MES_CORTO = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const CLAVE_STORAGE = 'aep-energia-v1';
  const TOTAL = '__total';

  /* Topología eléctrica del Aeroparque.
   * EDENOR 1/2 entra por el medidor cabecera SMEC, que alimenta las SET 01B, 03, 00, 04 y N3
   * (cada una con sus submedidores). EDENOR 3 alimenta la SET 02, que no pasa por SMEC.
   * Por eso: Total Aeroparque = SMEC + medidores de SET 02 (+ cualquier medidor no clasificado). */
  const GRUPOS = [
    { id: 'set01', nombre: 'SET 01B', color: '--s-set01', bajoSmec: true },
    { id: 'set03', nombre: 'SET 03', color: '--s-set03', bajoSmec: true },
    { id: 'set00', nombre: 'SET 00', color: '--s-set00', bajoSmec: true },
    { id: 'set04', nombre: 'SET 04', color: '--s-set04', bajoSmec: true },
    { id: 'setn3', nombre: 'SET N3', color: '--s-setn3', bajoSmec: true },
    { id: 'set02', nombre: 'SET 02', color: '--s-set02', bajoSmec: false },
    { id: 'otros', nombre: 'Otros', color: '--s-resto', bajoSmec: false },
  ];
  const RESTO = { id: 'resto', nombre: 'SMEC sin submedición', color: '--s-resto' };
  const GRUPO = Object.fromEntries(GRUPOS.map(g => [g.id, g]));

  const $ = sel => document.querySelector(sel);
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase();
  const esCabecera = medidor => norm(medidor) === 'smec';

  const nf1 = new Intl.NumberFormat('es-AR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const nf0 = new Intl.NumberFormat('es-AR', { maximumFractionDigits: 0 });
  const nf2 = new Intl.NumberFormat('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const mwh = kwh => nf1.format(kwh / 1000);
  const pct = x => nf1.format(x * 100) + ' %';

  /* ---------- Normalización ---------- */

  function mesAIndice(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return v.getMonth() + 1;
    if (typeof v === 'number') {
      if (v >= 1 && v <= 12) return Math.round(v);
      if (v > 20000 && window.XLSX) { const d = XLSX.SSF.parse_date_code(v); return d ? d.m : null; }
      return null;
    }
    const n = norm(v);
    if (/^\d{1,2}$/.test(n) && +n >= 1 && +n <= 12) return +n;
    if (n.startsWith('set')) return 9; // "Setiembre"
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

  function grupoDeUbicacion(ubicacion) {
    const u = norm(ubicacion).replace(/\s+/g, '');
    if (!u) return null;
    if (u.includes('n3') || u.includes('nucleo')) return 'setn3';
    const m = u.match(/set0?(\d)/);
    if (!m) return null;
    return { '0': 'set00', '1': 'set01', '2': 'set02', '3': 'set03', '4': 'set04' }[m[1]] || null;
  }

  function grupoDe(medidor) {
    if (esCabecera(medidor)) return 'set01';
    const meta = estado.store.medidores.find(x => norm(x.medidor) === norm(medidor));
    return (meta && grupoDeUbicacion(meta.ubicacion)) || grupoDeUbicacion(medidor) || 'otros';
  }

  /* ---------- Almacenamiento ---------- */

  function storeDeEjemplo() {
    const d = window.DATOS_EJEMPLO;
    const anio = d.anio;
    return {
      anios: {
        [anio]: {
          consumo: d.consumo.map(r => ({ mes: mesAIndice(r.mes), medidor: r.medidor, kwh: r.kwh })),
          pasajeros: Object.fromEntries(d.pasajeros.map(r => [mesAIndice(r.mes), r.pasajeros])),
          fuente: 'datos de ejemplo (BASE_POWER.xlsx)',
          cargado: null,
        },
      },
      medidores: d.medidores,
      gradosDias: d.gradosDias.map(r => ({ anio: r.anio, mes: mesAIndice(r.mes), cdd: r.cdd, hdd: r.hdd })),
      ejemplo: true,
    };
  }

  function leerStore() {
    try {
      const raw = localStorage.getItem(CLAVE_STORAGE);
      if (raw) {
        const s = JSON.parse(raw);
        if (s && s.anios && Object.keys(s.anios).length) return s;
      }
    } catch (e) { /* sin storage: seguimos con el ejemplo */ }
    return storeDeEjemplo();
  }

  function guardarStore() {
    try {
      if (estado.store.ejemplo) localStorage.removeItem(CLAVE_STORAGE);
      else localStorage.setItem(CLAVE_STORAGE, JSON.stringify(estado.store));
    } catch (e) { /* ignorar */ }
  }

  /* ---------- Lectura del Excel ---------- */

  function leerLibro(wb, anioPorDefecto) {
    const hojas = wb.SheetNames.map(n => ({
      nombre: n,
      filas: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: null, raw: true }),
    }));

    // Busca la primera tabla (en cualquier hoja, en las primeras 40 filas) que tenga las columnas pedidas.
    function buscarTabla(columnas) {
      for (const h of hojas) {
        for (let i = 0; i < Math.min(40, h.filas.length); i++) {
          const cab = (h.filas[i] || []).map(norm);
          const idx = {};
          let ok = true;
          for (const [clave, alias] of Object.entries(columnas)) {
            const opcional = clave.endsWith('?');
            const j = cab.findIndex(c => c && alias.some(a => c === a || c.startsWith(a)));
            if (j < 0 && !opcional) { ok = false; break; }
            idx[clave.replace('?', '')] = j;
          }
          if (!ok) continue;
          const datos = [];
          for (const fila of h.filas.slice(i + 1)) {
            if (!fila || fila.every(v => v == null || v === '')) break; // fin de la tabla
            const o = {};
            for (const [k, j] of Object.entries(idx)) o[k] = j >= 0 ? fila[j] : null;
            datos.push(o);
          }
          return { hoja: h.nombre, datos };
        }
      }
      return null;
    }

    const tConsumo = buscarTabla({ mes: ['mes'], medidor: ['medidor'], kwh: ['consumo', 'kwh'], 'anio?': ['ano', 'anio', 'year'] });
    if (!tConsumo) throw new Error('No encontré una hoja con las columnas Mes, Medidor y Consumo_kWh. Revisá los encabezados.');

    const anios = {};
    let descartadas = 0;
    for (const r of tConsumo.datos) {
      const mes = mesAIndice(r.mes);
      const kwh = aNumero(r.kwh);
      const medidor = r.medidor == null ? '' : String(r.medidor).trim();
      const anio = aNumero(r.anio) || anioPorDefecto;
      if (!mes || !medidor || kwh == null) { descartadas++; continue; }
      (anios[anio] = anios[anio] || { consumo: [], pasajeros: {} }).consumo.push({ mes, medidor, kwh });
    }
    if (!Object.keys(anios).length) throw new Error('La hoja "' + tConsumo.hoja + '" no tiene filas válidas de consumo.');

    const tPax = buscarTabla({ mes: ['mes'], pax: ['pasajeros'], 'anio?': ['ano', 'anio', 'year'] });
    if (tPax) {
      for (const r of tPax.datos) {
        const mes = mesAIndice(r.mes), pax = aNumero(r.pax), anio = aNumero(r.anio) || anioPorDefecto;
        if (mes && pax != null && anios[anio]) anios[anio].pasajeros[mes] = pax;
      }
    }

    const tMed = buscarTabla({ medidor: ['medidor'], ubicacion: ['ubicacion'], 'cargas?': ['cargas', 'carga'] });
    const medidores = tMed ? tMed.datos.filter(r => r.medidor).map(r => ({
      ubicacion: r.ubicacion == null ? '' : String(r.ubicacion), medidor: String(r.medidor).trim(), cargas: r.cargas == null ? '' : String(r.cargas),
    })) : [];

    const tGd = buscarTabla({ anio: ['ano', 'anio', 'year'], mes: ['mes'], cdd: ['cooling', 'cdd'], hdd: ['heating', 'hdd'] });
    const gradosDias = tGd ? tGd.datos.map(r => ({ anio: aNumero(r.anio), mes: mesAIndice(r.mes), cdd: aNumero(r.cdd), hdd: aNumero(r.hdd) }))
      .filter(r => r.anio && r.mes) : [];

    return { anios, medidores, gradosDias, descartadas, hoja: tConsumo.hoja };
  }

  function incorporar(res, nombreArchivo) {
    const s = estado.store.ejemplo ? { anios: {}, medidores: [], gradosDias: [] } : estado.store;
    for (const [anio, d] of Object.entries(res.anios)) {
      s.anios[anio] = { ...d, fuente: nombreArchivo, cargado: new Date().toISOString() };
    }
    if (res.medidores.length) {
      const porNombre = new Map(s.medidores.map(m => [norm(m.medidor), m]));
      res.medidores.forEach(m => porNombre.set(norm(m.medidor), m));
      s.medidores = [...porNombre.values()];
    }
    if (res.gradosDias.length) {
      const k = r => r.anio + '-' + r.mes;
      const mapa = new Map(s.gradosDias.map(r => [k(r), r]));
      res.gradosDias.forEach(r => mapa.set(k(r), r));
      s.gradosDias = [...mapa.values()];
    }
    delete s.ejemplo;
    estado.store = s;
    guardarStore();
  }

  /* ---------- Cálculos ---------- */

  function calcular(anio) {
    const d = estado.store.anios[anio];
    const valores = new Map(); // medidor -> [12] kWh | null
    for (const r of d.consumo) {
      if (!valores.has(r.medidor)) valores.set(r.medidor, new Array(12).fill(null));
      const arr = valores.get(r.medidor);
      arr[r.mes - 1] = (arr[r.mes - 1] || 0) + r.kwh;
    }
    // Orden de medidores: el de la hoja "Cargas Alimentadas" si existe, si no el de aparición.
    const orden = estado.store.medidores.map(m => m.medidor);
    const medidores = [...valores.keys()].sort((a, b) => {
      const ia = orden.findIndex(x => norm(x) === norm(a)), ib = orden.findIndex(x => norm(x) === norm(b));
      return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib);
    });
    const cabecera = medidores.find(esCabecera) || null;
    const meses = [];
    for (let i = 0; i < 12; i++) if (medidores.some(m => valores.get(m)[i] != null)) meses.push(i + 1);

    const grupoMes = {}; // id -> [12]
    [...GRUPOS, RESTO].forEach(g => { grupoMes[g.id] = new Array(12).fill(0); });
    const total = new Array(12).fill(null);
    const resto = new Array(12).fill(null);

    for (const mes of meses) {
      const i = mes - 1;
      let sub = 0, fuera = 0;
      for (const m of medidores) {
        if (m === cabecera) continue;
        const v = valores.get(m)[i];
        if (v == null) continue;
        const g = grupoDe(m);
        grupoMes[g][i] += v;
        if (cabecera && GRUPO[g].bajoSmec) sub += v; else fuera += v;
      }
      const vc = cabecera ? valores.get(cabecera)[i] : null;
      if (vc != null) {
        resto[i] = Math.max(0, vc - sub);
        grupoMes.resto[i] = resto[i];
        total[i] = vc + fuera;
      } else {
        total[i] = sub + fuera;
      }
    }

    const parcial = new Array(12).fill(false);
    for (const mes of meses) {
      const otros = meses.filter(x => x !== mes).map(x => total[x - 1]).sort((a, b) => a - b);
      if (otros.length < 2) continue;
      const med = otros.length % 2 ? otros[(otros.length - 1) / 2] : (otros[otros.length / 2 - 1] + otros[otros.length / 2]) / 2;
      parcial[mes - 1] = total[mes - 1] < 0.5 * med;
    }

    const serieDe = sel => {
      if (sel === TOTAL) return total;
      if (sel === RESTO.id) return resto;
      return valores.get(sel) || new Array(12).fill(null);
    };

    const gd = new Array(12).fill(null);
    estado.store.gradosDias.filter(r => +r.anio === +anio).forEach(r => { gd[r.mes - 1] = r; });

    return { anio, d, valores, medidores, cabecera, meses, grupoMes, total, resto, parcial, serieDe, gd, pasajeros: d.pasajeros || {} };
  }

  const suma = (arr, meses) => meses.reduce((a, m) => a + (arr[m - 1] || 0), 0);
  const mesesDelFiltro = c => estado.mes ? [estado.mes] : c.meses;

  /* ---------- Estado y colores ---------- */

  const estado = { store: null, anio: null, mes: 0, medidor: TOTAL, calc: null };
  const charts = {};

  function tokens() {
    const cs = getComputedStyle(document.documentElement);
    const t = n => cs.getPropertyValue(n).trim();
    return {
      ink: t('--ink'), ink2: t('--ink-2'), muted: t('--muted'), grid: t('--grid'), axis: t('--axis'),
      surface: t('--surface'), accent: t('--accent'), border: t('--border'), font: t('--font'), mono: t('--mono'),
      color: v => t(v),
    };
  }

  function conAlfa(hex, a) {
    const h = hex.replace('#', '');
    if (h.length !== 6) return hex;
    const n = parseInt(h, 16);
    return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${a})`;
  }

  function colorDe(sel, T) {
    if (sel === TOTAL) return T.accent;
    if (sel === RESTO.id) return T.color(RESTO.color);
    return T.color(GRUPO[grupoDe(sel)].color);
  }

  function nombreDe(sel) {
    if (sel === TOTAL) return 'Total Aeroparque';
    if (sel === RESTO.id) return RESTO.nombre;
    return sel;
  }

  function metaDe(medidor) {
    return estado.store.medidores.find(x => norm(x.medidor) === norm(medidor)) || null;
  }

  function opcionesBase(T) {
    return {
      responsive: true,
      maintainAspectRatio: false,
      animation: window.matchMedia('(prefers-reduced-motion: reduce)').matches ? false : { duration: 250 },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: T.surface, titleColor: T.ink, bodyColor: T.ink2, borderColor: T.axis, borderWidth: 1,
          padding: 10, cornerRadius: 8, boxPadding: 4, usePointStyle: true,
          titleFont: { family: T.font, weight: '600', size: 12 }, bodyFont: { family: T.mono, size: 12 },
        },
      },
    };
  }

  function ejes(T, { horizontal = false } = {}) {
    const valor = {
      beginAtZero: true, border: { display: false },
      grid: { color: T.grid, drawTicks: false },
      ticks: { color: T.muted, font: { family: T.mono, size: 11 }, padding: 6, callback: v => nf0.format(v) },
      title: { display: true, text: 'MWh', color: T.muted, font: { family: T.mono, size: 11 } },
    };
    const cat = {
      grid: { display: false }, border: { color: T.axis },
      ticks: { color: T.ink2, font: { family: T.font, size: 12 } },
    };
    return horizontal ? { x: valor, y: cat } : { x: cat, y: valor };
  }

  /* ---------- Render ---------- */

  function render() {
    const T = tokens();
    const c = estado.calc = calcular(estado.anio);
    if (estado.mes && !c.meses.includes(estado.mes)) estado.mes = 0;
    renderFuente(c);
    renderSelectores(c);
    renderKpis(c);
    renderMensual(c, T);
    renderRanking(c, T);
    renderDetalle(c, T);
    renderMatriz(c, T);
  }

  function renderFuente(c) {
    const d = c.d;
    const rango = c.meses.length ? MESES[c.meses[0] - 1] + ' a ' + MESES[c.meses[c.meses.length - 1] - 1] : 'sin meses';
    let txt = c.medidores.length + ' medidores · ' + rango + ' ' + c.anio + ' · fuente: ' + d.fuente;
    if (d.cargado) txt += ' (cargado el ' + new Date(d.cargado).toLocaleDateString('es-AR') + ')';
    $('#fuente').textContent = txt;
  }

  function renderSelectores(c) {
    const fa = $('#f-anio');
    const anios = Object.keys(estado.store.anios).sort();
    fa.innerHTML = anios.map(a => `<option value="${a}">${a}</option>`).join('');
    fa.value = estado.anio;

    const fm = $('#f-mes');
    fm.innerHTML = '<option value="0">Todo el período</option>' + c.meses.map(m =>
      `<option value="${m}">${MESES[m - 1]}${c.parcial[m - 1] ? ' (parcial)' : ''}</option>`).join('');
    fm.value = String(estado.mes);

    const fmed = $('#f-medidor');
    const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
    let html = `<option value="${TOTAL}">Total Aeroparque</option>`;
    const bloques = [
      { titulo: 'EDENOR 1/2', ids: ['set01', 'set03', 'set00', 'set04', 'setn3'] },
      { titulo: 'EDENOR 3', ids: ['set02'] },
      { titulo: 'Sin clasificar', ids: ['otros'] },
    ];
    for (const b of bloques) {
      const ms = c.medidores.filter(m => b.ids.includes(grupoDe(m)));
      if (!ms.length) continue;
      html += `<optgroup label="${b.titulo}">` + ms.map(m => {
        const g = GRUPO[grupoDe(m)];
        const et = esCabecera(m) ? m + ' — cabecera' : m + ' — ' + g.nombre;
        return `<option value="${esc(m)}">${esc(et)}</option>`;
      }).join('');
      if (b.titulo === 'EDENOR 1/2' && c.cabecera) html += `<option value="${RESTO.id}">${RESTO.nombre}</option>`;
      html += '</optgroup>';
    }
    fmed.innerHTML = html;
    if (estado.medidor !== TOTAL && estado.medidor !== RESTO.id && !c.valores.has(estado.medidor)) estado.medidor = TOTAL;
    fmed.value = estado.medidor;
  }

  function kpi(lab, val, unidad, pie) {
    return `<div class="kpi"><div class="k-lab">${lab}</div><div class="k-val">${val}<small>${unidad}</small></div><div class="k-pie">${pie}</div></div>`;
  }

  function deltaHtml(actual, previo, etiqueta) {
    if (actual == null || previo == null || previo === 0) return '';
    const d = actual / previo - 1;
    const cls = d > 0 ? 'sube' : 'baja';
    const signo = d > 0 ? '▲ +' : '▼ ';
    return `<span class="delta ${cls}">${signo}${nf1.format(d * 100)} %</span> ${etiqueta}`;
  }

  function renderKpis(c) {
    const meses = mesesDelFiltro(c);
    const periodo = estado.mes ? MESES[estado.mes - 1] + ' ' + c.anio : (c.meses.length + ' meses de ' + c.anio);
    const tot = suma(c.total, meses);

    let pie1 = periodo;
    if (estado.mes) {
      const i = c.meses.indexOf(estado.mes);
      if (c.parcial[estado.mes - 1]) pie1 = '<span class="pill pill-warn">parcial</span> lectura incompleta';
      else if (i > 0 && !c.parcial[c.meses[i - 1] - 1]) pie1 = deltaHtml(tot, c.total[c.meses[i - 1] - 1], 'vs ' + MESES[c.meses[i - 1] - 1].toLowerCase());
    } else {
      const completos = c.meses.filter(m => !c.parcial[m - 1]);
      if (completos.length) pie1 = 'Promedio ' + mwh(suma(c.total, completos) / completos.length) + ' MWh/mes' + (completos.length < c.meses.length ? ' (sin meses parciales)' : '');
    }

    const mesesPax = meses.filter(m => c.pasajeros[m] && c.total[m - 1] != null);
    const pax = mesesPax.reduce((a, m) => a + c.pasajeros[m], 0);
    const kwhPax = pax ? suma(c.total, mesesPax) / pax : null;

    const smec = c.cabecera ? suma(c.valores.get(c.cabecera), meses) : 0;
    const resto = suma(c.resto, meses);

    let mayor = null;
    for (const m of c.medidores) {
      if (m === c.cabecera) continue;
      const v = suma(c.valores.get(m), meses);
      if (!mayor || v > mayor.v) mayor = { m, v };
    }

    $('#kpis').innerHTML = [
      kpi('Consumo total Aeroparque', mwh(tot), 'MWh', pie1),
      kpi('Energía por pasajero', kwhPax != null ? nf2.format(kwhPax) : '—', 'kWh/pax',
        kwhPax != null ? nf0.format(pax) + ' pasajeros' : 'Sin datos de pasajeros'),
      kpi('Mayor submedidor', mayor ? mwh(mayor.v) : '—', 'MWh',
        mayor ? mayor.m + ' · ' + pct(mayor.v / tot) + ' del total' : ''),
      kpi('SMEC sin submedición', smec ? pct(resto / smec) : '—', '',
        smec ? mwh(resto) + ' MWh de ' + mwh(smec) + ' MWh de SMEC' : 'No hay medidor SMEC'),
    ].join('');
  }

  function renderMensual(c, T) {
    const ids = [...GRUPOS.map(g => g.id).filter(id => id !== 'otros' || c.grupoMes.otros.some(v => v)), ...(c.cabecera ? [RESTO.id] : [])];
    const grupos = ids.map(id => id === RESTO.id ? RESTO : GRUPO[id]).filter(g => c.meses.some(m => c.grupoMes[g.id][m - 1] > 0));
    const labels = c.meses.map(m => MES_CORTO[m - 1] + (c.parcial[m - 1] ? '*' : ''));
    const ultimo = grupos.length - 1;
    const datasets = grupos.map((g, gi) => {
      const base = T.color(g.color);
      return {
        label: g.nombre,
        data: c.meses.map(m => c.grupoMes[g.id][m - 1] / 1000),
        backgroundColor: c.meses.map(m => (estado.mes && m !== estado.mes) ? conAlfa(base, 0.3) : base),
        borderColor: T.surface,
        borderWidth: { top: gi === ultimo ? 0 : 2 },
        borderSkipped: false,
        borderRadius: gi === ultimo ? { topLeft: 4, topRight: 4 } : 0,
        maxBarThickness: 56,
        stack: 'e',
      };
    });

    $('#leyenda-mensual').innerHTML = grupos.map(g => `<li><i style="background:${T.color(g.color)}"></i>${g.nombre}</li>`).join('');

    charts.mensual && charts.mensual.destroy();
    const opt = opcionesBase(T);
    opt.scales = ejes(T);
    opt.scales.x.stacked = true;
    opt.scales.y.stacked = true;
    opt.interaction = { mode: 'index', intersect: false };
    opt.plugins.tooltip.itemSort = (a, b) => b.datasetIndex - a.datasetIndex;
    opt.plugins.tooltip.callbacks = {
      title: items => MESES[c.meses[items[0].dataIndex] - 1] + ' ' + c.anio + (c.parcial[c.meses[items[0].dataIndex] - 1] ? ' (parcial)' : ''),
      label: it => ' ' + it.dataset.label + ': ' + nf1.format(it.raw) + ' MWh',
      labelColor: it => ({ backgroundColor: T.color(grupos[it.datasetIndex].color), borderColor: 'transparent' }),
      footer: items => 'Total: ' + mwh(c.total[c.meses[items[0].dataIndex] - 1]) + ' MWh',
    };
    opt.plugins.tooltip.footerFont = { family: T.mono, size: 12, weight: '600' };
    opt.plugins.tooltip.footerColor = T.ink;
    opt.onClick = (e, els, chart) => {
      const pts = chart.getElementsAtEventForMode(e, 'index', { intersect: false }, false);
      if (!pts.length) return;
      const m = c.meses[pts[0].index];
      estado.mes = estado.mes === m ? 0 : m;
      // Se difiere: Chart.js sigue procesando este evento (otros plugins internos)
      // y destruir el gráfico dentro del propio handler dispara un error interno.
      setTimeout(render, 0);
    };
    opt.onHover = (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; };
    charts.mensual = new Chart($('#ch-mensual'), { type: 'bar', data: { labels, datasets }, options: opt });
  }

  function renderRanking(c, T) {
    const meses = mesesDelFiltro(c);
    const items = c.medidores.filter(m => m !== c.cabecera).map(m => ({ id: m, v: suma(c.valores.get(m), meses) }));
    if (c.cabecera) items.push({ id: RESTO.id, v: suma(c.resto, meses) });
    items.sort((a, b) => b.v - a.v);
    const tot = suma(c.total, meses);

    $('#t-ranking').textContent = 'Ranking de medidores · ' + (estado.mes ? MESES[estado.mes - 1] : 'período');
    $('#wrap-ranking').style.height = Math.max(220, items.length * 26 + 50) + 'px';

    const hayFoco = estado.medidor !== TOTAL;
    charts.ranking && charts.ranking.destroy();
    const opt = opcionesBase(T);
    opt.indexAxis = 'y';
    opt.scales = ejes(T, { horizontal: true });
    opt.scales.y.ticks.font = { family: T.mono, size: 11 };
    opt.plugins.tooltip.callbacks = {
      title: its => nombreDe(items[its[0].dataIndex].id),
      label: it => ' ' + nf1.format(it.raw) + ' MWh · ' + pct(items[it.dataIndex].v / tot) + ' del total',
      afterLabel: it => {
        const id = items[it.dataIndex].id;
        if (id === RESTO.id) return ' Diferencia entre SMEC y sus submedidores';
        const meta = metaDe(id);
        return meta && meta.cargas ? ' ' + meta.cargas : '';
      },
      labelColor: it => ({ backgroundColor: colorDe(items[it.dataIndex].id, T), borderColor: 'transparent' }),
    };
    opt.onClick = (e, els) => {
      if (!els.length) return;
      const id = items[els[0].index].id;
      estado.medidor = estado.medidor === id ? TOTAL : id;
      setTimeout(render, 0);
    };
    opt.onHover = (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; };
    charts.ranking = new Chart($('#ch-ranking'), {
      type: 'bar',
      data: {
        labels: items.map(i => nombreDe(i.id)),
        datasets: [{
          data: items.map(i => i.v / 1000),
          backgroundColor: items.map(i => {
            const col = colorDe(i.id, T);
            return hayFoco && i.id !== estado.medidor ? conAlfa(col, 0.3) : col;
          }),
          borderRadius: { topRight: 4, bottomRight: 4 },
          borderSkipped: false,
          barPercentage: 0.8,
          categoryPercentage: 0.9,
        }],
      },
      options: opt,
    });
  }

  function renderDetalle(c, T) {
    const sel = estado.medidor;
    const serie = c.serieDe(sel);
    const meses = mesesDelFiltro(c);
    const totPeriodo = suma(c.total, c.meses);
    const valPeriodo = suma(serie, c.meses);
    const completos = c.meses.filter(m => !c.parcial[m - 1] && serie[m - 1] != null);
    const prom = completos.length ? suma(serie, completos) / completos.length : null;

    $('#t-detalle').textContent = nombreDe(sel);
    let nota;
    if (sel === TOTAL) nota = 'SMEC (EDENOR 1/2) + medidores de SET 02 (EDENOR 3). Elegí un medidor en el filtro, el ranking o la tabla.';
    else if (sel === RESTO.id) nota = 'Energía medida por SMEC que no registra ningún submedidor: cargas no submedidas (p. ej. balizamiento sur) y pérdidas.';
    else if (sel === c.cabecera) nota = 'Medidor cabecera de EDENOR 1/2: incluye lo que miden los submedidores de SET 01B, 03, 00, 04 y N3.';
    else { const meta = metaDe(sel); nota = meta && meta.cargas ? 'Alimenta: ' + meta.cargas : ''; }
    $('#n-detalle').textContent = nota;

    const campos = [];
    if (sel !== TOTAL && sel !== RESTO.id) {
      const g = GRUPO[grupoDe(sel)];
      campos.push(['Subestación', (metaDe(sel) && metaDe(sel).ubicacion) || g.nombre]);
      campos.push(['Alimentador', g.bajoSmec ? 'EDENOR 1/2' : (g.id === 'set02' ? 'EDENOR 3' : '—')]);
    }
    campos.push(['Total ' + c.anio, mwh(valPeriodo) + ' MWh']);
    if (sel !== TOTAL) campos.push(['Participación', totPeriodo ? pct(valPeriodo / totPeriodo) : '—']);
    campos.push(['Promedio mensual', prom != null ? mwh(prom) + ' MWh' : '—']);
    if (estado.mes) {
      const v = serie[estado.mes - 1];
      const i = c.meses.indexOf(estado.mes);
      const prev = i > 0 ? serie[c.meses[i - 1] - 1] : null;
      let txt = v != null ? mwh(v) + ' MWh' : '—';
      if (v != null && prev && !c.parcial[estado.mes - 1] && !c.parcial[c.meses[i - 1] - 1]) {
        const d = v / prev - 1;
        txt += ` <span class="delta ${d > 0 ? 'sube' : 'baja'}">${d > 0 ? '+' : ''}${nf1.format(d * 100)} %</span>`;
      }
      campos.push([MESES[estado.mes - 1], txt]);
    }
    $('#ficha').innerHTML = campos.map(([k, v]) => `<div><dt>${k}</dt><dd>${v}</dd></div>`).join('');

    const col = colorDe(sel, T);
    charts.detalle && charts.detalle.destroy();
    const opt = opcionesBase(T);
    opt.scales = ejes(T);
    opt.interaction = { mode: 'index', intersect: false };
    opt.plugins.tooltip.callbacks = {
      title: its => MESES[c.meses[its[0].dataIndex] - 1] + ' ' + c.anio + (c.parcial[c.meses[its[0].dataIndex] - 1] ? ' (parcial)' : ''),
      label: it => ' ' + (it.raw == null ? 'sin dato' : nf1.format(it.raw) + ' MWh'),
      labelColor: () => ({ backgroundColor: col, borderColor: 'transparent' }),
    };
    charts.detalle = new Chart($('#ch-detalle'), {
      type: 'line',
      data: {
        labels: c.meses.map(m => MES_CORTO[m - 1] + (c.parcial[m - 1] ? '*' : '')),
        datasets: [{
          data: c.meses.map(m => serie[m - 1] == null ? null : serie[m - 1] / 1000),
          borderColor: col,
          backgroundColor: conAlfa(col.startsWith('#') ? col : '#2a78d6', 0.12),
          fill: 'origin',
          borderWidth: 2,
          tension: 0.25,
          pointRadius: c.meses.map(m => (m === estado.mes ? 6 : 3)),
          pointHoverRadius: 6,
          pointBackgroundColor: col,
          pointBorderColor: T.surface,
          pointBorderWidth: 2,
          spanGaps: false,
        }],
      },
      options: opt,
    });
  }

  function renderMatriz(c) {
    const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
    const T = tokens();
    const totPeriodo = suma(c.total, c.meses);
    const colSel = m => (m === estado.mes ? ' col-sel' : '');

    let h = '<thead><tr><th class="izq">Medidor</th><th class="izq">Carga</th>';
    h += c.meses.map(m => `<th class="mes${colSel(m)}" data-mes="${m}" title="Filtrar ${MESES[m - 1]}">${MES_CORTO[m - 1]}${c.parcial[m - 1] ? ' <span class="pill pill-warn" title="Mes parcial">P</span>' : ''}</th>`).join('');
    h += '<th>Total</th><th>% total</th></tr></thead><tbody>';

    const celdas = (serie, heat) => {
      const max = Math.max(0, ...c.meses.map(m => serie[m - 1] || 0));
      return c.meses.map(m => {
        const v = serie[m - 1];
        if (v == null) return `<td class="num${colSel(m)}">—</td>`;
        if (v === 0) return `<td class="num cero${colSel(m)}" title="Lectura en cero">0,0</td>`;
        const p = heat && max ? Math.round(8 + 92 * (v / max)) : 0;
        const bg = p ? ` style="background-color: color-mix(in srgb, var(--heat-1) ${p}%, transparent)"` : '';
        return `<td class="num${colSel(m)}"${bg} title="${nf0.format(v)} kWh">${mwh(v)}</td>`;
      }).join('');
    };
    const fila = (id, nombre, color, carga, serie, clases = 'fila') => {
      const tot = suma(serie, c.meses);
      const sel = id === estado.medidor ? ' sel' : '';
      return `<tr class="${clases}${sel}" data-id="${esc(id)}">` +
        `<td class="izq"><span class="medidor-nom"><i style="background:${color}"></i>${esc(nombre)}</span></td>` +
        `<td class="izq carga" title="${esc(carga)}">${esc(carga)}</td>` +
        celdas(serie, true) +
        `<td class="num">${mwh(tot)}</td><td class="num">${totPeriodo ? pct(tot / totPeriodo) : '—'}</td></tr>`;
    };
    const span = c.meses.length + 4;

    const bloques = [
      { titulo: 'EDENOR 1/2', ids: ['set01', 'set03', 'set00', 'set04', 'setn3'] },
      { titulo: 'EDENOR 3', ids: ['set02'] },
      { titulo: 'Sin clasificar', ids: ['otros'] },
    ];
    for (const b of bloques) {
      const ms = c.medidores.filter(m => b.ids.includes(grupoDe(m)))
        .sort((x, y) => (esCabecera(y) - esCabecera(x)));
      if (!ms.length) continue;
      h += `<tr class="grupo"><td colspan="${span}">${b.titulo}</td></tr>`;
      for (const m of ms) {
        const meta = metaDe(m);
        const nombre = esCabecera(m) ? m + ' (cabecera)' : m;
        h += fila(m, nombre, colorDe(m, T), meta ? meta.cargas : '', c.valores.get(m));
      }
      if (b.titulo === 'EDENOR 1/2' && c.cabecera) {
        h += fila(RESTO.id, RESTO.nombre, T.color(RESTO.color), 'SMEC − submedidores', c.resto, 'fila derivada');
      }
    }
    h += fila(TOTAL, 'Total Aeroparque', T.accent, 'SMEC + SET 02', c.total, 'fila total');

    // Filas de contexto
    const ctx = [];
    if (Object.keys(c.pasajeros).length) {
      ctx.push(['Pasajeros', m => c.pasajeros[m] != null ? nf0.format(c.pasajeros[m]) : '—', nf0.format(c.meses.reduce((a, m) => a + (c.pasajeros[m] || 0), 0))]);
      ctx.push(['kWh por pasajero', m => (c.pasajeros[m] && c.total[m - 1] != null) ? nf2.format(c.total[m - 1] / c.pasajeros[m]) : '—', (() => {
        const ms = c.meses.filter(m => c.pasajeros[m]);
        const p = ms.reduce((a, m) => a + c.pasajeros[m], 0);
        return p ? nf2.format(suma(c.total, ms) / p) : '—';
      })()]);
    }
    if (c.gd.some(Boolean)) {
      ctx.push(['Grados-día refrig. (base 18 °C)', m => c.gd[m - 1] && c.gd[m - 1].cdd != null ? nf1.format(c.gd[m - 1].cdd) : '—', '']);
      ctx.push(['Grados-día calef. (base 13 °C)', m => c.gd[m - 1] && c.gd[m - 1].hdd != null ? nf1.format(c.gd[m - 1].hdd) : '—', '']);
    }
    if (ctx.length) {
      h += `<tr class="grupo"><td colspan="${span}">Contexto</td></tr>`;
      for (const [nom, f, tot] of ctx) {
        h += `<tr><td class="izq" colspan="2">${nom}</td>` + c.meses.map(m => `<td class="num${colSel(m)}">${f(m)}</td>`).join('') + `<td class="num">${tot}</td><td></td></tr>`;
      }
    }
    h += '</tbody>';
    $('#matriz').innerHTML = h;
  }

  /* ---------- Eventos ---------- */

  function mensaje(txt, tipo) {
    const el = $('#msg-carga');
    el.textContent = txt;
    el.className = 'msg ' + (tipo || '');
  }

  function procesarArchivo(file) {
    if (!file) return;
    if (!window.XLSX) { mensaje('No se pudo cargar el lector de Excel (revisá la conexión a internet).', 'error'); return; }
    mensaje('Leyendo ' + file.name + '…');
    const lector = new FileReader();
    lector.onload = ev => {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array', cellDates: true });
        const anioDef = parseInt($('#anio-carga').value, 10) || new Date().getFullYear();
        const res = leerLibro(wb, anioDef);
        incorporar(res, file.name);
        const anios = Object.keys(res.anios).sort();
        estado.anio = anios[anios.length - 1];
        estado.mes = 0;
        estado.medidor = TOTAL;
        const filas = Object.values(res.anios).reduce((a, d) => a + d.consumo.length, 0);
        mensaje(`Listo: ${filas} lecturas de la hoja "${res.hoja}" para ${anios.join(', ')}` +
          (res.descartadas ? ` (${res.descartadas} filas ignoradas por datos incompletos)` : '') + '.', 'ok');
        render();
      } catch (err) {
        mensaje(err.message || 'No se pudo leer el archivo.', 'error');
      }
    };
    lector.onerror = () => mensaje('No se pudo leer el archivo.', 'error');
    lector.readAsArrayBuffer(file);
  }

  function iniciar() {
    estado.store = leerStore();
    const anios = Object.keys(estado.store.anios).sort();
    estado.anio = anios[anios.length - 1];
    $('#anio-carga').value = estado.anio;

    $('#btn-cargar').addEventListener('click', () => {
      const p = $('#panel-carga');
      p.hidden = !p.hidden;
      $('#btn-cargar').setAttribute('aria-expanded', String(!p.hidden));
    });
    $('#archivo').addEventListener('change', e => { procesarArchivo(e.target.files[0]); e.target.value = ''; });
    const drop = $('#drop');
    ['dragenter', 'dragover'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('sobre'); }));
    ['dragleave', 'drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('sobre'); }));
    drop.addEventListener('drop', e => procesarArchivo(e.dataTransfer.files[0]));
    $('#btn-ejemplo').addEventListener('click', () => {
      estado.store = storeDeEjemplo();
      guardarStore();
      estado.anio = String(estado.store.anios[Object.keys(estado.store.anios)[0]] ? Object.keys(estado.store.anios)[0] : '');
      estado.mes = 0;
      estado.medidor = TOTAL;
      mensaje('Se restauraron los datos de ejemplo. Los datos cargados se borraron de este navegador.', 'ok');
      render();
    });

    $('#f-anio').addEventListener('change', e => { estado.anio = e.target.value; estado.mes = 0; render(); });
    $('#f-mes').addEventListener('change', e => { estado.mes = +e.target.value; render(); });
    $('#f-medidor').addEventListener('change', e => { estado.medidor = e.target.value; render(); });
    $('#matriz').addEventListener('click', e => {
      const th = e.target.closest('th[data-mes]');
      if (th) { const m = +th.dataset.mes; estado.mes = estado.mes === m ? 0 : m; render(); return; }
      const tr = e.target.closest('tr[data-id]');
      if (tr) { estado.medidor = tr.dataset.id; render(); }
    });

    // Redibujar si cambia el tema (sistema o selector del visor).
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', render);
    new MutationObserver(render).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });

    Chart.defaults.font.family = tokens().font;
    render();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
