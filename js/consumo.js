/* Tablero de consumo mensual por año (gas, combustible) del Aeroparque.
 * Los datos vienen de js/datos-consumos.js (generado desde el Excel) o de un Excel cargado en el navegador.
 * Se configura con window.CONSUMO_CFG. Gráficos con Chart.js; sin servidor. */
(function () {
  'use strict';

  const CFG = window.CONSUMO_CFG;
  const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  const MES_CORTO = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const COLORES = ['--a-5', '--a-2', '--a-1', '--a-3', '--a-4', '--a-6'];

  const $ = s => document.querySelector(s);
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase();
  const nf0 = new Intl.NumberFormat('es-AR', { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat('es-AR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const nf2 = new Intl.NumberFormat('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const css = v => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
  const U = CFG.unidad;

  /* ---------- Datos ---------- */
  const original = () => JSON.parse(JSON.stringify({ ...window.DATOS_CONSUMOS[CFG.dato], fuente: 'Consumos_de_Fuentes_de_Energia_AEP.xlsx', original: true }));
  function leer() {
    try {
      const s = JSON.parse(localStorage.getItem(CFG.clave));
      if (s && s.consumo && Object.keys(s.consumo).length) return s;
    } catch (e) { /* sin storage */ }
    return original();
  }
  function guardar() {
    try {
      if (estado.d.original) localStorage.removeItem(CFG.clave);
      else localStorage.setItem(CFG.clave, JSON.stringify(estado.d));
    } catch (e) { /* ignorar */ }
  }

  /* Lee un libro con el mismo diseño que el Excel original: bloques con título arriba,
   * una celda "Mes" y los años en columnas (12 filas de meses debajo). */
  function numero(v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    if (v == null || v === '') return null;
    const n = Number(String(v).trim().replace(/\s/g, '').replace(/\./g, '').replace(',', '.'));
    return isFinite(n) ? n : null;
  }
  function leerBloques(ws) {
    const f = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const bloques = [];
    for (let r = 0; r < f.length; r++) {
      for (let c = 0; c < (f[r] || []).length; c++) {
        if (norm(f[r][c]) !== 'mes') continue;
        const titulo = r > 0 && f[r - 1] ? f[r - 1][c] : null;
        const anios = [];
        for (let k = c + 1; typeof f[r][k] === 'number' && f[r][k] > 1900 && f[r][k] < 2200; k++) anios.push([f[r][k], k]);
        if (!anios.length) continue;
        const datos = {};
        anios.forEach(([a, k]) => {
          const vals = Array(12).fill(null);
          for (let m = 0; m < 12; m++) {
            const fila = f[r + 1 + m];
            if (!fila) continue;
            const idx = MESES.findIndex(x => norm(x).slice(0, 3) === norm(fila[c]).slice(0, 3));
            if (idx >= 0) vals[idx] = numero(fila[k]);
          }
          datos[a] = vals;
        });
        bloques.push({ titulo: titulo == null ? '' : String(titulo), datos });
      }
    }
    return bloques;
  }
  function leerLibro(wb) {
    const nombre = wb.SheetNames.find(n => CFG.hoja.test(n));
    if (!nombre) throw new Error('No encontré una hoja que coincida con "' + CFG.lectura + '".');
    const bloques = leerBloques(wb.Sheets[nombre]);
    const res = { consumo: null };
    // El primer bloque con título es el consumo; los extras se buscan por título.
    const conTitulo = bloques.filter(b => b.titulo);
    res.consumo = (conTitulo[0] || bloques[0] || {}).datos;
    Object.entries(CFG.extras || {}).forEach(([k, re]) => {
      const b = conTitulo.find(x => re.test(x.titulo));
      if (b) res[k] = b.datos;
    });
    if (!res.consumo) throw new Error('No encontré una tabla con una celda "Mes" y los años en columnas.');
    return res;
  }

  const estado = { d: leer(), anio: null, cmp: null, mes: 0 };
  const anios = () => Object.keys(estado.d.consumo).map(Number).sort((a, b) => a - b);
  const serie = (clave, a) => (estado.d[clave] && estado.d[clave][a]) || Array(12).fill(null);
  const ok = v => v != null && isFinite(v);

  /* Meses con dato en `a` y en `b` (comparación de iguales). */
  const mesesComunes = (a, b) => serie('consumo', a).map((v, i) => ok(v) && ok(serie('consumo', b)[i]) ? i : -1).filter(i => i >= 0);
  const sumar = (s, idx) => idx.reduce((t, i) => t + (s[i] || 0), 0);

  function atipicos(a) {
    const s = serie('consumo', a), vals = s.filter(ok).sort((x, y) => x - y);
    const res = {};
    if (!vals.length) return res;
    const med = vals[Math.floor(vals.length / 2)];
    s.forEach((v, i) => { if (ok(v) && med > 0 && v > 2.2 * med && CFG.dato === 'combustible') res[i] = 'alto'; });
    s.forEach((v, i) => { if (i && ok(v) && v === s[i - 1] && v > 0) res[i] = 'repetido'; });
    return res;
  }

  /* ---------- Gráficos ---------- */
  const charts = {};
  function chart(id, cfg) { if (charts[id]) charts[id].destroy(); charts[id] = new Chart($('#' + id), cfg); }
  const colorAnio = a => css(COLORES[anios().indexOf(a) % COLORES.length]);
  function escalas(fmt) {
    return {
      x: { ticks: { color: css('--muted') }, grid: { display: false } },
      y: { ticks: { color: css('--muted'), callback: v => fmt(v) }, grid: { color: css('--grid') }, beginAtZero: true },
    };
  }
  function opts(fmt, extra) {
    return Object.assign({
      responsive: true, maintainAspectRatio: false, animation: false,
      interaction: { mode: 'index', intersect: false },
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${c.parsed.y == null ? 'sin dato' : fmt(c.parsed.y)}` } } },
      scales: escalas(fmt),
    }, extra || {});
  }
  const fmtU = v => nf0.format(v) + ' ' + CFG.unidadCorta;

  function kpi(lab, val, unidad, pie) {
    return `<div class="kpi"><div class="k-lab">${lab}</div><div class="k-val">${val}<small>${unidad || ''}</small></div><div class="k-pie">${pie || ''}</div></div>`;
  }
  function deltaTxt(a, b) {
    if (!b) return '—';
    const d = (a - b) / b;
    return `<span class="delta ${d > 0 ? 'sube' : 'baja'}">${d > 0 ? '▲' : '▼'} ${nf1.format(Math.abs(d) * 100)} %</span>`;
  }

  function render() {
    const A = estado.anio, B = estado.cmp, todos = anios().filter(a => a === A || a === B); // sólo el año elegido (y el de comparación, si hay)
    const sA = serie('consumo', A);
    const idxA = sA.map((v, i) => ok(v) ? i : -1).filter(i => i >= 0);
    const totalA = sumar(sA, idxA);
    const comunes = B ? mesesComunes(A, B) : [];
    const rango = comunes.length ? `${MES_CORTO[comunes[0]]}–${MES_CORTO[comunes[comunes.length - 1]]}` : '';
    const pico = idxA.reduce((b, i) => sA[i] > (b == null ? -1 : sA[b]) ? i : b, null);

    // KPIs
    const k = [];
    k.push(kpi('Consumo ' + A, nf0.format(totalA), U, idxA.length === 12 ? 'año completo' : `${idxA.length} meses con dato`));
    if (B) k.push(kpi(`Variación vs. ${B}`, comunes.length ? deltaTxt(sumar(sA, comunes), sumar(serie('consumo', B), comunes)) : '—', '', comunes.length ? `mismos meses (${rango})` : 'sin meses en común'));
    else k.push(kpi('Promedio mensual', idxA.length ? nf0.format(totalA / idxA.length) : '—', U + '/mes', ''));
    k.push(kpi('Mes pico', pico == null ? '—' : MESES[pico], '', pico == null ? '' : nf0.format(sA[pico]) + ' ' + U));
    if (CFG.extras) {
      const vu = serie('vuelos', A), pa = serie('pasajeros', A);
      const idx = idxA.filter(i => ok(vu[i]) && ok(pa[i]));
      const v = sumar(vu, idx), p = sumar(pa, idx), c = sumar(sA, idx);
      k.push(kpi('Litros por vuelo', v ? nf2.format(c / v) : '—', 'L/vuelo', p ? `${nf2.format(c / p * 1000)} L cada 1.000 pasajeros` : ''));
    } else if (B) {
      k.push(kpi('Promedio mensual', idxA.length ? nf0.format(totalA / idxA.length) : '—', U + '/mes', ''));
    }
    $('#kpis').innerHTML = k.join('');

    // Mensual por año
    $('#leyenda').innerHTML = todos.map(a => `<li><i style="background:${colorAnio(a)}"></i>${a}</li>`).join('');
    chart('ch-mensual', {
      type: 'bar',
      data: { labels: MES_CORTO, datasets: todos.map(a => ({
        label: String(a), data: serie('consumo', a), backgroundColor: colorAnio(a),
        borderColor: css('--ink'), borderWidth: a === A ? 0 : 0,
        hidden: false, order: a === A ? 0 : 1,
      })) },
      options: opts(fmtU, { onClick: (_, el) => { if (el.length) { estado.mes = el[0].index + 1 === estado.mes ? 0 : el[0].index + 1; render(); } } }),
    });

    // Acumulado
    chart('ch-acum', {
      type: 'line',
      data: { labels: MES_CORTO, datasets: todos.map(a => {
        let t = 0;
        return { label: String(a), borderColor: colorAnio(a), backgroundColor: colorAnio(a), pointRadius: 2, tension: 0.15, spanGaps: false,
          borderWidth: a === A ? 3 : 1.5,
          data: serie('consumo', a).map(v => { if (!ok(v)) return null; t += v; return t; }) };
      }) },
      options: opts(fmtU),
    });

    // Ratio (sólo combustible): litros por vuelo
    if (CFG.extras) {
      chart('ch-ratio', {
        type: 'line',
        data: { labels: MES_CORTO, datasets: todos.map(a => ({
          label: String(a), borderColor: colorAnio(a), backgroundColor: colorAnio(a), pointRadius: 2, tension: 0.15, spanGaps: false,
          borderWidth: a === A ? 3 : 1.5,
          data: serie('consumo', a).map((v, i) => ok(v) && ok(serie('vuelos', a)[i]) ? v / serie('vuelos', a)[i] : null),
        })) },
        options: opts(v => nf2.format(v) + ' L/vuelo'),
      });
    }

    // Tabla
    let h = '<thead><tr><th class="izq">Mes</th>' + todos.map(a => `<th class="${a === A ? 'col-sel' : ''}">${a}</th>`).join('') + (B ? '<th>Var. ' + A + ' vs ' + B + '</th>' : '') + '</tr></thead><tbody>';
    const at = Object.fromEntries(todos.map(a => [a, atipicos(a)]));
    MESES.forEach((m, i) => {
      const sel = estado.mes === i + 1;
      h += `<tr class="fila${sel ? ' sel' : ''}" data-mes="${i + 1}"><td class="izq">${m}</td>`;
      todos.forEach(a => {
        const v = serie('consumo', a)[i];
        const max = Math.max(...serie('consumo', a).filter(ok), 1);
        const pill = at[a][i] === 'alto' ? ' <span class="pill pill-warn" title="Más del doble de la mediana del año">alto</span>'
          : at[a][i] === 'repetido' ? ' <span class="pill pill-warn" title="Igual al mes anterior: revisar si es una lectura repetida">repetido</span>' : '';
        h += ok(v)
          ? `<td class="num${a === A ? ' col-sel' : ''}" style="background-color:color-mix(in srgb, var(--heat-1) ${Math.round(v / max * 100)}%, var(--heat-0))">${nf0.format(v)}${pill}</td>`
          : `<td class="num${a === A ? ' col-sel' : ''}">—</td>`;
      });
      const va = serie('consumo', A)[i], vb = serie('consumo', B)[i];
      h += (B ? `<td class="num">${ok(va) && ok(vb) ? deltaTxt(va, vb) : '—'}</td>` : '') + '</tr>';
    });
    h += '<tr class="total"><td class="izq">Total</td>' + todos.map(a => `<td class="num">${nf0.format(sumar(serie('consumo', a), serie('consumo', a).map((v, i) => ok(v) ? i : -1).filter(i => i >= 0)))}</td>`).join('') +
      (B ? `<td class="num">${comunes.length ? deltaTxt(sumar(sA, comunes), sumar(serie('consumo', B), comunes)) : '—'}</td>` : '') + '</tr></tbody>';
    $('#matriz').innerHTML = h;

    // Mes elegido: ficha
    const ficha = $('#ficha');
    if (estado.mes) {
      const i = estado.mes - 1;
      ficha.innerHTML = `<h3>${MESES[i]}</h3><dl class="ficha">` + todos.map(a => {
        const v = serie('consumo', a)[i];
        let extra = '';
        if (CFG.extras && ok(v) && ok(serie('vuelos', a)[i])) extra = `<dd class="k-pie">${nf2.format(v / serie('vuelos', a)[i])} L/vuelo</dd>`;
        return `<div><dt>${a}</dt><dd>${ok(v) ? nf0.format(v) + ' ' + U : 'sin dato'}</dd>${extra}</div>`;
      }).join('') + '</dl>';
      ficha.hidden = false;
    } else ficha.hidden = true;

    // Notas de calidad
    const notas = [];
    todos.forEach(a => {
      const faltan = serie('consumo', a).map((v, i) => ok(v) ? null : MES_CORTO[i]).filter(Boolean);
      const pasado = a < new Date().getFullYear();
      if (faltan.length && pasado) notas.push(`${a}: sin dato en ${faltan.join(', ')}.`);
      Object.entries(at[a]).forEach(([i, t]) => notas.push(`${a} ${MES_CORTO[i]}: valor ${t === 'alto' ? 'muy superior a la mediana del año' : 'igual al mes anterior'} (${nf0.format(serie('consumo', a)[i])}); conviene verificarlo.`));
    });
    $('#notas').innerHTML = notas.length ? '<strong>Para revisar</strong><ul>' + notas.map(n => `<li>${n}</li>`).join('') + '</ul>' : '';
  }

  function llenar() {
    const as = anios();
    if (!as.includes(estado.anio)) estado.anio = as[as.length - 1];
    if (!as.includes(estado.cmp) || estado.cmp === estado.anio) estado.cmp = 0;
    $('#f-anio').innerHTML = as.map(a => `<option ${a === estado.anio ? 'selected' : ''}>${a}</option>`).join('');
    $('#f-cmp').innerHTML = '<option value="0">Sin comparar</option>' + as.filter(a => a !== estado.anio).map(a => `<option ${a === estado.cmp ? 'selected' : ''}>${a}</option>`).join('');
    $('#fuente').textContent = 'Fuente: ' + estado.d.fuente;
  }
  const sync = () => { llenar(); render(); };

  $('#f-anio').onchange = e => { estado.anio = +e.target.value; sync(); };
  $('#f-cmp').onchange = e => { estado.cmp = +e.target.value; sync(); };
  $('#matriz').addEventListener('click', e => {
    const tr = e.target.closest('tr[data-mes]');
    if (tr) { estado.mes = +tr.dataset.mes === estado.mes ? 0 : +tr.dataset.mes; render(); }
  });

  const panel = $('#panel-carga'), btn = $('#btn-cargar'), msg = $('#msg-carga');
  btn.onclick = () => { panel.hidden = !panel.hidden; btn.setAttribute('aria-expanded', String(!panel.hidden)); };
  function cargar(file) {
    if (!file) return;
    const rd = new FileReader();
    rd.onload = ev => {
      try {
        const r = leerLibro(XLSX.read(ev.target.result, { type: 'array' }));
        estado.d = { ...r, fuente: file.name, original: false };
        estado.anio = null; estado.cmp = null; estado.mes = 0;
        guardar(); sync();
        msg.className = 'msg ok'; msg.textContent = 'Listo: años ' + anios().join(', ') + '.';
      } catch (err) { msg.className = 'msg error'; msg.textContent = err.message; }
    };
    rd.readAsArrayBuffer(file);
  }
  $('#archivo').onchange = e => { cargar(e.target.files[0]); e.target.value = ''; };
  const drop = $('#drop');
  ['dragover', 'dragenter'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('sobre'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('sobre'); }));
  drop.addEventListener('drop', e => cargar(e.dataTransfer.files[0]));
  $('#btn-ejemplo').onclick = () => { estado.d = original(); estado.anio = null; estado.cmp = null; estado.mes = 0; guardar(); sync(); msg.textContent = ''; };

  window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', render);
  sync();
})();
