/* Sección "Líneas base energéticas" (LBE).
 * Módulo independiente de app.js: no lee ni modifica su estado, sólo agrega
 * una sección nueva más abajo en la misma página, a partir de los datos de
 * "Linea_base_energetica_para_app.xlsx" (window.LINEA_BASE).
 *
 * Modelo (uno por SET, igual al de la planilla):
 *   LBEn (consumo esperado) = pendiente * CDD + intercepto
 *   Desvío = (Consumo real - LBEn) / LBEn
 *   Desv. estándar = desvío estándar muestral de los desvíos del año base
 *   LCS/LCI = LBEn ± LBEn * Desv. estándar
 * Un mes "excede" cuando el consumo real supera el LCS (alerta, dispara
 * análisis de causa raíz) o cae por debajo del LCI (informativo).
 */
(function () {
  'use strict';
  if (!window.LINEA_BASE) return;

  const $ = sel => document.querySelector(sel);
  const nf0 = new Intl.NumberFormat('es-AR', { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat('es-AR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const mwh = kwh => nf1.format(kwh / 1000);
  const pct = x => (x >= 0 ? '+' : '') + nf1.format(x * 100) + ' %';
  const MES_CORTO = { Enero: 'Ene', Febrero: 'Feb', Marzo: 'Mar', Abril: 'Abr', Mayo: 'May', Junio: 'Jun', Julio: 'Jul', Agosto: 'Ago', Septiembre: 'Sep', Octubre: 'Oct', Noviembre: 'Nov', Diciembre: 'Dic' };

  function calcular(set) {
    const base = set.filas.filter(f => f.anio === set.anioBase);
    const filas = set.filas.map(f => {
      const lbEn = set.pendiente * f.cdd + set.intercepto;
      const desvio = (f.consumo - lbEn) / lbEn;
      return { ...f, lbEn, desvio };
    });
    const desviosBase = filas.filter(f => f.anio === set.anioBase).map(f => f.desvio);
    const desvEstandar = stdevMuestral(desviosBase);
    for (const f of filas) {
      f.lcs = f.lbEn + f.lbEn * desvEstandar;
      f.lci = f.lbEn - f.lbEn * desvEstandar;
      f.estado = f.consumo > f.lcs ? 'alto' : (f.consumo < f.lci ? 'bajo' : 'ok');
    }
    const yBase = base.map(f => f.consumo);
    const yBar = yBase.reduce((a, b) => a + b, 0) / yBase.length;
    const filasBase = filas.filter(f => f.anio === set.anioBase);
    const ssRes = filasBase.reduce((a, f) => a + (f.consumo - f.lbEn) ** 2, 0);
    const ssTot = yBase.reduce((a, y) => a + (y - yBar) ** 2, 0);
    const r2 = 1 - ssRes / ssTot;
    const posteriores = filas.filter(f => f.anio > set.anioBase);
    const fueraControl = posteriores.filter(f => f.estado !== 'ok').length;
    return { filas, desvEstandar, r2, fueraControl, posteriores };
  }

  function stdevMuestral(arr) {
    const n = arr.length;
    const m = arr.reduce((a, b) => a + b, 0) / n;
    const ss = arr.reduce((a, v) => a + (v - m) ** 2, 0);
    return Math.sqrt(ss / (n - 1));
  }

  function calidadR2(r2) {
    if (r2 >= 0.9) return 'muy bueno';
    if (r2 >= 0.75) return 'bueno';
    if (r2 >= 0.5) return 'aceptable';
    return 'bajo';
  }

  const estado = { setId: 'set01' };
  const charts = {};

  function tokens() {
    const cs = getComputedStyle(document.documentElement);
    const t = n => cs.getPropertyValue(n).trim();
    return {
      ink: t('--ink'), ink2: t('--ink-2'), muted: t('--muted'), grid: t('--grid'), axis: t('--axis'),
      surface: t('--surface'), accent: t('--accent'), border: t('--border'), font: t('--font'), mono: t('--mono'),
      bad: t('--bad'), good: t('--good'), color: n => t(n),
    };
  }
  function conAlfa(hex, a) {
    const h = (hex || '').replace('#', '');
    if (h.length !== 6) return hex;
    const n = parseInt(h, 16);
    return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${a})`;
  }

  function render() {
    const set = window.LINEA_BASE.sets.find(s => s.id === estado.setId);
    const T = tokens();
    const c = calcular(set);

    $('#lbe-subtitulo').textContent = set.nombre + ' — ' + set.edificios + ' · alimentador propio, sin relación con los medidores SMEC/SET 02 de la sección anterior.';
    renderKpis(set, c);
    renderAjuste(set, c, T);
    renderControl(set, c, T);
    renderTabla(set, c);
  }

  function renderKpis(set, c) {
    const ultima = c.filas[c.filas.length - 1];
    const pillUltima = ultima.estado === 'alto'
      ? '<span class="pill pill-alerta">Atención</span>'
      : ultima.estado === 'bajo'
        ? '<span class="pill pill-info">Por debajo del LCI</span>'
        : '<span class="pill pill-ok">Dentro de control</span>';
    const kpi = (lab, val, pie) => `<div class="kpi"><div class="k-lab">${lab}</div><div class="k-val">${val}</div><div class="k-pie">${pie}</div></div>`;
    $('#lbe-kpis').innerHTML = [
      `<div class="kpi"><div class="k-lab">Ecuación de la línea base</div><div class="ecuacion">LBEn = ${nf0.format(set.pendiente)} × CDD + ${nf0.format(set.intercepto)}<small>kWh, ajustada con los 12 meses de ${set.anioBase}</small></div></div>`,
      kpi('Ajuste del modelo (R²)', nf1.format(c.r2 * 100) + ' %', 'Calidad del ajuste: ' + calidadR2(c.r2)),
      kpi('Desvío estándar', pct(c.desvEstandar).replace('+', ''), 'De los desvíos mensuales de ' + set.anioBase),
      `<div class="kpi"><div class="k-lab">${MES_CORTO[ultima.mes]} ${ultima.anio}</div><div class="k-val">${pct(ultima.desvio)}</div><div class="k-pie">${pillUltima}</div></div>`,
      kpi('Meses fuera de control', c.fueraControl + ' de ' + c.posteriores.length, 'Desde ' + (set.anioBase + 1) + ', consumo real fuera de LCS/LCI'),
    ].join('');
  }

  function renderAjuste(set, c, T) {
    const base = c.filas.filter(f => f.anio === set.anioBase);
    const posteriores = c.filas.filter(f => f.anio > set.anioBase);
    const cdds = c.filas.map(f => f.cdd);
    const x0 = Math.min(...cdds), x1 = Math.max(...cdds) * 1.05;
    const recta = [
      { x: x0, y: (set.pendiente * x0 + set.intercepto) / 1000 },
      { x: x1, y: (set.pendiente * x1 + set.intercepto) / 1000 },
    ];

    $('#lbe-leyenda-ajuste').innerHTML = [
      `<li><i style="background:${T.accent}"></i>Año base (${set.anioBase})</li>`,
      posteriores.length ? `<li><i style="background:${T.color('--s-set03')}"></i>Meses posteriores</li>` : '',
      `<li><i class="linea" style="background:${T.muted}"></i>Recta ajustada</li>`,
    ].join('');

    charts.ajuste && charts.ajuste.destroy();
    charts.ajuste = new Chart($('#lbe-ch-ajuste'), {
      type: 'scatter',
      data: {
        datasets: [
          {
            label: 'Recta ajustada', type: 'line', data: recta, borderColor: T.muted, borderDash: [5, 4],
            borderWidth: 1.5, pointRadius: 0, fill: false, order: 3,
          },
          {
            label: 'Año base (' + set.anioBase + ')',
            data: base.map(f => ({ x: f.cdd, y: f.consumo / 1000, mes: f.mes, anio: f.anio })),
            backgroundColor: T.accent, borderColor: T.surface, borderWidth: 1, pointRadius: 5, pointHoverRadius: 6, order: 1,
          },
          {
            label: 'Meses posteriores',
            data: posteriores.map(f => ({ x: f.cdd, y: f.consumo / 1000, mes: f.mes, anio: f.anio })),
            backgroundColor: T.color('--s-set03'), borderColor: T.surface, borderWidth: 1, pointRadius: 5,
            pointHoverRadius: 6, pointStyle: 'rectRot', order: 2,
          },
        ],
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        animation: window.matchMedia('(prefers-reduced-motion: reduce)').matches ? false : { duration: 250 },
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: T.surface, titleColor: T.ink, bodyColor: T.ink2, borderColor: T.axis, borderWidth: 1,
            padding: 10, cornerRadius: 8, titleFont: { family: T.font, weight: '600', size: 12 }, bodyFont: { family: T.mono, size: 12 },
            callbacks: {
              title: its => (its[0].raw.mes ? its[0].raw.mes + ' ' + its[0].raw.anio : 'Recta ajustada'),
              label: it => it.raw.mes
                ? [' CDD: ' + nf1.format(it.raw.x), ' Consumo real: ' + nf1.format(it.raw.y) + ' MWh']
                : ' LBEn en CDD ' + nf0.format(it.raw.x) + ': ' + nf1.format(it.raw.y) + ' MWh',
            },
          },
        },
        scales: {
          x: {
            title: { display: true, text: 'Grados-día de refrigeración (CDD, base 18 °C)', color: T.muted, font: { family: T.mono, size: 11 } },
            grid: { color: T.grid, drawTicks: false }, border: { color: T.axis },
            ticks: { color: T.muted, font: { family: T.mono, size: 11 } },
          },
          y: {
            title: { display: true, text: 'MWh', color: T.muted, font: { family: T.mono, size: 11 } },
            beginAtZero: true, grid: { color: T.grid, drawTicks: false }, border: { display: false },
            ticks: { color: T.muted, font: { family: T.mono, size: 11 }, callback: v => nf0.format(v) },
          },
        },
      },
    });
  }

  function renderControl(set, c, T) {
    const labels = c.filas.map(f => MES_CORTO[f.mes] + ' ' + String(f.anio).slice(2));
    $('#lbe-leyenda-control').innerHTML = [
      `<li><i style="background:${T.accent}"></i>Consumo real</li>`,
      `<li><i class="linea" style="background:${T.muted}"></i>Línea base (LBEn)</li>`,
      `<li><i class="banda" style="background:${T.accent}"></i>Banda LCI–LCS</li>`,
      `<li><i style="background:${T.bad}"></i>Fuera de control</li>`,
    ].join('');

    charts.control && charts.control.destroy();
    const opt = {
      responsive: true, maintainAspectRatio: false,
      animation: window.matchMedia('(prefers-reduced-motion: reduce)').matches ? false : { duration: 250 },
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: T.surface, titleColor: T.ink, bodyColor: T.ink2, borderColor: T.axis, borderWidth: 1,
          padding: 10, cornerRadius: 8, titleFont: { family: T.font, weight: '600', size: 12 }, bodyFont: { family: T.mono, size: 12 },
          filter: it => it.dataset.label !== 'LCI',
          callbacks: {
            title: its => c.filas[its[0].dataIndex].mes + ' ' + c.filas[its[0].dataIndex].anio,
            label: it => {
              const f = c.filas[it.dataIndex];
              if (it.dataset.label === 'LCS') return ' Banda de control: ' + nf1.format(f.lci / 1000) + ' – ' + nf1.format(f.lcs / 1000) + ' MWh';
              if (it.dataset.label === 'Línea base') return ' Línea base (LBEn): ' + nf1.format(f.lbEn / 1000) + ' MWh';
              return ' Consumo real: ' + nf1.format(f.consumo / 1000) + ' MWh (' + pct(f.desvio) + ' vs. LBEn)';
            },
          },
        },
      },
      scales: {
        x: { grid: { display: false }, border: { color: T.axis }, ticks: { color: T.ink2, font: { family: T.font, size: 11 } } },
        y: {
          beginAtZero: true, grid: { color: T.grid, drawTicks: false }, border: { display: false },
          ticks: { color: T.muted, font: { family: T.mono, size: 11 }, callback: v => nf0.format(v) },
          title: { display: true, text: 'MWh', color: T.muted, font: { family: T.mono, size: 11 } },
        },
      },
    };
    charts.control = new Chart($('#lbe-ch-control'), {
      type: 'line',
      data: {
        labels,
        datasets: [
          {
            label: 'LCI', data: c.filas.map(f => f.lci / 1000), borderWidth: 0, pointRadius: 0,
            fill: false, order: 4,
          },
          {
            label: 'LCS', data: c.filas.map(f => f.lcs / 1000), borderWidth: 0, pointRadius: 0,
            fill: '-1', backgroundColor: conAlfa(T.accent, 0.12), order: 3,
          },
          {
            label: 'Línea base', data: c.filas.map(f => f.lbEn / 1000), borderColor: T.muted, borderDash: [5, 4],
            borderWidth: 1.5, pointRadius: 0, fill: false, order: 2,
          },
          {
            label: 'Consumo real', data: c.filas.map(f => f.consumo / 1000), borderColor: T.accent, borderWidth: 2,
            tension: 0.2, pointRadius: 4, pointHoverRadius: 6, pointBorderWidth: 2, pointBorderColor: T.surface,
            pointBackgroundColor: c.filas.map(f => f.estado === 'ok' ? T.accent : T.bad),
            order: 1,
          },
        ],
      },
      options: opt,
    });
  }

  function renderTabla(set, c) {
    const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
    let h = '<thead><tr><th class="izq">Año</th><th class="izq">Mes</th><th>CDD</th><th>Consumo</th><th>LBEn</th><th>Desvío</th><th>LCS</th><th class="izq">Estado</th></tr></thead><tbody>';
    let anioActual = null;
    for (const f of c.filas) {
      if (f.anio !== anioActual) {
        anioActual = f.anio;
        h += `<tr class="grupo"><td colspan="8">${anioActual}${anioActual === set.anioBase ? ' — año base' : ''}</td></tr>`;
      }
      const clase = f.estado === 'alto' ? ' fila-alta' : (f.estado === 'bajo' ? ' fila-baja' : '');
      const pill = f.estado === 'alto' ? '<span class="pill pill-alerta">Atención</span>'
        : f.estado === 'bajo' ? '<span class="pill pill-info">Bajo LCI</span>'
          : '<span class="pill pill-ok">OK</span>';
      h += `<tr class="${clase}"><td class="izq">${f.anio}</td><td class="izq">${esc(f.mes)}</td>` +
        `<td class="num">${nf1.format(f.cdd)}</td><td class="num">${mwh(f.consumo)}</td>` +
        `<td class="num">${mwh(f.lbEn)}</td><td class="num">${pct(f.desvio)}</td>` +
        `<td class="num">${mwh(f.lcs)}</td><td class="izq">${pill}</td></tr>`;
    }
    h += '</tbody>';
    $('#lbe-tabla').innerHTML = h;
  }

  function iniciar() {
    const sel = $('#lbe-set');
    sel.innerHTML = window.LINEA_BASE.sets.map(s => `<option value="${s.id}">${s.nombre} — ${s.edificios}</option>`).join('');
    sel.value = estado.setId;
    sel.addEventListener('change', e => { estado.setId = e.target.value; render(); });
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', render);
    new MutationObserver(render).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });
    render();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
