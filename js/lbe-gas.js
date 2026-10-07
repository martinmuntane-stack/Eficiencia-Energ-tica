/* Línea base energética del gas contra HDD (grados-día de calefacción, base 13 °C).
 *   LBEn   = pendiente * HDD + intercepto   (regresión lineal de los meses del período base)
 *   Desvío = (real - LBEn) / LBEn
 *   LCS/LCI = LBEn ± LBEn * desvío estándar (muestral) de los desvíos del período base
 * Mismo criterio que la sección de líneas base eléctricas (js/lbe.js). */
(function () {
  'use strict';
  const el = document.getElementById('lbg');
  if (!el || !window.HDD13 || !window.CONSUMO_CFG || window.CONSUMO_CFG.dato !== 'gas') return;

  const BASE_DESDE = 2024, BASE_HASTA = 2025;
  const MES = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const $ = s => document.querySelector(s);
  const nf0 = new Intl.NumberFormat('es-AR', { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat('es-AR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const nf2 = new Intl.NumberFormat('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const pct = x => (x >= 0 ? '+' : '') + nf1.format(x * 100) + ' %';
  const css = v => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
  const ok = v => v != null && isFinite(v);
  const charts = {};

  function filas(consumo) {
    const out = [];
    Object.keys(consumo).map(Number).sort((a, b) => a - b).forEach(a => {
      (consumo[a] || []).forEach((v, i) => {
        const hdd = (window.HDD13[a] || [])[i];
        if (ok(v) && ok(hdd)) out.push({ anio: a, mes: i, real: v, hdd, base: a >= BASE_DESDE && a <= BASE_HASTA });
      });
    });
    return out;
  }

  function modelo(f) {
    const b = f.filter(r => r.base), n = b.length;
    if (n < 3) return null;
    const mx = b.reduce((t, r) => t + r.hdd, 0) / n, my = b.reduce((t, r) => t + r.real, 0) / n;
    const sxx = b.reduce((t, r) => t + (r.hdd - mx) ** 2, 0);
    if (!sxx) return null;
    const pend = b.reduce((t, r) => t + (r.hdd - mx) * (r.real - my), 0) / sxx;
    const inter = my - pend * mx;
    f.forEach(r => { r.lben = pend * r.hdd + inter; r.desvio = (r.real - r.lben) / r.lben; });
    const dv = b.map(r => r.desvio), md = dv.reduce((t, v) => t + v, 0) / n;
    const sd = Math.sqrt(dv.reduce((t, v) => t + (v - md) ** 2, 0) / (n - 1));
    f.forEach(r => {
      r.lcs = r.lben * (1 + sd); r.lci = r.lben * (1 - sd);
      r.estado = r.real > r.lcs ? 'alto' : r.real < r.lci ? 'bajo' : 'ok';
    });
    const ssr = b.reduce((t, r) => t + (r.real - r.lben) ** 2, 0), sst = b.reduce((t, r) => t + (r.real - my) ** 2, 0);
    return { pend, inter, sd, r2: 1 - ssr / sst, n };
  }

  function chart(id, cfg) { if (charts[id]) charts[id].destroy(); charts[id] = new Chart($('#' + id), cfg); }
  const kpi = (lab, val, unidad, pie) => `<div class="kpi"><div class="k-lab">${lab}</div><div class="k-val">${val}<small>${unidad || ''}</small></div><div class="k-pie">${pie || ''}</div></div>`;

  function render() {
    const f = filas(window.__consumoGas || {});
    const m = modelo(f);
    if (!m) { $('#lbg-kpis').innerHTML = ''; $('#lbg-tabla').innerHTML = '<tbody><tr><td>No hay suficientes meses con consumo y HDD en el período base.</td></tr></tbody>'; return; }
    const post = f.filter(r => !r.base), fuera = post.filter(r => r.estado !== 'ok');
    const etq = r => MES[r.mes] + ' ' + String(r.anio).slice(2);
    const colEstado = r => r.estado === 'alto' ? css('--bad') : r.estado === 'bajo' ? css('--good') : css('--accent');

    $('#lbg-ecuacion').textContent = `LBEn (m³) = ${nf1.format(m.pend)} × HDD ${m.inter < 0 ? '−' : '+'} ${nf0.format(Math.abs(m.inter))}`;
    $('#lbg-kpis').innerHTML =
      kpi('Pendiente', nf1.format(m.pend), 'm³/HDD', 'gas adicional por cada grado-día') +
      kpi('Intercepto', nf0.format(m.inter), 'm³', 'consumo con HDD = 0') +
      kpi('Ajuste (R²)', nf2.format(m.r2), '', `${m.n} meses, ene ${BASE_DESDE} – dic ${BASE_HASTA}`) +
      kpi('Meses fuera de la banda', fuera.length, `de ${post.length}`, post.length ? 'después del período base' : 'aún sin meses posteriores');

    // Dispersión HDD vs consumo + recta
    const maxH = Math.max(...f.map(r => r.hdd)) * 1.05;
    chart('lbg-ch-ajuste', {
      type: 'scatter',
      data: { datasets: [
        { label: 'Período base', data: f.filter(r => r.base).map(r => ({ x: r.hdd, y: r.real, r })), backgroundColor: css('--accent'), pointRadius: 4 },
        { label: 'Posteriores', data: post.map(r => ({ x: r.hdd, y: r.real, r })), backgroundColor: css('--s-set03'), pointStyle: 'rectRot', pointRadius: 5 },
        { label: 'LBEn', type: 'line', data: [{ x: 0, y: m.inter }, { x: maxH, y: m.pend * maxH + m.inter }], borderColor: css('--ink-2'), borderDash: [6, 4], pointRadius: 0, borderWidth: 1.5 },
      ] },
      options: {
        responsive: true, maintainAspectRatio: false, animation: false,
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => c.raw.r ? ` ${etq(c.raw.r)}: ${nf0.format(c.raw.y)} m³ con ${nf1.format(c.raw.x)} HDD` : ` LBEn: ${nf0.format(c.raw.y)} m³` } } },
        scales: {
          x: { title: { display: true, text: 'HDD base 13 °C', color: css('--muted') }, ticks: { color: css('--muted') }, grid: { color: css('--grid') }, beginAtZero: true },
          y: { title: { display: true, text: 'm³', color: css('--muted') }, ticks: { color: css('--muted'), callback: v => nf0.format(v) }, grid: { color: css('--grid') }, beginAtZero: true },
        },
      },
    });

    // Línea de tiempo: real vs LBEn y banda
    chart('lbg-ch-control', {
      type: 'line',
      data: { labels: f.map(etq), datasets: [
        { label: 'LCS', data: f.map(r => r.lcs), borderColor: 'transparent', backgroundColor: css('--grid'), pointRadius: 0, fill: '+1', order: 5 },
        { label: 'LCI', data: f.map(r => r.lci), borderColor: 'transparent', pointRadius: 0, fill: false, order: 5 },
        { label: 'LBEn', data: f.map(r => r.lben), borderColor: css('--ink-2'), borderDash: [6, 4], pointRadius: 0, borderWidth: 1.5, order: 2 },
        { label: 'Consumo real', data: f.map(r => r.real), borderColor: css('--accent'), backgroundColor: css('--accent'), tension: 0.1, borderWidth: 2, order: 1,
          pointRadius: f.map(r => r.estado === 'ok' ? 3 : 5), pointBackgroundColor: f.map(colEstado), pointBorderColor: f.map(colEstado) },
      ] },
      options: {
        responsive: true, maintainAspectRatio: false, animation: false, interaction: { mode: 'index', intersect: false },
        plugins: { legend: { display: false }, tooltip: { filter: i => i.dataset.label !== 'LCI', callbacks: { label: c => ` ${c.dataset.label}: ${nf0.format(c.parsed.y)} m³` } } },
        scales: {
          x: { ticks: { color: css('--muted'), maxRotation: 60, autoSkip: true }, grid: { display: false } },
          y: { ticks: { color: css('--muted'), callback: v => nf0.format(v) }, grid: { color: css('--grid') }, beginAtZero: true },
        },
      },
    });

    // Tabla
    let h = '<thead><tr><th class="izq">Mes</th><th>HDD</th><th>Real (m³)</th><th>LBEn (m³)</th><th>Desvío</th><th>LCI (m³)</th><th>LCS (m³)</th><th class="izq">Estado</th></tr></thead><tbody>';
    f.forEach(r => {
      const pill = r.estado === 'alto' ? '<span class="pill pill-alerta">sobre LCS</span>' : r.estado === 'bajo' ? '<span class="pill pill-ok">bajo LCI</span>' : '<span class="pill pill-info">en banda</span>';
      h += `<tr class="${r.estado === 'alto' ? 'fila-alta' : r.estado === 'bajo' ? 'fila-baja' : ''}"><td class="izq">${etq(r)}${r.base ? '' : ''}</td><td class="num">${nf1.format(r.hdd)}</td><td class="num">${nf0.format(r.real)}</td><td class="num">${nf0.format(r.lben)}</td><td class="num">${pct(r.desvio)}</td><td class="num">${nf0.format(Math.max(0, r.lci))}</td><td class="num">${nf0.format(r.lcs)}</td><td class="izq">${pill}${r.base ? ' <span class="pill pill-info">base</span>' : ''}</td></tr>`;
    });
    $('#lbg-tabla').innerHTML = h + '</tbody>';
    $('#lbg-sd').textContent = nf1.format(m.sd * 100) + ' %';
  }

  window.__lbgRender = render;
  render();
  window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', render);
})();
