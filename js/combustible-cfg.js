window.CONSUMO_CFG = {
  clave: 'aep-diesel-v1', unidad: 'litros', colValor: 'Consumo_litros', valorRe: /litro|lts|consumo/, picoInvierno: false,
  sectores: [
    { id: 'Generación', color: '--a-1', re: /grupo|electr|generador|ups/ },
    { id: 'Equipos de rampa', color: '--a-2', re: /rampa|gse|tractor|pushback|escalera/ },
    { id: 'Flota vehicular', color: '--a-3', re: /flota|camion|camioneta|vehic/ },
    { id: 'Contra incendio', color: '--a-4', re: /incend|bomba|bombero/ },
    { id: 'Otros', color: '--a-6' },
  ],
  // [punto, sector, base mensual, amplitud estacional] — ilustrativo
  puntos: [
    ['Grupos electrógenos', 'Generación', 4000, 0.3],
    ['Equipos de rampa (GSE)', 'Equipos de rampa', 12000, 0.2],
    ['Flota vehicular', 'Flota vehicular', 6500, 0.1],
    ['Bombas contra incendio', 'Contra incendio', 400, 0.1],
    ['Otros consumos', 'Otros', 900, 0.2],
  ],
};
