window.CONSUMO_CFG = {
  clave: 'aep-gas-v1', unidad: 'm³', colValor: 'Consumo_m3', valorRe: /m3|consumo/, picoInvierno: true,
  sectores: [
    { id: 'Calefacción', color: '--a-1', re: /calef|caldera|radiador/ },
    { id: 'Agua caliente', color: '--a-2', re: /agua|acs|termotanque/ },
    { id: 'Gastronomía', color: '--a-4', re: /gastro|cocina|comida|bar|restaur/ },
    { id: 'Hangares', color: '--a-5', re: /hangar/ },
    { id: 'Otros', color: '--a-6' },
  ],
  // [punto, sector, base mensual, amplitud estacional] — ilustrativo
  puntos: [
    ['Calderas Terminal', 'Calefacción', 9000, 1.7],
    ['Calefacción Hangares', 'Hangares', 3500, 1.5],
    ['Agua caliente sanitaria', 'Agua caliente', 2200, 0.5],
    ['Patio gastronómico', 'Gastronomía', 3000, 0.1],
    ['Edificio administrativo', 'Otros', 1200, 1.2],
  ],
};
