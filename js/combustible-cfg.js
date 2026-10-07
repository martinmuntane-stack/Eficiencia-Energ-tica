window.CONSUMO_CFG = {
  clave: 'aep-combustible-v2', dato: 'combustible', hoja: /combustible/i,
  unidad: 'litros', unidadCorta: 'L',
  titulo: 'Consumo de combustible (diésel)', lectura: 'Combustible',
  // Bloques de la hoja que dan contexto de actividad
  extras: { vuelos: /vuelos\s*totales/i, pasajeros: /pasajeros/i },
};
