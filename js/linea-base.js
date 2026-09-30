// Líneas base energéticas (LBE), tomadas de "Linea_base_energetica_para_app.xlsx".
// Cada SET tiene su propia regresión Consumo (kWh) = pendiente * CDD + intercepto,
// ajustada con los 12 meses de 2025 (año base). Con esa recta se predice el consumo
// esperado de cada mes (LBEn) a partir de los grados-día de refrigeración (CDD) y se
// compara contra el consumo real: eso es el desvío. LCS/LCI, el R² y el resto de las
// columnas del Excel se recalculan a partir de estos mismos datos en js/lbe.js, con
// las mismas fórmulas que la planilla (para no duplicar números que puedan desincronizarse).
window.LINEA_BASE = {
  sets: [
    {
      id: 'set01',
      nombre: 'SET 01',
      edificios: 'Edificios 5 y 6',
      pendiente: 2149,
      intercepto: 1362378,
      anioBase: 2025,
      filas: [
        { anio: 2025, mes: 'Enero', cdd: 218.5, consumo: 1837021 },
        { anio: 2025, mes: 'Febrero', cdd: 204.4, consumo: 1952184 },
        { anio: 2025, mes: 'Marzo', cdd: 148.2, consumo: 1595199 },
        { anio: 2025, mes: 'Abril', cdd: 31.1, consumo: 1367477 },
        { anio: 2025, mes: 'Mayo', cdd: 21.1, consumo: 1274985 },
        { anio: 2025, mes: 'Junio', cdd: 0, consumo: 1464033 },
        { anio: 2025, mes: 'Julio', cdd: 0, consumo: 1425087 },
        { anio: 2025, mes: 'Agosto', cdd: 1.6, consumo: 1416390 },
        { anio: 2025, mes: 'Septiembre', cdd: 9.9, consumo: 1391554 },
        { anio: 2025, mes: 'Octubre', cdd: 57.9, consumo: 1428643 },
        { anio: 2025, mes: 'Noviembre', cdd: 79, consumo: 1547026 },
        { anio: 2025, mes: 'Diciembre', cdd: 226.3, consumo: 1794106 },
        { anio: 2026, mes: 'Enero', cdd: 191.6, consumo: 1962588 },
        { anio: 2026, mes: 'Febrero', cdd: 179, consumo: 1986485 },
        { anio: 2026, mes: 'Marzo', cdd: 139, consumo: 1592372 },
      ],
    },
    {
      id: 'set05',
      nombre: 'SET 05',
      edificios: 'Edificio 4',
      pendiente: 1594,
      intercepto: 688045,
      anioBase: 2025,
      filas: [
        { anio: 2025, mes: 'Enero', cdd: 218.5, consumo: 1145107 },
        { anio: 2025, mes: 'Febrero', cdd: 204.4, consumo: 994435 },
        { anio: 2025, mes: 'Marzo', cdd: 148.2, consumo: 836587 },
        { anio: 2025, mes: 'Abril', cdd: 31.1, consumo: 709550 },
        { anio: 2025, mes: 'Mayo', cdd: 21.1, consumo: 700326 },
        { anio: 2025, mes: 'Junio', cdd: 0, consumo: 735240 },
        { anio: 2025, mes: 'Julio', cdd: 0, consumo: 693999 },
        { anio: 2025, mes: 'Agosto', cdd: 1.6, consumo: 680020 },
        { anio: 2025, mes: 'Septiembre', cdd: 9.9, consumo: 701894 },
        { anio: 2025, mes: 'Octubre', cdd: 57.9, consumo: 780653 },
        { anio: 2025, mes: 'Noviembre', cdd: 79, consumo: 861697 },
        { anio: 2025, mes: 'Diciembre', cdd: 226.3, consumo: 1008422 },
        { anio: 2026, mes: 'Enero', cdd: 191.6, consumo: 1050710 },
        { anio: 2026, mes: 'Febrero', cdd: 179, consumo: 919805 },
        { anio: 2026, mes: 'Marzo', cdd: 139, consumo: 933009 },
        { anio: 2026, mes: 'Abril', cdd: 63.3, consumo: 778800 },
        { anio: 2026, mes: 'Mayo', cdd: 6.5, consumo: 672246 },
        { anio: 2026, mes: 'Junio', cdd: 0.1, consumo: 711407 },
        { anio: 2026, mes: 'Julio', cdd: 3.4, consumo: 659204 },
      ],
    },
  ],
};
