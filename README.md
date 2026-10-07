# Eficiencia Energética — Aeroparque

Tablero web para visualizar el consumo eléctrico del Aeroparque Jorge Newbery, diferenciado por mes y por medidor.

## Uso

Es una aplicación estática (HTML + CSS + JS), sin backend ni build. Para usarla:

1. Abrí `index.html` en un navegador (doble clic, o servido por cualquier servidor estático — por ejemplo GitHub Pages).
2. Arranca con los datos de ejemplo de `data/BASE_POWER_ejemplo.xlsx`.
3. Para usar tus propios datos, tocá **Cargar Excel** y elegí (o arrastrá) tu archivo `.xlsx`.

### Formato del Excel

La app busca, en cualquier hoja del libro, una tabla con estas columnas (el nombre puede variar un poco, se detectan por coincidencia parcial):

| Columna | Obligatoria | Descripción |
|---|---|---|
| `Mes` | sí | Nombre del mes (Enero…Diciembre) o número 1–12 |
| `Medidor` | sí | Nombre del medidor (p. ej. `SMEC`, `M. A SET 03`) |
| `Consumo_kWh` | sí | Consumo del mes en kWh |
| `Año` | no | Si falta, se usa el año indicado en el panel de carga |

Opcionalmente, si existen estas otras hojas se incorporan automáticamente:

- **Cargas Alimentadas** (`Ubicación`, `Medidor`, `Cargas Alimenta`): para mostrar qué edificios/cargas alimenta cada medidor y agrupar por subestación.
- **Pasajeros** (`Mes`, `Pasajeros`): para calcular kWh por pasajero.
- **Grados Días** (`Año`, `Mes`, columnas con "Cooling"/"Heating"): se muestran como contexto climático en la tabla.

Cargar un año que ya existe reemplaza esos datos; cargar un año nuevo lo agrega para poder compararlos. Todo se guarda únicamente en el `localStorage` del navegador — no se sube a ningún servidor.

## Qué muestra

- **KPIs**: consumo total del período, energía por pasajero, mayor submedidor y el % de SMEC sin submedición.
- **Consumo mensual apilado por subestación**, con `Total Aeroparque = SMEC (EDENOR 1/2) + medidores de SET 02 (EDENOR 3)`.
- **Ranking de medidores** del período/mes elegido.
- **Detalle** del medidor o total seleccionado, con su evolución mensual.
- **Matriz medidor × mes** con mapa de calor por fila, más pasajeros y grados-día como contexto.

Los filtros de año, mes y medidor están enlazados entre sí: también se puede filtrar tocando una barra del gráfico mensual, una barra del ranking, o una fila/encabezado de la tabla.

## Estructura

```
gas.html, combustible.html   Páginas aparte (mismo estilo): consumo mensual de gas (m³) y de combustible diésel (litros) por año
css/consumo.css         Colores por año y navegación entre páginas
js/consumo.js           Lógica común: comparación entre años, acumulado, litros/vuelo, lectura del Excel
js/gas-cfg.js, js/combustible-cfg.js   Configuración de cada página
js/hdd.js               HDD (base 13 °C) mensuales, de la hoja GRADOS DIAS
js/lbe-gas.js           Línea base del gas contra HDD (período base ene 2024 – dic 2025)
js/datos-consumos.js    Datos (generados desde data/Consumos_de_Fuentes_de_Energia_AEP.xlsx)
index.html            Estructura de la página
css/app.css             Estilos del tablero de medidores (tema claro/oscuro, identidad de Aeropuertos Argentina)
css/lbe.css             Estilos adicionales de la sección "Líneas base energéticas" (no toca app.css)
js/app.js               Lógica del tablero de medidores: lectura de Excel, cálculos y gráficos
js/datos-ejemplo.js     Datos de ejemplo del tablero (generados desde data/BASE_POWER_ejemplo.xlsx)
js/linea-base.js        Datos de las líneas base (desde data/Linea_base_energetica_para_app.xlsx)
js/lbe.js               Lógica de la sección "Líneas base energéticas" (independiente de app.js)
js/vendor/               Chart.js y SheetJS (xlsx), vendorizados localmente
img/logo-negro.png       Isologo de Aeropuertos Argentina (tema claro)
img/logo-blanco.png      Isologo de Aeropuertos Argentina (tema oscuro)
data/BASE_POWER_ejemplo.xlsx               Excel de ejemplo del tablero de medidores
data/Linea_base_energetica_para_app.xlsx   Excel de origen de las líneas base
```

## Identidad visual

La paleta de colores y el isologo son los oficiales de Aeropuertos Argentina (kit de marca institucional/secundaria/terciaria). El verde institucional (`#2c8c95`) se usa como acento de marca (botones, foco, línea de "Total Aeroparque") y las subestaciones toman el resto de los tonos secundarios/terciarios del kit. La tipografía combina Poppins (títulos, en línea con el logo) con Work Sans (texto e interfaz) e IBM Plex Mono (cifras tabulares).

## Líneas base energéticas (EnPI)

Debajo del tablero de medidores hay una segunda sección, independiente, con las líneas base eléctricas
de SET 01 (Edificios 5 y 6) y SET 05 (Edificio 4). Estos datos vienen del PME (medición en tiempo real,
sólo desde 2026) para el consumo del tablero de arriba; las líneas base en cambio son un modelo estadístico
propio, construido con el consumo mensual real de 2025 (año base) contra los grados-día de refrigeración (CDD):

- **LBEn** (línea base esperada) = pendiente × CDD + intercepto, una regresión lineal ajustada con los 12
  meses de 2025.
- **Desvío** = (consumo real − LBEn) / LBEn, y su desvío estándar (de los 12 desvíos del año base) define
  el ancho de la banda de control: **LCS/LCI** = LBEn ± LBEn × desvío estándar.
- Cuando el consumo real supera el **LCS**, es la señal para que el equipo de gestión de la energía
  investigue la causa raíz de ese mes.

`js/lbe.js` recalcula todo esto (LBEn, desvío, LCS/LCI, R²) a partir de `js/linea-base.js` (pendiente,
intercepto y consumo/CDD mensual de cada SET), con las mismas fórmulas de la planilla original, para no
duplicar números que puedan desincronizarse.
