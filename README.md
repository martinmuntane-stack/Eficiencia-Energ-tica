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
index.html            Estructura de la página
css/app.css            Estilos (tema claro/oscuro)
js/app.js               Lógica: lectura de Excel, cálculos y gráficos
js/datos-ejemplo.js     Datos de ejemplo (generados desde data/BASE_POWER_ejemplo.xlsx)
js/vendor/               Chart.js y SheetJS (xlsx), vendorizados localmente
data/BASE_POWER_ejemplo.xlsx  Excel de ejemplo original
```
