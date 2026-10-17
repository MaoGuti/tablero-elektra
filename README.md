# Tablero de Elektra

¿Cuánto cuesta de verdad cada kilómetro de un híbrido enchufable en Bogotá? Este tablero lo responde con datos reales: la tarifa de la factura de energía de la casa, las cargas de la app My BMW y las tanqueadas.

## Qué muestra

- **Costo por kilómetro en eléctrico** frente a un carro a gasolina de referencia, con gasolina corriente (precio regulado de Bogotá, $16.377 en octubre de 2026, CREG) o extra (el de la última tanqueada). El carro de partida es una Renault Koleos 2.5: 30 km por galón en ciudad y 40 en carretera (los oficiales son 31,5 y 47). El rendimiento y el precio se pueden cambiar.
- **Precio real del kWh**, con la contribución del estrato, según el periodo de cada factura.
- **Consumo eléctrico** medido entre cargas, **km hechos con batería** y **ahorro aproximado**.
- **En Bogotá frente a las salidas**: separa cada tramo entre dos cargas en km con batería y km con gasolina. Los tramos de 120 km o más cuentan como salidas de la ciudad. Muestra el costo por km de cada uno, las salidas encontradas y el rendimiento con gasolina.
- **Mes a mes**: lo que se fue en energía y en gasolina.
- **El carro dentro del consumo de la casa**: qué parte de cada factura se lleva.
- **Hábito de carga**: con cuánta batería llega a enchufarse.

## Privacidad

El Excel se lee en el navegador y no se sube a ningún servidor. La página no guarda nada. El botón "Ver con datos de ejemplo" usa datos inventados.

## El Excel que necesita

Tres hojas, con la fila de encabezados en cualquier fila:

| Hoja | Columnas |
|---|---|
| Cargas | Enchufado, Odómetro km (u Odómetro mi), % al enchufar, % al desenchufar, kWh exactos (o Costo app y Tarifa app $/kWh) |
| Tarifas | Periodo desde, Periodo hasta, Precio final $/kWh (o Tarifa $/kWh y Contribución), Consumo kWh |
| Gasolina | Fecha y hora, Galones, Precio por galón, Valor, Valor neto (o Descuento TC e Intereses TC) |

## Cómo correrlo

Es una página estática: `index.html`, `styles.css` y `app.js`. Lee el Excel con [SheetJS](https://sheetjs.com/). Para verla en el computador:

```bash
python -m http.server 8765
```

y abrir `http://localhost:8765`.

## Cómo se hizo

Tarea de la sesión 5 de AI Second Brain (30X · Mycelium AI), con las 4D: descubrir, diseñar, desarrollar y desplegar. Lo construyó Claude Code a partir de un segundo cerebro en Obsidian; se publica con GitHub y Vercel.
