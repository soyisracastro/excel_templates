# Guion de Video Tutorial: Plantilla Deducción de Inversiones LISR

**Duración estimada:** 8-12 minutos
**Formato:** Screencast con voz en off
**Software:** Excel (escritorio)

---

## INTRO (0:00 - 0:30)

**[Pantalla: portada con título]**

> "En este video te voy a mostrar cómo usar la Plantilla de Deducción de Inversiones LISR paso a paso. Al terminar, vas a poder calcular la deducción fiscal actualizada de todos los activos de tu empresa en menos de 15 minutos."

**[Pantalla: abrir el archivo .xlsm en Excel]**

> "Lo primero: al abrir el archivo, Excel te va a pedir que habilites las macros. Haz clic en 'Habilitar contenido'. Esto es necesario para que funcione la hoja de instrucciones."

---

## ESCENA 1: Vista general del libro (0:30 - 1:30)

**[Pantalla: mostrar las pestañas del libro]**

> "El libro tiene 7 hojas. Vamos a recorrerlas rápidamente:"

**[Clic en cada pestaña mientras describes]**

> "**Instrucciones** — aquí tienes la guía completa de uso. Si en algún momento te pierdes, siempre puedes regresar aquí."
>
> "**Catálogo** — los porcentajes de deducción que marca la ley. Artículos 33, 34 y 35 de la LISR. No tienes que memorizarlos, la plantilla los busca automáticamente."
>
> "**Inversiones** — esta es la hoja principal donde vas a capturar tus activos y donde se hace todo el cálculo."
>
> "**Resumen** — te da los totales por categoría de activo. Útil para tu declaración anual."
>
> "**Baja_Activos** — si vendes o das de baja un activo, aquí calculas si hay ganancia o pérdida."
>
> "**INPC** — la tabla de Índices Nacionales de Precios al Consumidor desde 1984 hasta 2025."
>
> "**Config** — aquí configuras el ejercicio fiscal y los topes de automóviles."

---

## ESCENA 2: Configuración inicial (1:30 - 2:30)

**[Pantalla: hoja Config]**

> "Antes de capturar cualquier activo, ve a la hoja Config."

**[Señalar celda B2]**

> "Aquí tienes el campo 'Ejercicio'. Asegúrate de que tenga el año fiscal que estás calculando. Por ejemplo, si estás preparando tu declaración anual 2025, debe decir 2025."

> "Los topes de automóviles ya vienen configurados: 175 mil para combustión y 250 mil para eléctricos o híbridos. Si en algún momento la ley cambia estos montos, los actualizas aquí y todas las fórmulas se ajustan solas."

**[Señalar celda B3]**

> "El año base del INPC no lo toques — es para que las fórmulas sepan dónde buscar en la tabla de índices."

---

## ESCENA 3: Capturar un activo fijo (2:30 - 5:00)

**[Pantalla: hoja Inversiones, posicionarse en fila 5]**

> "Ahora vamos a la parte principal. Te voy a mostrar cómo capturar un activo con un ejemplo real."

### Ejemplo 1: Equipo de cómputo

**[Escribir en las celdas mientras narras]**

> "Supongamos que compraste una laptop para tu oficina el 15 de marzo de 2024 por 28 mil pesos más IVA."

| Celda | Valor | Narración |
|-------|-------|-----------|
| A5 | `EC-001` | "En número de control pones tu referencia interna." |
| B5 | `1206` | "La cuenta contable, si la manejas. Es opcional." |
| C5 | `Laptop Dell Latitude 5540` | "La descripción del bien." |
| D5 | `15/03/2024` | "La fecha de adquisición. Recuerda: sin IVA." |
| E5 | `28000` | "El Monto Original de la Inversión. Esto incluye el precio más fletes, instalación, lo que corresponda. Pero sin IVA, porque el IVA es acreditable." |

**[Hacer clic en G5, mostrar el dropdown]**

> "Ahora el tipo de bien. Haz clic en la celda y se abre un menú desplegable. Selecciona 'Equipo de cómputo'."

| Celda | Valor |
|-------|-------|
| G5 | `Equipo de cómputo` (seleccionar del dropdown) |

**[Pausa para que se vean las fórmulas calcularse]**

> "¿Ves lo que pasó? Automáticamente:"
>
> - "La columna F (MOI Deducible) muestra los mismos 28 mil porque no es automóvil, no tiene tope."
> - "La columna H (porcentaje) se llenó con 30%, que es lo que marca el Artículo 34 para equipo de cómputo."
> - "La columna I (meses de uso) calculó 9 meses. ¿Por qué? Porque marzo no cuenta como mes completo — se empieza a contar desde abril hasta diciembre."
> - "La columna J tiene la deducción del ejercicio: 28 mil por 30% por 9/12 igual a 6,300 pesos."
> - "Las columnas K y L buscaron los INPCs automáticamente en la tabla."
> - "La columna M tiene el factor de actualización a 4 decimales, como lo exige el Artículo 9 del Reglamento."
> - "Y la columna N es la deducción actualizada: la deducción multiplicada por el factor."

**[Señalar columna O — zoom in]**

> "Ahora, atención con la columna O: 'Dep. Acumulada'. Esta es la **única columna numérica que debes llenar a mano**. Aquí va la suma de todas las deducciones fiscales que ya aplicaste en años anteriores para ese activo."
>
> "¿De dónde sacas este dato? Tienes varias fuentes:"
>
> - "De tu **balanza de comprobación**: el saldo acumulado de la cuenta de depreciación fiscal."
> - "De tus **papeles de trabajo** de ejercicios anteriores."
> - "De tu **declaración anual** del año pasado."
> - "O si usaste **esta misma plantilla** el año pasado: sumas lo que tenías en Dep. Acumulada más la Deducción del Ejercicio de ese año. Ese total es lo que capturas aquí."
>
> "Para activos **nuevos** — que compraste este mismo ejercicio — déjala vacía o en cero."
>
> "Y para activos que **ya se depreciaron completamente**, la Dep. Acumulada debe ser igual al MOI Deducible. Así el saldo pendiente dará cero."
>
> "Si este dato está mal, el saldo pendiente de la columna P no va a cuadrar. Así que vale la pena tomarse unos minutos para verificarlo."

**[Señalar columna P]**

> "Y la columna P te dice cuánto te falta por deducir: el saldo pendiente. Es simplemente MOI Deducible menos Dep. Acumulada menos la Deducción del Ejercicio actual."

---

### Ejemplo 2: Automóvil con tope

**[Posicionarse en fila 6]**

> "Ahora un caso más interesante: un automóvil."

| Celda | Valor | Narración |
|-------|-------|-----------|
| C6 | `Nissan Sentra 2024` | "" |
| D6 | `01/06/2024` | "Comprado en junio de 2024." |
| E6 | `380000` | "El MOI son 380 mil pesos." |
| G6 | `Automóvil (combustión)` | "Seleccionamos del menú." |

**[Pausa dramática señalando columna F]**

> "Mira la columna F: en lugar de 380 mil, dice 175 mil. La plantilla aplicó automáticamente el tope del Artículo 36. Tú solo tienes que capturar el precio real y la plantilla hace el recorte."
>
> "Todo el cálculo de deducción, INPC y factor se hace sobre los 175 mil, no sobre los 380 mil. Exactamente como lo exige la ley."

---

### Ejemplo 3: Pick-up sin tope

**[Posicionarse en fila 7]**

> "¿Y si es una pick-up?"

| Celda | Valor |
|-------|-------|
| C7 | `Toyota Hilux 2024` |
| D7 | `15/01/2024` |
| E7 | `520000` |
| G7 | `Pick-up (camión de carga)` |

> "Mira: MOI deducible = 520 mil. Sin tope. Porque según el Criterio Normativo 27/ISR/N del SAT, las pick-up son camiones de carga, no automóviles. La plantilla ya lo sabe."

---

## ESCENA 4: Hoja de Resumen (5:00 - 6:00)

**[Pantalla: hoja Resumen]**

> "Conforme capturas activos, la hoja Resumen se actualiza automáticamente."
>
> "Te muestra por cada tipo de bien: cuántos activos tienes, el MOI total, la deducción del ejercicio, la deducción actualizada y el saldo pendiente."
>
> "Esto es lo que necesitas para llenar tu declaración anual o para entregar al contador."

---

## ESCENA 5: Catálogo de porcentajes (6:00 - 6:45)

**[Pantalla: hoja Catalogo]**

> "Si alguna vez tienes duda sobre qué porcentaje aplica, aquí lo tienes todo."
>
> "Está organizado en tres secciones:"
> - "Artículo 33: intangibles y diferidos."
> - "Artículo 34: activos fijos por tipo de bien."
> - "Artículo 35: maquinaria y equipo según la actividad de tu empresa."
>
> "Por ejemplo, si tienes un restaurante y compraste equipo de cocina, buscas 'Maq. - Restaurantes' y verás que el porcentaje es 20%."

---

## ESCENA 6: Baja de activos (6:45 - 8:00)

**[Pantalla: hoja Baja_Activos]**

> "Si vendes un activo antes de que se termine de deducir, necesitas calcular si hay ganancia acumulable o pérdida deducible. Para eso es esta hoja."

**[Capturar ejemplo]**

| Celda | Valor | Narración |
|-------|-------|-----------|
| A5 | `Laptop Dell` | "El bien que vendiste." |
| B5 | `28000` | "El MOI deducible original." |
| C5 | `14000` | "Lo que ya habías deducido en total." |
| E5 | `141.197` | "El INPC del mes en que vendiste. Lo consultas en la hoja INPC." |
| F5 | `134.065` | "El INPC del mes en que compraste." |
| I5 | `10000` | "El precio al que vendiste, sin IVA." |

**[Señalar resultados]**

> "La plantilla calcula:"
> - "Saldo pendiente: 28 mil menos 14 mil = 14 mil."
> - "Factor de actualización: el INPC de venta entre el de compra."
> - "Saldo actualizado: 14 mil por el factor."
> - "Ganancia o pérdida: precio de venta menos saldo actualizado."
> - "Y te dice si es 'Ganancia Acumulable' o 'Pérdida Deducible'."

---

## ESCENA 7: Actualizar el INPC (8:00 - 8:45)

**[Pantalla: hoja INPC]**

> "La tabla de INPC viene cargada desde 1984 hasta 2025."
>
> "Cuando se publiquen los índices de 2026, simplemente agregas una fila al final con el año y los valores mensuales conforme se vayan publicando. Las fórmulas de la hoja Inversiones los encontrarán automáticamente."
>
> "Los índices los publica el INEGI cada mes, generalmente los primeros 10 días del mes siguiente."

---

## ESCENA 8: Tips finales (8:45 - 9:30)

**[Pantalla: hoja Inversiones con varios activos capturados]**

> "Antes de cerrar, tres tips importantes:"
>
> "**Uno:** Si en un ejercicio anterior usaste un porcentaje menor al máximo, captúralo directamente en la columna H sobreescribiendo la fórmula. La ley te obliga a mantenerlo por 5 años."
>
> "**Dos:** La columna 'Dep. Acumulada' (columna O) es clave si ya tenías activos de ejercicios anteriores. Captura ahí la suma de todo lo que ya habías deducido. Sin este dato, el saldo pendiente no será correcto."
>
> "**Tres:** Si necesitas más de 50 activos, simplemente copia las fórmulas de la última fila hacia abajo. Las fórmulas son relativas y funcionarán igual."

---

## CIERRE (9:30 - 10:00)

**[Pantalla: vista general del libro]**

> "Eso es todo. Con esta plantilla, el cálculo de la deducción de inversiones que antes te tomaba horas, ahora lo resuelves en minutos."
>
> "Si tienes dudas, revisa la hoja de Instrucciones dentro del mismo archivo. Y si quieres profundizar en la teoría, en la descripción te dejo el enlace a la guía completa sobre deducción de inversiones en la LISR."
>
> "Nos vemos en el siguiente video."

---

## Notas de producción

### Resolución y formato
- Grabar en **1920×1080** (Full HD)
- Excel al **100% de zoom**, tema claro
- Fuente del sistema: Aptos (ya viene en la plantilla)

### Recursos en pantalla
- Resaltar celdas con un rectángulo amarillo semitransparente al mencionarlas
- Usar zoom in cuando muestres fórmulas o resultados específicos
- Mostrar brevemente la barra de fórmulas cuando expliques qué hace cada columna

### Música
- Música de fondo suave, tipo lo-fi o corporativa
- Bajar volumen durante las explicaciones

### Thumbnail sugerido
- Texto: "Deducción de Inversiones en 15 min"
- Captura de la hoja Inversiones con datos de ejemplo
- Logo de Excel en una esquina

### Plataformas
- YouTube (SEO: "deducción de inversiones LISR Excel plantilla tutorial")
- Embed en la página de venta de la plantilla
- Fragmentos cortos (30-60 seg) para Instagram/TikTok de cada escena
