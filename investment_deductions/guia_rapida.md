# Deducción de Inversiones LISR — Guía Rápida

**Plantilla de Cálculo Fiscal | Artículos 31 al 38 de la Ley del ISR**

---

## Contenido de la plantilla

Tu archivo contiene 7 hojas de trabajo:

| Hoja | Función |
|------|---------|
| **Instrucciones** | Guía de uso integrada en la plantilla |
| **Config** | Ejercicio fiscal y topes de automóviles |
| **Catálogo** | Porcentajes de deducción (Art. 33, 34 y 35) |
| **Inversiones** | Hoja principal de cálculo |
| **Resumen** | Totales agrupados por tipo de bien |
| **Baja_Activos** | Ganancia o pérdida por venta de activos |
| **INPC** | Índices Nacionales de Precios al Consumidor (1984–2025) |

---

## Paso 1 — Configura el ejercicio fiscal

Abre la hoja **Config** y verifica que el campo **Ejercicio** tenga el año correcto (por defecto: 2025).

| Parámetro | Valor | Descripción |
|-----------|-------|-------------|
| Ejercicio | 2025 | Año fiscal para todos los cálculos |
| TOPE_AUTO_COMBUSTION | 175,000 | Tope MOI deducible para autos de combustión |
| TOPE_AUTO_ELECTRICO | 250,000 | Tope MOI deducible para autos eléctricos/híbridos |

> Los topes se actualizan automáticamente en la hoja Inversiones. Si el SAT modifica estos valores en el futuro, solo necesitas cambiarlos aquí.

---

## Paso 2 — Registra tus inversiones

En la hoja **Inversiones**, llena únicamente estas columnas:

| Columna | Qué capturar |
|---------|-------------|
| **A** — No. | Número de control o cuenta contable |
| **B** — Cuenta Contable | Código contable del activo |
| **C** — Concepto | Descripción del bien |
| **D** — Fecha de Adquisición | Fecha en formato DD/MM/AAAA |
| **E** — M.O.I. | Monto Original de la Inversión (sin IVA) |
| **G** — Tipo de Bien | Selecciona del menú desplegable |

**Todo lo demás se calcula automáticamente**, incluyendo:
- MOI Deducible (aplica topes de automóviles)
- Porcentaje de deducción (del catálogo)
- Meses de uso en el ejercicio
- Deducción del ejercicio
- INPC y factor de actualización (a 4 decimales)
- Deducción actualizada
- Depreciación acumulada de ejercicios anteriores
- Saldo pendiente de deducir

---

## Ejemplo práctico — 5 activos

A continuación un ejemplo con 5 activos que muestran los diferentes escenarios que la plantilla maneja:

### Activo 1: Laptop Dell Latitude 5540

| Campo | Valor |
|-------|-------|
| Fecha de Adquisición | 15/03/2023 |
| M.O.I. | $32,000 |
| Tipo de Bien | Equipo de cómputo |

**Resultado**: La plantilla calcula automáticamente el 30% de deducción, 12 meses de uso (tercer año completo), y una depreciación acumulada de $16,800 por los 21 meses de uso en 2023–2024.

| Columna | Valor calculado |
|---------|----------------|
| MOI Deducible | $32,000 |
| % Deducción | 30% |
| Meses de Uso | 12 |
| Deducción del Ejercicio | $9,600 |
| Dep. Acumulada | $16,800 |
| Saldo Pendiente | $5,600 |

---

### Activo 2: Escritorio y silla ejecutiva

| Campo | Valor |
|-------|-------|
| Fecha de Adquisición | 01/06/2024 |
| M.O.I. | $18,500 |
| Tipo de Bien | Mobiliario y equipo de oficina |

**Resultado**: Porcentaje bajo del 10% (vida útil de 10 años). La depreciación acumulada corresponde a 6 meses de uso en 2024.

| Columna | Valor calculado |
|---------|----------------|
| MOI Deducible | $18,500 |
| % Deducción | 10% |
| Meses de Uso | 12 |
| Deducción del Ejercicio | $1,850 |
| Dep. Acumulada | $925 |
| Saldo Pendiente | $15,725 |

---

### Activo 3: Toyota Corolla 2024 (combustión)

| Campo | Valor |
|-------|-------|
| Fecha de Adquisición | 10/01/2024 |
| M.O.I. | $420,000 |
| Tipo de Bien | Automóvil (combustión) |

**Resultado**: El MOI de $420,000 excede el tope de $175,000 para automóviles de combustión. La plantilla recorta automáticamente el MOI Deducible. El contribuyente "pierde" $245,000 de deducción.

| Columna | Valor calculado |
|---------|----------------|
| MOI Deducible | **$175,000** (tope aplicado) |
| % Deducción | 25% |
| Meses de Uso | 12 |
| Deducción del Ejercicio | $43,750 |
| Dep. Acumulada | $40,104 |
| Saldo Pendiente | $91,146 |

> **Nota**: Si el vehículo fuera una pick-up (camión de carga), no tendría tope y el MOI Deducible sería $420,000 completos (Criterio 27/ISR/N).

---

### Activo 4: BYD Dolphin Mini 2025 (eléctrico)

| Campo | Valor |
|-------|-------|
| Fecha de Adquisición | 20/02/2025 |
| M.O.I. | $299,000 |
| Tipo de Bien | Automóvil (eléctrico/híbrido) |

**Resultado**: El tope para vehículos eléctricos/híbridos es de $250,000. Como es un activo nuevo adquirido en febrero 2025, solo tiene 10 meses de uso (marzo a diciembre) y la depreciación acumulada es $0.

| Columna | Valor calculado |
|---------|----------------|
| MOI Deducible | **$250,000** (tope eléctrico aplicado) |
| % Deducción | 25% |
| Meses de Uso | 10 (primer año parcial) |
| Deducción del Ejercicio | $52,083 |
| Dep. Acumulada | $0 |
| Saldo Pendiente | $197,917 |

---

### Activo 5: Servidor HPE ProLiant DL380

| Campo | Valor |
|-------|-------|
| Fecha de Adquisición | 01/07/2022 |
| M.O.I. | $95,000 |
| Tipo de Bien | Equipo de cómputo |

**Resultado**: Equipo de cómputo con vida útil de 40 meses (30%). Este servidor ya acumula 29 meses de uso en ejercicios anteriores, por lo que solo le quedan 11 meses de vida útil en 2025. Al terminar el ejercicio, queda totalmente depreciado.

| Columna | Valor calculado |
|---------|----------------|
| MOI Deducible | $95,000 |
| % Deducción | 30% |
| Meses de Uso | 11 (últimos meses de vida útil) |
| Deducción del Ejercicio | $26,125 |
| Dep. Acumulada | $68,875 |
| Saldo Pendiente | $0 |

---

## Paso 3 — Consulta el resumen

La hoja **Resumen** agrupa automáticamente los totales por tipo de bien. Con los 5 activos de ejemplo se ve así:

| Tipo de Bien | Cant. | MOI Deducible | Deducción del Ejercicio | Saldo Pendiente |
|--------------|-------|---------------|------------------------|-----------------|
| Equipo de cómputo | 2 | $127,000 | $35,725 | $5,600 |
| Mobiliario y equipo de oficina | 1 | $18,500 | $1,850 | $15,725 |
| Automóvil (combustión) | 1 | $175,000 | $43,750 | $91,146 |
| Automóvil (eléctrico/híbrido) | 1 | $250,000 | $52,083 | $197,917 |
| **TOTAL** | **5** | **$570,500** | **$133,408** | **$310,388** |

> Estos totales son útiles para tu declaración anual y reportes financieros.

---

## Paso 4 — Baja de activos (si aplica)

Si vendes o das de baja un activo, usa la hoja **Baja_Activos**. Ejemplo:

**Impresora Multifuncional HP — vendida en junio 2025**

| Campo | Valor |
|-------|-------|
| Concepto | Impresora Multifuncional HP LaserJet Pro |
| MOI Deducible | $15,000 |
| Deducciones Acumuladas | $15,000 |
| INPC Enajenación (Jun 2025) | 140.405 |
| INPC Adquisición (Mar 2021) | 111.824 |
| Precio de Venta (sin IVA) | $3,500 |

| Cálculo | Resultado |
|---------|-----------|
| Saldo Pendiente | $0 (totalmente depreciado) |
| Factor de Actualización | 1.2556 |
| Saldo Actualizado | $0 |
| Ganancia / Pérdida | **$3,500** |
| Resultado | **Ganancia Acumulable** |

> La ganancia es ingreso acumulable para ISR. Si el resultado fuera negativo, sería una pérdida deducible.

---

## Paso 5 — Actualizar el INPC

La hoja **INPC** contiene datos desde 1984 hasta 2025. Para agregar un año nuevo:

1. Ve a la última fila (2025)
2. Inserta una fila debajo
3. Escribe el año (2026) en la columna A
4. Llena los valores mensuales conforme el INEGI los publique

> Los valores de INPC se publican quincenalmente en [inegi.org.mx](https://www.inegi.org.mx/temas/inpc/).

---

## Topes de automóviles (Art. 36 LISR)

| Tipo de vehículo | Tope deducible |
|-------------------|---------------|
| Combustión interna | $175,000 MXN |
| Eléctrico o híbrido | $250,000 MXN |
| Pick-up (camión de carga) | Sin tope (100% deducible) |

Las pick-up se clasifican como camiones de carga conforme al Criterio 27/ISR/N del SAT.

---

## Notas importantes

- El **IVA no forma parte del MOI** (es acreditable), salvo que no tengas derecho al acreditamiento.
- Si **no deduces en el ejercicio de inicio de uso** ni en el siguiente, pierdes esos montos de forma permanente.
- Puedes aplicar un **porcentaje menor al máximo**, pero queda fijo por 5 años (Art. 66 RLISR).
- El **Factor de Actualización** se calcula a 4 decimales conforme al Art. 9 del Reglamento de la LISR.
- La **Dep. Acumulada se calcula automáticamente**. Si tu depreciación real difiere por ajustes o porcentajes menores, puedes sobreescribir la fórmula directamente en la celda.
- Para bienes de **energía renovable** (100% deducible), el sistema debe operar al menos 5 años continuos.

---

## Soporte

Si tienes dudas sobre cómo usar esta plantilla, revisa el video tutorial disponible en el blog o responde al correo de confirmación de compra.

---

*Versión 1.0 | Marzo 2026*
