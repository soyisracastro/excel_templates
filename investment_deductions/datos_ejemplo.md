# Datos de Ejemplo - Deducción de Inversiones LISR

Datos para usar en el video tutorial y verificar que la plantilla funciona correctamente.
Ejercicio fiscal: **2025**.

---

## Hoja Config (ya prellenada)

| Parámetro | Valor |
|-----------|-------|
| Ejercicio | 2025 |
| INPC_AÑO_BASE | 1984 |
| TOPE_AUTO_COMBUSTION | 175,000 |
| TOPE_AUTO_ELECTRICO | 250,000 |

---

## Hoja Inversiones - 5 activos

Solo necesitas capturar 5 columnas: No., Cuenta Contable, Concepto, Fecha de Adquisición, M.O.I. y Tipo de Bien.
Todo lo demás se calcula automáticamente (incluyendo Dep. Acumulada).

### Activo 1: Equipo de cómputo (año completo)

| Campo | Valor |
|-------|-------|
| No. | 1 |
| Cuenta Contable | 1240-001 |
| Concepto | Laptop Dell Latitude 5540 |
| Fecha de Adquisición | 15/03/2023 |
| M.O.I. | 32,000 |
| Tipo de Bien | Equipo de cómputo |

**Qué demuestra**: Activo en su 3er año, 12 meses completos de uso, % = 30%.
**Valores calculados esperados**:
- MOI Deducible = $32,000 (sin tope)
- % Deducción = 30%
- Meses de Uso = 12 (año completo, aún dentro de vida útil de 40 meses)
- Deducción del Ejercicio = 32,000 × 30% × 12/12 = $9,600
- Dep. Acumulada = 32,000 × 30% × 21/12 = $16,800 (21 meses: 9 en 2023 + 12 en 2024)
- Saldo Pendiente = 32,000 - 16,800 - 9,600 = $5,600

---

### Activo 2: Mobiliario (porcentaje bajo)

| Campo | Valor |
|-------|-------|
| No. | 2 |
| Cuenta Contable | 1240-002 |
| Concepto | Escritorio y silla ejecutiva |
| Fecha de Adquisición | 01/06/2024 |
| M.O.I. | 18,500 |
| Tipo de Bien | Mobiliario y equipo de oficina |

**Qué demuestra**: Activo en su 2do año, % = 10% (vida útil 10 años), deducción pequeña.
**Valores calculados esperados**:
- % Deducción = 10%
- Meses de Uso = 12
- Deducción del Ejercicio = 18,500 × 10% × 12/12 = $1,850
- Dep. Acumulada = 18,500 × 10% × 6/12 = $925 (6 meses en 2024: Jul-Dic)
- Saldo Pendiente = 18,500 - 925 - 1,850 = $15,725

---

### Activo 3: Automóvil con tope de combustión

| Campo | Valor |
|-------|-------|
| No. | 3 |
| Cuenta Contable | 1240-003 |
| Concepto | Toyota Corolla 2024 (combustión) |
| Fecha de Adquisición | 10/01/2024 |
| M.O.I. | 420,000 |
| Tipo de Bien | Automóvil (combustión) |

**Qué demuestra**: MOI $420,000 > tope $175,000. La columna MOI Deducible recorta automáticamente a $175,000.
**Valores calculados esperados**:
- MOI Deducible = $175,000 (tope combustión aplicado)
- % Deducción = 25%
- Meses de Uso = 12
- Deducción del Ejercicio = 175,000 × 25% × 12/12 = $43,750
- Dep. Acumulada = 175,000 × 25% × 11/12 = $40,104.17 (11 meses en 2024: Feb-Dic)
- Saldo Pendiente = 175,000 - 40,104.17 - 43,750 = $91,145.83
- El contribuyente "pierde" $245,000 de deducción por el tope

---

### Activo 4: Automóvil eléctrico con tope + primer año parcial

| Campo | Valor |
|-------|-------|
| No. | 4 |
| Cuenta Contable | 1240-004 |
| Concepto | BYD Dolphin Mini 2025 (eléctrico) |
| Fecha de Adquisición | 20/02/2025 |
| M.O.I. | 299,000 |
| Tipo de Bien | Automóvil (eléctrico/híbrido) |

**Qué demuestra**: MOI $299,000 > tope $250,000. Tope eléctrico aplica. Primer año = meses parciales.
**Valores calculados esperados**:
- MOI Deducible = $250,000 (tope eléctrico aplicado)
- % Deducción = 25%
- Meses de Uso = 10 (Mar-Dic, 12 - 2 = 10)
- Deducción del Ejercicio = 250,000 × 25% × 10/12 = $52,083.33
- Dep. Acumulada = $0 (adquirido en el ejercicio actual)
- Saldo Pendiente = 250,000 - 0 - 52,083.33 = $197,916.67

---

### Activo 5: Servidor casi totalmente depreciado

| Campo | Valor |
|-------|-------|
| No. | 5 |
| Cuenta Contable | 1240-005 |
| Concepto | Servidor HPE ProLiant DL380 |
| Fecha de Adquisición | 01/07/2022 |
| M.O.I. | 95,000 |
| Tipo de Bien | Equipo de cómputo |

**Qué demuestra**: Equipo de cómputo al 30%, vida útil 40 meses. Adquirido Jul 2022, para 2025 ya lleva 29 meses acumulados → aún dentro de vida útil pero cerca del fin.
**Valores calculados esperados**:
- % Deducción = 30%
- Meses de Uso = 12 (aún dentro de los 40 meses de vida útil: 29 previos + 12 = 41 > 40, se limita a 40-29=11)
- Nota: La fórmula de meses calcula MIN(12, MAX(0, 40 - 29)) = MIN(12, 11) = 11
- Deducción del Ejercicio = 95,000 × 30% × 11/12 = $26,125
- Dep. Acumulada = 95,000 × 30% × 29/12 = $68,875 (29 meses: 5 en 2022 + 12 en 2023 + 12 en 2024)
- Saldo Pendiente = MAX(0, 95,000 - 68,875 - 26,125) = $0

---

## Hoja Baja_Activos - 1 ejemplo

### Impresora totalmente depreciada vendida con ganancia

| Campo | Valor |
|-------|-------|
| Concepto | Impresora Multifuncional HP LaserJet Pro |
| MOI Deducible | 15,000 |
| Deducciones Acumuladas | 15,000 |
| INPC Enajenación (Jun 2025) | 140.405 |
| INPC Adquisición (Mar 2021) | 111.824 |
| Precio de Venta (sin IVA) | 3,500 |

**Qué demuestra**: Activo 100% depreciado que se vende. Saldo pendiente = 0, toda la venta es ganancia acumulable.
**Resultado esperado**:
- Saldo Pendiente = 15,000 - 15,000 = $0
- Factor = 140.405 / 111.824 = 1.2556
- Saldo Actualizado = 0 × 1.2556 = $0
- Ganancia/Pérdida = 3,500 - 0 = $3,500
- Resultado: "Ganancia Acumulable"

---

## Hoja Resumen (automática)

Con los 5 activos anteriores, el Resumen mostrará:

| Tipo de Bien | Cant. | MOI Total | MOI Deducible | Deducción Ejercicio | Saldo Pendiente |
|--------------|-------|-----------|---------------|---------------------|-----------------|
| Equipo de cómputo | 2 | $127,000 | $127,000 | ~$35,725 | ~$5,600 |
| Mobiliario y equipo de oficina | 1 | $18,500 | $18,500 | ~$1,850 | ~$15,725 |
| Automóvil (combustión) | 1 | $420,000 | $175,000 | ~$43,750 | ~$91,146 |
| Automóvil (eléctrico/híbrido) | 1 | $299,000 | $250,000 | ~$52,083 | ~$197,917 |

---

## Hoja INPC

Ya prellenada (1984-2025). En el video, mostrar cómo agregar una fila para 2026:
1. Ir a la última fila (2025)
2. Insertar fila debajo
3. Escribir "2026" en columna A
4. Ir llenando los INPC mensuales conforme el INEGI los publique

---

## Notas para el video

- Capturar primero Config (verificar ejercicio 2025)
- Luego ir a Inversiones y capturar los 5 activos uno por uno
- **Destacar que Dep. Acumulada se calcula sola** (ya no es captura manual)
- Hacer zoom en la columna MOI Deducible del Toyota para mostrar el tope
- Mostrar el dropdown de Tipo de Bien
- Ir a Resumen y mostrar cómo se agrupa automáticamente
- Ir a Baja_Activos y capturar el ejemplo de la impresora (aquí sí es manual)
- Cerrar mostrando la hoja INPC y cómo se actualiza
