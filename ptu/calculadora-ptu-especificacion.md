# Especificación Técnica: Calculadora de PTU con Cálculo de ISR

## Resumen Ejecutivo

Este documento contiene toda la información necesaria para desarrollar una calculadora de PTU (Participación de los Trabajadores en las Utilidades) que permita:
1. Calcular la PTU individual de cada trabajador
2. Determinar el ISR a retener usando dos mecánicas diferentes (Ley vs Reglamento)
3. Comparar ambas mecánicas para que el usuario elija la más conveniente

---

## 1. INPUTS DEL USUARIO

### 1.1 Datos de la Empresa
| Campo | Tipo | Descripción | Validación |
|-------|------|-------------|------------|
| `utilidad_fiscal` | number | Utilidad fiscal de la declaración anual | > 0 |
| `ptu_no_cobrada` | number | PTU de ejercicios anteriores no cobrada (prescribe en 1 año) | >= 0 |
| `ejercicio_fiscal` | number | Año del ejercicio (ej: 2024) | 2021-actual |
| `fecha_pago` | date | Fecha en que se pagará la PTU | Dentro de 60 días después de declaración anual |

### 1.2 Datos por Trabajador
| Campo | Tipo | Descripción | Validación |
|-------|------|-------------|------------|
| `nombre` | string | Nombre del trabajador | Requerido |
| `rfc` | string | RFC del trabajador | 13 caracteres |
| `curp` | string | CURP del trabajador | 18 caracteres |
| `fecha_inicio_laboral` | date | Fecha de inicio de la relación laboral | Fecha válida |
| `salario_diario_base` | number | Salario diario ordinario (sin variables ni prestaciones extraordinarias) | > 0 |
| `dias_trabajados` | number | Días efectivamente trabajados en el ejercicio | 1-365 |
| `percepcion_anual` | number | Total de salarios devengados en el año | > 0 |
| `ptu_año_1` | number | PTU recibida hace 3 años (ej: 2021 para pago 2024) | >= 0 |
| `ptu_año_2` | number | PTU recibida hace 2 años (ej: 2022 para pago 2024) | >= 0 |
| `ptu_año_3` | number | PTU recibida hace 1 año (ej: 2023 para pago 2024) | >= 0 |
| `es_trabajador_confianza` | boolean | Si es trabajador de confianza | true/false |
| `ingreso_mensual_ordinario` | number | Sueldo mensual del mes de pago de PTU | > 0 |
| `isr_mensual_ordinario` | number | ISR que corresponde al sueldo mensual ordinario | >= 0 |

### 1.3 Datos del Sistema (Constantes por Año)
| Campo | Valor 2025 | Descripción |
|-------|------------|-------------|
| `uma_diaria` | 113.14 | Unidad de Medida y Actualización diaria |
| `smg_diario` | 278.80 | Salario Mínimo General diario (zona no fronteriza) |
| `smg_frontera` | 419.88 | Salario Mínimo General zona fronteriza |
| `porcentaje_ptu` | 10% | Porcentaje fijo por ley |
| `dias_exencion` | 15 | Días de exención para ISR |

---

## 2. MECÁNICA DE CÁLCULO DE PTU (LEY FEDERAL DEL TRABAJO)

### 2.1 Base Legal
- **Constitución**: Art. 123, fracción IX
- **LFT**: Artículos 117 al 131
- **Porcentaje**: 10% de la renta gravable (Art. 120 LFT)

### 2.2 Fórmulas de Distribución

La PTU se divide en **dos partes iguales** (Art. 123 LFT):

#### Parte 1: Por Días Trabajados (50%)
```
PTU_dias_repartir = PTU_total * 0.50

factor_dias_trabajador = dias_trabajados_trabajador / suma_total_dias_todos_trabajadores

PTU_dias_trabajador = PTU_dias_repartir * factor_dias_trabajador
```

#### Parte 2: Por Salarios Devengados (50%)
```
PTU_salarios_repartir = PTU_total * 0.50

factor_salarios_trabajador = percepcion_anual_trabajador / suma_total_percepciones_todos_trabajadores

PTU_salarios_trabajador = PTU_salarios_repartir * factor_salarios_trabajador
```

#### PTU Total del Trabajador (antes de tope)
```
PTU_bruta_trabajador = PTU_dias_trabajador + PTU_salarios_trabajador
```

### 2.3 Tope Máximo (Art. 127, fracción VIII LFT - Reforma 2021)

El tope es el **mayor** entre:

**Opción A: 3 meses de salario**
```
tope_3_meses = salario_diario_base * 30.4 * 3
// Equivalente a: salario_diario_base * 91.2
```

**Opción B: Promedio de PTU de los últimos 3 años**
```
promedio_3_años = (ptu_año_1 + ptu_año_2 + ptu_año_3) / 3
```

**Tope máximo aplicable:**
```
monto_maximo = MAX(tope_3_meses, promedio_3_años)
```

**PTU Real a Repartir:**
```
PTU_real = MIN(PTU_bruta_trabajador, monto_maximo)
```

### 2.4 Trabajadores con menos de 3 años de antigüedad

Según criterio de SCJN (Amparo 633/2023):
- Se debe considerar el **promedio de PTU de los últimos 3 años de la categoría o puesto** que ocupa actualmente el trabajador
- La falta de antigüedad NO debe perjudicar al empleado

### 2.5 Trabajadores de Confianza (Art. 127, fracción II LFT)

Si el salario del trabajador de confianza es **mayor** al del trabajador sindicalizado (o de planta) de más alto salario:
```
salario_tope_confianza = salario_mas_alto_planta * 1.20  // +20%

// Para el cálculo de la parte proporcional por salarios, 
// se usa el salario_tope_confianza como máximo
```

---

## 3. MECÁNICA DE ISR SOBRE PTU

### 3.1 Exención (Art. 93, fracción XIV LISR)

**Parte Exenta de ISR:**
```
// Según SAT (criterio vigente):
PTU_exenta = uma_diaria * 15  // = 113.14 * 15 = $1,697.10 (2025)

// Según PRODECON (criterio más favorable):
PTU_exenta = smg_diario * 15  // = 278.80 * 15 = $4,182.00 (2025)
```

**Nota importante:** En 2024, PRODECON emitió criterio sustantivo indicando que debe usarse SMG, no UMA. Sin embargo, el SAT aún usa UMA en sus sistemas. La calculadora debe permitir seleccionar cuál criterio usar.

**Parte Gravada:**
```
PTU_gravada = PTU_real - PTU_exenta

// Si el resultado es negativo, PTU_gravada = 0
PTU_gravada = MAX(0, PTU_gravada)
```

### 3.2 Método 1: Procedimiento General (Art. 96 LISR) - OBLIGATORIO

Este método es el estándar para cualquier pago de nómina.

**Paso 1:** Calcular base gravable del mes
```
base_gravable_mes = ingreso_mensual_ordinario + PTU_gravada
```

**Paso 2:** Aplicar tarifa mensual Art. 96 LISR
```
// Ubicar en tarifa según base_gravable_mes
ISR_total_mes = calcular_isr_tarifa_96(base_gravable_mes)
```

**Paso 3:** Determinar ISR solo del sueldo ordinario
```
ISR_ordinario = calcular_isr_tarifa_96(ingreso_mensual_ordinario)
```

**Paso 4:** ISR a retener por PTU
```
ISR_PTU_art96 = ISR_total_mes - ISR_ordinario
```

**Características:**
- ✅ No genera diferencias en ajuste anual
- ❌ Retención alta en el momento del pago
- ❌ Puede empujar al trabajador a rangos superiores de tarifa

### 3.3 Método 2: Procedimiento Opcional (Art. 174 RLISR) - REGLAMENTO

Este método prorratea el efecto de la PTU a lo largo del año.

**Fracción I - Calcular PTU promedio mensual:**
```
PTU_promedio_mensual = (PTU_gravada / 365) * 30.4
```

**Fracción II - Sumar al ingreso mensual ordinario:**
```
base_promediada = ingreso_mensual_ordinario + PTU_promedio_mensual
```

**Fracción III - Aplicar tarifa Art. 96 (sin subsidio al empleo):**
```
ISR_base_promediada = calcular_isr_tarifa_96_sin_subsidio(base_promediada)
ISR_ordinario_sin_subsidio = calcular_isr_tarifa_96_sin_subsidio(ingreso_mensual_ordinario)

diferencia_ISR = ISR_base_promediada - ISR_ordinario_sin_subsidio
```

**Fracción IV - Calcular tasa efectiva:**
```
tasa_efectiva = (diferencia_ISR / PTU_promedio_mensual) * 100
// Expresada en porcentaje
```

**Fracción V - Aplicar tasa a PTU total gravada:**
```
ISR_PTU_art174 = PTU_gravada * (tasa_efectiva / 100)
```

**Características:**
- ✅ Retención menor en el momento del pago
- ✅ Mayor liquidez inmediata para el trabajador
- ❌ Puede generar ISR a cargo en declaración anual
- ❌ No considera subsidio al empleo en el cálculo

### 3.4 Comparación de Métodos

| Aspecto | Art. 96 LISR (Ley) | Art. 174 RLISR (Reglamento) |
|---------|--------------------|-----------------------------|
| Obligatoriedad | Obligatorio | Opcional |
| Retención inmediata | Alta | Baja |
| Ajuste anual | Sin diferencias | Posible ISR a cargo |
| Subsidio al empleo | Sí aplica | No aplica |
| Beneficio | Patrón (sin adeudos de empleados) | Trabajador (mayor neto inmediato) |

### 3.5 Ejemplo Numérico Comparativo

**Datos del trabajador:**
- Salario mensual ordinario: $8,780.43
- PTU Real a recibir: $83,845.46
- PTU Exenta (15 UMA): $1,697.10
- PTU Gravada: $82,148.36
- ISR mensual ordinario: $166.75

**Cálculo Art. 174 RLISR (del archivo ejemplo):**
```
1. PTU promedio mensual = (82,148.36 / 365) * 30.4 = $6,983.29

2. Base promediada = 8,780.43 + 6,983.29 = $15,763.73

3. ISR base promediada (tarifa Art. 96):
   - Límite inferior: $15,487.72
   - Excedente: $276.01
   - Tasa marginal: 21.36%
   - ISR marginal: $58.95
   - Cuota fija: $1,640.18
   - ISR total: $1,699.13

4. Diferencia ISR = 1,699.13 - 166.75 = $1,532.38

5. Tasa efectiva = (1,532.38 / 6,983.29) * 100 = 21.94%

6. ISR PTU = 82,148.36 * 0.2194 = $18,026.29

7. PTU Neta = 83,845.46 - 18,026.29 = $65,819.17
```

---

## 4. TARIFA ISR Art. 96 LISR (Mensual 2024-2025)

| Límite Inferior | Límite Superior | Cuota Fija | % s/excedente |
|-----------------|-----------------|------------|---------------|
| 0.01 | 746.04 | 0.00 | 1.92% |
| 746.05 | 6,332.05 | 14.32 | 6.40% |
| 6,332.06 | 11,128.01 | 371.83 | 10.88% |
| 11,128.02 | 12,935.82 | 893.63 | 16.00% |
| 12,935.83 | 15,487.71 | 1,182.88 | 17.92% |
| 15,487.72 | 31,236.49 | 1,640.18 | 21.36% |
| 31,236.50 | 49,233.00 | 5,004.12 | 23.52% |
| 49,233.01 | 93,993.90 | 9,236.89 | 30.00% |
| 93,993.91 | 125,325.20 | 22,665.17 | 32.00% |
| 125,325.21 | 375,975.61 | 32,691.18 | 34.00% |
| 375,975.62 | En adelante | 117,912.32 | 35.00% |

**Fórmula de cálculo:**
```
ISR = ((base_gravable - limite_inferior) * porcentaje) + cuota_fija
```

---

## 5. VALIDACIONES Y RESTRICCIONES LEGALES

### 5.1 Trabajadores con Derecho a PTU
- ✅ Trabajadores de planta
- ✅ Trabajadores eventuales con **mínimo 60 días trabajados** en el año
- ✅ Ex trabajadores (proporcional a días laborados)
- ✅ Trabajadores de confianza (con tope de salario)
- ✅ Madres en período pre/postnatal (se consideran en servicio activo)
- ✅ Trabajadores con incapacidad temporal por riesgo de trabajo

### 5.2 Trabajadores SIN Derecho a PTU
- ❌ Directores, administradores y gerentes generales
- ❌ Trabajadores del hogar
- ❌ Trabajadores eventuales con menos de 60 días trabajados

### 5.3 Empresas Exentas de Repartir PTU (Art. 126 LFT)
- Empresas de nueva creación (primer año)
- IMSS e instituciones públicas descentralizadas con fines culturales, asistenciales o de beneficencia
- Instituciones de asistencia privada
- Empresas con capital menor al fijado por la STPS
- Empresas sin utilidad fiscal

### 5.4 Fechas Límite de Pago
| Tipo de Patrón | Fecha Límite |
|----------------|--------------|
| Personas Morales | 30 de mayo |
| Personas Físicas | 29 de junio |

### 5.5 Prescripción
- La PTU no cobrada **prescribe en 1 año** contado a partir del día siguiente a la fecha límite de pago
- La PTU no cobrada debe sumarse al monto a repartir del siguiente ejercicio

---

## 6. ESTRUCTURA DE DATOS SUGERIDA

### 6.1 Modelo de Empresa
```typescript
interface Empresa {
  nombre: string;
  rfc: string;
  nss: string;
  ejercicio: number;
  fechaPago: Date;
  tipoPersona: 'moral' | 'fisica';
  
  // Montos PTU
  utilidadFiscal: number;
  ptuGenerada: number;        // utilidadFiscal * 0.10
  ptuNoCobrada: number;       // De ejercicios anteriores
  ptuARepartir: number;       // ptuGenerada + ptuNoCobrada
  ptuDiasTrabajados: number;  // ptuARepartir * 0.50
  ptuPercepcionAnual: number; // ptuARepartir * 0.50
  
  // Constantes del ejercicio
  umaDiaria: number;
  smgDiario: number;
  usarUmaParaExencion: boolean;  // true = SAT, false = PRODECON
}
```

### 6.2 Modelo de Trabajador
```typescript
interface Trabajador {
  // Datos personales
  nombre: string;
  rfc: string;
  curp: string;
  fechaInicioLaboral: Date;
  
  // Datos laborales
  salarioDiarioBase: number;
  diasTrabajados: number;
  percepcionAnual: number;
  esTrabajadorConfianza: boolean;
  
  // Historial PTU (para tope)
  ptuAño1: number;
  ptuAño2: number;
  ptuAño3: number;
  
  // Datos para ISR
  ingresoMensualOrdinario: number;
  isrMensualOrdinario: number;
}
```

### 6.3 Modelo de Resultado
```typescript
interface ResultadoPTU {
  trabajador: Trabajador;
  
  // Cálculo PTU
  factorDias: number;
  ptuPorDias: number;
  factorSalarios: number;
  ptuPorSalarios: number;
  ptuBruta: number;
  
  // Topes
  tope3Meses: number;
  promedioUltimos3Años: number;
  montoMaximo: number;
  ptuReal: number;
  
  // Exención ISR
  ptuExenta: number;
  ptuGravada: number;
  
  // ISR Art. 96 (Ley)
  ptuElevadaAlMes_art96: number;
  baseGravable_art96: number;
  isrTotal_art96: number;
  isrPtu_art96: number;
  tasaEfectiva_art96: number;
  ptuNeta_art96: number;
  
  // ISR Art. 174 (Reglamento)
  ptuPromedioMensual_art174: number;
  basePromediada_art174: number;
  isrBasePromediada_art174: number;
  diferenciaIsr_art174: number;
  tasaEfectiva_art174: number;
  isrPtu_art174: number;
  ptuNeta_art174: number;
  
  // Comparación
  metodoRecomendado: 'art96' | 'art174';
  diferenciaIsrMetodos: number;
}
```

---

## 7. FLUJO DE CÁLCULO

```
┌─────────────────────────────────────────┐
│    1. DATOS DE ENTRADA DE EMPRESA       │
│    - Utilidad fiscal                    │
│    - PTU no cobrada                     │
│    - Fecha de pago                      │
└───────────────────┬─────────────────────┘
                    ▼
┌─────────────────────────────────────────┐
│    2. CALCULAR PTU A REPARTIR           │
│    ptuGenerada = utilidad * 10%         │
│    ptuTotal = ptuGenerada + noCobrada   │
│    ptuDias = ptuTotal * 50%             │
│    ptuSalarios = ptuTotal * 50%         │
└───────────────────┬─────────────────────┘
                    ▼
┌─────────────────────────────────────────┐
│    3. POR CADA TRABAJADOR               │
│    - Calcular factores                  │
│    - Calcular PTU días + PTU salarios   │
│    - Aplicar tope (3 meses o promedio)  │
│    - Determinar PTU real                │
└───────────────────┬─────────────────────┘
                    ▼
┌─────────────────────────────────────────┐
│    4. CALCULAR EXENCIÓN ISR             │
│    - PTU exenta (15 UMA o 15 SMG)       │
│    - PTU gravada                        │
└───────────────────┬─────────────────────┘
                    ▼
┌─────────────────────────────────────────┐
│    5. CALCULAR ISR MÉTODO ART. 96       │
│    - Sumar PTU gravada a sueldo mes     │
│    - Aplicar tarifa                     │
│    - Restar ISR del sueldo ordinario    │
└───────────────────┬─────────────────────┘
                    ▼
┌─────────────────────────────────────────┐
│    6. CALCULAR ISR MÉTODO ART. 174      │
│    - Calcular PTU promedio mensual      │
│    - Sumar a sueldo ordinario           │
│    - Obtener tasa efectiva              │
│    - Aplicar tasa a PTU gravada         │
└───────────────────┬─────────────────────┘
                    ▼
┌─────────────────────────────────────────┐
│    7. COMPARAR Y MOSTRAR RESULTADOS     │
│    - PTU neta con cada método           │
│    - ISR con cada método                │
│    - Recomendar método más favorable    │
└─────────────────────────────────────────┘
```

---

## 8. FÓRMULAS COMPLETAS EN PSEUDOCÓDIGO

```javascript
// ===========================================
// CÁLCULO PRINCIPAL DE PTU
// ===========================================

function calcularPTU(empresa, trabajadores) {
  // Paso 1: Calcular totales
  const sumaDiasTrabajados = trabajadores.reduce((sum, t) => sum + t.diasTrabajados, 0);
  const sumaPercepciones = trabajadores.reduce((sum, t) => sum + t.percepcionAnual, 0);
  
  // Paso 2: PTU a repartir
  const ptuGenerada = empresa.utilidadFiscal * 0.10;
  const ptuTotal = ptuGenerada + empresa.ptuNoCobrada;
  const ptuPorDias = ptuTotal * 0.50;
  const ptuPorSalarios = ptuTotal * 0.50;
  
  // Paso 3: Calcular para cada trabajador
  return trabajadores.map(trabajador => {
    // Factor y PTU por días
    const factorDias = trabajador.diasTrabajados / sumaDiasTrabajados;
    const ptuDiasTrabajador = ptuPorDias * factorDias;
    
    // Factor y PTU por salarios
    const factorSalarios = trabajador.percepcionAnual / sumaPercepciones;
    const ptuSalariosTrabajador = ptuPorSalarios * factorSalarios;
    
    // PTU bruta
    const ptuBruta = ptuDiasTrabajador + ptuSalariosTrabajador;
    
    // Calcular topes
    const tope3Meses = trabajador.salarioDiarioBase * 30.4 * 3;
    const promedio3Años = (trabajador.ptuAño1 + trabajador.ptuAño2 + trabajador.ptuAño3) / 3;
    const montoMaximo = Math.max(tope3Meses, promedio3Años);
    
    // PTU real (aplicando tope)
    const ptuReal = Math.min(ptuBruta, montoMaximo);
    
    // Calcular ISR
    const resultadoISR = calcularISR(trabajador, ptuReal, empresa);
    
    return {
      trabajador,
      factorDias,
      ptuDiasTrabajador,
      factorSalarios,
      ptuSalariosTrabajador,
      ptuBruta,
      tope3Meses,
      promedio3Años,
      montoMaximo,
      ptuReal,
      ...resultadoISR
    };
  });
}

// ===========================================
// CÁLCULO DE ISR SOBRE PTU
// ===========================================

function calcularISR(trabajador, ptuReal, empresa) {
  // Exención
  const diasExencion = 15;
  const valorDiario = empresa.usarUmaParaExencion ? empresa.umaDiaria : empresa.smgDiario;
  const ptuExenta = valorDiario * diasExencion;
  const ptuGravada = Math.max(0, ptuReal - ptuExenta);
  
  // ============ MÉTODO ART. 96 (LEY) ============
  const baseGravable_art96 = trabajador.ingresoMensualOrdinario + ptuGravada;
  const isrTotal_art96 = calcularISRTarifa96(baseGravable_art96);
  const isrOrdinario = calcularISRTarifa96(trabajador.ingresoMensualOrdinario);
  const isrPtu_art96 = isrTotal_art96 - isrOrdinario;
  const tasaEfectiva_art96 = ptuGravada > 0 ? (isrPtu_art96 / ptuGravada) * 100 : 0;
  const ptuNeta_art96 = ptuReal - isrPtu_art96;
  
  // ============ MÉTODO ART. 174 (REGLAMENTO) ============
  // Fracción I: PTU promedio mensual
  const ptuPromedioMensual = (ptuGravada / 365) * 30.4;
  
  // Fracción II: Base promediada
  const basePromediada = trabajador.ingresoMensualOrdinario + ptuPromedioMensual;
  
  // Fracción III: ISR sin subsidio
  const isrBasePromediada = calcularISRTarifa96SinSubsidio(basePromediada);
  const isrOrdinarioSinSubsidio = calcularISRTarifa96SinSubsidio(trabajador.ingresoMensualOrdinario);
  const diferenciaIsr = isrBasePromediada - isrOrdinarioSinSubsidio;
  
  // Fracción IV: Tasa efectiva
  const tasaEfectiva_art174 = ptuPromedioMensual > 0 ? (diferenciaIsr / ptuPromedioMensual) * 100 : 0;
  
  // Fracción V: ISR sobre PTU
  const isrPtu_art174 = ptuGravada * (tasaEfectiva_art174 / 100);
  const ptuNeta_art174 = ptuReal - isrPtu_art174;
  
  // Comparación
  const metodoRecomendado = isrPtu_art174 < isrPtu_art96 ? 'art174' : 'art96';
  const diferenciaMetodos = Math.abs(isrPtu_art96 - isrPtu_art174);
  
  return {
    ptuExenta,
    ptuGravada,
    // Art. 96
    baseGravable_art96,
    isrTotal_art96,
    isrPtu_art96,
    tasaEfectiva_art96,
    ptuNeta_art96,
    // Art. 174
    ptuPromedioMensual,
    basePromediada,
    isrBasePromediada,
    diferenciaIsr,
    tasaEfectiva_art174,
    isrPtu_art174,
    ptuNeta_art174,
    // Comparación
    metodoRecomendado,
    diferenciaMetodos
  };
}

// ===========================================
// TARIFA ART. 96 LISR
// ===========================================

const TARIFA_MENSUAL = [
  { limiteInferior: 0.01, limiteSuperior: 746.04, cuotaFija: 0, porcentaje: 0.0192 },
  { limiteInferior: 746.05, limiteSuperior: 6332.05, cuotaFija: 14.32, porcentaje: 0.0640 },
  { limiteInferior: 6332.06, limiteSuperior: 11128.01, cuotaFija: 371.83, porcentaje: 0.1088 },
  { limiteInferior: 11128.02, limiteSuperior: 12935.82, cuotaFija: 893.63, porcentaje: 0.1600 },
  { limiteInferior: 12935.83, limiteSuperior: 15487.71, cuotaFija: 1182.88, porcentaje: 0.1792 },
  { limiteInferior: 15487.72, limiteSuperior: 31236.49, cuotaFija: 1640.18, porcentaje: 0.2136 },
  { limiteInferior: 31236.50, limiteSuperior: 49233.00, cuotaFija: 5004.12, porcentaje: 0.2352 },
  { limiteInferior: 49233.01, limiteSuperior: 93993.90, cuotaFija: 9236.89, porcentaje: 0.3000 },
  { limiteInferior: 93993.91, limiteSuperior: 125325.20, cuotaFija: 22665.17, porcentaje: 0.3200 },
  { limiteInferior: 125325.21, limiteSuperior: 375975.61, cuotaFija: 32691.18, porcentaje: 0.3400 },
  { limiteInferior: 375975.62, limiteSuperior: Infinity, cuotaFija: 117912.32, porcentaje: 0.3500 }
];

function calcularISRTarifa96(baseGravable) {
  const rango = TARIFA_MENSUAL.find(r => 
    baseGravable >= r.limiteInferior && baseGravable <= r.limiteSuperior
  );
  
  if (!rango) return 0;
  
  const excedente = baseGravable - rango.limiteInferior;
  const isrMarginal = excedente * rango.porcentaje;
  const isrTotal = isrMarginal + rango.cuotaFija;
  
  return isrTotal;
}
```

---

## 9. OUTPUTS ESPERADOS

### 9.1 Por Trabajador
| Campo | Descripción |
|-------|-------------|
| PTU por días trabajados | Monto de la primera mitad |
| PTU por salarios | Monto de la segunda mitad |
| PTU Bruta | Suma antes de tope |
| Tope 3 meses | Cálculo del tope |
| Promedio 3 años | Cálculo del promedio |
| PTU Real | Monto después de aplicar tope |
| PTU Exenta | Parte exenta de ISR |
| PTU Gravada | Parte gravada de ISR |
| ISR Art. 96 | ISR con método de Ley |
| ISR Art. 174 | ISR con método de Reglamento |
| PTU Neta (Art. 96) | PTU menos ISR Ley |
| PTU Neta (Art. 174) | PTU menos ISR Reglamento |
| Método recomendado | El que arroja mayor PTU neta |

### 9.2 Totales de Empresa
- Total PTU bruta
- Total PTU real (después de topes)
- Total ISR retenido (según método elegido)
- Total PTU neta a pagar
- Remanente de PTU (diferencia entre bruta y real por topes)

### 9.3 Para Timbrado CFDI
- Clave de percepción: 003 - Participación de los Trabajadores en las Utilidades (PTU)
- Tipo de régimen: Sueldos y salarios
- Separar parte gravada y exenta

---

## 10. CONSIDERACIONES ADICIONALES

### 10.1 Impuesto Sobre Nómina (ISN)
- La exención de ISR **NO aplica** para ISN
- Se debe pagar ISN sobre el total de la PTU según la legislación de cada estado

### 10.2 CFDI de Nómina
- Tipo de nómina: Extraordinaria
- Concepto: 003 - PTU
- Separar monto gravado y exento

### 10.3 Ajuste Anual de ISR
- Si se usa Art. 96: probablemente sin diferencias
- Si se usa Art. 174: posible ISR a cargo del trabajador en su declaración anual

### 10.4 Opciones de Configuración Recomendadas
1. Selección de método de exención (UMA vs SMG)
2. Selección de método de retención por defecto (Art. 96 vs Art. 174)
3. Opción de calcular ambos métodos y comparar
4. Zona geográfica (para SMG de zona fronteriza)

---

## 11. REFERENCIAS LEGALES

- **Constitución Política de los Estados Unidos Mexicanos**: Art. 123, fracción IX
- **Ley Federal del Trabajo (LFT)**: Arts. 117-131 (PTU), Art. 127 fracción VIII (topes)
- **Ley del Impuesto Sobre la Renta (LISR)**: Art. 93 fracción XIV (exención), Art. 96 (tarifa)
- **Reglamento de la LISR (RLISR)**: Art. 174 (procedimiento opcional)
- **Resolución Miscelánea Fiscal 2025**: Anexo 8 (tarifas actualizadas)
- **Criterio PRODECON 2024**: Uso de SMG vs UMA para exención de PTU
- **Amparo SCJN 633/2023**: Criterio sobre tope para trabajadores con < 3 años antigüedad

---

## 12. VALORES UMA HISTÓRICOS

| Año | Valor Diario |
|-----|--------------|
| 2015 | $70.10 |
| 2016 | $73.04 |
| 2017 | $75.49 |
| 2018 | $80.60 |
| 2019 | $84.49 |
| 2020 | $86.88 |
| 2021 | $89.62 |
| 2022 | $96.22 |
| 2023 | $103.74 |
| 2024 | $108.57 |
| 2025 | $113.14 |

---

*Documento generado el 13 de enero de 2026*
*Basado en legislación vigente a 2025*
