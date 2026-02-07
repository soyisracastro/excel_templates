# Documentación: Refactor del Módulo XML CFDI

## Fecha de Implementación
**Versión 2.0** - Refactor Completo del Módulo `ModuloXMLCFDI.bas`

---

## Resumen Ejecutivo

El módulo `ModuloXMLCFDI` ha sido completamente rediseñado para proporcionar una experiencia de usuario mejorada en la carga y consolidación de CFDI 4.0.

**Cambios principales:**
- ✅ Carga de XMLs **una fila por comprobante** (no consolidados)
- ✅ Soporte para **Ingresos (I) y Egresos (E)** simultáneamente
- ✅ Desglose detallado de IVA por **tasa (16%, 8%, 0%, exento)**
- ✅ **Modo append** - carga múltiple desde distintas carpetas sin duplicar
- ✅ **Deduplicación por UUID** - evita cargar el mismo XML dos veces
- ✅ Consolidación manual en hoja separada con **un clic**
- ✅ Limpieza de datos con **confirmación de usuario**

---

## Flujo de Uso

### 1️⃣ Preparación Inicial

Cree una hoja llamada **"Datos_Proveedores"** en su libro Excel con la siguiente estructura:

- **Filas 1-2**: Reservadas para botones
- **Fila 3**: Espacio en blanco (separador visual)
- **Fila 4**: Encabezados de columnas
- **Fila 5+**: Datos (se cargan aquí automáticamente)

### 2️⃣ Botón 1: "Cargar XML"

**Macro:** `CargarXMLProveedores`

**¿Qué hace?**
- Abre un diálogo para seleccionar una carpeta que contiene archivos XML
- Lee TODOS los archivos .xml de esa carpeta
- Filtra solo comprobantes tipo **Ingreso (I)** y **Egreso (E)**
- Procesa cada XML y escribe **una fila por comprobante** en la hoja
- Evita duplicados verificando el UUID
- Detiene automáticamente si encuentra errores en el XML

**Detalles técnicos:**
- Los **Egresos** se escriben con **valores negativos** (útil para devoluciones y descuentos)
- Extrae el desglose de IVA por tasa desde los nodos globales del CFDI
- Compatible con XMLs de múltiples PACs (Facturación Moderna, Globaltax, etc.)
- Puede usarse varias veces: cada invocación agrega al final de los datos existentes

**Resultado por fila:**

| RFC | Nombre | UUID | Fecha | Folio | Tipo | Método | Grav16% | Grav8% | Tasa0% | Exento | Desc | IVA16% | IVA8% | IVARet | Total |
|-----|--------|------|-------|-------|------|--------|---------|--------|--------|--------|------|--------|-------|--------|-------|
| AAA010101SA9 | EMPRESA ABC | 12AB-CD34-... | 2025-01-15 | A123 | I | PUE | 1000.00 | 500.00 | 0.00 | 100.00 | 0.00 | 160.00 | 40.00 | 0.00 | 1800.00 |
| BBB020202SB9 | PROVEEDORA XYZ | 56EF-GH78-... | 2025-01-16 | C456 | E | PPD | -500.00 | 0.00 | 0.00 | 0.00 | 0.00 | -80.00 | 0.00 | 0.00 | -580.00 |

### 3️⃣ Botón 2: "Concentrar Datos"

**Macro:** `ConcentrarDatos`

**¿Qué hace?**
- Lee TODOS los datos de la hoja "Datos_Proveedores"
- Agrupa por **RFC del emisor**
- Suma todas las columnas numéricas por RFC
- Cuenta el número de operaciones por proveedor
- Crea una nueva hoja llamada **"Datos_Concentrados"**
- Aplica formato automático (moneda, bordes, encabezados)

**Cuándo usarlo:**
- Después de cargar todos los XMLs que va a usar
- Antes de copiar la información a la plantilla DIOT oficial

**Resultado por fila concentrada:**

| RFC | Nombre | Num. Ops | Grav16% | Grav8% | Tasa0% | Exento | Desc | IVA16% | IVA8% | IVARet | Total |
|-----|--------|----------|---------|--------|--------|--------|------|--------|-------|--------|-------|
| AAA010101SA9 | EMPRESA ABC | 5 | 5000.00 | 2500.00 | 0.00 | 500.00 | 0.00 | 800.00 | 200.00 | 0.00 | 9000.00 |
| BBB020202SB9 | PROVEEDORA XYZ | 2 | -1000.00 | 0.00 | 0.00 | 0.00 | 0.00 | -160.00 | 0.00 | -50.00 | -1210.00 |

**Nota importante sobre Egresos:**
> Si cargó comprobantes de tipo **Egreso (E)**, estos aparecen con valores **negativos** en Datos_Proveedores.
> Al concentrar por RFC, las sumas se hacen automáticamente: Ingresos - Egresos = Neto.
> Esto es por diseño: el usuario puede determinar bajo principio de autodeterminación qué deducciones aplica.

### 4️⃣ Botón 3: "Limpiar Datos"

**Macro:** `LimpiarDatos`

**¿Qué hace?**
- Muestra un cuadro de confirmación (no se puede deshacer)
- Borra todos los datos desde la fila 5 en adelante
- Elimina la hoja "Datos_Concentrados" si existe
- Deja la hoja lista para una nueva carga

---

## Encabezados de Columnas

Copie y pegue estos encabezados en **fila 4** de la hoja "Datos_Proveedores":

| Columna | Encabezado |
|---------|-----------|
| A | RFC |
| B | Nombre del Emisor |
| C | UUID |
| D | Fecha |
| E | Serie-Folio |
| F | Tipo |
| G | Método de Pago |
| H | Valor Actos Gravados 16% |
| I | Valor Actos Gravados 8% |
| J | Valor Actos Tasa 0% |
| K | Valor Actos Exentos |
| L | Descuento |
| M | IVA Trasladado 16% |
| N | IVA Trasladado 8% |
| O | IVA Retenido |
| P | Total |

---

## Explicación de Columnas

### Columnas de Referencia
- **RFC**: Registro Federal de Contribuyentes del emisor (proveedor)
- **Nombre del Emisor**: Razón social o nombre comercial
- **UUID**: Identificador único del comprobante (TimbreFiscalDigital)
- **Fecha**: Fecha de emisión en formato YYYY-MM-DD
- **Serie-Folio**: Serie y número de folio concatenados
- **Tipo**: "I" para Ingresos, "E" para Egresos (notas de crédito)
- **Método de Pago**: "PUE" (Pago en una Exhibición) o "PPD" (Pago a Plazo)

### Columnas de Bases (Valores Actos)
Los valores en estas columnas corresponden a la **base neta de impuesto** para cada tasa del IVA:
- **Valor Actos Gravados 16%**: Monto sujeto a IVA al 16% (tasa estándar)
- **Valor Actos Gravados 8%**: Monto sujeto a IVA al 8% (región fronteriza)
- **Valor Actos Tasa 0%**: Monto con tasa de IVA del 0% (exportaciones)
- **Valor Actos Exentos**: Monto no sujeto a IVA (operaciones exentas)

> **Nota técnica**: Los valores en estas columnas vienen del atributo `Base` de los nodos globales `cfdi:Traslado` en el XML. Este Base ya es un monto neto (refleja descuentos a nivel concepto).

### Columnas de IVA Trasladado
- **IVA Trasladado 16%**: Impuesto causado al 16% (16% de la base de 16%)
- **IVA Trasladado 8%**: Impuesto causado al 8% (8% de la base de 8%)

> Estos valores corresponden al atributo `Importe` de los nodos `cfdi:Traslado` globales.

### Columnas de Retenciones e Impuestos
- **Descuento**: Total de descuentos aplicados al comprobante (atributo `Descuento` del Comprobante)
- **IVA Retenido**: Total de IVA retenido en el comprobante (de nodos `cfdi:Retencion`)
- **Total**: Monto total del comprobante

---

## Casos de Uso

### Caso 1: Carga Simple (Una carpeta)
1. Descarga XMLs desde el portal del SAT a una carpeta
2. Abre el libro Excel DIOT
3. Crea la hoja "Datos_Proveedores" con encabezados
4. Haz clic en botón "Cargar XML" → selecciona la carpeta
5. Los datos aparecen en la hoja

### Caso 2: Carga Múltiple (Varias carpetas)
1. Primera carga: XMLs de carpeta A → botón "Cargar XML"
2. Segunda carga: XMLs de carpeta B → botón "Cargar XML" (se agrega al final, sin duplicados)
3. Tercera carga: XMLs de carpeta C → botón "Cargar XML"
4. Resultado: Una sola hoja con todos los comprobantes

### Caso 3: Consolidación para Declaración
1. Cargar todos los XMLs necesarios (múltiples carpetas si es necesario)
2. Opcionalmente, agregar filas manualmente para proveedores sin XML
3. Botón "Concentrar Datos" → genera resumen por RFC
4. Copiar columnas de "Datos_Concentrados" a la plantilla DIOT oficial

### Caso 4: Devoluciones y Descuentos
1. Cargar XMLs de Ingresos (tipo I) normalmente
2. Cargar XMLs de Notas de Crédito (tipo E) - aparecen con valores negativos
3. Botón "Concentrar Datos" → sumas netas por RFC (Ingresos - Devoluciones)
4. El usuario decide qué porción de IVA retenido acredita (autodeterminación)

---

## Compatibilidad y Requisitos

### Requisitos del Sistema
- ✅ Microsoft Excel 2016 o superior
- ✅ Windows 7 o superior (VBA está habilitado)
- ✅ MSXML2.DOMDocument (Microsoft XML, 6.0) - incluido en Windows

### Tipos de Comprobantes Soportados
- ✅ **Ingresos (I)**: Facturas normales, tickets, recibos
- ✅ **Egresos (E)**: Notas de crédito, devoluciones
- ❌ **Pagos (P)**: Ya no se procesan (se pueden agregar manualmente si es necesario)
- ❌ **Nómina, Traslados, otros**: Se ignoran automáticamente

### Formatos XML
- ✅ CFDI 4.0 estándar
- ✅ Con Complemento TimbreFiscalDigital (TFD)
- ✅ Múltiples PACs: Facturación Moderna, Globaltax, Soluciones Fáciles, etc.

---

## Limitaciones Conocidas

1. **Egresos sin Método de Pago**: Si una nota de crédito no tiene el atributo `MetodoPago`, la columna G estará vacía
2. **UUIDs duplicados en el SAT**: Teóricamente imposible, pero si ocurre, solo se carga la primera ocurrencia
3. **Campos sin información**: Si el XML carece de algunos campos (ej: Nombre), la celda quedará vacía (Excel no lanza error)
4. **Descuentos a nivel concepto**: El XML puede tener descuentos por concepto individual que no aparecen en la columna "Descuento" (que es a nivel comprobante)

---

## Cómo Crear los Botones

### En Excel 2016+

1. **Inserta un botón:**
   - Ir a: `Insertar` > `Controles de formulario` > `Botón`
   - Dibuja un botón en la fila 1 o 2 de la hoja "Datos_Proveedores"

2. **Asigna la macro:**
   - Clic derecho en el botón → `Asignar macro`
   - Selecciona: `ModuloXMLCFDI.CargarXMLProveedores`
   - OK

3. **Repite para los otros dos botones:**
   - Segundo botón: `ModuloXMLCFDI.ConcentrarDatos`
   - Tercer botón: `ModuloXMLCFDI.LimpiarDatos`

4. **Personaliza el botón:**
   - Clic derecho > `Editar texto` → cambia el nombre
   - Clic derecho > `Formato de control` → ajusta color, tamaño, fuente

---

## Mensajes del Sistema

### Después de "Cargar XML"
```
Proceso completado.

XMLs cargados (Ingreso/Egreso): 23
Duplicados omitidos (UUID): 2
Ignorados (Pagos/Nómina/Traslado): 1
```

### Después de "Concentrar Datos"
```
Datos concentrados generados.
15 proveedores encontrados.
```

### Después de "Limpiar Datos"
```
Datos limpiados correctamente.
La hoja está lista para una nueva carga.
```

### Mensaje de Error - Carpeta no encontrada
```
No se pudo encontrar la ruta.

Ruta detectada: [ruta ingresada]

Estás intentando usar una ruta web de OneDrive/SharePoint.
Abre la carpeta en el explorador, copia la ruta local y asegúrate de que los
archivos estén 'Disponibles siempre en este dispositivo'.
```

### Mensaje de Error - Hoja no existe
```
La hoja 'Datos_Proveedores' no existe en este libro.
Cree la hoja con los encabezados en la fila 4 antes de ejecutar este macro.
```

---

## FAQ (Preguntas Frecuentes)

### P: ¿Puedo cargar la misma carpeta dos veces?
**R:** No. El sistema detecta UUIDs duplicados y omite los comprobantes ya cargados. Se le notificará cuántos fueron ignorados.

### P: ¿Qué pasa si algunos XMLs tienen errores?
**R:** Se omiten silenciosamente. El resumen al final le indicará cuántos se procesaron exitosamente.

### P: ¿Puedo editar manualmente la hoja Datos_Proveedores?
**R:** Sí. Puede agregar filas manualmente para proveedores sin XML (ej: gastos por otros medios). Solo respete la estructura de columnas.

### P: ¿Si concentro datos, se pierde la información detallada?
**R:** No. La hoja "Datos_Concentrados" es nueva; la original "Datos_Proveedores" permanece intacta. Puede consultarla en cualquier momento.

### P: ¿Los Egresos siempre aparecen en negativo?
**R:** Sí, por diseño. Esto permite que al concentrar por RFC, los valores negativos se resten automáticamente. Si desea cambiar esto, edite manualmente la fila.

### P: ¿Puedo usar esto sin la plantilla DIOT posterior?
**R:** Sí. Puede usar "Datos_Concentrados" directamente para cualquier análisis o reporte que necesite. Solo los IVA trasladados y retenidos son necesarios para la DIOT oficial.

### P: ¿Qué hago si OneDrive sincroniza "solo en la nube"?
**R:**
1. Abre el Explorador de Windows
2. Navega a la carpeta con XMLs
3. Haz clic derecho > "Siempre mantener en este dispositivo"
4. Espera a que se descarguen (puede tomar minutos)
5. Intenta nuevamente

---

## Cambios Respecto a Versión Anterior

| Aspecto | Versión 1.x | Versión 2.0 |
|--------|-----------|-----------|
| **Consolidación** | Automática (un RFC = una fila) | Manual (botón separado) |
| **Granularidad de datos** | Un RFC por fila | Un comprobante por fila |
| **IVA por tasa** | No | Sí (16%, 8%, 0%, exento) |
| **Egresos** | No soportados | Soportados (con signo negativo) |
| **Carga múltiple** | Reemplaza datos | Append (agrega al final) |
| **Deduplicación** | No | Sí, por UUID |
| **Limpieza** | Manual | Botón automático |
| **Hojas generadas** | CFDI_Importados | Datos_Concentrados (solo si se usa botón) |

---

## Soporte Técnico

Si encuentra algún problema o tiene preguntas:

1. **Revise que la hoja "Datos_Proveedores" exista** con los encabezados correctos en fila 4
2. **Verifique la ruta**: Si usa OneDrive, asegúrese de que los archivos estén "Disponibles siempre en este dispositivo"
3. **Inspect los XML**: Abra uno en un editor de texto y verifique que es un archivo XML válido (no corrupto)
4. **Reporte el error**: Si el problema persiste, recopile:
   - El mensaje de error exacto
   - La ruta de la carpeta XML
   - Versión de Excel
   - Nombre de uno o dos archivos XML problemáticos

---

## Versión del Módulo
- **Versión:** 2.0
- **Fecha de Implementación:** 2025-02
- **Macros:** CargarXMLProveedores, ConcentrarDatos, LimpiarDatos
- **Líneas de código:** ~575 líneas
- **Funciones auxiliares:** 5 (SeleccionarCarpeta, GetAttr, GetNodeAttr, CargarUUIDsExistentes, ObtenerSiguienteFila)

---

**Última actualización:** Febrero 2025
