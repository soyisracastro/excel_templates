# Guía: Instalación de Botones - Versión 2.0

## Paso a Paso

### PASO 1: Crear la hoja "Datos_Proveedores"

1. Abre tu libro Excel DIOT
2. **Haz clic derecho** en la pestaña de una hoja existente
3. Selecciona **"Insertar hoja"**
4. Nombre: `Datos_Proveedores`
5. Presiona OK

---

### PASO 2: Agregar encabezados (Fila 4)

En la fila 4, celda A4, comienza a escribir los siguientes encabezados:

```
A4:  RFC
B4:  Nombre del Emisor
C4:  UUID
D4:  Fecha
E4:  Serie-Folio
F4:  Tipo
G4:  Método de Pago
H4:  Valor Actos Gravados 16%
I4:  Valor Actos Gravados 8%
J4:  Valor Actos Tasa 0%
K4:  Valor Actos Exentos
L4:  Descuento
M4:  IVA Trasladado 16%
N4:  IVA Trasladado 8%
O4:  IVA Retenido
P4:  Total
```

**Opción rápida:** Copie y pegue esta línea en la fila 4:
```
RFC | Nombre del Emisor | UUID | Fecha | Serie-Folio | Tipo | Método de Pago | Valor Actos Gravados 16% | Valor Actos Gravados 8% | Valor Actos Tasa 0% | Valor Actos Exentos | Descuento | IVA Trasladado 16% | IVA Trasladado 8% | IVA Retenido | Total
```

Luego, en Excel: **Datos** > **Texto en columnas** > Separador **Tubería (|)** > Aceptar

---

### PASO 3: Insertar el PRIMER botón

1. Ve a la pestaña **"Insertar"**
2. En el grupo **"Controles de formulario"** (lado derecho), haz clic en el ícono de **"Botón"**

   > Si no ves este ícono, busca "Formulario" en el menú Insertar

3. **Dibuja un rectángulo** en la celda A1 (o donde quieras el botón)
   - Presiona el mouse, arrastra hasta crear un rectángulo
   - Suelta el mouse

4. Se abrirá automáticamente un cuadro: **"Asignar macro"**
   - En la lista, selecciona: `ModuloXMLCFDI.CargarXMLProveedores`
   - Presiona OK

5. **Edita el texto del botón:**
   - Clic derecho en el botón
   - Selecciona **"Editar texto"**
   - Borra todo y escribe: `Cargar XML`
   - Haz clic fuera del botón

---

### PASO 4: Insertar el SEGUNDO botón

1. Repite los pasos 1-2 del PASO 3
2. **Dibuja el botón** en la celda C1 (a la derecha del primero)
3. En el cuadro **"Asignar macro"**:
   - Selecciona: `ModuloXMLCFDI.ConcentrarDatos`
   - OK
4. **Edita el texto:** `Concentrar Datos`

---

### PASO 5: Insertar el TERCER botón

1. Repite los pasos 1-2 del PASO 3
2. **Dibuja el botón** en la celda E1 (a la derecha del segundo)
3. En el cuadro **"Asignar macro"**:
   - Selecciona: `ModuloXMLCFDI.LimpiarDatos`
   - OK
4. **Edita el texto:** `Limpiar Datos`

---

### PASO 6: Ajustar tamaño y formato (opcional)

Para que los botones se vean mejor:

1. **Haz clic en el primer botón** (Cargar XML)
2. Clic derecho > **"Propiedades"** (o **"Formato de control"**)
3. Ajusta:
   - **Fuente:** Tamaño 11
   - **Color de relleno:** Azul claro
   - **Color de texto:** Blanco
4. Presiona OK
5. Repite para los otros dos botones

---

## ✅ Verificación

Para verificar que todo funciona:

1. **Haz clic en botón "Cargar XML"**
   - Debe abrirse un diálogo para seleccionar carpeta
   - Si no abre: verifica que hayas asignado la macro correctamente

2. **Cancela ese diálogo** (no necesitas cargar XMLs ahora)

3. Los botones están listos para usar

---

## 🔧 Solución de Problemas

### Problema: El botón no hace nada

**Solución:**
1. Clic derecho en el botón
2. **"Asignar macro"**
3. Verifica que esté asignada la macro correcta:
   - Botón 1: `ModuloXMLCFDI.CargarXMLProveedores`
   - Botón 2: `ModuloXMLCFDI.ConcentrarDatos`
   - Botón 3: `ModuloXMLCFDI.LimpiarDatos`

### Problema: No aparece la opción "Asignar macro"

**Solución:**
1. Verifica que el botón sea del tipo **"Formulario"** (no ActiveX)
2. Si es ActiveX, borra y crea uno nuevo desde **Insertar** > **Controles de formulario**

### Problema: La hoja "Datos_Proveedores" no existe

**Solución:**
1. Crea la hoja manualmente (ver PASO 1)
2. Verifica que se llame exactamente **"Datos_Proveedores"** (sin mayúsculas adicionales)

---

## 🎯 Configuración Opcional Recomendada

### Proteger encabezados (filas 1-4) de cambios accidentales

1. Selecciona filas 5 en adelante: Clic en **5** en el encabezado de filas
2. Ve a **Formato** > **Celdas** > **Protección**
3. Marca **"Bloqueado"** (generalmente ya está)
4. Ahora:
   - Ve a **Revisar** > **Proteger hoja**
   - Opciones: Deja todo marcado
   - Presiona OK (sin contraseña, o con contraseña si lo prefieres)

Esto previene que se cierren columnas accidentalmente.

### Ancho de columnas

Para que el encabezado se vea bien:
1. Selecciona la fila 4 completa (clic en **4**)
2. Haz doble clic en la línea divisoria entre dos columnas en el encabezado
3. Excel ajusta automáticamente el ancho

---

## 📞 Si algo falla

- Verifica que el archivo Excel tenga habilitadas las **Macros**
- Asegúrate de que **no esté en Modo Seguro**
- Si aparece error, apunta el número exacto del error

---

**Versión de esta guía:** 2.0
**Última actualización:** Febrero 2025
