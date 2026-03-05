# 📢 ACTUALIZACIÓN VERSIÓN 2.0

## Módulo XML CFDI - Refactor Completo

---

## ✨ Principales Mejoras

### 1. Carga más detallada
- Ahora **una fila por comprobante** (antes: consolidado por RFC)
- Desglose automático de IVA por **4 tasas distintas** (16%, 8%, 0%, exento)
- Incluye fecha, serie-folio, método de pago para referencia

### 2. Soporte para Egresos (Notas de Crédito)
- Procesa tanto **Ingresos (I)** como **Egresos (E)**
- Los egresos aparecen con **valores negativos** (resta automática al consolidar)
- Útil para devoluciones, descuentos y ajustes

### 3. Carga múltiple sin duplicados
- Puede cargar XMLs desde **varias carpetas**
- Sistema automático de **deduplicación por UUID**
- Cada invocación del botón **agrega al final** (no reemplaza)

### 4. Consolidación flexible
- Nueva hoja separada **"Datos_Concentrados"**
- Consolida por RFC solo cuando usted haga clic en el botón
- Permite revisar detalle antes de consolidar

### 5. Limpieza segura
- Botón para borrar datos **con confirmación**
- No hay sorpresas: pregunta antes de eliminar

---

## 🔄 Nuevo Flujo de Trabajo

```
1. Crear hoja "Datos_Proveedores"
   ↓
2. Cargar XML (botón) → una fila por comprobante
   ↓
3. Verificar datos (opcional) → puede editar manualmente
   ↓
4. Concentrar Datos (botón) → resumen por RFC
   ↓
5. Copiar a plantilla DIOT → usar columnas IVA para declaración
```

---

## 📋 Nuevas Columnas

| Columna | Dato |
|---------|------|
| A-C | RFC, Nombre, UUID |
| D-G | Fecha, Folio, Tipo (I/E), Método de Pago |
| H-K | **Bases por tasa:** 16%, 8%, 0%, Exento |
| L | Descuento |
| M-N | **IVA Trasladado:** 16%, 8% |
| O | IVA Retenido |
| P | Total |

---

## ⚙️ Instalación

1. **Crear la hoja "Datos_Proveedores"** en su libro
2. **Copiar encabezados** en fila 4 (ver documentación)
3. **Insertar 3 botones** en filas 1-2:
   - Botón 1 → `CargarXMLProveedores` (Cargar XML)
   - Botón 2 → `ConcentrarDatos` (Concentrar Datos)
   - Botón 3 → `LimpiarDatos` (Limpiar Datos)

---

## ⚠️ Cambios que Afectan Usuarios

| Cambio | Antes | Ahora |
|--------|-------|-------|
| **Granularidad** | 1 RFC = 1 fila | 1 Comprobante = 1 fila |
| **IVA detallado** | No | Sí, por tasa |
| **Egresos** | No | Sí (valores negativos) |
| **Carga múltiple** | Reemplaza | Agrega (append) |
| **Consolidación** | Automática | Manual (botón) |
| **Hoja resultado** | CFDI_Importados | Datos_Concentrados |

---

## ❓ Preguntas Rápidas

**P: ¿Pierdo la hoja anterior "CFDI_Importados"?**
R: Sí. La nueva versión usa "Datos_Concentrados". Conserve un backup si necesita datos históricos.

**P: ¿Puedo cargar XMLs varias veces?**
R: Sí. Solo asegúrese de no cargar la misma carpeta dos veces (el sistema detecta duplicados por UUID).

**P: ¿Los Egresos siempre en negativo?**
R: Sí, por diseño. Permite que sumas automáticas causen el efecto de deducción.

**P: ¿Necesito la plantilla DIOT para usar esto?**
R: No. Puede usar solo esta herramienta como análisis de comprobantes.

---

## 📖 Documentación Completa

Vea el archivo **DOCUMENTACION_REFACTOR_MODULO_XML.md** para:
- Explicación detallada de cada columna
- Cómo crear los botones paso a paso
- Casos de uso con ejemplos
- Solución de problemas

---

## 🐛 Reporte de Bugs

Si encuentra algún problema durante QA:
1. Anote el **mensaje de error exacto**
2. Indique la **ruta de la carpeta XML**
3. Adjunte **uno o dos XMLs de ejemplo**
4. Mencione su **versión de Excel**

---

**Versión:** 2.0
**Fecha de Implementación:** Febrero 2025
**Macros actualizadas:** 3 (CargarXMLProveedores, ConcentrarDatos, LimpiarDatos)
