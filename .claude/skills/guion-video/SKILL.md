---
name: guion-video
description: Genera un guion de video tutorial para una plantilla Excel del proyecto
argument-hint: [carpeta-del-proyecto]
---

# Guion de Video Tutorial para Plantilla Excel

Genera un guion de video tutorial en español (México) para la plantilla Excel ubicada en `$ARGUMENTS`.

## Proceso

1. **Explorar el proyecto**: Lee todos los archivos relevantes de la carpeta indicada (.bas, .cls, .py, .md) para entender:
   - Qué hace la plantilla
   - Qué hojas tiene
   - Qué macros o fórmulas incluye
   - Cuál es el flujo de trabajo del usuario

2. **Generar el guion** con la estructura descrita abajo y guardarlo como `guion_video_tutorial.md` dentro de la carpeta del proyecto.

## Estructura del guion

El guion debe seguir esta estructura de secciones. Adapta el número de escenas y duración según la complejidad de la plantilla:

```markdown
# Guion de Video Tutorial: [Nombre de la plantilla]

**Duración estimada:** X-Y minutos
**Formato:** Screencast con voz en off
**Software:** Excel (escritorio)

---

## INTRO (0:00 - 0:30)

**[Pantalla: portada con título]**

> "En este video te voy a mostrar cómo usar [nombre de la plantilla] paso a paso. Al terminar, vas a poder [beneficio principal para el usuario]."

**[Pantalla: abrir el archivo .xlsm en Excel]**

> "Lo primero: al abrir el archivo, Excel te va a pedir que habilites las macros. Haz clic en 'Habilitar contenido'."

---

## ESCENA 1: Vista general del libro

> Recorrer cada hoja/pestaña con una frase descriptiva de su propósito.

---

## ESCENA 2: Configuración inicial (si aplica)

> Mostrar qué debe configurar el usuario antes de empezar.

---

## ESCENA 3+: Flujo de trabajo paso a paso

> Una escena por cada paso del flujo principal.
> Incluir tablas con:
> | Celda | Valor | Narración |
> Para que el locutor sepa exactamente qué escribir y decir.
> Incluir pausas para mostrar resultados automáticos.

---

## ESCENA N: Tips finales

> 3-5 consejos prácticos numerados.

---

## CIERRE

> Despedida breve, referencia a documentación o blog.

---

## Notas de producción

### Resolución y formato
- Grabar en 1920×1080 (Full HD)
- Excel al 100% de zoom, tema claro

### Recursos en pantalla
- Resaltar celdas con rectángulo semitransparente al mencionarlas
- Zoom in para fórmulas o resultados específicos
- Mostrar barra de fórmulas cuando expliques columnas calculadas

### Thumbnail sugerido
- Texto principal y captura de pantalla relevante

### Plataformas
- YouTube con SEO relevante
- Embed en página de venta (si aplica)
- Fragmentos cortos (30-60 seg) para redes sociales
```

## Reglas de estilo

- **Tono**: Profesional pero accesible. Como un colega que te explica algo en la oficina.
- **Narración**: Siempre en segunda persona ("tú"). Sin tutear de "usted".
- **Tecnicismos**: Explicar la primera vez que aparecen. Después usarlos con naturalidad.
- **Ejemplos**: Usar datos realistas pero genéricos (no datos reales de una empresa específica).
- **Acotaciones**: Usar `**[Pantalla: ...]**` para indicar lo que se ve, y `> "..."` para lo que se dice.
- **Tiempos**: Incluir timestamps estimados por escena entre paréntesis.
- **Tablas de captura**: Cuando el usuario deba escribir algo, usar tablas con columnas Celda, Valor, Narración.
- **Sin emojis** en el guion (el texto se lee en voz alta).
- **Idioma**: Español de México. Sin acentos en el código VBA referenciado. Usar "computadora" no "ordenador", "clic" no "click".
