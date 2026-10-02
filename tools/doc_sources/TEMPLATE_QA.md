# Control de calidad de las plantillas

## Propósito

Esta lista verifica que `entidades_portrait.docx` y `entidades_landscape.docx` sean ejemplos válidos del contrato 1.0 y que su apariencia siga siendo estable al cambiar el generador. Las plantillas son datos de prueba del compilador, no documentos que se cargan en Oracle APEX.

## Generación reproducible

```flujo
# Figura 1. Ciclo de control de calidad de las plantillas
Informática | Ejecutar el generador en build\qa\templates | tools\generate_sample_templates.py
? QA | ¿Idénticas byte a byte a templates\? | No: el generador cambió; revisar si fue intencional
QA | Inspección estructural | un párrafo de título, una tabla de dos filas, un pie
QA | Compilar cada plantilla con un proyecto válido | template.json se genera automáticamente
QA | Exportar a PDF e inspeccionar | build\qa, al 100 %
? QA | ¿Todos los puntos conformes? | No: corregir el generador y repetir el ciclo completo
QA | Publicar en templates\ y copiar a examples\entidades.docx | regenerar examples\build_expected
```

1. Desde la raíz del proyecto, con el entorno `.venv` preparado:

```
# Genera las dos plantillas en una carpeta de QA e imprime sus SHA-256
.\.venv\Scripts\python.exe tools\generate_sample_templates.py --output-dir build\qa\templates
```

2. Guardar los hashes SHA-256 impresos por el script.
3. Comparar byte a byte con las plantillas versionadas:

```
fc.exe /b build\qa\templates\entidades_portrait.docx templates\entidades_portrait.docx
fc.exe /b build\qa\templates\entidades_landscape.docx templates\entidades_landscape.docx
```

4. Confirmar que los únicos archivos generados son: `entidades_portrait.docx` y `entidades_landscape.docx`.
5. `examples\entidades.docx` debe ser idéntico a `templates\entidades_landscape.docx`.
6. Solo si se cambió el generador a propósito, regenerar en `templates` (directorio predeterminado del script) y repetir esta lista completa.

## Inspección estructural

Comprobar con el lector del propio compilador y, cuando sea necesario, inspeccionando el OOXML:

- el archivo es `.docx`, no `.docm`;
- existe una sola sección;
- el encabezado real de Word está vacío;
- existe un único párrafo con contenido antes de la tabla;
- ese párrafo contiene `{{REPORT_TITLE}}` exactamente una vez y, tras un salto de línea manual, `Unidad: {{FIELD:UNIDAD}}`;
- existe una sola tabla principal;
- la tabla tiene exactamente dos filas y el mismo número de celdas por fila;
- la fila 1 contiene solo etiquetas literales;
- cada celda de la fila 2 contiene únicamente un marcador `{{COLUMN:ALIAS}}`;
- no hay aliases duplicados;
- el footer real contiene un único párrafo con `{{APP_USER}}` y `{{GENERATED_AT}}`;
- en conjunto se usan los cinco tipos de marcador: `{{REPORT_TITLE}}`, `{{FIELD:...}}`, `{{COLUMN:...}}`, `{{APP_USER}}` y `{{GENERATED_AT}}`;
- no hay párrafos con contenido después de la tabla;
- no hay imágenes, relaciones externas, macros, formas, objetos OLE, campos Word, comentarios, cambios controlados, tablas anidadas ni celdas combinadas.

## Contenido esperado

### Plantilla vertical

| Posición | Etiqueta | Alias |
|---|---|---|
| 1 | DNI | VDNI |
| 2 | Nombre completo | VNOM |
| 3 | Departamento | DEPARTAMENTO |

La página debe ser A4 vertical. Los anchos visuales deben favorecer el nombre completo y conservar el DNI como columna compacta.

### Plantilla horizontal

| Posición | Etiqueta | Alias |
|---|---|---|
| 1 | DNI | VDNI |
| 2 | Nombre completo | VNOM |
| 3 | Dirección actual | VDIREC_ACTUAL |
| 4 | Departamento | DEPARTAMENTO |
| 5 | Distrito | DISTRITO |
| 6 | Teléfono | VNRO_TLF1 |

La página debe ser A4 horizontal. `VDIREC_ACTUAL` debe seguir siendo una de las columnas más anchas. `DEPARTAMENTO` y `VNRO_TLF1` reciben espacio adicional para que sus marcadores técnicos no se dividan de forma innecesaria. El modo `AUTO`, si se desea probar, se declara en el manifiesto del proyecto y no se codifica dentro del DOCX.

## Estilos esperados

- título y línea de unidad centrados, Arial 15 pt, negrita, color `#2F343A` (un solo estilo para todo el párrafo);
- fila de encabezado centrada, Arial 9 pt, negrita, texto blanco y fondo `#4A4F55`;
- fila prototipo Arial 8.5 pt en vertical y 8 pt en horizontal, peso normal, texto `#25282B` y fondo blanco;
- cuerpo alineado de acuerdo con el tipo de dato: texto a la izquierda y identificadores o teléfono centrados;
- bordes internos y externos uniformes de 0.5 pt, color `#BFC3C7`;
- footer centrado, Arial 8 pt, peso normal, color `#666666`;
- sin cursiva, subrayado, tachado ni formato mixto dentro de una zona.

## Compilación esperada

Ejecutar cada plantilla con un proyecto válido y verificar. El proyecto debe declarar el campo de la plantilla, por ejemplo:

```
"fields": [
  {"name": "UNIDAD", "source": "CONSTANT", "value": "Oficina de Informática"}  // {{FIELD:UNIDAD}}
]
```

«Nuevo proyecto desde DOCX» crea automáticamente un proyecto válido con `UNIDAD` como campo `ITEM`; cualquiera de las dos formas sirve. Verificar:

- el diagnóstico no contiene errores;
- la lectura valida que el DOCX fuente use papel A4, sin emitir una opción de tamaño de papel en el proyecto ni en `template.json`;
- `report.orientation` de `template.json` es `PORTRAIT` o `LANDSCAPE`, según el archivo, cuando el proyecto no declara `orientation`;
- los nombres, etiquetas, alineaciones y orden de columnas coinciden con las tablas de «Contenido esperado»;
- las columnas sin entrada en `column_widths` reciben `WEIGHT` 1.0; los anchos dibujados en Word son solo una ayuda visual y no alteran silenciosamente el contrato compilado;
- el estilo del título queda normalizado a `CENTER`;
- `report.header_template` es `{{REPORT_TITLE}}` + salto de línea + `Unidad: {{FIELD:UNIDAD}}`;
- el footer conserva ambos marcadores integrados;
- repetir la compilación produce JSON y SQL normalizados idénticos.

## Renderizado e inspección visual

Abrir cada DOCX de `templates` en Microsoft Word y exportarlo a PDF con *Archivo > Exportar > Crear PDF/XPS* en `build\qa`. Como alternativa sin Word:

```
& "C:\Program Files\LibreOffice\program\soffice.exe" --headless --convert-to pdf --outdir build\qa templates\entidades_portrait.docx templates\entidades_landscape.docx
```

Revisar al 100 % todas las páginas del PDF. Confirmar:

- orientación y tamaño de página correctos;
- título visible, centrado y sin línea decorativa;
- tabla centrada y contenida dentro de los márgenes;
- ninguna etiqueta o marcador cortado, superpuesto o pegado al borde;
- en la plantilla horizontal, `{{COLUMN:DEPARTAMENTO}}` y `{{COLUMN:VNRO_TLF1}}` deben permanecer completos y con un ajuste de línea visualmente razonable; no se permite partir los aliases manualmente;
- anchos visualmente distintos y coherentes con el tipo de columna;
- bordes internos y externos visibles y uniformes;
- footer centrado, legible y dentro del área imprimible;
- sin sustitución visible de fuente ni glifos faltantes;
- ninguna página en blanco adicional.

Si falla un punto visual o estructural, corregir el generador, volver a generar ambos DOCX y repetir toda esta sección. Los PDF son evidencia interna de QA, quedan en `build\qa` y no forman parte de los artefactos finales.
