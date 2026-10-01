# Control de calidad de las plantillas de referencia

## Propósito

Esta lista verifica que `entidades_portrait.docx` y
`entidades_landscape.docx` sean ejemplos válidos del contrato 1.0 y que su
apariencia siga siendo estable al cambiar el generador. Las plantillas son
datos de prueba del compilador, no documentos que se cargan en Oracle APEX.

## Generación reproducible

1. Ejecutar el marcador de autoría requerido por el flujo de documentos antes
   de generar los DOCX por primera vez.
2. Usar el Python del runtime de artefactos:

   ```bash
   "$CODEX_PRIMARY_RUNTIME_PYTHON" tools/generate_sample_templates.py
   ```

3. Guardar los hashes SHA-256 impresos por el script.
4. Ejecutar el generador por segunda vez en otro directorio y comprobar que
   ambos hashes coinciden byte por byte para cada nombre de archivo.
5. Confirmar que los únicos archivos generados son:
   `entidades_portrait.docx` y `entidades_landscape.docx`.

## Inspección estructural

Comprobar con el lector del propio compilador y, cuando sea necesario,
inspeccionando el OOXML:

- el archivo es `.docx`, no `.docm`;
- existe una sola sección;
- el encabezado real de Word está vacío;
- existe un único párrafo con contenido antes de la tabla;
- ese párrafo contiene `{{REPORT_TITLE}}` exactamente una vez;
- existe una sola tabla principal;
- la tabla tiene exactamente dos filas y el mismo número de celdas por fila;
- la fila 1 contiene solo etiquetas literales;
- cada celda de la fila 2 contiene únicamente un marcador
  `{{COLUMN:ALIAS}}`;
- no hay aliases duplicados;
- el footer real contiene un único párrafo con
  `{{APP_USER}}` y `{{GENERATED_AT}}`;
- no hay párrafos con contenido después de la tabla;
- no hay imágenes, relaciones externas, macros, formas, objetos OLE, campos
  Word, comentarios, cambios controlados, tablas anidadas ni celdas combinadas.

## Contenido esperado

### Plantilla vertical

| Posición | Etiqueta | Alias |
|---:|---|---|
| 1 | DNI | `VDNI` |
| 2 | Nombre completo | `VNOM` |
| 3 | Departamento | `DEPARTAMENTO` |

La página debe ser A4 vertical. Los anchos visuales deben favorecer el nombre
completo y conservar el DNI como columna compacta.

### Plantilla horizontal

| Posición | Etiqueta | Alias |
|---:|---|---|
| 1 | DNI | `VDNI` |
| 2 | Nombre completo | `VNOM` |
| 3 | Dirección actual | `VDIREC_ACTUAL` |
| 4 | Departamento | `DEPARTAMENTO` |
| 5 | Distrito | `DISTRITO` |
| 6 | Teléfono | `VNRO_TLF1` |

La página debe ser A4 horizontal. `VDIREC_ACTUAL` debe seguir siendo una de las
columnas más anchas. `DEPARTAMENTO` y `VNRO_TLF1` reciben espacio adicional
para que sus marcadores técnicos no se dividan de forma innecesaria. El modo
`AUTO`, si se desea probar, se declara en el manifiesto del proyecto y no se
codifica dentro del DOCX.

## Estilos esperados

- título centrado, Arial 15 pt, negrita, color `#2F343A`;
- fila de encabezado centrada, Arial 9 pt, negrita, texto blanco y fondo
  `#4A4F55`;
- fila prototipo Arial 8.5 pt en vertical y 8 pt en horizontal, peso normal,
  texto `#25282B` y fondo blanco;
- cuerpo alineado de acuerdo con el tipo de dato: texto a la izquierda y
  identificadores o teléfono centrados;
- bordes internos y externos uniformes de 0.5 pt, color `#BFC3C7`;
- footer centrado, Arial 8 pt, peso normal, color `#666666`;
- sin cursiva, subrayado, tachado ni formato mixto dentro de una zona.

## Compilación esperada

Ejecutar cada plantilla con un proyecto válido y verificar:

- el diagnóstico no contiene errores;
- la lectura valida que el DOCX fuente use papel A4, sin emitir una opción de
  tamaño de papel en el proyecto ni en `template.json`;
- `template_orientation` es `PORTRAIT` o `LANDSCAPE`, según el archivo;
- los nombres, etiquetas, alineaciones y orden de columnas coinciden con las
  tablas anteriores;
- las columnas sin entrada en `column_widths` reciben `WEIGHT` 1.0; los anchos
  dibujados en Word son solo una ayuda visual y no alteran silenciosamente el
  contrato compilado;
- el estilo del título queda normalizado a `CENTER`;
- el footer conserva ambos marcadores integrados;
- repetir la compilación produce JSON y SQL normalizados idénticos.

## Renderizado e inspección visual

Renderizar cada DOCX con el script canónico de la habilidad de documentos y
emitir PDF solo como apoyo de diagnóstico:

```bash
"$CODEX_PRIMARY_RUNTIME_PYTHON" \
  /root/.codex/skills/builtins/documents/render_docx.py \
  examples/templates/entidades_portrait.docx \
  --output_dir build/qa/entidades_portrait \
  --emit_pdf

"$CODEX_PRIMARY_RUNTIME_PYTHON" \
  /root/.codex/skills/builtins/documents/render_docx.py \
  examples/templates/entidades_landscape.docx \
  --output_dir build/qa/entidades_landscape \
  --emit_pdf
```

Abrir y revisar al 100 % todas las imágenes `page-*.png`. Confirmar:

- orientación y tamaño de página correctos;
- título visible, centrado y sin línea decorativa;
- tabla centrada y contenida dentro de los márgenes;
- ninguna etiqueta o marcador cortado, superpuesto o pegado al borde;
- en la plantilla horizontal, `{{COLUMN:DEPARTAMENTO}}` y
  `{{COLUMN:VNRO_TLF1}}` deben permanecer completos y con un ajuste de línea
  visualmente razonable; no se permite partir los aliases manualmente;
- anchos visualmente distintos y coherentes con el tipo de columna;
- bordes internos y externos visibles y uniformes;
- footer centrado, legible y dentro del área imprimible;
- sin sustitución visible de fuente ni glifos faltantes;
- ninguna página en blanco adicional.

Si falla un punto visual o estructural, corregir el generador, volver a generar
ambos DOCX y repetir toda esta sección. Los PNG y PDF son evidencia interna de
QA; no forman parte de los artefactos finales salvo solicitud expresa.
