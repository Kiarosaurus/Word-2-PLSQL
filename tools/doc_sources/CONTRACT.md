# Contrato funcional y técnico

## Compilador Word restringido para reportes Oracle APEX

**Versión del contrato:** 1.0  
**APEX objetivo:** 24.2  
**Motor de salida en el servidor:** `APEX_DATA_EXPORT`  
**Herramienta de autoría:** Microsoft Word, usada únicamente por personal de Informática  
**Dependencias de ejecución adicionales en el servidor:** ninguna

## 1. Objetivo y frontera del producto

El producto permite que personal de Informática diseñe en Word la apariencia de un reporte tabular, compile localmente esa plantilla y copie en APEX una llamada PL/SQL breve. En ejecución, APEX consulta Oracle con SQL enlazado, sustituye campos escalares con valores autorizados y genera el archivo con `APEX_DATA_EXPORT`.

El DOCX es un artefacto de diseño local. **No se carga, interpreta ni almacena en el servidor APEX.** El compilador produce una definición JSON y un bloque PL/SQL. Por ello, Word y Python no son dependencias del ambiente productivo.

La versión 1.0 tiene deliberadamente el alcance de un reporte tabular corporativo:

- un título;
- campos escalares opcionales dentro del encabezado o footer textual;
- una sola tabla dinámica;
- un pie de página;
- tipografía, tamaño, negrita, color, fondos, borde, alineación, papel y orientación;
- una consulta SQL de solo lectura (`SELECT` o `WITH`);
- parámetros lógicos enlazados a Page Items;
- columnas ordenadas, excluibles y dimensionadas por nombre.

No es un conversor general de Word ni un reemplazo completo de Oracle Reports, BI Publisher o AOP.

## 2. Flujo contractual

1. Informática crea un DOCX conforme al lenguaje restringido de la sección 3.
2. Informática proporciona al compilador la consulta SQL y un manifiesto de mapeos.
3. El compilador valida el DOCX, la consulta y el manifiesto.
4. El compilador emite:
  - `template.json`, definición canónica y versionada;
  - `apex_process.sql`, llamada mínima a `PKG_CORPORATE_REPORTS.DOWNLOAD_QUERY`;
  - diagnóstico de advertencias y errores.
5. El desarrollador copia el bloque generado en un proceso APEX de descarga.
6. En ejecución, el paquete:
  - obtiene Page Items desde estado de sesión;
  - convierte sus tipos;
  - los añade como parámetros de `APEX_EXEC`;
  - ejecuta la consulta;
  - valida las columnas solicitadas;
  - sustituye los campos escalares;
  - aplica estilos, exclusiones y anchos;
  - genera y descarga PDF o XLSX.

```flujo
# Figura 1. Flujo contractual: fase local y fase de ejecución
Informática | DOCX conforme a la sección 3 | marcadores {{REPORT_TITLE}}, {{FIELD:…}}, {{COLUMN:…}}, {{APP_USER}}, {{GENERATED_AT}}
Compilador | Manifiesto .report.json y consulta .sql | generados automáticamente desde el DOCX; editables a mano
? Compilador | ¿DOCX, SQL y manifiesto válidos? | No: diagnóstico con código; no se emite ningún artefacto
Compilador | Emisión automática | template.json, validation.json y apex_process.sql
Desarrollador APEX | Proceso de descarga | contiene el bloque apex_process.sql sin edición
PKG_CORPORATE_REPORTS | DOWNLOAD_QUERY | sesión → binds tipados → APEX_EXEC → columnas → campos → estilos
APEX_DATA_EXPORT | PDF o XLSX | descarga al navegador
```

Los artefactos del paso 4 se emiten siempre de forma automática a partir de las tres entradas; ninguno se redacta a mano. Toda modificación de la plantilla se propaga recompilando. El manifiesto y la consulta permanecen bajo control directo de Informática: pueden generarse como esqueleto desde el DOCX y editarse libremente después.

## 3. Lenguaje DOCX permitido

### 3.1 Estructura

La plantilla debe usar una única sección de Word y contener, en este orden:

1. un único párrafo de encabezado del reporte que contenga `{{REPORT_TITLE}}` exactamente una vez y, opcionalmente, campos escalares en el mismo párrafo;
2. una única tabla principal de dos filas de plantilla;
3. opcionalmente, un único párrafo con contenido en el footer real de la sección de Word.

Se permiten saltos de línea manuales dentro del párrafo de encabezado, pero toda la zona usa un estilo uniforme porque `APEX_DATA_EXPORT` solo ofrece un `page_header` textual y un estilo global. No se admiten párrafos de contenido libre entre el encabezado y la tabla.

Ese párrafo se compila como `page_header`, no como contenido de cuerpo exclusivo de la primera página; el renderizador puede repetirlo en cada página del PDF.

La tabla principal tiene esta forma:

- **fila 1, encabezado:** una celda por columna, con la etiqueta visible literal; no contiene marcadores;
- **fila 2, prototipo del cuerpo:** una celda por columna, con exactamente un marcador `{{COLUMN:ALIAS}}`.

Ejemplo conceptual:

| Contenido de la celda de encabezado | Contenido de la celda prototipo |
|---|---|
| DNI | {{COLUMN:VDNI}} |
| Nombre completo | {{COLUMN:VNOM}} |

El marcador técnico no aparece en el reporte generado. La fila 1 aporta la etiqueta y el estilo global del encabezado; la fila 2 declara el alias SQL y aporta el estilo global del cuerpo. Cada etiqueta debe ser texto literal no vacío y no puede superar 255 bytes UTF-8 (por ejemplo, 255 letras sin tilde o 127 letras acentuadas).

La tabla admite entre 1 y 50 columnas.

### 3.2 Placeholders

Los identificadores no distinguen mayúsculas y minúsculas al validarse y se normalizan a mayúsculas. Deben cumplir:

```
^[A-Z][A-Z0-9_$#]{0,29}$
```

Placeholders válidos:

| Forma | Uso | Reglas |
|---|---|---|
| {{REPORT_TITLE}} | Ubicación del título configurado | Exactamente uno, en el único párrafo previo a la tabla |
| {{FIELD:NOMBRE}} | Valor escalar | Solo en el párrafo de encabezado o en el footer; debe tener un mapeo |
| {{COLUMN:ALIAS}} | Identidad SQL de una columna y prototipo de celda de datos | Exactamente una vez por celda de la fila 2 |
| {{GENERATED_AT}} | Fecha/hora de generación | Azúcar sintáctico para un campo de sistema |
| {{APP_USER}} | Usuario APEX | Azúcar sintáctico para un campo de contexto |

`REPORT_TITLE`, `APP_USER` y `GENERATED_AT` son nombres reservados: `{{FIELD:APP_USER}}` es error; se escribe `{{APP_USER}}`.

Ejemplo normativo que usa las cinco formas (las plantillas de `templates/` siguen exactamente este esquema):

```
Párrafo de encabezado (un solo párrafo; salto de línea manual entre líneas):
  {{REPORT_TITLE}}
  Unidad: {{FIELD:UNIDAD}}

Tabla principal:
  fila 1:  DNI              | Nombre completo  | Departamento
  fila 2:  {{COLUMN:VDNI}}  | {{COLUMN:VNOM}}  | {{COLUMN:DEPARTAMENTO}}

Footer de la sección:
  Usuario: {{APP_USER}} | Fecha: {{GENERATED_AT}}
```

Resultado compilado automáticamente en `template.json` (extracto):

```
"header_template": "{{REPORT_TITLE}}\nUnidad: {{FIELD:UNIDAD}}",   // texto literal del párrafo
"footer_template": "Usuario: {{APP_USER}} | Fecha: {{GENERATED_AT}}",
"fields":  [{"name": "UNIDAD", "source": "CONSTANT", "value": "..."}], // del manifiesto
"columns": [{"name": "VDNI", "heading": "DNI", ...}, ...]          // filas 1 y 2 de la tabla
```

No se admite lógica dentro del DOCX: no hay bucles, condiciones, expresiones, SQL, JavaScript ni PL/SQL embebido.

### 3.3 Formato admitido

El compilador extrae exclusivamente:

- familia tipográfica;
- tamaño tipográfico;
- peso normal o negrita;
- color del texto;
- alineación izquierda, centro o derecha;
- fondo uniforme de la fila de encabezado;
- fondo uniforme del cuerpo;
- color y grosor uniforme del borde;
- tamaño de papel;
- orientación vertical u horizontal.

La alineación del encabezado del reporte se normaliza siempre a centrada. Su título y campos escalares comparten un único estilo. El footer también tiene un único estilo. Si Word contiene formato mixto dentro de una de esas zonas, el compilador emite error en modo estricto.

Familias finales soportadas por `APEX_DATA_EXPORT`:

| Familia Word de ejemplo | Familia compilada |
|---|---|
| Arial, Calibri, Aptos, Helvetica | HELVETICA |
| Times New Roman, Cambria, Georgia | TIMES |
| Courier New, Consolas | COURIER |

Una fuente no reconocida produce error en modo estricto. En modo compatible produce una advertencia y usa `HELVETICA`.

El peso se reduce a `NORMAL` o `BOLD`. Cursiva, subrayado y tachado no son representables en la salida nativa; producen advertencia o error según el modo del compilador. Otros efectos de Word (espaciado entre caracteres, versalitas, resaltado, sombras) no se detectan: simplemente no se trasladan a la salida.

Los anchos de columna dibujados en Word son solo una ayuda visual. El ancho contractual se declara en `column_widths` (sección 6); una columna sin entrada usa `WEIGHT` 1.

### 3.4 Elementos prohibidos

La compilación falla si detecta:

- imágenes, iconos, WordArt o marcas de agua;
- cuadros de texto, formas, gráficos o SmartArt;
- más de una tabla principal;
- celdas combinadas horizontal o verticalmente;
- tablas anidadas;
- más de dos filas de plantilla en la tabla principal;
- encabezados o footers diferentes por primera página, páginas pares o secciones múltiples;
- campos de Word, macros o documentos `.docm`;
- comentarios, notas al pie o al final, o cambios controlados pendientes (incluidos los cambios de formato);
- texto oculto (`w:vanish`), en el documento o en un estilo;
- envoltorios cuyo texto Word muestra pero el compilador no puede interpretar con seguridad: etiquetas inteligentes, XML personalizado, texto ruby, bloques de dirección de texto y contenido alternativo de compatibilidad;
- símbolos insertados, números de página, ecuaciones y subdocumentos;
- cualquier estructura del cuerpo distinta de párrafos, la tabla y marcadores de posición (bookmarks o permisos de edición), y tablas en el encabezado o pie de Word;
- marcadores desconocidos con forma `{{...}}`, o con espacios no ASCII (por ejemplo, un espacio de no separación) dentro de las llaves.

## 4. Consulta SQL y parámetros

### 4.1 SQL admitido

Cada reporte define una sola sentencia SQL de hasta 32 767 bytes UTF-8 cuyo primer token efectivo es `SELECT` o `WITH`. Antes de ese token solo se admiten comentarios y espacios, tabuladores o saltos de línea ASCII. No se admite punto y coma terminal, SQL*Plus, bloque PL/SQL, DML, DDL, `SELECT ... INTO` ni `FOR UPDATE`.

Tampoco se admite la sintaxis de sustitución de APEX `&ITEM.` fuera de literales y comentarios: APEX la reemplazaría por el valor de sesión antes de ejecutar el proceso, lo que equivale a concatenar. Para impedirlo en cualquier caso, el código generado emite cada carácter `&` de cualquier literal como `chr(38)`; por eso un `&ITEM.` dentro de un literal o comentario se conserva como texto y solo produce una advertencia.

Ejemplo:

```
select e.vdni,
       e.vnom,
       d.descripcion as departamento
  from entidad e
  left join departamento d
    on d.id = e.departamento_id
 where (:DNI is null or e.vdni = :DNI)
   and (:ACTIVO is null or e.activo = :ACTIVO)
 order by e.vnom
```

Las variables `:DNI` y `:ACTIVO` son **binds lógicos**, no nombres de Page Items. Nunca se concatenan valores en el SQL.

El compilador realiza una validación léxica conservadora. El paquete vuelve a validar la forma básica y ejecuta con `APEX_EXEC.OPEN_QUERY_CONTEXT`. Esta validación no sustituye los privilegios de Oracle: el parsing schema de APEX debe tener solo los permisos necesarios.

### 4.2 Mapeo de binds lógicos a Page Items

El manifiesto contiene un mapeo explícito por cada bind:

```
[
  {
    "bind": "DNI",
    "item": "P42_DNI",
    "type": "VARCHAR2",
    "required": false
  },
  {
    "bind": "ACTIVO",
    "item": "P42_ACTIVO",
    "type": "VARCHAR2",
    "required": false
  }
]
```

Tipos permitidos: `VARCHAR2`, `NUMBER`, `DATE` y `TIMESTAMP`. Para `DATE` y `TIMESTAMP`, `format_mask` es obligatorio y debe coincidir con la máscara del Page Item. En `NUMBER` la conversión usa la configuración del item y de la sesión. Los valores se leen mediante `APEX_SESSION_STATE`, se convierten antes de abrir el contexto y se añaden con `APEX_EXEC.ADD_PARAMETER` en su tipo real.

Reglas:

- todos los binds encontrados en el SQL deben estar mapeados;
- no puede haber mapeos sin bind correspondiente;
- un bind repetido en SQL se mapea una sola vez;
- el nombre del Page Item debe cumplir la convención `P0_...` o `P<page>_...`;
- un bind no puede usar un nombre reservado de APEX (`APP_USER`, `APP_ID`, `APP_PAGE_ID`, `APP_SESSION`, `REQUEST`, `DEBUG`, `SESSION`, entre otros): su valor provendría de un Page Item modificable y no del contexto autenticado. Para imprimir el usuario use `{{APP_USER}}` o un campo `CONTEXT`;
- un valor requerido vacío produce error funcional antes de ejecutar SQL;
- una conversión inválida produce un mensaje que identifica el bind, no el valor sensible.

El compilador no inserta el contenido de los Page Items en el SQL generado.

Para conservar equivalencia con los tipos internos del package, que miden en bytes, el título se limita a 255 bytes UTF-8, el nombre base del archivo a 180 caracteres ASCII, cada constante o máscara a 4000 bytes UTF-8 y cada nombre de Page Item a 128 caracteres. El encabezado, con el título ya sustituido, no puede superar 4000 bytes; el package vuelve a comprobarlo con los valores reales.

### 4.3 Campos escalares

Un `{{FIELD:NOMBRE}}` se resuelve mediante uno de estos orígenes:

| Origen | Propiedades | Ejemplo |
|---|---|---|
| ITEM | item, type, format opcional | P42_NOMBRE_REPORTE |
| CONTEXT | key en lista permitida | APP_USER, APP_ID, APP_PAGE_ID |
| SYSTEM | key en lista permitida, format opcional | GENERATED_AT |
| CONSTANT | value | nombre de oficina o clasificación |

Ejemplo:

```
[
  {
    "name": "SOLICITANTE",
    "source": "ITEM",
    "item": "P42_SOLICITANTE",
    "type": "VARCHAR2"
  },
  {
    "name": "UNIDAD",
    "source": "CONSTANT",
    "value": "Oficina de Informática"
  }
]
```

Solo se sustituyen placeholders declarados. Un placeholder sin mapeo o un mapeo no usado es error de compilación. El paquete limita la longitud final del encabezado y footer a la capacidad de `APEX_DATA_EXPORT` y rechaza residuos `{{...}}`.

### 4.4 Valores LOV

En el modo `DOWNLOAD_QUERY` el paquete no intenta reconstruir automáticamente los LOV declarativos de la página. La consulta debe devolver la etiqueta que se desea imprimir, normalmente mediante `JOIN`, vista o función de dominio autorizada.

```
select e.departamento_id,
       d.nombre as departamento
  from entidad e
  join departamento d on d.id = e.departamento_id
```

La plantilla usa `{{COLUMN:DEPARTAMENTO}}`, no `DEPARTAMENTO_ID`. Esta decisión evita depender de metadatos internos de APEX y hace que la consulta sea comprobable en SQL Workshop.

## 5. Columnas, exclusiones y formatos

El orden contractual es el orden de las celdas del DOCX. Cada alias `COLUMN` debe tener una columna con ese nombre en el contexto de `APEX_EXEC`.

- una columna requerida por la plantilla pero ausente en SQL produce error;
- columnas adicionales de SQL se ignoran;
- aliases duplicados o no válidos producen error;
- las exclusiones se aplican después de validar el contrato y antes de calcular anchos;
- excluir todas las columnas produce error;
- una exclusión inexistente genera advertencia del compilador y se ignora en ejecución.

El manifiesto puede definir `format_mask` y alineación por columna. Si no hay alineación explícita, números se alinean a la derecha y los demás tipos a la izquierda.

## 6. Anchos por nombre

Los anchos se configuran por alias, nunca por posición:

```
[
  {"column": "VDNI", "mode": "FIXED_PERCENT", "value": 14},
  {"column": "VNOM", "mode": "WEIGHT", "value": 2},
  {"column": "DEPARTAMENTO", "mode": "AUTO"}
]
```

Semántica:

- `FIXED_PERCENT`: porcentaje solicitado del ancho imprimible;
- `WEIGHT`: participación relativa en el espacio disponible después de los `FIXED_PERCENT`;
- `AUTO`: omite `p_width` para esa columna y delega su ajuste al motor de APEX; es una indicación, no una garantía de que absorberá una cantidad exacta de espacio;
- una columna sin entrada se trata como `WEIGHT` con valor `1`;
- se permite como máximo una columna `AUTO`, en cualquier posición;
- `FIXED_PERCENT` debe estar entre 1 y 95;
- `WEIGHT` debe ser mayor que cero y no superar 1000;
- la suma de `FIXED_PERCENT` debe ser menor que 99;
- una referencia a una columna inexistente es error, para detectar errores tipográficos.

El presupuesto técnico es 99 %, dejando 1 % de tolerancia al renderizador. Cuando existe `AUTO`, el package reserva internamente 20 % como presupuesto heurístico para calcular los anchos numéricos de las demás columnas. Esta constante no forma parte del proyecto, del JSON compilado ni de la API pública. El package sigue omitiendo `p_width` en la columna `AUTO`; por ello, el motor puede asignarle un ancho distinto del reservado. Si no existe `AUTO`, los pesos comparten todo el espacio restante.

Los anchos son indicaciones al motor PDF; saltos de línea, fuentes y contenido pueden obligar al renderizador a ajustar el resultado.

## 7. Definición JSON compilada

El formato canónico es UTF-8 y se serializa de manera determinista: claves ordenadas donde corresponda, arrays en orden semántico y sin datos de ejecución.

`template.json` se genera **automáticamente** en cada compilación; nunca es una entrada. Cuando la plantilla cambia, basta recompilar para que la definición refleje las nuevas etiquetas, columnas, textos de encabezado y pie, estilos y orientación. El archivo puede abrirse y editarse directamente para inspección o pruebas, pero la siguiente compilación lo reemplaza; los ajustes permanentes se expresan en el manifiesto `.report.json` (por ejemplo, `style_overrides` prevalece sobre los estilos extraídos del DOCX).

El manifiesto `.report.json` también puede crearse automáticamente desde el DOCX («Nuevo proyecto desde DOCX»): un binding opcional por cada `{{COLUMN:...}}` y un campo `ITEM` por cada `{{FIELD:...}}`. El DOCX se copia a `proyectos/<nombre>/` y el manifiesto y la consulta se escriben en `proyectos/<nombre>/generado/`. Volver a cargar un DOCX con el mismo nombre exige confirmación y regenera ambos, conservando los anteriores como `generado/<archivo>.bak`; los demás archivos de la carpeta no se modifican. Entre cargas, el manifiesto se mantiene editándolo directamente.

Ubicación de los artefactos de cada reporte:

| Archivo | Ubicación | Origen |
|---|---|---|
| <nombre>.docx | proyectos/<nombre>/ | copia del DOCX cargado |
| apex_process.sql | proyectos/<nombre>/ | compilación (único archivo que se pega en APEX) |
| <nombre>.sql | proyectos/<nombre>/generado/ | esqueleto automático; su texto se incrusta en apex_process.sql |
| <nombre>.report.json | proyectos/<nombre>/generado/ | esqueleto automático; editable |
| template.json, validation.json | proyectos/<nombre>/generado/ | compilación |

### 7.1 Propiedades del proyecto `.report.json`

| Propiedad | Obligatoria | Regla |
|---|---|---|
| schema | sí | exactamente corporate-report-project/1.0 |
| report_id | sí | identificador de 1 a 30 caracteres |
| template, query_file | sí | rutas relativas al .report.json que no salen de la carpeta del proyecto: proyectos/<nombre>/ si el manifiesto está en generado/, o la carpeta del propio manifiesto en otro caso |
| title | sí | texto JSON, máximo 255 bytes UTF-8 |
| file_name | no | letras, números, _ o -; máximo 180; por defecto report_id en minúsculas |
| max_rows | no | entero de 1 a 100 000; por defecto 1000 |
| orientation | no | AUTO, PORTRAIT o LANDSCAPE; si falta, se usa la orientación de la página Word |
| format_item | no | Page Item cuyo valor (PDF o XLSX) elige el formato en ejecución |
| orientation_item | no | Page Item cuyo valor (AUTO, PORTRAIT, LANDSCAPE) elige la orientación |
| bindings, fields, excluded_columns, column_widths | no | secciones 4 a 6 |
| columns | no | array {name, alignment, format_mask} que ajusta una columna por alias |
| style_overrides | no | sustituye estilos extraídos del DOCX (ver abajo) |

Todas las propiedades textuales deben ser cadenas JSON; un número, `null` u objeto en su lugar es error. No se admiten claves JSON duplicadas.

`style_overrides` admite estas zonas y claves, idénticas a las que acepta el package:

| Zona | Claves |
|---|---|
| title | font_family, font_size, font_weight, font_color, alignment (solo CENTER) |
| table_header | las anteriores y background_color |
| table_body | font_family, font_size, font_weight, font_color, background_color (la alineación es por columna; alignment se ignora con advertencia) |
| border | width, color |
| footer | font_family, font_size, font_weight, font_color, alignment |

Si `orientation` es `PORTRAIT` o `LANDSCAPE` y contradice la página del DOCX, el compilador emite una advertencia. La orientación `AUTO` solo se aplica cuando se declara explícitamente en el proyecto o la elige el Page Item `orientation_item`.

### 7.2 Ejemplo de `template.json`

Ejemplo abreviado:

```
{
  "schema_version": "1.0",
  "compiler": {
    "name": "apex-word-report-compiler",
    "version": "1.0.0"
  },
  "source": {
    "file_name": "entidades.docx",
    "sha256": "..."
  },
  "report": {
    "id": "ENTIDADES",
    "header_template": "{{REPORT_TITLE}}\nUnidad: {{FIELD:UNIDAD}}",
    "title_value": "Relación de entidades",
    "footer_template": "Usuario: {{APP_USER}} | Fecha: {{GENERATED_AT}}",
    "orientation": "AUTO",
    "max_rows": 1000,
    "file_name": "reporte_entidades"
  },
  "query": {
    "sql": "select ... where (:DNI is null or vdni = :DNI)",
    "bindings": [
      {"bind": "DNI", "item": "P42_DNI", "type": "VARCHAR2", "required": false}
    ]
  },
  "fields": [
    {"name": "UNIDAD", "source": "CONSTANT", "value": "Oficina de Informática"}
  ],
  "columns": [
    {
      "name": "VDNI",
      "heading": "DNI",
      "alignment": "START",
      "format_mask": null,
      "width": {"mode": "FIXED_PERCENT", "value": 14}
    },
    {
      "name": "VNOM",
      "heading": "Nombre completo",
      "alignment": "START",
      "format_mask": null,
      "width": {"mode": "AUTO"}
    }
  ],
  "excluded_columns": [],
  "styles": {
    "title": {
      "font_family": "HELVETICA",
      "font_size": 15,
      "font_weight": "BOLD",
      "font_color": "#2F343A",
      "alignment": "CENTER"
    },
    "table_header": {
      "background_color": "#4A4F55",
      "font_family": "HELVETICA",
      "font_size": 9,
      "font_weight": "BOLD",
      "font_color": "#FFFFFF",
      "alignment": "CENTER"
    },
    "table_body": {
      "background_color": "#FFFFFF",
      "font_family": "HELVETICA",
      "font_size": 8.5,
      "font_weight": "NORMAL",
      "font_color": "#25282B"
    },
    "border": {"width": 0.5, "color": "#BFC3C7"},
    "footer": {
      "font_family": "HELVETICA",
      "font_size": 8,
      "font_weight": "NORMAL",
      "font_color": "#666666",
      "alignment": "CENTER"
    }
  }
}
```

Restricciones de estilo:

- color: `^#[0-9A-F]{6}$`;
- título: 8 a 24 puntos;
- encabezado: 6 a 16 puntos;
- cuerpo: 6 a 14 puntos;
- footer: 6 a 12 puntos;
- borde: 0 a 5 puntos;
- alineación: `START`, `CENTER` o `END`;
- fondos y bordes son uniformes para toda la tabla, pues `APEX_DATA_EXPORT` no reproduce estilos arbitrarios por celda.

## 8. API pública del paquete

La operación pública que invoca el código generado es `DOWNLOAD_QUERY`. Esta firma está congelada para la versión 1.0:

```
PROCEDURE download_query(
    p_sql_query              IN VARCHAR2,
    p_columns_json           IN CLOB,
    p_bindings_json          IN CLOB DEFAULT NULL,
    p_fields_json            IN CLOB DEFAULT NULL,
    p_style_json             IN CLOB DEFAULT NULL,
    p_header_template        IN VARCHAR2 DEFAULT '{{REPORT_TITLE}}',
    p_footer_template        IN VARCHAR2 DEFAULT NULL,
    p_title                  IN VARCHAR2 DEFAULT 'Reporte',
    p_file_name              IN VARCHAR2 DEFAULT 'reporte',
    p_format                 IN VARCHAR2 DEFAULT 'PDF',
    p_orientation            IN VARCHAR2 DEFAULT 'AUTO',
    p_max_rows               IN PLS_INTEGER DEFAULT 1000,
    p_excluded_columns_json  IN CLOB DEFAULT NULL,
    p_column_widths_json     IN CLOB DEFAULT NULL
);
```

El compilador genera una llamada con literales `q'~...~'` (y `chr(38)` en lugar de cada `&`). El código generado no incluye valores de Page Items; incluye únicamente sus nombres autorizados. Cuando el proyecto declara `format_item` u `orientation_item`, el argumento se emite como `coalesce(:ITEM, valor_compilado)`.

La separación de argumentos evita que el paquete tenga que confiar ciegamente en un JSON monolítico y permite revisar el SQL directamente en APEX. `template.json` sigue siendo el artefacto canónico del compilador; el generador lo descompone automáticamente en los argumentos de esta llamada. El papel es A4 y la heurística de distribución `AUTO` es interna: ninguno de los dos se expone como parámetro público.

Correspondencia exacta de los argumentos JSON:

| Argumento | Fragmento canónico |
|---|---|
| p_bindings_json | array query.bindings |
| p_fields_json | array fields |
| p_columns_json | array columns sin la propiedad width |
| p_style_json | objeto styles |
| p_excluded_columns_json | array de strings excluded_columns |
| p_column_widths_json | array de objetos {column, mode, value?} derivado de columns[*].width |

Los valores `null` o arrays vacíos se emiten como `NULL` cuando el procedimiento admite un valor predeterminado. El generador debe escoger un delimitador de quoting alternativo que no aparezca en el contenido; si no existe uno seguro, debe fragmentar el literal CLOB de forma determinista.

### 8.1 Comportamiento de ejecución

`DOWNLOAD_QUERY` debe:

1. validar formato, límites, orientación y JSON;
2. verificar que la sentencia comienza con `SELECT` o `WITH` y no contiene delimitadores de varias sentencias;
3. leer y convertir cada Page Item;
4. poblar `APEX_EXEC.T_PARAMETERS` con `APEX_EXEC.ADD_PARAMETER`;
5. abrir el contexto local con `APEX_EXEC.OPEN_QUERY_CONTEXT` y `p_auto_bind_items => FALSE`;
6. resolver las columnas por nombre con `APEX_EXEC.GET_COLUMN_POSITION`;
7. construir `APEX_DATA_EXPORT.T_COLUMNS` solo con las columnas contratadas y no excluidas;
8. sustituir campos escalares en encabezado y footer mediante una lista cerrada;
9. construir `APEX_DATA_EXPORT.T_PRINT_CONFIG` con valores validados;
10. exportar, cerrar el contexto aun ante error y descargar el resultado.

Ningún valor de usuario se concatena al SQL, a un identificador SQL ni a un nombre de columna.

```flujo
# Figura 2. Ejecución de DOWNLOAD_QUERY
? Paquete | ¿Formato, límites, orientación y JSON válidos? | No: error controlado sin valores sensibles
? Paquete | ¿La sentencia empieza con SELECT o WITH y es única? | No: error controlado
Paquete | Leer y convertir cada Page Item | VARCHAR2, NUMBER, DATE, TIMESTAMP con su máscara
APEX_EXEC | Abrir el contexto con parámetros tipados | p_auto_bind_items => FALSE
? Paquete | ¿Existen todas las columnas contratadas? | No: error controlado; el contexto se cierra
Paquete | Construir columnas, encabezado, pie y estilos | sustituye REPORT_TITLE, FIELD, APP_USER y GENERATED_AT
APEX_DATA_EXPORT | Exportar y descargar | el contexto se cierra también ante error
```

## 9. Instalación del paquete

El paquete se instala con `CREATE OR REPLACE PACKAGE` y `PACKAGE BODY` mediante `sql/install.sql`, que detiene la ejecución ante un error, consulta `USER_ERRORS` y exige que ambos objetos queden `VALID`. Se instala una sola vez por parsing schema; los procesos APEX generados para cada reporte no requieren edición cuando se reinstala el mismo archivo.

## 10. Seguridad

### 10.1 Confianza y autoría

- Solo Informática puede crear DOCX, manifiestos y procesos de reporte.
- Los usuarios finales solo proporcionan valores mediante Page Items controlados.
- La herramienta local no requiere credenciales de Oracle ni conexión a la base de datos.
- El DOCX no llega al servidor.

### 10.2 SQL

- solo `SELECT`/`WITH`;
- binds obligatorios, nunca concatenación;
- `p_auto_bind_items => FALSE`;
- lista exacta de bind mappings;
- parsing schema con privilegio mínimo;
- vistas de negocio preferidas para ocultar tablas sensibles;
- `p_max_rows` obligatorio y limitado por una cota institucional; el motor exporta como máximo esa cantidad y puede truncar filas adicionales;
- tiempo y costo de consultas controlados por políticas Oracle, no por el compilador;
- no se acepta SQL enviado por el navegador.

La validación textual no puede demostrar que una función invocada por un `SELECT` carezca de efectos laterales. La defensa real es restringir quién instala reportes, los privilegios del parsing schema y las funciones ejecutables.

### 10.3 Salida

- nombres de archivo saneados;
- textos de encabezado y footer limitados;
- colores, fuentes, tamaños, formatos y alineaciones en listas permitidas;
- JSON con `ERROR ON ERROR` donde una omisión deba bloquear, y mensajes funcionales sin valores sensibles;
- cierre de `APEX_EXEC` en excepción;
- no registrar SQL con valores ni contenido sensible en mensajes de usuario;
- PDF accesible cuando el formato lo soporte.

## 11. Errores y advertencias

Un **error** impide compilar o descargar. Ejemplos:

- estructura DOCX no admitida;
- placeholder desconocido, duplicado o sin mapeo;
- bind sin mapeo;
- columna de plantilla ausente en la consulta;
- JSON inválido;
- color, fuente, tipo o ancho fuera de rango;
- SQL que no comienza con `SELECT`/`WITH`;
- cero columnas finales;
- Page Item requerido vacío o con tipo inválido;
- sustitución `&ITEM.`, `INTO` o `FOR UPDATE` en el SQL;
- bind con nombre reservado de APEX;
- carpeta de salida que sobrescribiría un archivo fuente del proyecto.

Una **advertencia** permite continuar, pero queda en el diagnóstico. Ejemplos:

- fuente Word mapeada a una familia aproximada;
- orientación del proyecto distinta de la página del DOCX;
- estilo Word no representable e ignorado;
- exclusión que no coincide con una columna;
- ancho que puede ser reajustado por el motor PDF;
- XLSX que no conserva todos los estilos de impresión PDF.

## 12. Criterios de aceptación

### 12.1 Compilador local

1. Compila un DOCX A4 vertical válido con tres columnas y produce JSON determinista.
2. Compila un DOCX horizontal con una columna `AUTO` intermedia.
3. Detecta imágenes, celdas combinadas, segunda tabla, campos Word y placeholders desconocidos.
4. Detecta SQL no permitido, binds sin mapeo, mapeos sobrantes y aliases duplicados.
5. Mapea correctamente las tres familias tipográficas finales.
6. Extrae encabezado de reporte, encabezado de tabla, cuerpo, borde y footer dentro de las tolerancias documentadas.
7. Genera un bloque PL/SQL sintácticamente estable, sin valores de sesión incrustados.
8. Repetir la compilación con las mismas entradas produce bytes equivalentes en JSON y SQL normalizados.
9. La suite local no requiere Oracle, Word automatizado, LibreOffice ni Docker.

### 12.2 Paquete PL/SQL

1. `PKG_CORPORATE_REPORTS` compila en APEX 24.2 sin errores.
2. `DOWNLOAD_QUERY` enlaza `VARCHAR2`, `NUMBER`, `DATE` y `TIMESTAMP` con `APEX_EXEC.ADD_PARAMETER`.
3. Un filtro nulo y uno informado producen los conjuntos esperados.
4. SQL mal formado, una columna faltante o un Page Item requerido vacío producen error controlado.
5. Una consulta que devuelve códigos y descripciones imprime la columna descriptiva elegida por la plantilla.
6. Exclusiones y anchos se aplican por nombre, independientemente del orden del `SELECT`.
7. `AUTO` puede declararse en una columna intermedia, solo se permite una y se emite sin `p_width`; la prueba no exige una absorción porcentual exacta del motor.
8. La orientación `AUTO` elige vertical para hasta tres columnas finales y horizontal para cuatro o más, salvo definición de plantilla o elección explícita según la precedencia acordada.
9. Encabezado centrado con título y campos `{{FIELD:...}}`, y footer con usuario/fecha, se generan con las familias soportadas.
10. El contexto se cierra ante éxito y error; no quedan colecciones o archivos temporales.
11. `p_max_rows` impide exportaciones ilimitadas.

### 12.3 Prueba manual APEX a cargo del integrador

1. Crear Page Items de filtro y asegurar que se envían a estado de sesión antes de la descarga.
2. Copiar el proceso generado en una página de prueba.
3. Probar PDF, XLSX, orientación automática/vertical/horizontal y cero filas.
4. Confirmar que los LOV se muestran mediante la descripción seleccionada en el SQL.
5. Comparar visualmente la salida con la plantilla dentro de las limitaciones nativas.
6. Ejecutar APEX Debug y comprobar que no aparecen valores sensibles en mensajes propios del paquete.

## 13. Precedencias configurables

Para evitar ambigüedades, la precedencia efectiva del código generado es:

1. valor no nulo del Page Item `format_item` u `orientation_item`, si el proyecto lo declara (`coalesce(:ITEM, ...)`);
2. valor compilado: para la orientación, `orientation` del proyecto o, si falta, la orientación de la página Word; para el formato, `PDF`;
3. valor predeterminado del paquete, solo si se invoca el procedimiento sin ese argumento.

Los estilos siguen la misma idea: `style_overrides` del proyecto prevalece sobre el estilo extraído del DOCX.

Las exclusiones institucionales del proceso siempre prevalecen sobre la plantilla y el usuario. El usuario final no puede volver visible una columna excluida por Informática.

## 14. Limitaciones conocidas

- La salida no es una reproducción pixel-perfect del DOCX.
- El DOCX actúa como diseñador restringido de propiedades que sí entiende `APEX_DATA_EXPORT`.
- No hay imágenes, logos, múltiples tablas, subreportes, gráficos, agrupaciones visuales ni totales arbitrarios en 1.0.
- Encabezado y footer se reducen a las capacidades textuales globales del motor nativo; no hay formato rico por fragmento ni bloques Word arbitrarios.
- No hay resolución genérica de LOV en modo SQL; la consulta debe devolver el display value.
- Los anchos son indicaciones y pueden variar por contenido y renderizador.
- El compilador no verifica contra Oracle que la consulta compile; esa prueba se realiza en SQL Workshop/APEX.
- XLSX puede ignorar propiedades que solo aplican a impresión PDF.
- La seguridad de la consulta depende también del parsing schema, vistas, funciones ejecutables y gobierno de cambios.

Estas limitaciones son parte del contrato y no defectos de la versión 1.0.
