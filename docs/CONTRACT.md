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

## 3. Lenguaje DOCX permitido

### 3.1 Estructura

La plantilla debe usar una única sección de Word y contener, en este orden:

1. un único párrafo de encabezado del reporte que contenga `{{REPORT_TITLE}}` exactamente una vez y, opcionalmente, campos escalares en el mismo párrafo;
2. una única tabla principal de dos filas de plantilla;
3. un único párrafo en el footer real de la sección de Word.

Se permiten saltos de línea manuales dentro del párrafo de encabezado, pero toda la zona usa un estilo uniforme porque `APEX_DATA_EXPORT` solo ofrece un `page_header` textual y un estilo global. No se admiten párrafos de contenido libre entre el encabezado y la tabla.

Ese párrafo se compila como `page_header`, no como contenido de cuerpo exclusivo
de la primera página; el renderizador puede repetirlo en cada página del PDF.

La tabla principal tiene esta forma:

- **fila 1, encabezado:** una celda por columna, con la etiqueta visible literal; no contiene marcadores;
- **fila 2, prototipo del cuerpo:** una celda por columna, con exactamente un marcador `{{COLUMN:ALIAS}}`.

Ejemplo conceptual:

| Contenido de la celda de encabezado | Contenido de la celda prototipo |
|---|---|
| `DNI` | `{{COLUMN:VDNI}}` |
| `Nombre completo` | `{{COLUMN:VNOM}}` |

El marcador técnico no aparece en el reporte generado. La fila 1 aporta la etiqueta y el estilo global del encabezado; la fila 2 declara el alias SQL y aporta el estilo global del cuerpo. Cada etiqueta debe ser texto literal no vacío y no puede superar 255 caracteres.

La tabla admite entre 1 y 50 columnas.

### 3.2 Placeholders

Los identificadores no distinguen mayúsculas y minúsculas al validarse y se normalizan a mayúsculas. Deben cumplir:

```text
^[A-Z][A-Z0-9_$#]{0,29}$
```

Placeholders válidos:

| Forma | Uso | Reglas |
|---|---|---|
| `{{REPORT_TITLE}}` | Ubicación del título configurado | Exactamente uno, en el único párrafo previo a la tabla |
| `{{FIELD:NOMBRE}}` | Valor escalar | Solo en el párrafo de encabezado o en el footer; debe tener un mapeo |
| `{{COLUMN:ALIAS}}` | Identidad SQL de una columna y prototipo de celda de datos | Exactamente una vez por celda de la fila 2 |
| `{{GENERATED_AT}}` | Fecha/hora de generación | Azúcar sintáctico para un campo de sistema |
| `{{APP_USER}}` | Usuario APEX | Azúcar sintáctico para un campo de contexto |

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
| Arial, Calibri, Aptos, Helvetica | `HELVETICA` |
| Times New Roman, Cambria, Georgia | `TIMES` |
| Courier New, Consolas | `COURIER` |

Una fuente no reconocida produce error en modo estricto. En modo compatible produce una advertencia y usa `HELVETICA`.

El peso se reduce a `NORMAL` o `BOLD`. Cursiva, subrayado, tachado, espaciado entre caracteres y efectos de Word no son representables en la salida nativa; producen advertencia o error según el modo del compilador.

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
- comentarios o cambios controlados pendientes;
- marcadores desconocidos con forma `{{...}}`.

## 4. Consulta SQL y parámetros

### 4.1 SQL admitido

Cada reporte define una sola sentencia SQL de hasta 32 767 caracteres cuyo primer token efectivo es `SELECT` o `WITH`. No se admite punto y coma terminal, SQL*Plus, bloque PL/SQL, DML ni DDL.

Ejemplo:

```sql
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

```json
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

Tipos permitidos: `VARCHAR2`, `NUMBER`, `DATE` y `TIMESTAMP`. Para `DATE` y `TIMESTAMP`, el manifiesto debe proporcionar una máscara de entrada cuando el Page Item no tenga un valor canónico. Los valores se leen mediante `APEX_SESSION_STATE`, se convierten antes de abrir el contexto y se añaden con `APEX_EXEC.ADD_PARAMETER` en su tipo real.

Reglas:

- todos los binds encontrados en el SQL deben estar mapeados;
- no puede haber mapeos sin bind correspondiente;
- un bind repetido en SQL se mapea una sola vez;
- el nombre del Page Item debe cumplir la convención `P0_...` o `P<page>_...`;
- un valor requerido vacío produce error funcional antes de ejecutar SQL;
- una conversión inválida produce un mensaje que identifica el bind, no el valor sensible.

El compilador no inserta el contenido de los Page Items en el SQL generado.

Para conservar equivalencia con los tipos internos del package, el título se
limita a 255 caracteres, el nombre base del archivo a 180 caracteres, cada
constante o máscara a 4000 bytes UTF-8 y cada nombre de Page Item a 128
caracteres totales.

### 4.3 Campos escalares

Un `{{FIELD:NOMBRE}}` se resuelve mediante uno de estos orígenes:

| Origen | Propiedades | Ejemplo |
|---|---|---|
| `ITEM` | `item`, `type`, `format` opcional | `P42_NOMBRE_REPORTE` |
| `CONTEXT` | `key` en lista permitida | `APP_USER`, `APP_ID`, `APP_PAGE_ID` |
| `SYSTEM` | `key` en lista permitida, `format` opcional | `GENERATED_AT` |
| `CONSTANT` | `value` | nombre de oficina o clasificación |

Ejemplo:

```json
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

En el modo `DOWNLOAD_QUERY` el paquete no intenta reconstruir automáticamente los LOV declarativos de una página o IG. La consulta debe devolver la etiqueta que se desea imprimir, normalmente mediante `JOIN`, vista o función de dominio autorizada.

```sql
select e.departamento_id,
       d.nombre as departamento
  from entidad e
  join departamento d on d.id = e.departamento_id
```

La plantilla usa `{{COLUMN:DEPARTAMENTO}}`, no `DEPARTAMENTO_ID`. Esta decisión evita depender de metadatos internos del IG y hace que la consulta sea comprobable en SQL Workshop.

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

Los anchos dejan de depender de posiciones anónimas. Se configuran por alias:

```json
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
- `WEIGHT` debe ser mayor que cero;
- la suma de `FIXED_PERCENT` debe ser menor que 99;
- una referencia a una columna inexistente es error, para detectar errores tipográficos.

El presupuesto técnico es 99 %, dejando 1 % de tolerancia al renderizador. Cuando existe `AUTO`, el package reserva internamente 20 % como presupuesto heurístico para calcular los anchos numéricos de las demás columnas. Esta constante no forma parte del proyecto, del JSON compilado ni de la API pública. El package sigue omitiendo `p_width` en la columna `AUTO`; por ello, el motor puede asignarle un ancho distinto del reservado. Si no existe `AUTO`, los pesos comparten todo el espacio restante.

Los anchos son indicaciones al motor PDF; saltos de línea, fuentes y contenido pueden obligar al renderizador a ajustar el resultado.

Para llamadas heredadas de IG, `p_column_spans_json` mantiene su semántica posicional y el marcador `"*"`. No se reinterpreta como el nuevo contrato por nombre.

## 7. Definición JSON compilada

El formato canónico es UTF-8 y se serializa de manera determinista: claves ordenadas donde corresponda, arrays en orden semántico y sin datos de ejecución.

Ejemplo abreviado:

```json
{
  "schema_version": "1.0",
  "compiler": {
    "name": "apex-word-report-compiler",
    "version": "1.0.0"
  },
  "report": {
    "id": "ENTIDADES",
    "header_template": "{{REPORT_TITLE}}",
    "title_value": "Relación de entidades",
    "footer_template": "Usuario: {{APP_USER}} | Fecha: {{GENERATED_AT}}",
    "orientation": "AUTO",
    "max_rows": 1000
  },
  "query": {
    "sql": "select ... where (:DNI is null or vdni = :DNI)",
    "bindings": [
      {"bind": "DNI", "item": "P42_DNI", "type": "VARCHAR2", "required": false}
    ]
  },
  "fields": [
    {"name": "SOLICITANTE", "source": "ITEM", "item": "P42_SOLICITANTE", "type": "VARCHAR2"}
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

## 8. API propuesta del paquete

Se añade un procedimiento nuevo sin alterar las firmas actuales:

```plsql
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

El compilador puede generar una llamada con literales `q'~...~'`. El código generado no incluye valores de Page Items; incluye únicamente sus nombres autorizados.

La separación de argumentos evita que el paquete tenga que confiar ciegamente en un JSON monolítico y permite revisar el SQL directamente en APEX. `template.json` sigue siendo el artefacto canónico del compilador; el generador lo descompone en la llamada anterior. El papel es A4 y la heurística de distribución `AUTO` es interna: ninguno de los dos se expone como parámetro público.

Correspondencia exacta de los argumentos JSON:

| Argumento | Fragmento canónico |
|---|---|
| `p_bindings_json` | array `query.bindings` |
| `p_fields_json` | array `fields` |
| `p_columns_json` | array `columns` sin la propiedad `width` |
| `p_style_json` | objeto `styles` |
| `p_excluded_columns_json` | array de strings `excluded_columns` |
| `p_column_widths_json` | array de objetos `{column, mode, value?}` derivado de `columns[*].width` |

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

## 9. Compatibilidad con `PKG_CORPORATE_REPORTS` v7

Las siguientes firmas y comportamientos se conservan:

- `DOWNLOAD_IG`;
- `DOWNLOAD_IG_PDF`;
- uso de `APEX_REGION.EXPORT_DATA` para conservar los display values LOV del IG;
- columnas visibles recibidas desde el navegador;
- exclusiones por nombre;
- spans posicionales y `"*"` adaptativo;
- PDF/XLSX;
- `p_display_columns_json` obsoleto pero aceptado.

`DOWNLOAD_QUERY` comparte helpers internos de validación, impresión y descarga, pero no reemplaza ni redirige automáticamente `DOWNLOAD_IG`. Así se evita cambiar reportes que ya funcionan.

La ampliación debe instalarse con `CREATE OR REPLACE PACKAGE` y `PACKAGE BODY`. Los procesos APEX existentes no requieren edición. El script de instalación debe documentar la versión y ofrecer consultas de verificación de estado del paquete.

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
- Page Item requerido vacío o con tipo inválido.

Una **advertencia** permite continuar, pero queda en el diagnóstico. Ejemplos:

- fuente Word mapeada a una familia aproximada;
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
2. Las llamadas existentes a `DOWNLOAD_IG` y `DOWNLOAD_IG_PDF` siguen funcionando sin cambios.
3. `DOWNLOAD_QUERY` enlaza `VARCHAR2`, `NUMBER`, `DATE` y `TIMESTAMP` con `APEX_EXEC.ADD_PARAMETER`.
4. Un filtro nulo y uno informado producen los conjuntos esperados.
5. SQL mal formado, una columna faltante o un Page Item requerido vacío producen error controlado.
6. Una consulta que devuelve códigos y descripciones imprime la columna descriptiva elegida por la plantilla.
7. Exclusiones y anchos se aplican por nombre, independientemente del orden del `SELECT`.
8. `AUTO` puede declararse en una columna intermedia, solo se permite una y se emite sin `p_width`; la prueba no exige una absorción porcentual exacta del motor.
9. La orientación `AUTO` elige vertical para hasta tres columnas finales y horizontal para cuatro o más, salvo definición de plantilla o elección explícita según la precedencia acordada.
10. Encabezado centrado y footer con usuario/fecha se generan con las familias soportadas.
11. El contexto se cierra ante éxito y error; no quedan colecciones o archivos temporales.
12. `p_max_rows` impide exportaciones ilimitadas.

### 12.3 Prueba manual APEX a cargo del integrador

1. Crear Page Items de filtro y asegurar que se envían a estado de sesión antes de la descarga.
2. Copiar el proceso generado en una página de prueba.
3. Probar PDF, XLSX, orientación automática/vertical/horizontal y cero filas.
4. Confirmar que los LOV se muestran mediante la descripción seleccionada en el SQL.
5. Comparar visualmente la salida con la plantilla dentro de las limitaciones nativas.
6. Ejecutar APEX Debug y comprobar que no aparecen valores sensibles en mensajes propios del paquete.

## 13. Precedencias configurables

Para evitar ambigüedades, la precedencia es:

1. elección explícita del usuario en un Page Item autorizado, si el proceso generado decide pasarla;
2. argumento explícito del proceso APEX;
3. valor compilado desde el DOCX/manifiesto;
4. valor predeterminado del paquete.

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
