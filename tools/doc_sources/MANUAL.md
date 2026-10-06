# Manual de uso

## Compilador local Word → reportes Oracle APEX 24.2

**Versión:** 1.0  
**Destinatarios:** personal de Informática que diseña reportes y desarrolladores Oracle APEX  
**Salida:** PDF o XLSX mediante `APEX_DATA_EXPORT`

## 1. Qué resuelve esta herramienta

La herramienta permite diseñar en Microsoft Word un reporte tabular corporativo, validar la plantilla en una computadora local y generar el bloque PL/SQL mínimo que se copiará en APEX. Word no genera el PDF y el DOCX no se carga al servidor: la apariencia compatible se traduce a configuración declarativa y el paquete común `PKG_CORPORATE_REPORTS` genera la descarga con APIs nativas de APEX.

En la VM de APEX/ORDS no se instalan Word, LibreOffice, Python, Docker ni un motor de impresión. Python y Word se usan únicamente en el equipo del diseñador.

El flujo está pensado para reportes de una consulta SQL, un título, una tabla y un pie. No es un conversor general de Word ni reproduce diseños libres de Oracle Reports.

```flujo
# Figura 1. Flujo completo de un reporte
Informática | Diseñar la plantilla Word | sección 4: título, tabla de dos filas y pie con marcadores {{...}}
Compilador | Nuevo proyecto desde DOCX | copia el Word a proyectos\<nombre>\ y crea automáticamente generado\<nombre>.sql y .report.json (sección 2.2)
Informática | Completar SQL y proyecto | secciones 5 y 6; ambos archivos se pueden editar a mano en cualquier momento
Compilador | Validar | revisa DOCX, SQL y proyecto sin escribir archivos (sección 7)
? Compilador | ¿Sin errores? | No: corregir según el código del diagnóstico (sección 11) y validar otra vez
Compilador | Compilar | genera automáticamente apex_process.sql junto al DOCX, y template.json y validation.json en generado\
APEX | Instalar el package (una vez) y pegar apex_process.sql | secciones 8 y 9
Usuario final | Pulsar Descargar | APEX_DATA_EXPORT entrega el PDF o el XLSX
```

## 2. Conceptos esenciales

| Elemento | Dónde vive | Responsabilidad |
|---|---|---|
| Plantilla .docx | Equipo local y repositorio de Informática | Define título, columnas y estilos admitidos |
| Consulta .sql | Proyecto del reporte | Devuelve los datos y los textos que se imprimirán |
| Proyecto .report.json | Proyecto del reporte | Une plantilla, SQL, Page Items, exclusiones y anchos |
| Compilador | Equipo local | Valida y genera template.json, validation.json y apex_process.sql |
| PKG_CORPORATE_REPORTS | Base de datos de APEX | Ejecuta el SQL con binds y descarga PDF/XLSX |
| Page Items | Página APEX | Proporcionan filtros mediante estado de sesión |

### 2.1 Qué se genera automáticamente y qué se edita a mano

La herramienta automatiza todo lo que puede deducirse de la plantilla Word. Ningún archivo queda bloqueado: el usuario siempre puede abrir y editar directamente el `.report.json` y el `.sql` con un editor de texto.

| Archivo | Cómo se crea | Cuándo se actualiza | Edición directa |
|---|---|---|---|
| <nombre>.docx | A mano en Word; «Nuevo proyecto desde DOCX» lo copia a proyectos\<nombre>\ | Cuando el diseñador la cambia o se vuelve a cargar | Sí: es la fuente del diseño |
| generado\<nombre>.report.json | **Automáticamente** con «Nuevo proyecto desde DOCX», a partir de los marcadores del Word | Al cargar el DOCX; después se ajusta a mano | Sí, en cualquier momento |
| generado\<nombre>.sql | **Automáticamente** (esqueleto con un filtro por columna) con «Nuevo proyecto desde DOCX» | Al cargar el DOCX; después se ajusta a mano | Sí, en cualquier momento |
| generado\template.json | **Automáticamente** en cada compilación, leyendo el DOCX, el SQL y el proyecto | En cada compilación | Posible, pero la siguiente compilación lo reemplaza |
| generado\validation.json | **Automáticamente** en cada compilación | En cada compilación | No tiene sentido editarlo |
| apex_process.sql | **Automáticamente** en cada compilación, junto al DOCX | En cada compilación | Posible antes de pegarlo en APEX, pero la siguiente compilación lo reemplaza |

Al cambiar la plantilla Word **no hace falta rehacer ningún JSON a mano**: basta con volver a compilar y `template.json` y `apex_process.sql` se regeneran solos con las nuevas etiquetas, estilos, orden de columnas, título, pie y orientación. Solo se toca el proyecto o el SQL cuando el cambio de Word introduce algo que no puede deducirse del documento:

| Cambio en la plantilla Word | Qué hacer |
|---|---|
| Etiquetas de la fila 1, fuentes, tamaños, colores, fondos, bordes, orientación, textos fijos del título o del pie, orden de columnas | Nada: recompilar. Todo se toma automáticamente del DOCX |
| Nueva columna {{COLUMN:ALIAS}} | Añadir ALIAS al SELECT del .sql; opcionalmente su ancho en column_widths |
| Columna eliminada | Quitar su alias de column_widths (si no, PROJECT-052) y de excluded_columns (si no, advertencia PROJECT-042) |
| Nuevo {{FIELD:NOMBRE}} | Añadir su entrada en fields (sección 6.2) |
| {{FIELD:NOMBRE}} eliminado | Quitar su entrada de fields |
| Plantilla rehecha por completo | Volver a usar «Nuevo proyecto desde DOCX» con el mismo nombre: regenera el proyecto y guarda los anteriores como .bak (sección 2.2) |

El validador local indica qué falta en el proyecto (por ejemplo `TOKEN-020` para un `{{FIELD:...}}` sin entrada en `fields`, o `PROJECT-052` para un ancho de una columna que ya no está en el Word). El compilador no ejecuta el SQL: un alias de `{{COLUMN:...}}` ausente en el `SELECT` lo detecta el package al descargar, con un error controlado. Por eso conviene probar el SQL en SQL Workshop después de añadir columnas.

Si se desea un estilo distinto del que tiene el Word sin abrir Word, puede escribirse directamente en `style_overrides` del `.report.json` (sección 6); ese valor prevalece sobre el extraído automáticamente del DOCX.

### 2.2 Carpeta de proyectos

Cada reporte vive en `proyectos\<nombre>\`, donde `<nombre>` es el nombre del DOCX. En la carpeta del reporte queda solo lo que se mira o se sube; el material de trabajo va en la subcarpeta `generado\`, que existe siempre:

```
proyectos\
  ventas\
    ventas.docx               <- plantilla (se edita en Word)
    apex_process.sql          <- lo único que se pega en APEX
    generado\                 <- no se sube a APEX
      ventas.sql              <- consulta fuente; su texto se incrusta en apex_process.sql
      ventas.report.json      <- binds, campos, anchos, estilos
      template.json           <- definición compilada
      validation.json         <- diagnósticos
```

El `.sql` no se sube a APEX como archivo: al compilar, la consulta se copia dentro de `apex_process.sql`. El package común `sql\modo_simple\pkg_corporate_reports.sql` se instala aparte, una sola vez por esquema (sección 8).

Reglas de la carpeta:

- Si el DOCX elegido está en otra carpeta (Descargas, correo, red), se **copia** a `proyectos\<nombre>\<nombre>.docx`. Si ya está allí, se usa en su sitio.
- Si se vuelve a cargar un DOCX con el **mismo nombre**, la herramienta pide confirmación y reemplaza solo sus archivos conocidos: el DOCX, `apex_process.sql` y los cuatro de `generado\`. De cada uno guarda antes una copia `generado\<archivo>.bak` (la siguiente carga sobrescribe esa copia).
- Cualquier **otro archivo** que se guarde en la carpeta del reporte o en `generado\` (notas, capturas, documentación) no se toca nunca.
- La compilación reemplaza únicamente `apex_process.sql`, `generado\template.json` y `generado\validation.json`.
- Las rutas del `.report.json` no pueden salir de `proyectos\<nombre>\` (`PROJECT-005`).

Los proyectos de demostración `proyectos\demo\entidades\` (modo simple) y `proyectos\demo\estado_cuenta\` (modo layout) siguen la misma organización; su documentación común es `proyectos\demo\README.docx`.

```flujo
# Figura 2. Cargar un DOCX en la carpeta de proyectos
Usuario | Elegir el .docx | desde cualquier carpeta
Compilador | Validar la plantilla | si no es válida, se informa y no se toca ningún archivo
? Compilador | ¿proyectos\<nombre>\ está vacío o no existe? | No: pedir confirmación; si se acepta, mover cada archivo conocido a generado\<archivo>.bak
Compilador | Copiar el Word a proyectos\<nombre>\<nombre>.docx | se omite si el DOCX ya estaba allí
Compilador | Escribir generado\<nombre>.sql y generado\<nombre>.report.json | esqueleto: un alias y un filtro opcional por columna; un campo por {{FIELD:…}}
Usuario | Ajustar SQL y proyecto, validar y compilar | apex_process.sql aparece junto al DOCX
```

## 3. Requisitos locales

- Windows 10/11 o un sistema con Python compatible.
- Python 3.12 o superior.
- Microsoft Word para editar plantillas. Word no es necesario para ejecutar el compilador.
- Acceso al proyecto del compilador.
- Un editor de texto para SQL y JSON.

### 3.1 Forma rápida: acceso directo

En la carpeta raíz del proyecto haga doble clic en **Compilador de reportes APEX** (acceso directo) o en **Abrir compilador.bat**. La primera vez en un equipo nuevo, el lanzador crea la carpeta `.venv` e instala `python-docx` (requiere Python 3.12 o superior y conexión a Internet); las siguientes veces abre la interfaz directamente.

El acceso directo funciona aunque la carpeta se mueva, se renombre o se entregue comprimida a otra persona: no guarda ninguna ruta absoluta. Abre PowerShell de Windows, que localiza la carpeta del propio acceso directo y ejecuta **Abrir compilador.bat** que está a su lado. Si una política corporativa bloquea PowerShell, use directamente **Abrir compilador.bat**. El acceso directo debe permanecer en la carpeta raíz junto a ese archivo: no lo copie al Escritorio y extraiga el ZIP antes de abrirlo. Para regenerar el acceso directo:

```
# Vuelve a crear «Compilador de reportes APEX.lnk» en la carpeta raíz
powershell -ExecutionPolicy Bypass -File .\tools\crear_acceso_directo.ps1
```

### 3.2 Instalación manual

Instalación inicial en PowerShell, desde la raíz del proyecto:

```
py -3.13 -m venv .venv                 # 1. crea el entorno virtual
.\.venv\Scripts\Activate.ps1           # 2. lo activa en esta ventana
python -m pip install --upgrade pip    # 3. actualiza pip
python -m pip install -e .             # 4. instala el compilador y python-docx
```

Si PowerShell bloquea `Activate.ps1` por la directiva de ejecución, use `Set-ExecutionPolicy -Scope Process Bypass` en esa ventana o invoque directamente `.\.venv\Scripts\apex-report-compiler.exe`.

Para confirmar la instalación:

```
apex-report-compiler --version         # debe imprimir: apex-report-compiler 1.0.0
```

La interfaz gráfica también puede abrirse con:

```
apex-report-compiler-gui
```

### 3.3 Qué ofrece la interfaz gráfica

| Botón | Resultado |
|---|---|
| Nuevo proyecto desde DOCX… | Copia el Word a proyectos\<nombre>\ y crea **automáticamente** en generado\ el <nombre>.report.json y el <nombre>.sql, con un filtro opcional por columna. Cada filtro se enlaza a un Page Item PXX_<ALIAS> (o P42_<ALIAS> si indica el número de página) y cada {{FIELD:NOMBRE}} a PXX_<NOMBRE>. Si la carpeta ya existe, pide confirmación y guarda copias .bak (sección 2.2). |
| Validar | Revisa DOCX, SQL y proyecto sin escribir archivos. |
| Compilar | Genera **automáticamente** apex_process.sql en proyectos\<nombre>\ y template.json y validation.json en generado\, y muestra qué subir a APEX y dónde: el package, la tabla de Page Items con su tipo sugerido, el botón, el proceso y la ruta exacta de apex_process.sql. |
| Copiar código APEX | Copia apex_process.sql al portapapeles para pegarlo en el proceso. |
| Abrir carpeta de salida | Abre proyectos\<nombre>\, donde está apex_process.sql. |

Ejemplo de lo que genera automáticamente en `generado\entidades.sql` para `entidades.docx` con página 42 (extracto comentado; los comentarios `--` son válidos en SQL):

```
-- entidades.sql generado: reemplace tabla_origen y t."ALIAS" por los objetos reales
select t."VDNI" as "VDNI",                 -- una línea por {{COLUMN:...}} del Word
       t."VNOM" as "VNOM",
       ...
  from tabla_origen t
 where (:F_VDNI is null or t."VDNI" = :F_VDNI)   -- filtro opcional por columna
   and (:F_VNOM is null or t."VNOM" = :F_VNOM)   -- borre los que no necesite
   ...
```

El esqueleto es un punto de partida: reemplace `XX` por el número real de página, la tabla `tabla_origen` y las columnas por las reales, y elimine los filtros que no necesite (y su entrada en `bindings`).

## 4. Crear la plantilla Word

### 4.1 Estructura obligatoria

La plantilla usa papel A4, una sola sección y exactamente una tabla principal de 1 a 50 columnas. Debe contener:

1. Un párrafo previo a la tabla con `{{REPORT_TITLE}}` exactamente una vez.
2. Una tabla de dos filas.
3. Un pie de página real de Word, opcional pero recomendado.

El primer párrafo no se convierte en contenido exclusivo de la primera página: se compila como el encabezado de página nativo de `APEX_DATA_EXPORT` y puede repetirse en todas las páginas del PDF.

La tabla se construye así:

| Fila | Contenido | Ejemplo |
|---|---|---|
| 1 | Etiquetas literales que verá el usuario | DNI, Nombre completo, Departamento |
| 2 | Un marcador de columna por celda | {{COLUMN:VDNI}}, {{COLUMN:VNOM}}, {{COLUMN:DEPARTAMENTO}} |

La fila 2 es un prototipo técnico: no aparece como una fila del PDF. Su orden establece el orden contractual de las columnas.

### 4.2 Marcadores admitidos

| Marcador | Dónde se escribe | Uso |
|---|---|---|
| {{REPORT_TITLE}} | Párrafo previo a la tabla (obligatorio, una vez) | Título recibido por el proceso APEX (title del proyecto) |
| {{FIELD:NOMBRE}} | Párrafo de título o pie de página | Campo escalar declarado en fields |
| {{COLUMN:ALIAS}} | Fila 2 de la tabla, uno por celda | Columna tabular; ALIAS debe coincidir con el alias SQL |
| {{APP_USER}} | Párrafo de título o pie de página | Usuario actual de APEX |
| {{GENERATED_AT}} | Párrafo de título o pie de página | Fecha y hora de generación |

Los identificadores se normalizan a mayúsculas y deben comenzar con una letra. No coloque SQL, fórmulas, condiciones, bucles ni código en un marcador.

Las plantillas de referencia de `templates/` y `proyectos/demo/entidades/entidades.docx` usan los cinco tipos a la vez. Así se ven en Word (los comentarios a la derecha no forman parte del documento):

```
{{REPORT_TITLE}}                  <- título; un salto de línea manual (Mayús+Intro)...
Unidad: {{FIELD:UNIDAD}}          <- ...mantiene un único párrafo con un campo escalar

| DNI             | Nombre completo  | Departamento            |   <- fila 1: texto literal
| {{COLUMN:VDNI}} | {{COLUMN:VNOM}}  | {{COLUMN:DEPARTAMENTO}} |   <- fila 2: alias SQL

Pie de página de Word:
Usuario: {{APP_USER}} | Fecha: {{GENERATED_AT}}                   <- integrados, sin fields
```

Y así se resuelven al descargar el PDF:

```
Relación de entidades                       <- valor de "title" del .report.json
Unidad: Oficina de Informática              <- fields: UNIDAD, origen CONSTANT

| DNI      | Nombre completo   | Departamento |
| 40123456 | ANA PÉREZ RUIZ    | LIMA         |   <- una fila por registro del SELECT
| ...      | ...               | ...          |

Usuario: JPEREZ | Fecha: 01/10/2026 10:15    <- usuario APEX y fecha de generación
```

Escriba `{{REPORT_TITLE}}`, `{{FIELD:...}}`, `{{APP_USER}}` y `{{GENERATED_AT}}` en el **mismo párrafo** de título (separando líneas con Mayús+Intro) o en el pie. Un segundo párrafo antes de la tabla se rechaza.

### 4.3 Formato admitido

- Fuentes: Arial, Aptos, Calibri o Helvetica (se imprimen como Helvetica); Times New Roman, Cambria o Georgia (Times); Courier New o Consolas (Courier).
- Tamaño entre 6 y 24 puntos.
- Peso normal o negrita.
- Color de texto y fondo uniforme.
- Alineación izquierda, centrada o derecha.
- Bordes simples y uniformes.
- Orientación vertical u horizontal; también puede decidirse en APEX con `AUTO`.

El título se normaliza centrado y se repite como encabezado de página. El motor usa A4. El resultado no es una copia exacta de Word: Word solo expresa la intención visual y el PDF lo dibuja `APEX_DATA_EXPORT` con estilos uniformes por zona.

Todos estos estilos se extraen **automáticamente** del DOCX en cada compilación y quedan en la sección `styles` de `template.json`; no hay que copiarlos a mano a ningún JSON. Si se prefiere fijarlos sin abrir Word, se escriben en `style_overrides` del `.report.json`.

Los anchos de columna que dibuje en Word **no** se trasladan al reporte: son una ayuda visual. Defina los anchos en `column_widths` del proyecto (sección 6.4); una columna sin entrada recibe `WEIGHT` 1.

Las etiquetas de la fila 1 admiten hasta 255 bytes (unas 127 letras acentuadas) y el encabezado y el pie hasta 4000 bytes.

### 4.4 Elementos no admitidos

No use imágenes, macros, `.docm`, cuadros de texto, WordArt, formas, gráficos, tablas anidadas, celdas combinadas, varias tablas, varias secciones, campos Word, hipervínculos externos, controles ActiveX, comentarios, notas al pie, control de cambios (también los cambios de formato), texto oculto, etiquetas inteligentes ni controles de contenido. El compilador los rechaza para evitar resultados ambiguos y reducir riesgos.

Si Word inserta un espacio de no separación dentro de un marcador (`{{ APP_USER}}` escrito con Ctrl+Mayús+Espacio), el marcador se rechaza: borre ese espacio y escriba el marcador sin espacios.

## 5. Escribir la consulta SQL

Cada reporte usa un archivo SQL con una sola consulta `SELECT` o `WITH ... SELECT`. Los aliases seleccionados deben coincidir con los marcadores `COLUMN` de Word. «Nuevo proyecto desde DOCX» crea un esqueleto automáticamente; luego se edita a mano.

```
-- Reporte ENTIDADES: un alias por cada {{COLUMN:...}} de la plantilla
select e.vdni          as vdni,           -- {{COLUMN:VDNI}}
       e.vnom          as vnom,           -- {{COLUMN:VNOM}}
       d.descripcion   as departamento    -- {{COLUMN:DEPARTAMENTO}}: texto, no el ID
  from entidad e
  left join mae_departamento d
    on d.id = e.departamento_id
 where (:DNI is null or e.vdni = :DNI)    -- filtro opcional: vacío = sin filtro
   and (:DEPARTAMENTO_ID is null
        or e.departamento_id = :DEPARTAMENTO_ID)
 order by e.vnom
```

Los valores variables siempre usan binds (`:DNI`, `:DEPARTAMENTO_ID`). Nunca concatene el contenido de un Page Item dentro del texto SQL.

Reglas adicionales que el compilador comprueba:

- puede empezar con comentarios (`/* ... */` o `-- ...`) seguidos de saltos de línea;
- no use la sintaxis de sustitución `&P42_ITEM.`: es error fuera de literales. Dentro de un literal o comentario solo genera una advertencia, porque el código generado la conserva como texto y APEX no la sustituye. Use `:BIND` y su mapping;
- no use `SELECT ... INTO`, `FOR UPDATE`, punto y coma ni la barra `/` de SQL*Plus;
- no use como bind un nombre reservado de APEX (`:APP_USER`, `:APP_ID`, `:REQUEST`…). Para imprimir el usuario use `{{APP_USER}}`.

### 5.1 LOV: imprimir el display value

El paquete no intenta leer la configuración interna de un LOV de la página. La consulta debe devolver explícitamente el texto que se imprimirá:

```
select e.departamento_id,                   -- return value: sirve para filtrar
       d.descripcion as departamento        -- display value: es lo que se imprime
  from entidad e
  join mae_departamento d
    on d.id = e.departamento_id
```

La plantilla usa `{{COLUMN:DEPARTAMENTO}}`; el ID puede seguir usándose como bind o para lógica interna, pero no tiene que formar parte de la plantilla.

## 6. El archivo de proyecto

El archivo `nombre.report.json` relaciona el DOCX y el SQL. Se crea **automáticamente** con «Nuevo proyecto desde DOCX» a partir de los marcadores de la plantilla y luego puede editarse directamente con cualquier editor de texto. Vive en `proyectos\<nombre>\generado\`; sus rutas son relativas a ese archivo y deben permanecer dentro de `proyectos\<nombre>\`.

Versión comentada del ejemplo `proyectos/demo/entidades/generado/entidades.report.json`. JSON **no admite comentarios**: las líneas `//` son solo explicativas y deben eliminarse si copia este texto; el archivo real, sin comentarios, está en esa carpeta.

```
{
  "schema": "corporate-report-project/1.0",   // fijo
  "report_id": "ENTIDADES",                   // identificador; nombra el botón DOWNLOAD_ENTIDADES
  "template": "../entidades.docx",            // plantilla Word, relativa a este archivo
  "query_file": "entidades.sql",              // consulta, relativa a este archivo
  "title": "Relación de entidades",           // reemplaza {{REPORT_TITLE}}
  "file_name": "reporte_entidades",           // nombre del archivo descargado (sin extensión)
  "orientation": "LANDSCAPE",                 // AUTO, PORTRAIT o LANDSCAPE; si falta, la del Word
  "format_item": "P0_REPORT_FORMAT",          // opcional: el usuario elige PDF o XLSX
  "orientation_item": "P0_REPORT_ORIENTATION",// opcional: el usuario elige la orientación
  "bindings": [                               // un mapping por cada :BIND del SQL
    {"bind":"DNI", "item":"P42_DNI", "type":"VARCHAR2", "required":false},
    {"bind":"DEPARTAMENTO_ID", "item":"P42_DEPARTAMENTO_ID", "type":"NUMBER", "required":false}
  ],
  "fields": [                                 // un objeto por cada {{FIELD:...}} del Word
    {"name":"UNIDAD", "source":"CONSTANT", "value":"Oficina de Informática"}
  ],
  "excluded_columns": [],                     // aliases que no deben imprimirse
  "column_widths": [                          // anchos por alias; el resto recibe WEIGHT 1
    {"column":"VDNI", "mode":"FIXED_PERCENT", "value":14},
    {"column":"VNOM", "mode":"WEIGHT", "value":2},
    {"column":"DEPARTAMENTO", "mode":"AUTO"}
  ]
}
```

`format_item` y `orientation_item` son opcionales. Si los declara, el usuario podrá elegir formato u orientación en la página (sección 9); si los omite, el reporte usa PDF y la orientación compilada. Sin `orientation`, se usa la orientación de la página Word; declare `"AUTO"` para que APEX elija vertical hasta tres columnas y horizontal con cuatro o más.

`style_overrides` permite ajustar estilos sin editar el Word; sus zonas y claves están en el contrato técnico (sección 7.1). Por ejemplo, para que el cuerpo de la tabla use 9 pt y el encabezado de tabla un fondo azul, sin tocar el DOCX:

```
"style_overrides": {
  "table_body":   {"font_size": 9},                 // prevalece sobre el tamaño del Word
  "table_header": {"background_color": "#283848"}   // color en formato #RRGGBB
}
```

Límites que el compilador valida antes de generar PL/SQL:

- `title`: máximo 255 bytes UTF-8;
- `file_name`: máximo 180 caracteres y solo letras, números, `_` o `-`;
- constantes y máscaras de formato: máximo 4000 bytes UTF-8;
- nombres de Page Items: máximo 128 caracteres en total.

### 6.1 Bindings

Cada bind detectado en SQL debe tener exactamente un mapping en `bindings`.

| Propiedad | Significado |
|---|---|
| bind | Nombre sin dos puntos, por ejemplo FECHA_DESDE |
| item | Page Item, por ejemplo P42_FECHA_DESDE |
| type | VARCHAR2, NUMBER, DATE o TIMESTAMP |
| required | Si true, un valor vacío detiene el reporte |
| format_mask | Obligatorio para DATE y TIMESTAMP, por ejemplo DD/MM/YYYY |

Los nombres de Page Items deben usar el patrón `P<número>_NOMBRE`, con letras, números y guion bajo después del prefijo. No use `$` ni `#` en Page Items del proyecto.

No existe límite de filas: el reporte exporta todas las filas que devuelve la consulta. El volumen se controla con los filtros, que Informática revisa antes de publicar el reporte. Antes de publicar, pruebe el reporte con los filtros vacíos para conocer el volumen máximo real.

Ejemplo de fecha (comentado):

```
{
  "bind": "FECHA_DESDE",          // en el SQL: e.fecha_registro >= :FECHA_DESDE
  "item": "P42_FECHA_DESDE",      // Date Picker de la página 42
  "type": "DATE",                 // se convierte con TO_DATE antes de ejecutar
  "format_mask": "DD/MM/YYYY",    // la misma máscara que el Date Picker
  "required": false               // vacío = sin filtro; true = error si está vacío
}
```

### 6.2 Campos escalares

Un `{{FIELD:NOMBRE}}` requiere una entrada en `fields`. «Nuevo proyecto desde DOCX» la crea automáticamente con origen `ITEM`; se cambia a mano si el valor debe venir de otro origen. Los orígenes son `ITEM`, `CONTEXT`, `SYSTEM` y `CONSTANT`:

```
[
  // {{FIELD:SOLICITANTE}} <- valor del Page Item P42_SOLICITANTE
  {"name":"SOLICITANTE", "source":"ITEM", "item":"P42_SOLICITANTE", "type":"VARCHAR2"},
  // {{FIELD:UNIDAD}} <- texto fijo
  {"name":"UNIDAD", "source":"CONSTANT", "value":"Oficina de Informática"},
  // {{FIELD:APLICACION}} <- contexto APEX: APP_USER, APP_ID o APP_PAGE_ID
  {"name":"APLICACION", "source":"CONTEXT", "key":"APP_ID"},
  // {{FIELD:EMITIDO}} <- fecha de generación con un formato propio
  {"name":"EMITIDO", "source":"SYSTEM", "key":"GENERATED_AT", "format":"DD/MM/YYYY"}
]
```

Cada nombre declarado debe aparecer en la plantilla y viceversa: un `{{FIELD:...}}` sin entrada (`TOKEN-020`) o una entrada que el Word no usa (`TOKEN-021`) detienen la validación. No declare `APP_USER` ni `GENERATED_AT`: son campos incorporados.

### 6.3 Excluir columnas

Para impedir que una columna técnica se imprima, use aliases, no posiciones:

```
"excluded_columns": ["PROCESSES", "ID_INTERNO"]   // siguen en el SELECT, no en el PDF
```

No se pueden excluir todas las columnas. Una exclusión no sustituye el control de acceso: el SQL no debe seleccionar datos sensibles innecesarios.

### 6.4 Definir anchos

| Modo | Comportamiento |
|---|---|
| FIXED_PERCENT | Solicita un porcentaje del ancho imprimible |
| WEIGHT | Reparte el espacio restante proporcionalmente |
| AUTO | Omite el ancho y delega el ajuste al motor de APEX |

Solo puede existir una columna `AUTO`, en cualquier posición. Una columna sin configuración se comporta como `WEIGHT` de valor 1. `WEIGHT` admite valores mayores que 0 y hasta 1000. `AUTO` es una sugerencia al renderizador, no una garantía matemática.

```
"column_widths": [
  {"column":"VDNI", "mode":"FIXED_PERCENT", "value":12},  // 12 % del ancho útil
  {"column":"VNOM", "mode":"WEIGHT", "value":2},          // doble que una columna WEIGHT 1
  {"column":"VDIREC_ACTUAL", "mode":"AUTO"}               // APEX ajusta esta columna
]                                                         // las demás: WEIGHT 1
```

## 7. Validar y compilar

```flujo
# Figura 3. Validación y compilación
Compilador | Leer DOCX, SQL y .report.json | el DOCX se inspecciona antes de abrirlo
Compilador | Validar plantilla, SQL y proyecto | marcadores, aliases, binds, campos, estilos y límites
? Compilador | ¿Validación sin errores? | No: se informan con su código, no se escribe nada y la última salida válida queda intacta
Compilador | Generar en una carpeta temporal | template.json, validation.json y apex_process.sql
Compilador | Publicar | apex_process.sql en proyectos\<nombre>\; template.json y validation.json en generado\; solo esos tres archivos se reemplazan
```

La interfaz gráfica hace todo esto con los botones. Desde la línea de comandos:

```
# Cargar el Word en proyectos\ventas\ (desde cualquier carpeta)
apex-report-compiler new --docx C:\Descargas\ventas.docx --page 42

# Volver a cargarlo con el mismo nombre: --replace guarda copias .bak en generado\
apex-report-compiler new --docx C:\Descargas\ventas.docx --page 42 --replace
```

Valide, sin escribir artefactos:

```
# Solo revisa; útil después de cada cambio en Word, SQL o proyecto
apex-report-compiler validate `
  --project .\proyectos\ventas\generado\ventas.report.json
```

Compile cuando la validación sea correcta:

```
# Sin --output: apex_process.sql va a proyectos\ventas\ y el resto a generado\
apex-report-compiler compile `
  --project .\proyectos\ventas\generado\ventas.report.json
```

Otras opciones:

```
# Acepta aproximaciones documentadas (p. ej. una fuente no reconocida) como advertencias
apex-report-compiler validate --compatible --project .\proyectos\ventas\generado\ventas.report.json

# Diagnósticos en JSON, para scripts o integración continua
apex-report-compiler validate --json --project .\proyectos\ventas\generado\ventas.report.json

# Enviar los tres archivos a otra carpeta (por ejemplo, para comparar versiones)
apex-report-compiler compile --project .\proyectos\ventas\generado\ventas.report.json --output .\build\ventas
```

El modo predeterminado es estricto. `--compatible` admite únicamente aproximaciones documentadas y las informa como advertencias.

La compilación genera automáticamente:

| Archivo | Uso |
|---|---|
| generado\template.json | Interpretación canónica y auditable de la plantilla |
| generado\validation.json | Diagnósticos reproducibles |
| apex_process.sql | Proceso PL/SQL listo para revisar y copiar en APEX (junto al DOCX) |

Estos tres archivos se regeneran en cada compilación a partir del DOCX, el SQL y el proyecto; los demás archivos de la carpeta no se tocan. Si una compilación falla, se conserva intacta la última salida válida. Con `--output`, la carpeta indicada no puede sobrescribir el `.sql` ni el `.report.json` (`IO-003`).

## 8. Instalar el paquete común

La instalación se realiza una sola vez por esquema de aplicación. Desde SQLcl o SQL*Plus:

```
-- Conectado como el parsing schema, desde la raíz del proyecto
@sql/modo_simple/install.sql
```

Desde SQL Workshop cargue y ejecute `sql/modo_simple/pkg_corporate_reports.sql`, y compruebe:

```
-- No debe devolver filas
select type, line, position, text
  from user_errors
 where name = 'PKG_CORPORATE_REPORTS'
 order by sequence;
```

Tanto el package como su body deben figurar como `VALID`. No conceda `EXECUTE` a `PUBLIC` y mantenga `AUTHID CURRENT_USER`.

## 9. Integrar en una página APEX

Solo dos cosas llegan a APEX: `sql/modo_simple/pkg_corporate_reports.sql` (una vez por esquema, sección 8) y el contenido de `apex_process.sql` (una vez por reporte). El DOCX, el `.report.json`, el `.sql` fuente, `template.json` y `validation.json` se quedan en el repositorio. La interfaz gráfica muestra esta misma receta con los nombres reales del reporte al compilar.

```flujo
# Figura 4. Qué ocurre cuando el usuario pulsa Descargar
Usuario final | Completa los filtros y pulsa Descargar | botón con Action: Submit Page
APEX | Envía la página completa | los Page Items pasan al estado de sesión
? APEX | ¿Se pulsó el botón de descarga? | No: el proceso no se ejecuta (Server-side Condition)
APEX | Ejecuta el proceso de Processing | el bloque pegado desde apex_process.sql
PKG_CORPORATE_REPORTS | DOWNLOAD_QUERY lee y convierte los Page Items | binds tipados con APEX_EXEC.ADD_PARAMETER
PKG_CORPORATE_REPORTS | Ejecuta el SQL y aplica columnas, estilos y anchos | sustituye {{REPORT_TITLE}}, {{FIELD:…}}, {{APP_USER}} y {{GENERATED_AT}}
APEX_DATA_EXPORT | Genera y envía el PDF o el XLSX | la página no se recarga; no se ejecutan branches
```

### 9.1 Page Items

| Item | Dónde | Tipo recomendado | Valores |
|---|---|---|---|
| Cada bindings[].item (por ejemplo P42_DNI) | página del reporte | Text Field, Number Field o Select List (con LOV: el return value) | según el filtro |
| Item DATE/TIMESTAMP | página del reporte | Date Picker | misma máscara que format_mask |
| Cada fields[].item de origen ITEM | página del reporte | el que corresponda | texto para encabezado o pie |
| format_item (por ejemplo P0_REPORT_FORMAT) | página 0 (Global Page) | Select List o Radio Group | retorno exacto PDF o XLSX; por defecto PDF. El XLSX lleva solo títulos de columna y datos |
| orientation_item (por ejemplo P0_REPORT_ORIENTATION) | página 0 (Global Page) | Select List o Radio Group | retorno exacto AUTO, PORTRAIT o LANDSCAPE; si queda vacío se usa la orientación del proyecto o del DOCX |

Si la aplicación no tiene Global Page, créela (página 0) o quite `format_item` y `orientation_item` del proyecto y recompile. Active **Session State Protection** en los items de filtro cuando corresponda.

Ejemplo de LOV estática para `P0_REPORT_FORMAT` (Select List):

```
STATIC:PDF;PDF,Excel;XLSX
   ^ pares «display;return» separados por comas; el return debe ser exactamente PDF o XLSX
```

### 9.2 Botón y proceso

1. Cree un botón **Descargar** con *Action: Submit Page*. La descarga necesita un envío completo de la página; no use una Dynamic Action Ajax ni *Execute Server-side Code* para descargar.
2. En los atributos de la página, *Advanced > Reload on Submit*, elija **Always**. Con el valor predeterminado *Only for Success* el navegador espera una respuesta JSON y la descarga falla con un error de sintaxis.
3. Cree un proceso en *Processing*: *Type: Execute Code*, *Language: PL/SQL*.
4. En *PL/SQL Code* pegue el contenido completo de `apex_process.sql` (el bloque `declare … end;`), sin editarlo.
5. En *Server-side Condition* elija *When Button Pressed* = el botón de descarga.
6. Asigne al botón y al proceso el mismo *Authorization Scheme* que protege la página.
7. Al enviar la página, los filtros pasan al estado de sesión antes del proceso; la descarga detiene el procesamiento de APEX, por lo que no se ejecutan branches posteriores.
8. Pruebe con un usuario real autorizado y con otro no autorizado.

Alternativa: un proceso *Before Header* en una página que reciba una request concreta, siempre que los filtros ya estén en el estado de sesión.

El código generado automáticamente llama a esta API (extracto comentado; el valor real de cada argumento es el que genera el compilador, por ejemplo `coalesce(:P0_REPORT_FORMAT, 'PDF')`):

```
pkg_corporate_reports.download_query(
    p_sql_query             => l_sql,                -- contenido del .sql
    p_columns_json          => l_columns_json,       -- {{COLUMN:...}} y etiquetas de la fila 1
    p_bindings_json         => l_bindings_json,      -- "bindings" del proyecto
    p_fields_json           => l_fields_json,        -- "fields" del proyecto
    p_style_json            => l_style_json,         -- estilos del Word + style_overrides
    p_header_template       => '{{REPORT_TITLE}}
Unidad: {{FIELD:UNIDAD}}',                          -- párrafo de título del Word
    p_footer_template       => 'Usuario: {{APP_USER}} | Fecha: {{GENERATED_AT}}',
    p_title                 => 'Relación de entidades',
    p_file_name             => 'reporte_entidades',
    p_format                => coalesce(:P0_REPORT_FORMAT, 'PDF'),
    p_orientation           => coalesce(:P0_REPORT_ORIENTATION, 'LANDSCAPE'), -- Page Item o valor compilado
    p_excluded_columns_json => l_excluded_columns,
    p_column_widths_json    => l_column_widths
);
```

El usuario se obtiene del contexto de APEX y el papel es siempre A4.

## 10. Pruebas obligatorias en APEX

Antes de producción pruebe:

- compilación del package sin errores;
- PDF y XLSX (el XLSX debe traer solo títulos de columna y datos, sin encabezado ni pie);
- consulta con cero, una y muchas filas;
- volumen grande con los filtros vacíos (todas las filas se exportan);
- filtros vacíos y requeridos;
- `VARCHAR2`, `NUMBER`, `DATE` y `TIMESTAMP`;
- orientación `AUTO`, `PORTRAIT` y `LANDSCAPE`;
- columnas fijas, ponderadas, automática y excluidas;
- LOV mostrando descripción, no ID;
- título, campos `{{FIELD:...}}`, usuario y fecha en encabezado y pie del PDF;
- permisos de un usuario autorizado y otro no autorizado.

## 11. Diagnóstico de problemas

```flujo
# Figura 5. La descarga no ocurre
? Página | ¿El botón usa Action: Submit Page? | No: cambiarlo; una Dynamic Action Ajax no puede descargar
? Página | ¿Reload on Submit está en Always? | No: cambiarlo; con Only for Success aparece «SyntaxError … JSON»
? Proceso | ¿Server-side Condition = When Button Pressed del botón correcto? | No: corregir la condición
? Proceso | ¿El usuario tiene la autorización del proceso? | No: revisar el Authorization Scheme
Diagnóstico | Revisar APEX Debug | el mensaje del package indica el bind, columna o límite afectado
```

| Síntoma | Causa probable | Acción |
|---|---|---|
| Falta un bind | El SQL contiene :NOMBRE sin mapping | Añadirlo a bindings |
| Page Item inválido | Nombre mal escrito o fuera del patrón | Usar P0_... o P<número>_... |
| El PDF imprime un ID de LOV | El SQL selecciona el return value | Hacer JOIN y devolver el display value |
| Columna inexistente | Alias Word y SQL no coinciden | Alinear COLUMN:ALIAS y SELECT ... AS alias |
| TOKEN-020 | {{FIELD:NOMBRE}} en el Word sin entrada en fields | Añadir la entrada en fields |
| TOKEN-021 | Entrada en fields que el Word ya no usa | Quitarla de fields |
| Fecha inválida | Valor de sesión no coincide con la máscara | Corregir format_mask o el formato del item |
| No se descarga | Flujo de página o estado de sesión incorrecto | Seguir la figura 5 |
| Tabla demasiado comprimida | Demasiadas columnas o pesos inadecuados | Excluir columnas, usar landscape o ajustar anchos |
| DOCX rechazado | Contiene estructura fuera del contrato | Simplificar usando una plantilla de ejemplo |
| SQL-006 con un comentario inicial | Hay un espacio no ASCII antes del SELECT | Reescribir el salto de línea o el espacio |
| SQL-013 | El SQL usa &ITEM. | Reemplazar por :BIND y su mapping |
| PROJECT-017 | Bind con nombre reservado de APEX | Renombrar el bind o usar {{APP_USER}} |
| IO-003 | La salida sobrescribiría el SQL o el proyecto | Elegir otra carpeta de salida |
| GUIDE-002 | Se volvió a cargar un DOCX con el mismo nombre sin confirmar | Confirmar en la ventana o usar --replace; se guardan copias .bak |
| PROJECT-005 | Una ruta del .report.json sale de proyectos\<nombre>\ | Mantener la plantilla en ../<nombre>.docx y el SQL en generado\ |
| «SyntaxError … JSON» al descargar | Reload on Submit en Only for Success | Cambiarlo a Always en la página |
| TOKEN-005 | {{FIELD:APP_USER}} u otro nombre reservado | Escribir {{APP_USER}} sin FIELD: |

## 12. Mantenimiento recomendado

- Mantenga cada reporte en su carpeta `proyectos\<nombre>\`; guarde allí cualquier material adicional, que nunca se borra.
- Antes de volver a cargar un DOCX con el mismo nombre, recuerde que el `.sql` y el `.report.json` se regeneran: recupere sus cambios desde los `.bak`.
- Versione cada cambio y revise SQL/configuración como código.
- Recompile después de cambiar Word, SQL, aliases, campos, exclusiones o anchos: los artefactos se regeneran automáticamente.
- Si edita `apex_process.sql` directamente, refleje el mismo cambio en el proyecto fuente; de lo contrario la siguiente compilación lo deshará.
- Instale una sola copia central del paquete; no genere un package distinto por reporte.
- Revise rendimiento y permisos del SQL antes de promover.

La carpeta `proyectos\demo` contiene dos proyectos completos, uno por modo, con su documentación en `proyectos\demo\README.docx`, que sirven como punto de partida; `templates` contiene dos plantillas de referencia, vertical y horizontal, que usan los cinco tipos de marcador.

## 13. Modo layout: varias tablas y maquetación libre

El modo descrito en las secciones 4 a 12 usa `APEX_DATA_EXPORT` y admite una sola tabla. Para reportes como los de Oracle Reports con varias tablas, cuadrículas, recuadros, totales y número de página existe el **modo layout**: el compilador traduce el Word completo y lo imprime un motor PDF escrito en PL/SQL (`RPT_PDF` + `RPT_LAYOUT`), sin Java, sin servidores de impresión y sin licencias. El compilador elige el modo según el proyecto: `corporate-layout-project/1.0` activa el modo layout.

```flujo
# Figura 6. Modo layout
Informática | Diseñar el Word con varias tablas | maquetación con bordes visibles, blancos o sin borde
Compilador | Nuevo proyecto desde DOCX | crea generado\q_<consulta>.sql por cada consulta y el .report.json
Informática | Pegar en cada q_<consulta>.sql la Query del reporte original | mismos binds :P_NOMBRE; edición mínima
Compilador | Compilar | genera rpt_<reporte>.sql y apex_process.sql junto al DOCX
APEX | Instalar el motor (una vez) y el package del reporte | sql\modo_layout\install.sql
Usuario final | Descargar | el package ejecuta las consultas y dibuja el PDF
```

### 13.1 Qué admite el Word

| Elemento | Cómo se escribe en Word | Resultado |
|---|---|---|
| Tabla de datos | Una fila con {{COLUMN:CONSULTA.COLUMNA}}; las filas anteriores son cabecera y las posteriores, totales | Una fila por registro; la cabecera se repite en cada página |
| Totales | {{SUM:CONSULTA.COLUMNA}} en las filas posteriores | Suma de la columna en la tabla |
| Campo de una consulta | {{FIELD:CONSULTA.COLUMNA}} en cualquier párrafo o celda | Primera fila de la consulta (grupo maestro) |
| Parámetro o constante | {{FIELD:NOMBRE}} | Valor del parámetro o de constants |
| Número de página | {{PAGE}} y {{PAGES}}, solo en el encabezado o pie de Word | 1 DE 3 |
| Formato | Barra vertical y máscara al final del marcador (ver el ejemplo comentado) | TO_CHAR con esa máscara, por ejemplo FM999G999G990D00 |
| Maquetación | Tablas sin {{COLUMN:...}}, con bordes por celda: visibles, blancos o sin borde; rellenos; celdas combinadas en horizontal | Se dibujan tal cual |
| Encabezado y pie | Párrafos y tablas en el encabezado/pie de Word | Se repiten en cada página |

Se mantienen las restricciones de seguridad del modo simple (sin imágenes, macros, campos de Word ni contenido externo). No se admiten celdas combinadas en vertical ni tablas anidadas, y el texto de una celda no se parte en varias líneas automáticamente: use saltos de línea manuales. Las fuentes se imprimen como Helvetica.

Ejemplo comentado (extracto de `proyectos\demo\estado_cuenta\estado_cuenta.docx`, un ejemplo ficticio):

```
Encabezado de Word:  {{FIELD:SISTEMA}}         ...        {{GENERATED_AT|DD/MM/YYYY}}
                     {{REPORT_TITLE}}                      <- se repite en cada página

Cuerpo:
| Alumno : {{FIELD:ALUMNO.COD_ALUMNO}}  {{FIELD:ALUMNO.NOMBRE}}            |  <- celda combinada, fondo gris
| Código :      | {{FIELD:ALUMNO.COD_ALUMNO}}     | Costo Programa : | ...  |  <- borde blanco entre etiqueta y valor

| CUOTAS POR PAGAR (banda gris)        |                                       <- cabecera, fila 1
| Año | Mes | N° Cuota | Pensión | TOTAL |                                    <- cabecera, fila 2
| {{COLUMN:CUOTAS.ANIO}} | ... | {{COLUMN:CUOTAS.TOTAL|FM999G999G990D00}} |    <- se repite por registro
|     |     |          | {{SUM:CUOTAS.PENSION|...}} | {{SUM:CUOTAS.TOTAL|...}} |  <- totales

Pie de Word:  Generado por {{APP_USER}}   [{{PAGE}}] DE [{{PAGES}}]
```

### 13.2 El proyecto: una consulta por archivo, como el Data Model

```
{
  "schema": "corporate-layout-project/1.0",
  "report_id": "ESTADO_CUENTA",                // nombra el package: RPT_ESTADO_CUENTA
  "template": "../estado_cuenta.docx",
  "title": "ESTADO DE CUENTA DEL ALUMNO",
  "parameters": [                              // User Parameters del reporte
    {"name": "P_COD_ALUMNO", "item": "P71_COD_ALUMNO", "type": "VARCHAR2", "required": true}
  ],
  "queries": {                                 // Queries del Data Model, una por archivo
    "ALUMNO":   "q_alumno.sql",
    "CUOTAS":   "q_cuotas.sql"
  },
  "constants": {"SISTEMA": "ACADEMIA DEMO"}    // textos fijos usados como {{FIELD:SISTEMA}}
}
```

| Oracle Reports | Modo layout |
|---|---|
| User Parameter | parameters (Page Item y tipo); en SQL sigue siendo :P_NOMBRE |
| System Parameter | {{APP_USER}}, {{GENERATED_AT}}, {{PAGE}}, {{PAGES}} |
| Query | Un archivo generado\q_<nombre>.sql con el mismo SELECT |
| Group maestro | {{FIELD:CONSULTA.COLUMNA}} (primera fila) |
| Group repetitivo | Tabla de datos con {{COLUMN:CONSULTA.COLUMNA}} |
| Formula Column | Expresión en el SELECT (o función PL/SQL llamada desde el SELECT) |
| Summary Column | {{SUM:...}} en la tabla, o SUM en una consulta de resumen |
| Data Link | El mismo parámetro :P_NOMBRE en la consulta hija |

### 13.3 Instalación y publicación

1. Una vez por esquema: `sql\modo_layout\install.sql` (o los cuatro archivos `rpt_pdf.*` y `rpt_layout.*` en SQL Workshop).
2. Por reporte: ejecute `proyectos\<nombre>\rpt_<reporte>.sql` en SQL Workshop cada vez que recompile.
3. Cree los Page Items de `parameters`, el botón (Submit Page, Reload on Submit: Always) y el proceso con `apex_process.sql`, igual que en la sección 9.

El ejemplo ficticio `proyectos\demo\estado_cuenta` (estado de cuenta de un alumno de una academia de demostración) usa las tablas de prueba `PRUEBAP_ALUMNO*`; `proyectos\demo\estado_cuenta\datos_prueba.sql` las crea solo si no existen y nunca borra datos. `proyectos\demo\estado_cuenta\ejemplo_estado_cuenta.pdf` muestra el resultado.
