# Manual de uso

## Compilador local Word → reportes Oracle APEX 24.2

**Versión:** 1.0  
**Destinatarios:** personal de Informática que diseña reportes y desarrolladores Oracle APEX  
**Salida:** PDF o XLSX mediante `APEX_DATA_EXPORT`

## 1. Qué resuelve esta herramienta

La herramienta permite diseñar en Microsoft Word un reporte tabular corporativo, validar la plantilla en una computadora local y generar el bloque PL/SQL mínimo que se copiará en APEX. Word no genera el PDF y el DOCX no se carga al servidor: la apariencia compatible se traduce a configuración declarativa y el paquete común `PKG_CORPORATE_REPORTS` genera la descarga con APIs nativas de APEX.

En la VM de APEX/ORDS no se instalan Word, LibreOffice, Python, Docker ni un motor de impresión. Python y Word se usan únicamente en el equipo del diseñador.

El flujo está pensado para reportes de una consulta SQL, un título, una tabla y un pie. No es un conversor general de Word ni reproduce diseños libres de Oracle Reports.

## 2. Conceptos esenciales

| Elemento | Dónde vive | Responsabilidad |
|---|---|---|
| Plantilla `.docx` | Equipo local y repositorio de Informática | Define título, columnas y estilos admitidos |
| Consulta `.sql` | Proyecto del reporte | Devuelve los datos y los textos que se imprimirán |
| Proyecto `.report.json` | Proyecto del reporte | Une plantilla, SQL, Page Items, exclusiones y anchos |
| Compilador | Equipo local | Valida y genera `template.json`, `validation.json` y `apex_process.sql` |
| `PKG_CORPORATE_REPORTS` | Base de datos de APEX | Ejecuta el SQL con binds y descarga PDF/XLSX |
| Page Items | Página APEX | Proporcionan filtros mediante estado de sesión |

## 3. Requisitos locales

- Windows 10/11 o un sistema con Python compatible.
- Python 3.12 o superior.
- Microsoft Word para editar plantillas. Word no es necesario para ejecutar el compilador.
- Acceso al proyecto del compilador.
- Un editor de texto para SQL y JSON.

Instalación inicial en PowerShell, desde la raíz del proyecto:

```powershell
py -3.13 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e .
```

Para confirmar la instalación:

```powershell
apex-report-compiler --version
```

La interfaz gráfica es opcional:

```powershell
apex-report-compiler-gui
```

## 4. Crear la plantilla Word

### 4.1 Estructura obligatoria

La plantilla usa papel A4, una sola sección y exactamente una tabla principal de 1 a 50 columnas. Debe contener:

1. Un párrafo previo a la tabla con `{{REPORT_TITLE}}` exactamente una vez.
2. Una tabla de dos filas.
3. Un pie de página real de Word, opcional pero recomendado.

El primer párrafo no se convierte en contenido exclusivo de la primera página:
se compila como el encabezado de página nativo de `APEX_DATA_EXPORT` y puede
repetirse en todas las páginas del PDF.

La tabla se construye así:

| Fila | Contenido | Ejemplo |
|---|---|---|
| 1 | Etiquetas literales que verá el usuario | `DNI`, `Nombre completo`, `Departamento` |
| 2 | Un marcador de columna por celda | `{{COLUMN:VDNI}}`, `{{COLUMN:VNOM}}`, `{{COLUMN:DEPARTAMENTO}}` |

La fila 2 es un prototipo técnico: no aparece como una fila del PDF. Su orden establece el orden contractual de las columnas.

### 4.2 Marcadores admitidos

| Marcador | Uso |
|---|---|
| `{{REPORT_TITLE}}` | Título recibido por el proceso APEX |
| `{{COLUMN:ALIAS}}` | Columna tabular; `ALIAS` debe coincidir con el alias SQL |
| `{{FIELD:NOMBRE}}` | Campo escalar declarado en `fields` |
| `{{APP_USER}}` | Usuario actual de APEX |
| `{{GENERATED_AT}}` | Fecha y hora de generación |

Los identificadores se normalizan a mayúsculas y deben comenzar con una letra. No coloque SQL, fórmulas, condiciones, bucles ni código en un marcador.

Ejemplo de encabezado:

```text
{{REPORT_TITLE}}
Unidad: {{FIELD:UNIDAD}}
```

Ejemplo recomendado de pie:

```text
Usuario: {{APP_USER}} | Fecha: {{GENERATED_AT}}
```

### 4.3 Formato admitido

- Fuentes equivalentes a Helvetica, Times o Courier.
- Tamaño entre 6 y 24 puntos.
- Peso normal o negrita.
- Color de texto y fondo uniforme.
- Alineación izquierda, centrada o derecha.
- Bordes simples y uniformes.
- Orientación vertical u horizontal; también puede decidirse en APEX con `AUTO`.

El título se normaliza centrado. El motor usa A4. El ajuste final puede diferir ligeramente de Word porque Word solo expresa intención visual; el PDF lo renderiza APEX.

### 4.4 Elementos no admitidos

No use imágenes, macros, `.docm`, cuadros de texto, WordArt, formas, gráficos, tablas anidadas, celdas combinadas, varias tablas, varias secciones, campos Word, hipervínculos externos, controles ActiveX, comentarios pendientes ni control de cambios. El compilador los rechaza para evitar resultados ambiguos y reducir riesgos.

## 5. Escribir la consulta SQL

Cada reporte usa un archivo SQL con una sola consulta `SELECT` o `WITH ... SELECT`. Los aliases seleccionados deben coincidir con los marcadores `COLUMN` de Word.

```sql
select e.vdni          as vdni,
       e.vnom          as vnom,
       d.descripcion   as departamento
  from entidad e
  left join mae_departamento d
    on d.id = e.departamento_id
 where (:DNI is null or e.vdni = :DNI)
   and (:DEPARTAMENTO_ID is null
        or e.departamento_id = :DEPARTAMENTO_ID)
 order by e.vnom
```

Los valores variables siempre usan binds. Nunca concatene el contenido de un Page Item dentro del texto SQL.

### 5.1 LOV: imprimir el display value

El modo SQL no intenta leer la configuración interna de un LOV del Interactive Grid. La consulta debe devolver explícitamente el texto que se imprimirá:

```sql
select e.departamento_id,
       d.descripcion as departamento
  from entidad e
  join mae_departamento d
    on d.id = e.departamento_id
```

La plantilla usa `{{COLUMN:DEPARTAMENTO}}`; el ID puede seguir usándose como bind o para lógica interna, pero no tiene que formar parte de la plantilla.

## 6. Crear el archivo de proyecto

El archivo `nombre.report.json` relaciona el DOCX y el SQL. Las rutas son relativas y deben permanecer dentro de la carpeta del proyecto.

```json
{
  "schema": "corporate-report-project/1.0",
  "report_id": "ENTIDADES",
  "template": "entidades.docx",
  "query_file": "entidades.sql",
  "title": "Relación de entidades",
  "file_name": "reporte_entidades",
  "orientation": "LANDSCAPE",
  "format_item": "P0_REPORT_FORMAT",
  "orientation_item": "P0_REPORT_ORIENTATION",
  "max_rows": 2000,
  "bindings": [
    {"bind":"DNI", "item":"P42_DNI", "type":"VARCHAR2", "required":false},
    {"bind":"DEPARTAMENTO_ID", "item":"P42_DEPARTAMENTO_ID", "type":"NUMBER", "required":false}
  ],
  "fields": [],
  "excluded_columns": [],
  "column_widths": [
    {"column":"VDNI", "mode":"FIXED_PERCENT", "value":14},
    {"column":"VNOM", "mode":"WEIGHT", "value":2},
    {"column":"DEPARTAMENTO", "mode":"AUTO"}
  ]
}
```

Límites que el compilador valida antes de generar PL/SQL:

- `title`: máximo 255 caracteres;
- `file_name`: máximo 180 caracteres y solo letras, números, `_` o `-`;
- constantes y máscaras de formato: máximo 4000 bytes UTF-8;
- nombres de Page Items: máximo 128 caracteres en total.

### 6.1 Bindings

Cada bind detectado en SQL debe tener exactamente un mapping en `bindings`.

| Propiedad | Significado |
|---|---|
| `bind` | Nombre sin dos puntos, por ejemplo `FECHA_DESDE` |
| `item` | Page Item, por ejemplo `P42_FECHA_DESDE` |
| `type` | `VARCHAR2`, `NUMBER`, `DATE` o `TIMESTAMP` |
| `required` | Si `true`, un valor vacío detiene el reporte |
| `format_mask` | Obligatorio para `DATE` y `TIMESTAMP`, por ejemplo `DD/MM/YYYY` |

Los nombres de Page Items deben usar el patrón `P<número>_NOMBRE`, con letras, números y guion bajo después del prefijo. No use `$` ni `#` en Page Items del proyecto.

`max_rows` es un límite de salida, no una comprobación de totalidad: si la consulta devuelve más filas, APEX exportará como máximo esa cantidad. Defina un valor acorde con el reporte y filtros que eviten resultados ambiguamente truncados.

Ejemplo de fecha:

```json
{
  "bind": "FECHA_DESDE",
  "item": "P42_FECHA_DESDE",
  "type": "DATE",
  "format_mask": "DD/MM/YYYY",
  "required": false
}
```

### 6.2 Campos escalares

Un `{{FIELD:NOMBRE}}` requiere una entrada en `fields`. Los orígenes son `ITEM`, `CONTEXT`, `SYSTEM` y `CONSTANT`.

```json
[
  {"name":"SOLICITANTE", "source":"ITEM", "item":"P42_SOLICITANTE", "type":"VARCHAR2"},
  {"name":"UNIDAD", "source":"CONSTANT", "value":"Oficina de Informática"}
]
```

No declare `APP_USER` ni `GENERATED_AT`: son campos incorporados.

### 6.3 Excluir columnas

Para impedir que una columna técnica se imprima, use aliases, no posiciones:

```json
"excluded_columns": ["PROCESSES", "ID_INTERNO"]
```

No se pueden excluir todas las columnas. Una exclusión no sustituye el control de acceso: el SQL no debe seleccionar datos sensibles innecesarios.

### 6.4 Definir anchos

| Modo | Comportamiento |
|---|---|
| `FIXED_PERCENT` | Solicita un porcentaje del ancho imprimible |
| `WEIGHT` | Reparte el espacio restante proporcionalmente |
| `AUTO` | Omite el ancho y delega el ajuste al motor de APEX |

Solo puede existir una columna `AUTO`, en cualquier posición. Una columna sin configuración se comporta como `WEIGHT` de valor 1. `AUTO` es una sugerencia al renderizador, no una garantía matemática.

## 7. Validar y compilar

Valide primero, sin escribir artefactos:

```powershell
apex-report-compiler validate `
  --project .\mi_reporte\mi_reporte.report.json
```

Compile cuando la validación sea correcta:

```powershell
apex-report-compiler compile `
  --project .\mi_reporte\mi_reporte.report.json `
  --output .\build\mi_reporte
```

El modo predeterminado es estricto. `--compatible` admite únicamente aproximaciones documentadas y las informa como advertencias. Para integrar la herramienta en automatización, `--json` emite diagnósticos estructurados.

La compilación genera:

| Archivo | Uso |
|---|---|
| `template.json` | Interpretación canónica y auditable de la plantilla |
| `validation.json` | Diagnósticos reproducibles |
| `apex_process.sql` | Proceso PL/SQL listo para revisar y copiar |

Si una compilación falla, no reemplaza una salida válida anterior.

## 8. Instalar el paquete común

La instalación se realiza una sola vez por esquema de aplicación. Desde SQLcl o SQL*Plus:

```sql
@sql/install.sql
```

Desde SQL Workshop cargue y ejecute `sql/pkg_corporate_reports.sql`, y compruebe:

```sql
select type, line, position, text
  from user_errors
 where name = 'PKG_CORPORATE_REPORTS'
 order by sequence;
```

Tanto el package como su body deben figurar como `VALID`. No conceda `EXECUTE` a `PUBLIC` y mantenga `AUTHID CURRENT_USER`.

## 9. Integrar en una página APEX

1. Cree los Page Items usados por `bindings` y, si aplica, `P0_REPORT_FORMAT` y `P0_REPORT_ORIENTATION`.
2. Use valores de orientación `AUTO`, `PORTRAIT` o `LANDSCAPE`; el valor predeterminado recomendado es `AUTO`.
3. Asegure que los filtros estén en estado de sesión antes de descargar. En una Dynamic Action Ajax deben aparecer en **Items to Submit**; un submit completo procesa los items normalmente.
4. Cree un proceso de descarga adecuado al flujo de la página y copie el contenido revisado de `apex_process.sql`.
5. Proteja el botón y el proceso con las autorizaciones de la aplicación.
6. Pruebe con un usuario real autorizado.

El código generado llama a esta API congelada:

```plsql
pkg_corporate_reports.download_query(
    p_sql_query             => l_sql,
    p_columns_json          => l_columns_json,
    p_bindings_json         => l_bindings_json,
    p_fields_json           => l_fields_json,
    p_style_json            => l_style_json,
    p_header_template       => '{{REPORT_TITLE}}',
    p_footer_template       => 'Usuario: {{APP_USER}} | Fecha: {{GENERATED_AT}}',
    p_title                 => 'Relación de entidades',
    p_file_name             => 'reporte_entidades',
    p_format                => 'PDF',
    p_orientation           => 'AUTO',
    p_max_rows              => 2000,
    p_excluded_columns_json => l_excluded_columns,
    p_column_widths_json    => l_column_widths
);
```

No existen en esta API los parámetros `p_template_json`, `p_generated_by`, `p_paper_size` ni `p_auto_reserve_pct`. El usuario se obtiene del contexto de APEX y el papel es A4.

## 10. Pruebas obligatorias en APEX

Antes de producción pruebe:

- compilación del package sin errores;
- PDF y XLSX;
- consulta con cero, una y muchas filas;
- límite `p_max_rows`;
- filtros vacíos y requeridos;
- `VARCHAR2`, `NUMBER`, `DATE` y `TIMESTAMP`;
- orientación `AUTO`, `PORTRAIT` y `LANDSCAPE`;
- columnas fijas, ponderadas, automática y excluidas;
- LOV mostrando descripción, no ID;
- usuario y fecha en el pie;
- permisos de un usuario autorizado y otro no autorizado;
- compatibilidad de `DOWNLOAD_IG` y `DOWNLOAD_IG_PDF` si ya se usan.

## 11. Diagnóstico de problemas

| Síntoma | Causa probable | Acción |
|---|---|---|
| Falta un bind | El SQL contiene `:NOMBRE` sin mapping | Añadirlo a `bindings` |
| Page Item inválido | Nombre mal escrito o fuera del patrón | Usar `P0_...` o `P<número>_...` |
| El PDF imprime un ID de LOV | El SQL selecciona el return value | Hacer `JOIN` y devolver el display value |
| Columna inexistente | Alias Word y SQL no coinciden | Alinear `COLUMN:ALIAS` y `SELECT ... AS alias` |
| Fecha inválida | Valor de sesión no coincide con la máscara | Corregir `format_mask` o el formato del item |
| No se descarga | Flujo de página o estado de sesión incorrecto | Revisar request, proceso, branch e Items to Submit |
| Tabla demasiado comprimida | Demasiadas columnas o pesos inadecuados | Excluir columnas, usar landscape o ajustar anchos |
| DOCX rechazado | Contiene estructura fuera del contrato | Simplificar usando una plantilla de ejemplo |

## 12. Mantenimiento recomendado

- Mantenga juntos el DOCX, `.sql`, `.report.json` y una salida de referencia.
- Versione cada cambio y revise SQL/configuración como código.
- Recompile después de cambiar Word, SQL, aliases, campos, exclusiones o anchos.
- No edite `apex_process.sql` sin reflejar el cambio en el proyecto fuente.
- Instale una sola versión central del paquete; no genere un package distinto por reporte.
- Revise rendimiento y permisos del SQL antes de promover.

El directorio `examples` contiene un proyecto completo que sirve como punto de partida.
