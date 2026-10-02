# Ejemplo ejecutable ENTIDADES

Este directorio contiene un proyecto completo que el compilador puede validar y compilar sin mover archivos. Los nombres de tablas son ilustrativos; solo la ejecución posterior en APEX exige sustituirlos por objetos que existan en el esquema de la aplicación.

## Archivos

- `entidades.docx`: plantilla Word restringida, A4 horizontal, con seis columnas y los cinco tipos de marcador.
- `entidades.sql`: consulta de solo lectura con exactamente los seis aliases de la plantilla.
- `entidades.report.json`: configuración, binds, campos, anchos y estilos admitidos.
- `build_expected/`: salida de referencia generada automáticamente por el compilador.

El DOCX se utiliza solamente durante la compilación local. En APEX se instala el paquete una vez y se pega el bloque `apex_process.sql` que genere al compilar. Las tablas del SQL (`entidad`, `mae_departamento`, `mae_distrito`) son ilustrativas: el proceso solo funcionará en APEX después de adaptar el SQL a objetos reales y recompilar.

## Contrato de la plantilla

El párrafo de título contiene el título integrado y un campo escalar, separados por un salto de línea manual (Mayús+Intro) dentro del mismo párrafo:

```
{{REPORT_TITLE}}            <- se reemplaza por "title": "Relación de entidades"
Unidad: {{FIELD:UNIDAD}}    <- se reemplaza por el campo UNIDAD de "fields"
```

El pie de página incluye los valores integrados:

```
Usuario: {{APP_USER}} | Fecha: {{GENERATED_AT}}   <- usuario APEX y fecha DD/MM/YYYY HH24:MI
```

La tabla contiene una fila de etiquetas y otra de marcadores. Los seis aliases coinciden con el `SELECT`:

| Encabezado | Marcador Word | Alias SQL |
|---|---|---|
| DNI | {{COLUMN:VDNI}} | VDNI |
| Nombre completo | {{COLUMN:VNOM}} | VNOM |
| Dirección actual | {{COLUMN:VDIREC_ACTUAL}} | VDIREC_ACTUAL |
| Departamento | {{COLUMN:DEPARTAMENTO}} | DEPARTAMENTO |
| Distrito | {{COLUMN:DISTRITO}} | DISTRITO |
| Teléfono | {{COLUMN:VNRO_TLF1}} | VNRO_TLF1 |

El campo `{{FIELD:UNIDAD}}` se declara en `fields` como constante:

```
"fields": [
  {"name": "UNIDAD", "source": "CONSTANT", "value": "Oficina de Informática"}
]
```

Para que el valor provenga de la página en lugar de una constante, cámbielo a `{"name": "UNIDAD", "source": "ITEM", "item": "P42_UNIDAD", "type": "VARCHAR2"}` y cree ese Page Item.

## Binds lógicos y Page Items

| Bind SQL | Page Item | Tipo |
|---|---|---|
| :DNI | P42_DNI | VARCHAR2 |
| :DEPARTAMENTO_ID | P42_DEPARTAMENTO_ID | NUMBER |
| :FECHA_DESDE | P42_FECHA_DESDE | DATE, máscara DD/MM/YYYY |

Cada entrada de `bindings` relaciona un `:BIND` del SQL con un Page Item mediante las claves `bind` e `item`. El proceso generado no concatena estos valores en el SQL: `PKG_CORPORATE_REPORTS` los obtiene del estado de sesión y los añade como parámetros tipados de `APEX_EXEC`.

Antes de la descarga, los tres Page Items deben estar en el estado de sesión. El botón de descarga usa *Submit Page*, que los envía automáticamente; no use una Dynamic Action Ajax para descargar.

El proyecto declara además `P0_REPORT_FORMAT` y `P0_REPORT_ORIENTATION` en la página global (página 0), con valores de retorno `PDF`/`XLSX` y `AUTO`/`PORTRAIT`/`LANDSCAPE`. Si no desea que el usuario elija, elimine `format_item` y `orientation_item` del proyecto y recompile.

## LOV: imprimir display, no return

`P42_DEPARTAMENTO_ID` puede almacenar el identificador numérico del LOV. Ese valor sirve solo como filtro. El reporte imprime `d.descripcion AS departamento`, por lo que el PDF recibe el texto visible y no el ID.

La misma regla debe aplicarse a cualquier otro LOV: la consulta del reporte debe hacer el `JOIN` correspondiente y seleccionar el display con el alias declarado en Word.

## Anchos y exclusiones

- `VDNI` y `VNRO_TLF1` usan `FIXED_PERCENT`.
- `VNOM`, `DEPARTAMENTO` y `DISTRITO` distribuyen el espacio restante mediante `WEIGHT`.
- `VDIREC_ACTUAL` usa `AUTO`; solo puede existir una columna con ese modo.
- `excluded_columns` está presente y vacío. Para una exclusión institucional se agrega allí uno de los seis aliases.

Los estilos avanzados están en `style_overrides`, con las zonas `title`, `table_header`, `table_body`, `border` y `footer`. Sus valores prevalecen sobre los que el compilador extrae automáticamente del DOCX; en este ejemplo coinciden con la plantilla y sirven como muestra para editarlos directamente.

## Validar y compilar

Desde la raíz del proyecto, con el entorno virtual activado:

```
# 1. Validar: no escribe archivos
python -m corporate_report_compiler validate `
  --project .\examples\entidades.report.json

# 2. Compilar en build\entidades (nunca sobre examples\build_expected)
python -m corporate_report_compiler compile `
  --project .\examples\entidades.report.json `
  --output .\build\entidades
```

Compare la salida con la referencia; los tres comandos deben indicar que no hay diferencias:

```
# Cada comando debe responder «no se encontraron diferencias»
fc.exe /b .\build\entidades\template.json .\examples\build_expected\template.json
fc.exe /b .\build\entidades\validation.json .\examples\build_expected\validation.json
fc.exe /b .\build\entidades\apex_process.sql .\examples\build_expected\apex_process.sql
```

No compile sobre `examples\build_expected`: es la referencia que detecta regresiones.

En Linux o macOS:

```
PYTHONPATH=src python -m corporate_report_compiler compile \
  --project examples/entidades.report.json \
  --output build/entidades
cmp build/entidades/apex_process.sql examples/build_expected/apex_process.sql
```

Una validación correcta no escribe archivos. La compilación genera automáticamente:

- `template.json`: contrato normalizado y estilos extraídos del DOCX;
- `validation.json`: diagnóstico reproducible de la compilación;
- `apex_process.sql`: llamada mínima a `PKG_CORPORATE_REPORTS.DOWNLOAD_QUERY`.

## Adaptación a una página real

```flujo
# Figura 1. De este ejemplo a un reporte real
Informática | Copiar entidades.docx y editarlo en Word | etiquetas, estilos y columnas; mantener los marcadores {{...}}
Compilador | Nuevo proyecto desde DOCX | crea automáticamente <nombre>.report.json y <nombre>.sql desde el Word
Informática | Adaptar SQL y proyecto | tablas reales, aliases del DOCX, Page Items, título, anchos y max_rows
Compilador | Validar y compilar | template.json y apex_process.sql se regeneran automáticamente
APEX | Pegar el proceso generado y probar | Manual de uso, sección 9
```

Alternativa: copiar también `entidades.sql` y `entidades.report.json`, renombrarlos y editarlos directamente.

1. Cambiar las tablas y joins del SQL manteniendo los aliases del DOCX.
2. Cambiar los Page Items de `bindings` por los de la página.
3. Si se añade o quita una columna en Word, añadir o quitar su alias en el `SELECT` y, si figura, en `column_widths`. Las etiquetas, estilos y el orden se toman automáticamente del DOCX al recompilar.
4. Ajustar título, nombre de archivo, límite de filas, campos, exclusiones y anchos directamente en el `.report.json`.
5. Validar y compilar localmente.
6. Revisar `template.json`, `validation.json` y `apex_process.sql`.
7. Copiar el proceso generado a APEX (Manual de uso, sección 9) y probarlo allí. Una validación local correcta no equivale a aceptación en APEX.

El SQL, el JSON y el proceso son artefactos administrados por Informática. No deben recibirse desde el navegador ni ser editables por el usuario final.
