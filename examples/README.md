# Ejemplo ejecutable `ENTIDADES`

Este directorio contiene un proyecto completo que el compilador puede validar y compilar sin mover archivos. Los nombres de tablas son ilustrativos; solo la ejecución posterior en APEX exige sustituirlos por objetos que existan en el esquema de la aplicación.

## Archivos

- `entidades.docx`: plantilla Word restringida, A4 horizontal y con seis columnas.
- `entidades.sql`: consulta de solo lectura con exactamente los seis aliases de la plantilla.
- `entidades.report.json`: configuración, binds, anchos y estilos admitidos.
- `build_expected/`: salida de referencia generada por el compilador.

El DOCX se utiliza solamente durante la compilación local. En APEX se instala el paquete una vez y se copia el bloque generado en `build_expected/apex_process.sql`.

## Contrato de la plantilla

La plantilla incluye el título integrado:

```text
{{REPORT_TITLE}}
```

El pie de página incluye los valores integrados:

```text
Usuario: {{APP_USER}} | Fecha: {{GENERATED_AT}}
```

La tabla contiene una fila de etiquetas y otra de marcadores. Los seis aliases coinciden con el `SELECT`:

| Encabezado | Marcador Word | Alias SQL |
|---|---|---|
| DNI | `{{COLUMN:VDNI}}` | `VDNI` |
| Nombre completo | `{{COLUMN:VNOM}}` | `VNOM` |
| Dirección actual | `{{COLUMN:VDIREC_ACTUAL}}` | `VDIREC_ACTUAL` |
| Departamento | `{{COLUMN:DEPARTAMENTO}}` | `DEPARTAMENTO` |
| Distrito | `{{COLUMN:DISTRITO}}` | `DISTRITO` |
| Teléfono | `{{COLUMN:VNRO_TLF1}}` | `VNRO_TLF1` |

No se declararon campos `{{FIELD:...}}`; por eso `fields` está vacío.

## Binds lógicos y Page Items

| Bind SQL | Page Item | Tipo |
|---|---|---|
| `:DNI` | `P42_DNI` | `VARCHAR2` |
| `:DEPARTAMENTO_ID` | `P42_DEPARTAMENTO_ID` | `NUMBER` |
| `:FECHA_DESDE` | `P42_FECHA_DESDE` | `DATE`, máscara `DD/MM/YYYY` |

El JSON usa la clave vigente `bindings`; cada entrada usa `bind`, no `name`. El proceso generado no concatena estos valores en el SQL: `PKG_CORPORATE_REPORTS` los obtiene del estado de sesión y los añade como parámetros tipados de `APEX_EXEC`.

Antes de la descarga, la página debe enviar los tres Page Items al estado de sesión. En una Dynamic Action Ajax deben aparecer en **Items to Submit**; en un submit normal deben ser procesados por la página.

## LOV: imprimir display, no return

`P42_DEPARTAMENTO_ID` puede almacenar el identificador numérico del LOV. Ese valor sirve solo como filtro. El reporte imprime `d.descripcion AS departamento`, por lo que el PDF recibe el texto visible y no el ID.

La misma regla debe aplicarse a cualquier otro LOV: la consulta del reporte debe hacer el `JOIN` correspondiente y seleccionar el display con el alias declarado en Word.

## Anchos y exclusiones

- `VDNI` y `VNRO_TLF1` usan `FIXED_PERCENT`.
- `VNOM`, `DEPARTAMENTO` y `DISTRITO` distribuyen el espacio restante mediante `WEIGHT`.
- `VDIREC_ACTUAL` usa `AUTO`; solo puede existir una columna con ese modo.
- `excluded_columns` está presente y vacío. Para una exclusión institucional se agrega allí uno de los seis aliases; no se usan claves antiguas como `exclude_columns`.

Los estilos avanzados están en `style_overrides`. La zona del título se denomina `title`; `header` es una compatibilidad antigua y no se utiliza en este ejemplo.

## Validar y compilar

Desde la raíz del proyecto, con el entorno virtual activado:

```powershell
python -m corporate_report_compiler validate `
  --project .\examples\entidades.report.json

python -m corporate_report_compiler compile `
  --project .\examples\entidades.report.json `
  --output .\examples\build_expected
```

En Linux o macOS:

```bash
PYTHONPATH=src python -m corporate_report_compiler validate \
  --project examples/entidades.report.json

PYTHONPATH=src python -m corporate_report_compiler compile \
  --project examples/entidades.report.json \
  --output examples/build_expected
```

Una validación correcta no escribe archivos. La compilación genera:

- `template.json`: contrato normalizado y estilos extraídos del DOCX;
- `validation.json`: diagnóstico reproducible de la compilación;
- `apex_process.sql`: llamada mínima a `PKG_CORPORATE_REPORTS.DOWNLOAD_QUERY`.

## Adaptación a una página real

1. Cambiar las tablas y joins del SQL manteniendo los aliases del DOCX.
2. Cambiar los Page Items de `bindings` por los de la página.
3. Si se cambia una columna, actualizar coordinadamente Word, el `SELECT` y `column_widths`.
4. Ajustar título, nombre de archivo, límite de filas, exclusiones y anchos.
5. Validar y compilar localmente.
6. Revisar `template.json`, `validation.json` y `apex_process.sql`.
7. Copiar el proceso generado a APEX y probarlo allí.

El SQL, el JSON y el proceso son artefactos administrados por Informática. No deben recibirse desde el navegador ni ser editables por el usuario final.
