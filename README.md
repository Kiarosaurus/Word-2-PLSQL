# Compilador local Word → reportes Oracle APEX

Este proyecto permite que Informática diseñe un reporte tabular en Microsoft
Word, lo valide localmente y genere el bloque PL/SQL mínimo que Oracle APEX
24.2 necesita para descargar PDF o XLSX mediante `APEX_DATA_EXPORT`.

Word y Python se usan solamente en la computadora del desarrollador. En el
servidor no se requieren Word, LibreOffice, Python, Docker ni un motor de
impresión adicional. En Oracle se instala una sola vez
`PKG_CORPORATE_REPORTS`.

## Alcance de la versión 1.0

- una consulta SQL `SELECT` o `WITH` por reporte;
- una plantilla DOCX A4 con un título, una tabla de dos filas y un pie;
- fila 1: encabezados literales;
- fila 2: un marcador `{{COLUMN:ALIAS}}` por celda;
- marcadores escalares `{{REPORT_TITLE}}`, `{{APP_USER}}`,
  `{{GENERATED_AT}}` y `{{FIELD:NOMBRE}}`;
- bindings tipados a Page Items sin concatenar valores en el SQL;
- anchos `FIXED_PERCENT`, `WEIGHT` y una columna `AUTO`;
- orientación `AUTO`, `PORTRAIT` o `LANDSCAPE`;
- salida PDF o XLSX.

El SQL debe devolver el texto que se desea imprimir. Para un LOV, use un
`JOIN` y seleccione el *display value* con el alias declarado en Word.

## Inicio rápido en Windows

```powershell
py -3.13 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e .
```

Validar el ejemplo sin escribir artefactos:

```powershell
apex-report-compiler validate `
  --project .\examples\entidades.report.json
```

Compilarlo:

```powershell
apex-report-compiler compile `
  --project .\examples\entidades.report.json `
  --output .\build\entidades
```

Interfaz gráfica opcional:

```powershell
apex-report-compiler-gui
```

También puede ejecutar la CLI sin instalar el paquete editable:

```powershell
python .\launch_compiler.py validate `
  --project .\examples\entidades.report.json
```

## Artefactos generados

- `template.json`: interpretación canónica de la plantilla y configuración;
- `validation.json`: diagnósticos reproducibles;
- `apex_process.sql`: proceso PL/SQL listo para revisar y copiar en APEX.

La compilación es transaccional: si la validación falla, no sustituye una
salida anterior válida.

## Instalación en Oracle APEX

1. Lea `sql/README.md`.
2. Instale `sql/pkg_corporate_reports.sql` en el *parsing schema*.
3. Confirme que el package y su body estén `VALID`.
4. Compile localmente el proyecto del reporte.
5. Revise y copie `apex_process.sql` en un proceso de descarga de APEX.
6. Pruebe PDF/XLSX, seguridad, Page Items, límites y datos reales en APEX
   24.2 antes de promover a producción.

El package conserva las operaciones anteriores `DOWNLOAD_IG` y
`DOWNLOAD_IG_PDF`; la nueva operación SQL declarativa es `DOWNLOAD_QUERY`.

## Documentación

- `deliverables/Manual_de_uso.docx`: guía completa para diseñadores y
  desarrolladores APEX.
- `deliverables/Informe_de_alcance_y_limitaciones.docx`: alcance aprobado,
  arquitectura, seguridad y límites.
- `docs/CONTRACT.md`: contrato técnico normativo.
- `examples/`: proyecto ejecutable y resultado esperado.
- `tests/`: pruebas automáticas del compilador y del contrato estático PL/SQL.

## Límite de validación local

La suite local verifica el compilador, la seguridad del DOCX, la generación
determinista y el contrato estático del package. No reemplaza la compilación
real contra las APIs instaladas de APEX ni una descarga PDF/XLSX en Oracle;
esas pruebas finales deben realizarse en la aplicación APEX 24.2.
