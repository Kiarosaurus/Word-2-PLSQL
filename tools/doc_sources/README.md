# Compilador local Word → reportes Oracle APEX

Este proyecto permite que Informática diseñe un reporte tabular en Microsoft Word, lo valide localmente y genere el bloque PL/SQL mínimo que Oracle APEX 24.2 necesita para descargar PDF o XLSX mediante `APEX_DATA_EXPORT`.

Word y Python se usan solamente en la computadora del desarrollador. En el servidor no se requieren Word, LibreOffice, Python, Docker ni un motor de impresión adicional. En Oracle se instala una sola vez `PKG_CORPORATE_REPORTS`.

## Alcance de la versión 1.0

- una consulta SQL `SELECT` o `WITH` por reporte;
- una plantilla DOCX A4 con un título, una tabla de dos filas y un pie;
- fila 1: encabezados literales;
- fila 2: un marcador `{{COLUMN:ALIAS}}` por celda;
- marcadores escalares `{{REPORT_TITLE}}`, `{{APP_USER}}`, `{{GENERATED_AT}}` y `{{FIELD:NOMBRE}}`;
- bindings tipados a Page Items sin concatenar valores en el SQL;
- anchos `FIXED_PERCENT`, `WEIGHT` y una columna `AUTO`;
- orientación `AUTO`, `PORTRAIT` o `LANDSCAPE`;
- salida PDF o XLSX.

El SQL debe devolver el texto que se desea imprimir. Para un LOV, use un `JOIN` y seleccione el *display value* con el alias declarado en Word.

La plantilla de referencia (`templates/` y `proyectos/entidades/entidades.docx`) usa los cinco tipos de marcador:

```
{{REPORT_TITLE}}                  <- título: párrafo previo a la tabla
Unidad: {{FIELD:UNIDAD}}          <- campo escalar en la misma zona del título

DNI              Nombre completo   ...   <- fila 1: etiquetas literales
{{COLUMN:VDNI}}  {{COLUMN:VNOM}}   ...   <- fila 2: un alias SQL por celda

Usuario: {{APP_USER}} | Fecha: {{GENERATED_AT}}   <- pie de página de Word
```

## Flujo de trabajo

```flujo
# Figura 1. Del Word a la descarga en APEX
Informática | Diseñar la plantilla Word | título, tabla de dos filas y pie con marcadores {{...}}
Compilador | Nuevo proyecto desde DOCX | copia el Word a proyectos\<nombre>\ y crea automáticamente generado\<nombre>.sql y .report.json
Informática | Ajustar SQL y proyecto | tablas reales, filtros y anchos; editables a mano en cualquier momento
Compilador | Validar | revisa DOCX, SQL y proyecto sin escribir archivos
? Compilador | ¿Sin errores? | No: corregir según el código de diagnóstico y validar otra vez
Compilador | Compilar | genera automáticamente apex_process.sql junto al DOCX, y template.json y validation.json en la carpeta generado
APEX | Pegar apex_process.sql en el proceso de descarga | el package se instala una sola vez por esquema
Usuario final | Descargar PDF o XLSX | APEX_DATA_EXPORT genera el archivo en el servidor
```

Cada vez que cambia la plantilla Word basta con volver a compilar: `template.json` y `apex_process.sql` se regeneran automáticamente a partir del DOCX, el SQL y el `.report.json`. El `.report.json` y el `.sql` son archivos de texto que el usuario puede editar directamente en cualquier momento.

## Inicio rápido en Windows

Haga doble clic en **Compilador de reportes APEX** (acceso directo de esta carpeta) o en **Abrir compilador.bat**. La primera vez prepara `.venv` automáticamente; requiere Python 3.12 o superior. El acceso directo sigue funcionando si la carpeta se mueve o se entrega comprimida.

Instalación manual equivalente:

```
py -3.13 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e .
```

Validar el ejemplo sin escribir artefactos:

```
# Revisa DOCX + SQL + proyecto; no escribe nada en disco
apex-report-compiler validate `
  --project .\proyectos\entidades\generado\entidades.report.json
```

Compilarlo:

```
# apex_process.sql -> proyectos\entidades\ ; template.json y validation.json -> generado\
apex-report-compiler compile `
  --project .\proyectos\entidades\generado\entidades.report.json
```

Cargar un Word propio (se copia a `proyectos\<nombre>\`):

```
# Crea proyectos\ventas\ventas.docx y proyectos\ventas\generado\ventas.sql/.report.json
apex-report-compiler new --docx C:\Descargas\ventas.docx --page 42
```

Interfaz gráfica:

```
# Abre la ventana «Compilador de reportes APEX»
apex-report-compiler-gui
```

La interfaz puede crear un proyecto inicial desde un DOCX en `proyectos\<nombre>\` (Page Items `PXX_` mapeados a los marcadores del Word) y, al compilar, muestra exactamente qué archivo subir a APEX, dónde pegarlo y qué Page Items crear.

También puede ejecutar la CLI sin instalar el paquete editable:

```
# Mismo comando validate usando el código de src\, sin instalar el paquete
python .\launch_compiler.py validate `
  --project .\proyectos\entidades\generado\entidades.report.json
```

## Artefactos generados

```
proyectos\ventas\
  ventas.docx            <- plantilla
  apex_process.sql       <- proceso PL/SQL listo para revisar y copiar en APEX
  generado\
    ventas.sql           <- consulta fuente (su texto queda dentro de apex_process.sql)
    ventas.report.json   <- proyecto: binds, campos, anchos, estilos
    template.json        <- interpretación canónica de la plantilla y configuración
    validation.json      <- diagnósticos reproducibles
```

Los tres se generan automáticamente en cada compilación; no hay que escribirlos a mano. Pueden editarse directamente si hace falta, pero la siguiente compilación los reemplaza: los cambios permanentes se hacen en el DOCX, el `.sql` o el `.report.json`.

La compilación es transaccional: si la validación falla, se conserva intacta la última salida válida.

## Instalación en Oracle APEX

1. Lea `sql/README.docx`.
2. Instale `sql/pkg_corporate_reports.sql` en el *parsing schema*.
3. Confirme que el package y su body estén `VALID`.
4. Compile localmente el proyecto del reporte.
5. Revise y copie `apex_process.sql` en un proceso de descarga de APEX.
6. Pruebe PDF/XLSX, seguridad, Page Items, límites y datos reales en APEX 24.2 antes de promover a producción.

El proceso generado llama a la operación pública `PKG_CORPORATE_REPORTS.DOWNLOAD_QUERY`.

## Documentación

- `deliverables/Manual_de_uso.docx`: guía completa para diseñadores y desarrolladores APEX.
- `deliverables/Informe_de_alcance_y_limitaciones.docx`: alcance aprobado, arquitectura, seguridad y límites.
- `docs/CONTRACT.docx`: contrato técnico normativo.
- `docs/TEMPLATE_QA.docx`: control de calidad de las plantillas de referencia.
- `docs/RELEASE_REVIEW.docx`: evidencia y veredicto de la revisión de liberación.
- `sql/README.docx`: instalación y seguridad del package.
- `proyectos/`: una carpeta por reporte, con el mismo nombre que su DOCX. Junto al DOCX queda `apex_process.sql` (lo que se pega en APEX); el material de trabajo va en `generado/`. Volver a cargar un DOCX con el mismo nombre reemplaza solo esos archivos, con copias `.bak`; nada más se borra.
- `proyectos/entidades/`: proyecto de ejemplo ejecutable; es el único que incluye además su documentación (`README.docx`).
- `templates/`: plantillas Word de referencia, vertical y horizontal.
- `tools/doc_sources/`: fuentes de esta documentación; `python tools\build_documentation.py` regenera todos los Word.
- `tests/`: pruebas automáticas del compilador y del contrato estático PL/SQL.

## Límite de validación local

La suite local verifica el compilador, la seguridad del DOCX, la generación determinista y el contrato estático del package. No reemplaza la compilación real contra las APIs instaladas de APEX ni una descarga PDF/XLSX en Oracle; esas pruebas finales deben realizarse en la aplicación APEX 24.2.
