# Estudio de reportes, primera versión

## Estado comprobado

El proyecto mantiene dos contratos de compilación. El modo simple acepta un Word tabular y una consulta y emite configuración para `APEX_DATA_EXPORT`. El modo layout compila párrafos, tablas, marcadores `FIELD`, `COLUMN`, `SUM`, `GROUP`, encabezados y pies a `layout.json` y un paquete PL/SQL que llama a `RPT_LAYOUT` y `RPT_PDF`. La lectura de Oracle Reports usa `rwconverter` para convertir un RDF binario a XML. El modelo existente lee consultas, columnas, grupos, enlaces, parámetros, fórmulas y resúmenes; la conversión de fórmulas PL/SQL puede producir funciones que devuelven `NULL` con un diagnóstico.

El corpus bajo `proyectos` contiene ocho proyectos: `demo/entidades` (simple), `demo/estado_cuenta`, `FCR108_B`, `FCR118_E`, `PFR201_3`, `PSR308W_1I`, `PSR308Y` y `RPC004`. Seis conservan el XML de Reports. La carpeta externa `D:\CHAMBERCITO\Migración`, proporcionada después del primer inventario, contiene seis RDF y nueve PDF de despliegue u originales. Las pruebas detectan esas referencias de forma opcional y nunca las modifican.

## Arquitectura elegida

`studio_model.json` guarda el grafo importado (consultas, columnas, grupos y enlaces) para inspección; `.report.json` sigue siendo el contrato compatible del compilador. El Word conserva la presentación editable y los marcadores de repetición. `validate_project` produce una definición común que consume el emisor Oracle y la vista local. Así se comparten interpretación de Word, resolución de nombres y diagnósticos. La vista HTML tiene un renderizador estructural local, explícitamente aproximado: no calcula PL/SQL ni reproduce la medición de texto y paginación de `RPT_LAYOUT`. Para equivalencia visual real hace falta ejecutar el mismo paquete en Oracle o portar su motor de layout a una biblioteca común. Esa segunda opción evita Oracle en el diseño, pero duplicaría la lógica de PDF y exige una validación extensa; se elige la primera como siguiente incremento.

La importación RDF crea un Word inicial a partir del modelo de datos, lo pasa por el importador existente y guarda `studio_findings.json` con las omisiones identificadas. No reescribe proyectos existentes. El proyecto SQL nuevo genera un Word, copia las consultas y declara sus binds como Page Items de texto para ajustar tipo y página antes de publicar.

## Uso

Desde la raíz, con `python -m pip install -e .`:

```powershell
python -m corporate_report_compiler studio-rdf --rdf C:\ruta\reporte.rdf --projects-dir proyectos --page 70
python -m corporate_report_compiler studio-queries --name VENTAS --query VENTAS=C:\ruta\ventas.sql --projects-dir proyectos --page 70
python -m corporate_report_compiler studio-preview --project proyectos\VENTAS\generado\VENTAS.report.json --output tmp\ventas.html
python -m corporate_report_compiler validate --project proyectos\VENTAS\generado\VENTAS.report.json --compatible
python -m corporate_report_compiler compile --project proyectos\VENTAS\generado\VENTAS.report.json --compatible
```

Abra el HTML en un navegador y seleccione los valores azules. El panel muestra SQL, parámetros, fórmula original cuando existe, grupo, ámbito, repetición y comportamiento de filas vacías/página. Edite el Word y el JSON en sus aplicaciones habituales, vuelva a validar y regenere la vista. Para importar un RDF binario se requiere `rwconverter.exe` de Oracle Reports configurado en `RWCONVERTER`; los XML de referencia permiten probar localmente la misma ruta sin ese ejecutable.

En este equipo `rwconverter.exe` ya está en `D:\oracle\Middleware\FR_Home\bin\rwconverter.exe` y LibreOffice en `C:\Program Files\LibreOffice\program\soffice.exe`. Para renderizar un Word en PowerShell sin cambiar el `PATH` del sistema:

```powershell
$env:PATH = 'C:\Program Files\LibreOffice\program;' + $env:PATH
& '.\.venv\Scripts\python.exe' C:\Users\Usuario\.codex\plugins\cache\openai-primary-runtime\documents\26.1025.11240\skills\documents\render_docx.py proyectos\FCR118_E\FCR118_E.docx --output_dir tmp\revision_fcr118 --emit_pdf
```

Use el ejecutable de Python de `.venv` explícitamente: LibreOffice también instala un `python.exe` en su carpeta y, al anteponerla al `PATH`, puede reemplazar al del proyecto.

## Criterios de aceptación y estado

| Criterio medible | Estado de esta versión |
|---|---|
| Los ocho proyectos existentes validan y compilan sin Oracle en modo compatible | Cumple en pruebas locales |
| Los seis XML de Reports producen proyecto y Word que validan y se previsualizan | Cumple en pruebas locales |
| Un conjunto de queries SQL produce Word, proyecto y vista con trazabilidad | Cumple para SELECT con aliases reconocibles |
| Toda pérdida de RDF queda identificada por elemento o columna | Parcial; identifica identificadores incompatibles, columnas omitidas y recuentos de marcos y reglas; no interpreta toda semántica de formato |
| Vista con varios registros, cambios de grupo, nulos y totales reproducibles | Parcial; genera seis filas y nulos, señala grupo y total, pero no ejecuta fórmulas ni agregaciones reales |
| Vista y PDF APEX tienen igual paginación y geometría | Pendiente; el HTML es estructural y el PDF real requiere Oracle/APEX |
| Un usuario define visualmente relaciones, grupos, fórmulas y paginación desde cero | Pendiente; esta versión permite modificar modelo y Word, sin editor visual de reglas |
| Comparación visual con PDF original en cada proyecto | Parcial; ahora hay PDFs en `Migración`, pero falta un PDF APEX comparable por cada juego de datos |

## Próximo incremento

Agregar un editor de reglas sobre el grafo de `studio_model.json` con migración explícita a un esquema nuevo, validación de relaciones y cambios de grupo; luego ofrecer una vista previa alimentada por un servicio Oracle que ejecute `RPT_LAYOUT` con datos de prueba reproducibles. Incorporar capturas comparables y métricas de páginas, bloques y posiciones para cada PDF original que se consiga.

## Verificación realizada el 28 de octubre de 2026

`python -m pytest -q`: 221 pruebas y 219 subpruebas aprobadas antes de añadir las pruebas de referencias externas. La nueva prueba recorre los ocho proyectos para validación, compilación en carpeta temporal y vista HTML; importa los seis XML de Reports (incluidos FCR108_B y FCR118_E), y verifica el flujo SQL nuevo. Las pruebas añadidas después revisan los PDF de `Migración` y convierten el RDF binario FCR118_E cuando `rwconverter` está disponible. No modifican los PDF de referencia.

Una revisión posterior encontró LibreOffice instalado fuera de `PATH`. Con la ruta añadida a la sesión se renderizaron el Word inicial de FCR108_B y el Word existente de FCR118_E. El Word inicial FCR108_B quedó demasiado apretado y no reproduce el diseño de Reports; el Word existente FCR118_E sí conserva el esquema de encabezados y columnas. También se inspeccionó la primera página de `FCR118_E - original.pdf` y `FCR118_E - emitido.pdf`: comparten la estructura general, pero tienen 9 y 4 páginas respectivamente, distinto tamaño de papel (Letter frente a A4) y valores diferentes. La comparación de números extraídos encontró 880 coincidencias con multiplicidad entre 1161 valores del original y 1155 del emitido. No se debe interpretar como equivalencia de datos ni de paginación. La prueba estructural del archivo externo verifica páginas, texto extraíble y un umbral de solapamiento numérico cuando ambos PDF FCR118_E están presentes.

El Word inicial se ajustó luego a A4 horizontal cuando una consulta tiene más de seis columnas, marcadores pequeños y título sin borde. Su nueva versión se renderizó e inspeccionó: ya no desborda la página, pero sigue siendo un esqueleto lógico, no una reproducción fiel del formato de FCR108_B. El RDF binario real FCR118_E se convirtió, validó, previsualizó y compiló en `tmp/studio_actual_rdf`, con advertencia por la función de base de datos `MES_ANTERIOR`.
