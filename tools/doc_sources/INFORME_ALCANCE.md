# Informe de alcance y limitaciones

## Compilador local Word → reportes Oracle APEX 24.2

**Versión evaluada:** 1.0  
**Decisión técnica:** Word como editor local restringido; APEX como motor de ejecución y exportación  
**Dependencias nuevas en el servidor:** ninguna

## 1. Resumen ejecutivo

El proyecto proporciona una alternativa interna para generar reportes corporativos tabulares sin adquirir licencias de AOP o BI Publisher y sin instalar Microsoft Word, LibreOffice, Python ni Docker en la VM de APEX/ORDS.

Microsoft Word se usa únicamente para expresar una plantilla visual restringida. Un compilador local inspecciona el DOCX, valida su seguridad y traduce los elementos compatibles a configuración declarativa. En producción, un único paquete PL/SQL ejecuta una consulta SQL revisada mediante `APEX_EXEC` y genera PDF o XLSX con `APEX_DATA_EXPORT`.

La solución no convierte Word a PDF. Su alcance deliberado es un título, una consulta tabular, una tabla y un pie, con formato corporativo uniforme. Esta frontera reduce dependencias, superficie de ataque y mantenimiento, a cambio de no reproducir documentos Word arbitrarios.

## 2. Objetivos alcanzados

- Autoría visual básica en Word para personal de Informática.
- Un compilador local con CLI e interfaz gráfica, abrible con un acceso directo portable, que indica tras compilar qué subir a APEX y dónde, y que puede generar un proyecto inicial desde el DOCX con Page Items `PXX_`.
- Consulta SQL de solo lectura con parámetros tipados enlazados a Page Items.
- Configuración por reporte sin crear un package PL/SQL distinto para cada diseño.
- PDF y XLSX mediante capacidades nativas de APEX 24.2.
- Exclusiones y anchos de columnas configurados por alias.
- Resolución explícita de display values de LOV mediante SQL.
- Generación automática del proyecto inicial (`.report.json` y `.sql`) desde los marcadores del DOCX, y de los artefactos de APEX en cada compilación.
- Validación defensiva de DOCX, SQL, configuración y consistencia de columnas.
- Generación determinista y transaccional de artefactos.

## 3. Arquitectura final

### 3.1 Diseño local

Informática administra cuatro elementos por reporte: plantilla DOCX, consulta SQL, archivo `.report.json` y resultados compilados. El compilador produce:

| Artefacto | Propósito |
|---|---|
| template.json | Definición canónica de plantilla, columnas y estilos |
| validation.json | Errores, advertencias y trazabilidad de validación |
| apex_process.sql | Bloque mínimo que llama al paquete común |

El DOCX nunca se interpreta en APEX. El compilador no ejecuta la consulta y no necesita conectarse a Oracle.

```flujo
# Figura 1. Arquitectura: fase local (equipo de Informática) y fase de ejecución (APEX)
Equipo local | Plantilla Word (.docx) | única fuente del diseño
Equipo local | Proyecto .report.json y consulta .sql | creados automáticamente desde el DOCX; editables a mano
Equipo local | Compilador | valida y genera automáticamente template.json, validation.json y apex_process.sql
Repositorio | Versionado de fuentes y artefactos | el DOCX no sale del repositorio
Base de datos | PKG_CORPORATE_REPORTS | instalado una vez por esquema
APEX | Proceso de descarga | contiene apex_process.sql
APEX | APEX_EXEC + APEX_DATA_EXPORT | PDF o XLSX, sin dependencias nuevas en el servidor
```

La automatización cubre todo lo que se deduce de la plantilla: al cambiar el Word basta recompilar para que la definición y el proceso de APEX se regeneren. El manifiesto y la consulta pueden editarse directamente en cualquier momento; el manifiesto solo necesita cambios cuando la plantilla incorpora información que no está en el documento, como un campo `{{FIELD:...}}` nuevo o el ancho de una columna.

### 3.2 Ejecución en APEX

1. La página somete los filtros autorizados al estado de sesión.
2. El proceso invoca `PKG_CORPORATE_REPORTS.DOWNLOAD_QUERY`.
3. El paquete lee los Page Items declarados y convierte tipos.
4. `APEX_EXEC` abre el contexto de la consulta con parámetros enlazados.
5. El paquete valida y configura las columnas.
6. `APEX_DATA_EXPORT` genera y descarga PDF o XLSX.
7. Los contextos se cierran tanto en éxito como ante error.

El papel de salida es A4. La orientación pública puede ser `AUTO`, `PORTRAIT` o `LANDSCAPE`.

## 4. Alcance funcional

### 4.1 Incluido

- Una consulta `SELECT` o `WITH ... SELECT` por reporte.
- Un título centrado.
- Campos escalares controlados en encabezado o pie.
- Una tabla principal de 1 a 50 columnas dinámicas.
- Un pie textual, recomendado para usuario y fecha.
- Fuentes equivalentes a Helvetica, Times y Courier.
- Tamaños, negrita, colores, fondos, alineación y bordes simples.
- Binds `VARCHAR2`, `NUMBER`, `DATE` y `TIMESTAMP`.
- Page Items globales o de página.
- Máximo configurable de filas.
- Columnas excluidas por alias.
- Anchos `FIXED_PERCENT`, `WEIGHT` y una columna `AUTO`.
- Formatos PDF y XLSX dentro de las capacidades de APEX.
- Modo estricto y modo compatible del compilador.

### 4.2 Excluido

- Conversión fiel de DOCX a PDF.
- Word o LibreOffice automatizado en servidor.
- Imágenes, logotipos, gráficos, formas y WordArt.
- Varias tablas, tablas anidadas, celdas combinadas o subreportes.
- Maestros-detalles y múltiples datasets independientes.
- Macros, VBA, campos ejecutables y expresiones dentro de Word.
- PL/SQL arbitrario como fuente de datos.
- DML, DDL o SQL construido desde texto aportado por usuarios.
- Compatibilidad completa con Oracle Reports, BI Publisher o AOP.
- Deducción automática del display value de LOV desde metadatos de la página APEX.
- Edición de plantillas por usuarios finales.
- Papel seleccionable: el papel es siempre A4.

## 5. Contrato de autoría Word

La plantilla contiene una sola sección, un párrafo de título con `{{REPORT_TITLE}}`, una tabla de dos filas y un pie opcional. La primera fila contiene etiquetas literales; la segunda, exactamente un `{{COLUMN:ALIAS}}` por celda.

Los marcadores adicionales son `{{FIELD:NOMBRE}}`, `{{APP_USER}}` y `{{GENERATED_AT}}`. No existe lenguaje de programación en el DOCX. Las plantillas de referencia usan los cinco tipos:

```
{{REPORT_TITLE}}                                  <- título del reporte
Unidad: {{FIELD:UNIDAD}}                          <- campo escalar (fields)
{{COLUMN:VDNI}} | {{COLUMN:VNOM}} | ...           <- fila 2 de la tabla
Usuario: {{APP_USER}} | Fecha: {{GENERATED_AT}}   <- pie de página
```

El compilador rechaza contenido OOXML peligroso o ambiguo, incluidos macros, relaciones externas, `customXml`, objetos incrustados, campos Word, múltiples tablas y traversal de rutas dentro del ZIP.

## 6. Contrato de datos

El SQL es código administrado por Informática. Los valores variables no alteran el texto SQL: se añaden como parámetros de `APEX_EXEC` a partir de mappings explícitos en `bindings`.

La firma pública es:

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

El usuario se obtiene del contexto de APEX, el papel es A4 y la heurística interna de `AUTO` no es configurable públicamente.

## 7. LOV y valores mostrados

Una consulta SQL independiente no conoce cómo la página muestra un LOV. El proyecto evita depender de metadatos internos de APEX: el SQL debe hacer el `JOIN` o consultar una vista de dominio y devolver el texto que se desea imprimir.

```
select e.departamento_id,
       d.descripcion as departamento
  from entidad e
  join departamento d on d.id = e.departamento_id
```

Este enfoque es explícito, comprobable en SQL Workshop y estable ante cambios de la interfaz. La plantilla referencia `DEPARTAMENTO`, no el ID.

## 8. Anchos de columnas

Los anchos se configuran por alias:

| Modo | Semántica |
|---|---|
| FIXED_PERCENT | Porcentaje solicitado del ancho imprimible |
| WEIGHT | Participación relativa en el espacio restante |
| AUTO | Sin p_width; APEX decide el ajuste |

Solo se permite una columna `AUTO`. Si una columna no se configura, se trata como `WEIGHT` 1. La reserva utilizada internamente para calcular el resto es heurística; no garantiza que APEX otorgue a `AUTO` una fracción exacta. Fuentes, contenido y saltos de línea pueden obligar ajustes.

## 9. Modelo de seguridad

### 9.1 Fronteras de confianza

- DOCX, SQL y JSON son artefactos administrados por Informática.
- El usuario final solo aporta valores mediante Page Items autorizados.
- Ningún usuario final puede suministrar el texto SQL, estilos, columnas o plantilla.

### 9.2 Controles incorporados

- Inspección defensiva del contenedor ZIP/OOXML antes de analizar Word.
- Rechazo de macros, contenido externo y estructuras no admitidas.
- SQL limitado a una consulta y revisión adicional en servidor.
- Valores enlazados, nunca concatenados.
- Validación de nombres de binds, items, columnas, tipos y formatos.
- Rechazo de la sustitución `&ITEM.` en el SQL y emisión de `&` como `chr(38)`, para que APEX no incruste valores de sesión en el código del proceso.
- Rechazo de nombres reservados de APEX como binds lógicos.
- Límites de longitud medidos en bytes UTF-8, igual que los tipos del package.
- Lista cerrada de estilos, fuentes, colores y orientaciones.
- Límite de filas y tamaños de entrada.
- Saneamiento del nombre de archivo.
- `AUTHID CURRENT_USER` y cierre de contextos ante excepción.
- Errores funcionales que evitan revelar valores sensibles.

### 9.3 Riesgo residual

Una consulta `SELECT` puede ser costosa o acceder a datos que el esquema puede leer. La validación sintáctica no sustituye privilegios mínimos, vistas seguras, revisión por pares, autorización APEX, índices y pruebas de rendimiento. El package no debe exponerse a `PUBLIC`.

## 10. Limitaciones técnicas

### 10.1 Fidelidad visual

APEX_DATA_EXPORT no es Word. La herramienta traduce un conjunto limitado de propiedades y puede aproximar medidas. El resultado debe validarse en APEX con datos reales.

### 10.2 Estado de sesión

Un valor visible en el navegador que no se haya enviado al servidor no está disponible para un bind. La página debe someter los items antes de iniciar la descarga.

### 10.3 Validación de base de datos

La compilación local valida contrato y coherencia, pero no comprueba objetos, sinónimos, permisos ni la especificación exacta instalada de APEX. Esa verificación requiere el ambiente APEX 24.2.

### 10.4 Volumen

Los límites reales dependen de Oracle, ORDS, memoria, complejidad SQL y configuración de la aplicación. Cada reporte debe tener `p_max_rows`, consultas indexables y pruebas con volumen representativo. `p_max_rows` limita la salida; si existen filas adicionales, la exportación queda truncada a ese máximo y no se promete un error de desbordamiento.

### 10.5 XLSX

PDF es el formato visual principal. Algunos atributos de página o estilo no tienen un equivalente idéntico en XLSX. En particular, XLSX no garantiza encabezado o pie de página, orientación ni anchos idénticos a los observados en el PDF.

## 11. Operación y mantenimiento

- Instalar un solo paquete común por esquema.
- Versionar plantilla, SQL, configuración y resultados compilados.
- Revisar cada cambio de SQL y manifest como código.
- Recompilar después de cualquier cambio de diseño o contrato; los artefactos se regeneran automáticamente.
- Conservar el último resultado válido; la compilación fallida no lo sustituye.
- Distribuir juntos el compilador y el package de una misma entrega.

Generar un package distinto por plantilla no se recomienda: duplicaría seguridad, validaciones y correcciones, y elevaría el costo operativo.

## 12. Verificación realizada y pendiente

La verificación local incluye pruebas unitarias e integrales del compilador, casos de seguridad OOXML, parsing SQL, determinismo, preservación de la última salida válida y contrato estático del package. Las plantillas de ejemplo se generan de forma reproducible y se revisan visualmente.

Quedan necesariamente a cargo del ambiente APEX:

- compilar el package y body contra APEX 24.2;
- generar PDF y XLSX reales;
- validar cero filas, máximo de filas y conversiones erróneas;
- confirmar orientación y anchos con datos representativos;
- confirmar session state y autorizaciones.

## 13. Criterios de aceptación en APEX

| Área | Criterio |
|---|---|
| Instalación | Package y body VALID, sin filas en USER_ERRORS |
| Seguridad | Usuario no autorizado no puede ejecutar la descarga |
| Datos | Filtros tipados funcionan sin concatenación |
| LOV | PDF muestra display value, no return value |
| Diseño | Título y campos centrados; usuario y fecha en pie; tabla legible |
| Columnas | Orden, exclusiones y anchos coinciden con configuración |
| Formatos | PDF y XLSX descargables y válidos |
| Límites | Nunca se exportan más filas que p_max_rows; la entrada inválida produce error controlado |
## 14. Conclusión

La arquitectura propuesta satisface el objetivo de reportes corporativos tabulares con una experiencia de autoría accesible para Informática, sin incorporar un motor de documentos al servidor ni pagar una licencia de impresión. Su sostenibilidad depende de respetar la frontera: Word define una plantilla restringida, SQL define datos, el compilador traduce y un paquete común controla ejecución y exportación.

La solución es apropiada cuando se acepta uniformidad corporativa y una única tabla. Para documentos contractuales complejos, varias secciones, gráficos, imágenes o fidelidad exacta a Word, será necesario evaluar un motor de documentos especializado.
