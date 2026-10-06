# Revisión de liberación

## Compilador Word restringido para Oracle APEX 24.2

**Fecha de revisión:** 2026-10-01  
**Versión revisada:** 1.0.0 (compilador y package `PKG_CORPORATE_REPORTS` de esta entrega)  
**Veredicto local:** **GO para empaquetado y aceptación en APEX**  
**Producción:** pendiente; no debe promoverse hasta completar las pruebas de la sección 7 en APEX 24.2.

## 1. Resumen ejecutivo

La auditoría se hizo en tres rondas con revisores independientes de arquitectura, compilador Python, seguridad OOXML, SQL y binds, PL/SQL y APEX, pruebas adversariales y documentación. Cada hallazgo se reprodujo antes de corregirlo y cada corrección tiene una prueba de regresión que falla sin ella. Las pruebas se ejecutaron en Windows, la plataforma del diseñador.

```flujo
# Figura 1. Proceso de revisión aplicado a cada hallazgo
Revisor | Reproducir el hallazgo | caso mínimo en tests/test_audit_regressions.py
? Revisor | ¿La prueba falla sin la corrección? | No: la prueba no demuestra el defecto; reescribirla
Desarrollo | Corregir compilador, package o documentación | el mismo criterio en Python y en PL/SQL
? Revisor | ¿Suite completa en verde? | No: volver a corregir
Revisor | Registrar en la sección 3 | severidad, hallazgo y corrección
```

La revisión no certifica que el package compile ni que el PDF o el XLSX se rendericen en la instalación real: no hubo conexión a Oracle Database ni a APEX.

## 2. Evidencia ejecutada

Desde la raíz del proyecto, en Windows 11, con Python 3.13.2 y python-docx 1.2.0:

```
.\.venv\Scripts\python.exe -m unittest discover -s tests
.\.venv\Scripts\python.exe -m compileall -q src tests tools launch_compiler.py launch_gui.pyw
.\.venv\Scripts\apex-report-compiler.exe --version
```

| Puerta | Resultado |
|---|---|
| Suite completa | todas las pruebas OK |
| compileall | sin errores |
| --version | apex-report-compiler 1.0.0 |
| Validación del ejemplo | 0 errores, 0 advertencias, sin escribir archivos |
| Compilación del ejemplo, dos veces | los tres artefactos son idénticos byte a byte entre sí y con los versionados en proyectos/entidades |
| Generador de plantillas | reproduce templates/*.docx byte a byte |
| ZIP de entrega | pruebas y compilación del ejemplo correctas desde una extracción limpia; SHA-256 en el archivo .sha256 adjunto |

Pruebas nuevas:

- `tests/test_audit_regressions.py`: una o más pruebas por hallazgo. Se verificó que fallan sin las correcciones y pasan con ellas.
- `tests/test_package_semantics.py`: contrato estático del package sobre el código efectivo, sin comentarios, con firmas ordenadas exactas. Se valida contra siete mutaciones semánticas: auto-bind comentado, cierre de contexto eliminado, valor por defecto ampliado, parámetros intercambiados, cierre ausente en `e_stop_apex_engine`, SQL dinámico y `LTRIM` de un solo argumento.
- `tests/test_apex_guide.py`: guía de publicación en APEX, esqueleto de proyecto desde DOCX, salida de la GUI y acceso directo portable.
- `tests/test_workspace.py`: carpeta `proyectos/<nombre>/`, copia del DOCX, reemplazo con `.bak` sin borrar otros archivos, salida dividida y confinamiento de rutas.
- `tests/test_documentation.py`: los entregables Word contienen las instrucciones críticas y no mencionan nombres ajenos a esta versión.

## 3. Hallazgos corregidos

| Severidad | Hallazgo | Corrección |
|---|---|---|
| ALTO | background_color en título o pie se aceptaba localmente y el package lo rechazaba (-20145) | Claves de estilo por zona idénticas a PARSE_STYLE |
| ALTO | Comentario inicial seguido de salto de línea: válido en local, -20104 en ejecución (LTRIM sin conjunto) | El package elimina espacio, tabulador, LF y CR; el compilador acepta solo esos caracteres |
| ALTO | Límites en caracteres en Python frente a VARCHAR2 en bytes en el package (título, etiquetas, encabezado, pie): ORA-06502 sin control | Límites en bytes UTF-8 en ambos lados, LENGTHB/SUBSTRB y comprobación de encabezado con título |
| ALTO | &ITEM. en el SQL o en textos: APEX lo sustituiría en el código del proceso, concatenando valores de sesión | Error fuera de literales; todo & se emite como chr(38) |
| ALTO | --output podía sobrescribir el SQL o el proyecto fuente | Diagnóstico IO-003 |
| ALTO | Pruebas estáticas del package con falsos positivos (código comentado o subcadenas) | Pruebas semánticas con mutaciones |
| ALTO | Receta APEX ambigua o contradictoria (Ajax frente a Before Header; Reload on Submit) | Receta concreta en el Manual y en la GUI |
| MEDIO | Nombre \ en el ZIP no detectado en Windows | Se inspecciona orig_filename; fixture corregido |
| MEDIO | DTD en UTF-16, [Content_Types].xml en mayúsculas, parte principal renombrada, tipos de relación renombrados | Detección con el parser, content types y tipos de relación |
| MEDIO | ZIP corrupto o con compresión no soportada provocaba una excepción | Diagnósticos ZIP-010/ZIP-011 |
| MEDIO | Texto oculto, envoltorios en línea y de bloque, símbolos, ecuaciones, notas, revisiones de formato y tablas en el encabezado o pie: Word los mostraba y el reporte no | Rechazo explícito |
| MEDIO | Espacio no ASCII dentro de {{…}} aceptado en local y rechazado en el package | Solo espacios ASCII |
| MEDIO | Tipos JSON no textuales convertidos en silencio ("title": null → «None») | PROJECT-103 |
| MEDIO | Errores de APEX_SESSION_STATE propagados sin sanear (posible valor en el mensaje) | Lista blanca de códigos propios; el resto pasa a -20121/-20132 |
| MEDIO | Binds con nombre reservado de APEX (:APP_USER desde un Page Item modificable) | PROJECT-017 y -20173 |
| MEDIO | Salida de la CLI con traza al redirigirla en Windows | stdout/stderr en UTF-8 |
| MEDIO | Publicación interrumpida podía mezclar artefactos | Rollback ante BaseException; el respaldo se conserva si la restauración falla |
| MEDIO | La GUI congelaba resultados viejos ante errores y compilaba en la carpeta de otro proyecto | Errores visibles y salida ligada al proyecto |
| BAJO | INTO/FOR UPDATE, claves duplicadas, JSON muy anidado, WEIGHT sin cota, alias duplicados, orientación por dimensiones, pgSz ausente, instalador sin comprobar VALID y otros | Corregidos con prueba |

## 4. Seguridad

No se concatenan valores de Page Items en el SQL. El package conserva `AUTHID CURRENT_USER`, `p_auto_bind_items => FALSE` y binds tipados con `APEX_EXEC.ADD_PARAMETER`. El código generado no contiene ningún `&` crudo. El DOCX se inspecciona antes de abrirlo con python-docx. Las rutas del proyecto quedan confinadas a su carpeta. Las instrucciones no conceden `EXECUTE` a `PUBLIC`.

Riesgo residual: restringir el primer token a `SELECT`/`WITH` no es una frontera de autorización. La protección real depende de revisar el SQL, de la autorización de página y de los privilegios mínimos del parsing schema.

## 5. Herramientas para el diseñador

- Acceso directo portable **Compilador de reportes APEX** y lanzador **Abrir compilador.bat**. El lanzador prepara `.venv` la primera vez en un equipo nuevo.
- La GUI genera un proyecto inicial desde el DOCX con Page Items `PXX_` y, al compilar, indica qué archivo subir a APEX, dónde, y qué Page Items crear.

## 6. Riesgos residuales y decisiones

- Un Page Item mal escrito en `bindings` devuelve NULL en ejecución. Con `required: false`, el filtro queda anulado y se exportan todas las filas. Verifíquelo en la aceptación.
- `TO_DATE` sin el modificador `FX` acepta años de dos dígitos con la máscara `YYYY`.
- Los tamaños de fuente decimales podrían depender del separador decimal de la sesión.
- El acceso directo requiere PowerShell de Windows. Si una política lo bloquea, se usa el `.bat`.

## 7. Puerta obligatoria en APEX real

Antes de producción, el responsable APEX debe registrar evidencia de:

1. ejecución de `sql/modo_simple/install.sql` (o del package en SQL Workshop) en el parsing schema;
2. package y body en estado `VALID`, sin filas de error en `USER_ERRORS`;
3. descarga PDF y XLSX con la receta del Manual, sección 9 (*Submit Page*, *Reload on Submit: Always*, proceso en *Processing*);
4. binds `VARCHAR2`, `NUMBER`, `DATE` y `TIMESTAMP`, con valores nulos, obligatorios y conversiones inválidas (mensaje sin el valor);
5. cero filas, una fila y muchas filas (todas se exportan, sin límite);
6. orientación `AUTO`, `PORTRAIT` y `LANDSCAPE` en A4;
7. anchos fijos, ponderados, una columna `AUTO` intermedia y exclusiones;
8. LOV impreso con el display value;
9. un usuario sin autorización no puede descargar;
10. un SQL con comentario inicial y salto de línea, y un título con letras acentuadas y `&`;
11. un Page Item inexistente mapeado en `bindings`, para observar el comportamiento;
12. título con `{{FIELD:...}}` y pie con `{{APP_USER}}` y `{{GENERATED_AT}}` resueltos en el PDF.

Un fallo en los puntos 1 a 3 cambia el veredicto operativo a **NO-GO**.

## 8. Condiciones de empaquetado

El ZIP incluye `src`, `tests`, `sql`, `templates`, `proyectos`, `deliverables`, `docs`, `tools`, la documentación Word, `README.docx`, `pyproject.toml`, `requirements.txt`, los lanzadores y el acceso directo. Excluye `build/`, `qa/`, `.venv/`, `__pycache__`, `.pyc`, `.egg-info` y las salidas compiladas locales.
