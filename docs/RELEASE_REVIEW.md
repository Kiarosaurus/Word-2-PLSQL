# Revisión independiente de liberación

## Compilador Word restringido para Oracle APEX 24.2

**Fecha de revisión:** 2026-10-01  
**Versión revisada:** 1.0.0  
**Veredicto:** **GO para empaquetado y prueba de aceptación en APEX**  
**Producción:** pendiente; no debe promoverse hasta completar las pruebas indicadas en la sección 7.

## 1. Resumen ejecutivo

No quedan hallazgos locales abiertos que bloqueen la entrega del candidato 1.0.0. El compilador Python, el contrato publicado y la operación `PKG_CORPORATE_REPORTS.DOWNLOAD_QUERY` coinciden en los elementos que pueden verificarse sin una instancia Oracle: nombres y tipos de argumentos, JSON emitidos, límites, placeholders, binds, estilos, alineaciones, exclusiones, anchos, orientación, A4 y cierre del contexto.

La revisión no certifica que el package compile ni que el PDF/XLSX se renderice correctamente en la instalación real. Ese límite es material: no hubo conexión a Oracle Database/APEX 24.2 y las pruebas estáticas no sustituyen `USER_ERRORS` ni una descarga HTTP real.

## 2. Evidencia ejecutada

### 2.1 Suite automática

Comando ejecutado desde la raíz del proyecto:

```bash
PYTHONPATH=src python -m unittest discover -v
```

Resultado:

```text
Ran 92 tests in 10.397s
OK
```

La suite cubre:

- CLI y API de validación/compilación;
- compilación transaccional y preservación de la última salida válida;
- determinismo byte a byte;
- DOCX válido, estructura restringida y estilos;
- alineación no especificada conservada como `null` para inferencia por tipo en APEX;
- placeholders y gramática `FIELD`;
- seguridad ZIP/OOXML, relaciones externas, macros, ActiveX, objetos, imágenes, DTD y entidades;
- análisis SQL, binds, comentarios, literales Oracle y límites UTF-8;
- manifiesto, Page Items, campos, estilos, exclusiones y anchos;
- enteros JSON extremos devueltos como diagnóstico, sin excepción no controlada;
- contrato estático y compatibilidad de la API pública del package.

### 2.2 Ejemplo reproducible

Se validó y compiló `examples/entidades.report.json`. La validación produjo cero errores y cero advertencias. Los tres artefactos generados fueron idénticos a `examples/build_expected`:

```text
validation.json   idéntico
template.json     idéntico
apex_process.sql  idéntico
```

También se ejecutó `python -m compileall` sobre `src`, `tests`, `tools` y `launch_compiler.py`, sin errores de sintaxis Python.

## 3. Trazabilidad compilador → runtime

| Contrato aceptado localmente | Emisión Python | Consumo en `DOWNLOAD_QUERY` | Resultado |
|---|---|---|---|
| SQL `SELECT`/`WITH`, máximo 32 767 bytes | `l_sql VARCHAR2(32767)` en fragmentos seguros | validación y `APEX_EXEC.OPEN_QUERY_CONTEXT` | Conforme |
| Binds lógicos tipados | JSON `query.bindings` | `APEX_SESSION_STATE` + `APEX_EXEC.ADD_PARAMETER` | Conforme |
| Sin auto-bind implícito | llamada generada no incrusta valores | `p_auto_bind_items => FALSE` | Conforme |
| Campos `ITEM`, `CONTEXT`, `SYSTEM`, `CONSTANT` | `p_fields_json` | resolución restringida de header/footer | Conforme |
| `REPORT_TITLE`, `APP_USER`, `GENERATED_AT`, `FIELD` | texto validado | parser de marcadores equivalente | Conforme |
| 1 a 50 columnas | `p_columns_json` | validación por nombre contra `APEX_EXEC` | Conforme |
| Alineación no definida | JSON `null` | número a la derecha; otros tipos a la izquierda | Conforme |
| Exclusiones por alias | array de nombres | exclusión posterior a validar el contrato | Conforme |
| `FIXED_PERCENT`, `WEIGHT`, una `AUTO` | JSON por nombre | presupuesto 99 %, reserva AUTO interna de 20 % | Conforme |
| Columna sin ancho configurado | `WEIGHT` 1.0 | distribución ponderada | Conforme |
| A4; `AUTO`, `PORTRAIT`, `LANDSCAPE` | orientación compilada o Page Item autorizado | A4 fijo y orientación validada | Conforme |
| PDF/XLSX | valor fijo o Page Item | lista cerrada de formatos | Conforme |
| Cierre de recursos | no aplica | cierre antes de descargar y en excepciones | Conforme estáticamente |

## 4. Revisión de seguridad

No se encontró concatenación de valores de Page Items en el SQL. Los nombres de bind, columnas y Page Items se restringen a gramáticas cerradas. La consulta, los JSON técnicos y la plantilla compilada se consideran configuración administrada por Informática, no entrada del usuario final.

Controles presentes:

- `AUTHID CURRENT_USER`;
- SQL y configuración separados de los valores de sesión;
- `p_auto_bind_items => FALSE`;
- binds tipados mediante `APEX_EXEC.ADD_PARAMETER`;
- límite de filas entre 1 y 100 000, con recomendación de usar límites institucionales menores;
- inspección defensiva del contenedor DOCX antes de abrirlo con `python-docx`;
- rechazo de rutas externas al proyecto y de elementos OOXML activos o ambiguos;
- mensajes de conversión que identifican el bind sin reproducir su valor sensible;
- documentación explícita de que un `SELECT` puede invocar funciones y debe revisarse como código.

Riesgo residual: limitar el primer token a `SELECT`/`WITH` no es una frontera de autorización. La protección real depende de revisión de código, autorización de página y privilegios mínimos del esquema.

## 5. Compatibilidad con las APIs de APEX 24.2

Se contrastaron las llamadas principales con la referencia oficial de Oracle APEX 24.2:

- `APEX_EXEC.OPEN_QUERY_CONTEXT` admite `p_sql_parameters`, `p_auto_bind_items`, `p_first_row` y `p_max_rows`:  
  <https://docs.oracle.com/en/database/oracle/apex/24.2/aeapi/APEX_EXEC.OPEN_QUERY_CONTEXT-Function-1.html>
- `APEX_DATA_EXPORT.ADD_COLUMN` admite encabezado, máscara, alineaciones y ancho:  
  <https://docs.oracle.com/en/database/oracle/apex/24.2/aeapi/APEX_DATA_EXPORT-ADD_COLUMN-Procedure.html>
- `APEX_DATA_EXPORT.DOWNLOAD` admite `p_content_disposition` y `p_add_file_extension`:  
  <https://docs.oracle.com/en/database/oracle/apex/24.2/aeapi/APEX_DATA_EXPORT-DOWNLOAD-Procedure.html>
- `APEX_EXEC` 24.2 declara `c_data_type_number` y `c_data_type_binary_number`, utilizados para inferir alineación numérica:  
  <https://docs.oracle.com/en/database/oracle/apex/24.2/aeapi/APEX_EXEC.Global-Constants.html>

Esta comparación reduce el riesgo de una firma evidentemente incompatible; solo la compilación en la instancia objetivo confirma sin ambigüedad los sinónimos, grants y especificaciones instaladas.

## 6. Hallazgos cerrados durante la auditoría

| Hallazgo | Riesgo | Cierre verificado |
|---|---|---|
| Alineación Word no especificada se convertía en `START` | anulaba la inferencia numérica del runtime | ahora se conserva como `null`; pruebas unitarias e integral |
| Enteros JSON extremos podían propagar `OverflowError` | caída de CLI/GUI fuera del sistema de diagnósticos | conversión finita defensiva; prueba de constantes, estilos y anchos extremos |
| Columnas sin `column_widths` heredaban proporciones del DOCX | contradicción con el contrato `WEIGHT = 1` | emisor cambiado a 1.0; prueba integral con anchos Word desiguales |

## 7. Puerta obligatoria en APEX real

Antes de producción, el responsable APEX debe registrar evidencia de:

1. ejecución de `sql/pkg_corporate_reports.sql` en el parsing schema;
2. package y package body en estado `VALID`, sin filas en `USER_ERRORS`;
3. descarga PDF y XLSX desde un proceso **Before Header**;
4. prueba de binds `VARCHAR2`, `NUMBER`, `DATE` y `TIMESTAMP`, incluidos nulos, obligatorios y conversiones inválidas;
5. cero filas, una fila, volumen representativo y truncamiento por `p_max_rows`;
6. orientación `AUTO`, `PORTRAIT` y `LANDSCAPE` en A4;
7. anchos fijos, ponderados, una columna `AUTO` intermedia y exclusiones;
8. LOV impreso mediante el display value devuelto por `JOIN` o vista;
9. usuario sin autorización incapaz de ejecutar la descarga;
10. regresión de `DOWNLOAD_IG` y `DOWNLOAD_IG_PDF` en páginas existentes.

Un fallo en cualquiera de los puntos 1 a 3 cambia el veredicto operativo a **NO-GO** hasta corregirlo.

## 8. Condiciones de empaquetado

El ZIP de entrega debe incluir `src`, `tests`, `sql`, `templates`, `examples`, `deliverables`, los documentos finales, `README.md`, `pyproject.toml`, `requirements.txt` y el iniciador.

Debe excluir:

- `build/`, porque contiene copias históricas que pueden quedar desactualizadas respecto de `src`;
- `qa/` y salidas temporales de renderizado;
- `__pycache__`, `.pyc`, `.egg-info` y entornos virtuales;
- borradores internos (`SCOPE_DRAFT`, `REVIEW_ROUND_0`, `PACKAGE_AUDIT`, `TEST_STRATEGY`, `COMPILER_DESIGN`) que no forman parte del contrato final.

El hash del ZIP y una ejecución de la suite desde una extracción limpia deben conservarse como evidencia de liberación.

## 9. Conclusión

El candidato es coherente y reproducible en la frontera verificable localmente. Se autoriza su empaquetado y entrega para aceptación. No se autoriza afirmar compatibilidad productiva completa hasta compilar el package y descargar archivos reales en la instancia Oracle APEX 24.2 de destino.
