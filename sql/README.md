# Instalación de `PKG_CORPORATE_REPORTS`

## Requisitos

- Oracle APEX 24.2.
- Una aplicación con impresión nativa disponible mediante `APEX_DATA_EXPORT`.
- Acceso al *parsing schema* de la aplicación para crear un package.
- Privilegios directos de lectura sobre los objetos consultados. No dependa de
  privilegios recibidos mediante roles para el código almacenado.

No se necesita Word, Python, LibreOffice ni un contenedor en el servidor. Esas
herramientas solo intervienen localmente al compilar una plantilla.

## Instalación con SQLcl o SQL*Plus

Conéctese como el *parsing schema*, cambie a esta carpeta y ejecute:

```sql
@install.sql
```

El instalador detiene la ejecución ante un error y consulta `USER_ERRORS` para
evitar aceptar silenciosamente un package inválido.

## Instalación desde SQL Workshop

`install.sql` usa la directiva `@@`, propia de SQLcl/SQL*Plus. En SQL Workshop:

1. abra **SQL Workshop > SQL Scripts**;
2. cargue `pkg_corporate_reports.sql`;
3. ejecute el script completo;
4. confirme que tanto el package como su body figuren como `VALID`;
5. revise los errores con:

```sql
select type, line, position, text
from user_errors
where name = 'PKG_CORPORATE_REPORTS'
order by sequence;
```

## Controles de seguridad obligatorios

- El package usa `AUTHID CURRENT_USER`; no lo cambie a derechos del definidor.
- No conceda `EXECUTE` sobre el package a `PUBLIC`.
- `p_sql_query`, los JSON de columnas/estilos y los mapeos son configuración de
  Informática. Nunca los reciba desde un Page Item, URL, JavaScript o archivo
  aportado por un usuario final.
- Los valores variables deben llegar mediante `p_bindings_json`; el package los
  añade con `APEX_EXEC.ADD_PARAMETER`. No concatene valores de sesión en el SQL.
- La validación `SELECT/WITH` reduce errores, pero no convierte SQL arbitrario en
  una barrera de autorización. Una consulta `SELECT` puede invocar funciones de
  base de datos. Revise el SQL generado antes de instalarlo.
- Otorgue al *parsing schema* solo los privilegios de lectura indispensables y
  use vistas controladas cuando el reporte abarque datos sensibles.
- Fije `p_max_rows` conforme al volumen aprobado. El límite técnico del paquete
  no sustituye límites funcionales más bajos.
- Mantenga el proceso de descarga protegido por la autorización de página o
  componente correspondiente.

## API nueva y compatibilidad

- `DOWNLOAD_QUERY`: ejecuta una única consulta declarativa con parámetros,
  columnas y estilos compilados desde Word.
- `DOWNLOAD_IG` y `DOWNLOAD_IG_PDF`: conservan el flujo anterior de Interactive
  Grid, incluidas exclusiones y spans posicionales.

`DOWNLOAD_QUERY` usa siempre papel `A4` y admite orientación `AUTO`, `PORTRAIT`
o `LANDSCAPE`. Si existe una única columna `AUTO`, el package aplica internamente
una reserva heurística fija de 20 % al distribuir el ancho restante. El tamaño de
papel y esa reserva no forman parte de la API pública.

## Comprobaciones que requieren APEX

La revisión local no puede compilar contra las especificaciones instaladas de
APEX ni generar un PDF real. Antes de promover a producción, pruebe en APEX 24.2:

1. compilación del package y body sin filas en `USER_ERRORS`;
2. descarga PDF y XLSX desde un proceso **Before Header**;
3. Page Items `VARCHAR2`, `NUMBER`, `DATE` y `TIMESTAMP`;
4. consulta sin filas, con el máximo permitido y con conversión inválida;
5. papel A4 en orientación `AUTO`, `PORTRAIT` y `LANDSCAPE`;
6. una columna `AUTO`, columnas fijas y ponderadas, y columnas excluidas;
7. cierre correcto de la petición después de `APEX_DATA_EXPORT.DOWNLOAD`.
