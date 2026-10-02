# Instalación de PKG_CORPORATE_REPORTS

## Requisitos

- Oracle APEX 24.2.
- Una aplicación con impresión nativa disponible mediante `APEX_DATA_EXPORT`.
- Acceso al *parsing schema* de la aplicación para crear un package.
- Privilegios directos de lectura sobre los objetos consultados. No dependa de privilegios recibidos mediante roles para el código almacenado.

No se necesita Word, Python, LibreOffice ni un contenedor en el servidor. Esas herramientas solo intervienen localmente al compilar una plantilla.

## Instalación con SQLcl o SQL*Plus

Conéctese como el *parsing schema*, cambie a esta carpeta y ejecute:

```
-- Ejemplo con SQLcl, conectado como el parsing schema de la aplicación
sql APP_SCHEMA@//servidor:1521/servicio
cd sql                       -- carpeta sql/ de este proyecto
@install.sql                 -- crea package y body, verifica USER_ERRORS y VALID
```

El instalador detiene la ejecución ante un error, consulta `USER_ERRORS` y exige que el package y su body figuren como `VALID` en `USER_OBJECTS`, para evitar aceptar silenciosamente un package inválido. Al iniciar muestra la versión del compilador con la que se distribuye (1.0.0); el package no expone su versión en ejecución, así que reinstálelo cuando cambie este archivo.

```flujo
# Figura 1. Instalación y verificación del package
DBA / Informática | Conectarse como el parsing schema | nunca como SYS ni con un usuario distinto del de la aplicación
? Instalación | ¿Se dispone de SQLcl o SQL*Plus? | No: SQL Workshop > SQL Scripts y ejecutar pkg_corporate_reports.sql
Instalación | Ejecutar @install.sql | CREATE OR REPLACE PACKAGE y PACKAGE BODY
? Verificación | ¿Package y body VALID, sin filas en USER_ERRORS? | No: revisar privilegios directos y la versión de APEX; reinstalar
Seguridad | Mantener AUTHID CURRENT_USER y no conceder EXECUTE a PUBLIC | una sola instalación sirve a todos los reportes
APEX | Pegar en cada página el apex_process.sql generado | se genera automáticamente al compilar cada reporte
```

## Instalación desde SQL Workshop

`install.sql` usa la directiva `@@`, propia de SQLcl/SQL*Plus. En SQL Workshop:

1. abra **SQL Workshop > SQL Scripts**;
2. cargue `pkg_corporate_reports.sql`;
3. ejecute el script completo;
4. confirme que tanto el package como su body figuren como `VALID`;
5. revise los errores con:

```
-- Errores de compilación: no debe devolver filas
select type, line, position, text
from user_errors
where name = 'PKG_CORPORATE_REPORTS'
order by sequence;

-- Estado: deben aparecer PACKAGE y PACKAGE BODY con status VALID
select object_type, status
from user_objects
where object_name = 'PKG_CORPORATE_REPORTS';
```

## Controles de seguridad obligatorios

- El package usa `AUTHID CURRENT_USER`; no lo cambie a derechos del definidor.
- No conceda `EXECUTE` sobre el package a `PUBLIC`.
- `p_sql_query`, los JSON de columnas/estilos y los mapeos son configuración de Informática. Nunca los reciba desde un Page Item, URL, JavaScript o archivo aportado por un usuario final.
- Los valores variables deben llegar mediante `p_bindings_json`; el package los añade con `APEX_EXEC.ADD_PARAMETER`. No concatene valores de sesión en el SQL ni use la sustitución `&ITEM.`: el compilador la rechaza y emite cada `&` como `chr(38)`.
- Los nombres reservados de APEX (`APP_USER`, `APP_ID`, `REQUEST`…) no pueden usarse como binds lógicos; el package los rechaza con -20173. Para filtrar por el usuario autenticado use `SYS_CONTEXT('APEX$SESSION', 'APP_USER')`.
- La validación `SELECT/WITH` reduce errores, pero no convierte SQL arbitrario en una barrera de autorización. Una consulta `SELECT` puede invocar funciones de base de datos. Revise el SQL generado antes de instalarlo.
- Otorgue al *parsing schema* solo los privilegios de lectura indispensables y use vistas controladas cuando el reporte abarque datos sensibles.
- Fije `p_max_rows` conforme al volumen aprobado. El límite técnico del paquete no sustituye límites funcionales más bajos.
- Mantenga el proceso de descarga protegido por la autorización de página o componente correspondiente.

## API pública

`DOWNLOAD_QUERY` ejecuta una única consulta declarativa con parámetros, columnas y estilos compilados desde Word. Es la operación que invoca el `apex_process.sql` que el compilador genera automáticamente para cada reporte; no hace falta escribir la llamada a mano. Extracto comentado de un proceso generado:

```
declare
    l_sql clob := to_clob(q'~select ... where (:DNI is null or e.vdni = :DNI)~');
    -- l_bindings_json, l_fields_json, l_columns_json, l_style_json, ...
    -- se generan desde el .report.json y la plantilla Word
begin
    pkg_corporate_reports.download_query(
        p_sql_query       => l_sql,
        p_bindings_json   => l_bindings_json,   -- :DNI <- P42_DNI (VARCHAR2)
        p_header_template => q'~{{REPORT_TITLE}}
Unidad: {{FIELD:UNIDAD}}~',                     -- párrafo de título del Word
        p_footer_template => q'~Usuario: {{APP_USER}} | Fecha: {{GENERATED_AT}}~',
        p_format          => coalesce(:P0_REPORT_FORMAT, 'PDF'),
        ...
    );
end;
```

`DOWNLOAD_QUERY` usa siempre papel `A4` y admite orientación `AUTO`, `PORTRAIT` o `LANDSCAPE`. Si existe una única columna `AUTO`, el package aplica internamente una reserva heurística fija de 20 % al distribuir el ancho restante. El tamaño de papel y esa reserva no forman parte de la API pública.

## Comprobaciones que requieren APEX

La revisión local no puede compilar contra las especificaciones instaladas de APEX ni generar un PDF real. Antes de promover a producción, pruebe en APEX 24.2:

1. compilación del package y body sin filas en `USER_ERRORS`;
2. descarga PDF y XLSX desde el proceso descrito en el Manual de uso, sección 9 (botón *Submit Page*, *Reload on Submit: Always* y proceso en *Processing*), y desde un proceso **Before Header** si la aplicación usa ese patrón;
3. Page Items `VARCHAR2`, `NUMBER`, `DATE` y `TIMESTAMP`;
4. consulta sin filas, con el máximo permitido y con conversión inválida;
5. papel A4 en orientación `AUTO`, `PORTRAIT` y `LANDSCAPE`;
6. una columna `AUTO`, columnas fijas y ponderadas, y columnas excluidas;
7. cierre correcto de la petición después de `APEX_DATA_EXPORT.DOWNLOAD`.
