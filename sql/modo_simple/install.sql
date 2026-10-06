-- Motor del modo simple (PKG_CORPORATE_REPORTS): se instala una sola vez por parsing
-- schema y lo usan todos los reportes del modo simple.
-- Instalador para SQLcl o SQL*Plus, desde esta carpeta: @install.sql
-- Ejecútelo conectado como el parsing schema de la aplicación APEX.
WHENEVER SQLERROR EXIT SQL.SQLCODE ROLLBACK

SET DEFINE OFF
SET SERVEROUTPUT ON

PROMPT Instalando PKG_CORPORATE_REPORTS (compilador 1.0.0)...
@@pkg_corporate_reports.sql

SHOW ERRORS PACKAGE pkg_corporate_reports
SHOW ERRORS PACKAGE BODY pkg_corporate_reports

DECLARE
    l_error_count PLS_INTEGER;
    l_valid_count PLS_INTEGER;
BEGIN
    SELECT COUNT(*)
      INTO l_error_count
      FROM user_errors
     WHERE name = 'PKG_CORPORATE_REPORTS'
       AND type IN ('PACKAGE', 'PACKAGE BODY')
       AND attribute = 'ERROR';

    SELECT COUNT(*)
      INTO l_valid_count
      FROM user_objects
     WHERE object_name = 'PKG_CORPORATE_REPORTS'
       AND object_type IN ('PACKAGE', 'PACKAGE BODY')
       AND status = 'VALID';

    IF l_error_count > 0 OR l_valid_count <> 2 THEN
        raise_application_error(
            -20000,
            'PKG_CORPORATE_REPORTS quedó con errores de compilación. ' ||
            'Consulte USER_ERRORS.'
        );
    END IF;
END;
/

SELECT object_type, status, last_ddl_time
  FROM user_objects
 WHERE object_name = 'PKG_CORPORATE_REPORTS'
 ORDER BY object_type;

PROMPT PKG_CORPORATE_REPORTS instalado: package y body VALID, sin errores registrados.
