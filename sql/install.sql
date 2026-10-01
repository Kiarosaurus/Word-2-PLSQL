-- Instalador para SQLcl o SQL*Plus.
-- Ejecútelo conectado como el parsing schema de la aplicación APEX.
WHENEVER SQLERROR EXIT SQL.SQLCODE ROLLBACK

SET DEFINE OFF
SET SERVEROUTPUT ON

PROMPT Instalando PKG_CORPORATE_REPORTS...
@@pkg_corporate_reports.sql

SHOW ERRORS PACKAGE pkg_corporate_reports
SHOW ERRORS PACKAGE BODY pkg_corporate_reports

DECLARE
    l_error_count PLS_INTEGER;
BEGIN
    SELECT COUNT(*)
      INTO l_error_count
      FROM user_errors
     WHERE name = 'PKG_CORPORATE_REPORTS'
       AND type IN ('PACKAGE', 'PACKAGE BODY');

    IF l_error_count > 0 THEN
        raise_application_error(
            -20000,
            'PKG_CORPORATE_REPORTS quedó con errores de compilación. ' ||
            'Consulte USER_ERRORS.'
        );
    END IF;
END;
/

PROMPT PKG_CORPORATE_REPORTS instalado sin errores registrados.
