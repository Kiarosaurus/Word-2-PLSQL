-- Motor PDF del modo layout: se instala una sola vez por parsing schema.
-- SQLcl o SQL*Plus desde esta carpeta: @install_motor.sql
-- SQL Workshop > SQL Scripts: ejecute los cuatro archivos en este orden.
WHENEVER SQLERROR EXIT FAILURE
PROMPT RPT_PDF (dibujo de PDF en PL/SQL)...
@@rpt_pdf.pks
@@rpt_pdf.pkb
PROMPT RPT_LAYOUT (intérprete del layout compilado desde Word)...
@@rpt_layout.pks
@@rpt_layout.pkb
SHOW ERRORS
SELECT object_name, object_type, status
  FROM user_objects
 WHERE object_name IN ('RPT_PDF', 'RPT_LAYOUT')
 ORDER BY object_name, object_type;
