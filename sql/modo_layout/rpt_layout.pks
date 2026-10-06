-- RPT_LAYOUT: intérprete del layout compilado desde Word (modo layout).
-- Lee la descripción JSON que genera el compilador, ejecuta las consultas del
-- reporte con binds tipados y dibuja el PDF con RPT_PDF. Se instala una vez.
CREATE OR REPLACE PACKAGE rpt_layout AUTHID CURRENT_USER AS

    TYPE t_bind IS RECORD (
        name     VARCHAR2(128),
        kind     VARCHAR2(1),          -- T texto, N número, D fecha
        v_text   VARCHAR2(4000),
        v_number NUMBER,
        v_date   DATE
    );
    TYPE t_binds IS TABLE OF t_bind;

    FUNCTION bind_text(p_name IN VARCHAR2, p_value IN VARCHAR2) RETURN t_bind;
    FUNCTION bind_number(p_name IN VARCHAR2, p_value IN NUMBER) RETURN t_bind;
    FUNCTION bind_date(p_name IN VARCHAR2, p_value IN DATE) RETURN t_bind;

    -- p_layout : layout compilado (zonas header, body, footer y página).
    -- p_queries: {"NOMBRE": {"sql": "...", "binds": ["P_X", ...]}, ...}
    -- p_values : textos fijos {"REPORT_TITLE": "...", "APP_USER": "...", "<CONSTANTE>": "..."}
    FUNCTION render(
        p_layout  IN CLOB,
        p_queries IN CLOB,
        p_values  IN CLOB,
        p_binds   IN t_binds
    ) RETURN BLOB;
END rpt_layout;
/
