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
    -- p_queries: {"NOMBRE": {"sql": "...", "binds": ["P_X", ...], "columns": [...],
    --             "formulas": ["CF_X", ...], "filters": [{"c": "HIJA", "op": "eq", "p": "PADRE"}]}, ...}
    -- p_values : textos fijos {"REPORT_TITLE": "...", "APP_USER": "...", "<CONSTANTE>": "..."}
    -- p_model  : código convertido de Oracle Reports (opcional):
    --            {"package": "RPT_X", "formulas": {"CF_X": {"q": "CONSULTA", "t": "N", "f": "cf_xformula"}},
    --             "placeholders": {"CP_X": {"t": "T", "by": ["CF_X"]}},
    --             "summaries": {"CS_X": {"q": "CONSULTA", "s": "COLUMNA", "f": "sum"}}}
    FUNCTION render(
        p_layout  IN CLOB,
        p_queries IN CLOB,
        p_values  IN CLOB,
        p_binds   IN t_binds,
        p_model   IN CLOB DEFAULT NULL
    ) RETURN BLOB;

    -- Valores del reporte para el código convertido de Oracle Reports (equivalen a :NOMBRE):
    -- columnas de la fila actual, parámetros, fórmulas, marcadores de posición y totales.
    FUNCTION num(p_name IN VARCHAR2) RETURN NUMBER;
    FUNCTION txt(p_name IN VARCHAR2) RETURN VARCHAR2;
    FUNCTION dat(p_name IN VARCHAR2) RETURN DATE;
    -- Asignación a un marcador de posición (:CP_X := valor).
    PROCEDURE set_num(p_name IN VARCHAR2, p_value IN NUMBER);
    PROCEDURE set_txt(p_name IN VARCHAR2, p_value IN VARCHAR2);
    PROCEDURE set_dat(p_name IN VARCHAR2, p_value IN DATE);
END rpt_layout;
/
