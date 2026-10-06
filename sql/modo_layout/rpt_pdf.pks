-- RPT_PDF: primitivas de dibujo de PDF en PL/SQL puro (las usa RPT_LAYOUT).
-- No usa librerías externas, servidores de impresión ni Java: escribe el PDF
-- byte a byte con las fuentes estándar (Helvetica) que todo lector trae.
-- Coordenadas en milímetros desde la esquina superior izquierda de la página.
CREATE OR REPLACE PACKAGE rpt_pdf AUTHID CURRENT_USER AS

    -- Estilos de fuente (Helvetica estándar, sin incrustar)
    c_regular     CONSTANT PLS_INTEGER := 1;
    c_bold        CONSTANT PLS_INTEGER := 2;
    c_italic      CONSTANT PLS_INTEGER := 3;
    c_bold_italic CONSTANT PLS_INTEGER := 4;

    -- ---------------------------------------------------------- documento
    PROCEDURE new_document(p_landscape IN BOOLEAN DEFAULT FALSE);
    PROCEDURE new_page;
    FUNCTION  page_count   RETURN PLS_INTEGER;
    FUNCTION  current_page RETURN PLS_INTEGER;
    PROCEDURE set_page(p_page IN PLS_INTEGER);    -- para dibujar encabezados/pies al final
    FUNCTION  page_width_mm  RETURN NUMBER;
    FUNCTION  page_height_mm RETURN NUMBER;
    FUNCTION  get_pdf RETURN BLOB;

    -- ---------------------------------------------------------- dibujo
    PROCEDURE set_font(p_style IN PLS_INTEGER, p_size_pt IN NUMBER);
    PROCEDURE set_text_color(p_rgb IN VARCHAR2);  -- 'RRGGBB'
    FUNCTION  text_width_mm(p_text IN VARCHAR2) RETURN NUMBER;
    -- p_y es la línea base del texto. p_align: L, R o C dentro de [p_x, p_x + p_width].
    PROCEDURE text(
        p_x     IN NUMBER,
        p_y     IN NUMBER,
        p_text  IN VARCHAR2,
        p_align IN VARCHAR2 DEFAULT 'L',
        p_width IN NUMBER   DEFAULT 0
    );
    PROCEDURE line(
        p_x1 IN NUMBER, p_y1 IN NUMBER, p_x2 IN NUMBER, p_y2 IN NUMBER,
        p_width_pt IN NUMBER DEFAULT 0.5, p_rgb IN VARCHAR2 DEFAULT '000000'
    );
    PROCEDURE rect(
        p_x IN NUMBER, p_y IN NUMBER, p_w IN NUMBER, p_h IN NUMBER,
        p_fill     IN VARCHAR2 DEFAULT NULL,       -- NULL = sin relleno
        p_stroke   IN VARCHAR2 DEFAULT '000000',   -- NULL = sin borde
        p_width_pt IN NUMBER   DEFAULT 0.5
    );
END rpt_pdf;
/
