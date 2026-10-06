-- RPT_PDF: motor mínimo de PDF en PL/SQL puro (prototipo).
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
    PROCEDURE new_document(
        p_landscape      IN BOOLEAN DEFAULT FALSE,
        p_body_top_mm    IN NUMBER  DEFAULT 20,   -- donde empieza el cuerpo en cada página
        p_body_bottom_mm IN NUMBER  DEFAULT 20    -- margen inferior reservado al pie
    );
    PROCEDURE new_page;
    FUNCTION  page_count   RETURN PLS_INTEGER;
    FUNCTION  current_page RETURN PLS_INTEGER;
    PROCEDURE set_page(p_page IN PLS_INTEGER);    -- para dibujar encabezados/pies al final
    FUNCTION  page_width_mm  RETURN NUMBER;
    FUNCTION  page_height_mm RETURN NUMBER;
    FUNCTION  body_top_mm    RETURN NUMBER;
    FUNCTION  body_limit_mm  RETURN NUMBER;        -- última coordenada útil del cuerpo
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

    -- ---------------------------------------------------------- tablas
    TYPE t_column IS RECORD (
        heading    VARCHAR2(200),
        width_mm   NUMBER,
        align      VARCHAR2(1),     -- L, R, C
        format     VARCHAR2(60),    -- máscara TO_CHAR para números/fechas
        summary    VARCHAR2(10),    -- SUM o NULL (Summary Column de Oracle Reports)
        gap_mm     NUMBER           -- espacio libre antes de la columna
    );
    TYPE t_columns IS TABLE OF t_column;

    FUNCTION col(
        p_heading  IN VARCHAR2,
        p_width_mm IN NUMBER,
        p_align    IN VARCHAR2 DEFAULT 'L',
        p_format   IN VARCHAR2 DEFAULT NULL,
        p_summary  IN VARCHAR2 DEFAULT NULL,
        p_gap_mm   IN NUMBER   DEFAULT 0
    ) RETURN t_column;

    -- Imprime una tabla a partir de un cursor (una columna del cursor por
    -- columna declarada, en el mismo orden). Salta de página cuando no cabe,
    -- repite el título y la cabecera, y al final imprime los totales SUM.
    -- p_y entra con la posición actual y sale con la posición siguiente libre.
    PROCEDURE print_table(
        p_cursor     IN OUT SYS_REFCURSOR,
        p_columns    IN t_columns,
        p_y          IN OUT NUMBER,
        p_x          IN NUMBER   DEFAULT 15,
        p_title      IN VARCHAR2 DEFAULT NULL,
        p_title_w_mm IN NUMBER   DEFAULT 78,
        p_font_pt    IN NUMBER   DEFAULT 7,
        p_row_mm     IN NUMBER   DEFAULT 3.1
    );
END rpt_pdf;
/
