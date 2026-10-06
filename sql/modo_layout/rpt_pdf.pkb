CREATE OR REPLACE PACKAGE BODY rpt_pdf AS

    c_pt_per_mm CONSTANT NUMBER := 72 / 25.4;
    c_nls       CONSTANT VARCHAR2(40) := 'NLS_NUMERIC_CHARACTERS=''.,''';

    -- Anchos de Helvetica / Helvetica-Bold (WinAnsi 32..255, milésimas de em).
    -- La cursiva usa los mismos anchos que su variante recta.
    c_widths_regular CONSTANT VARCHAR2(1000) :=
        '278,278,355,556,556,889,667,191,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,' ||
        '556,278,278,584,584,584,556,1015,667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,667,' ||
        '778,722,667,611,722,667,944,667,667,611,278,278,278,469,556,333,556,556,500,556,556,278,556,556,222,' ||
        '222,500,222,833,556,556,556,556,333,500,278,556,500,722,500,500,500,334,260,334,584,0,556,0,222,556,' ||
        '333,1000,556,556,333,1000,667,333,1000,0,611,0,0,222,222,333,333,350,556,1000,333,1000,500,333,944,' ||
        '0,500,667,278,333,556,556,556,556,260,556,333,737,370,556,584,333,737,552,400,549,333,333,333,576,' ||
        '537,333,333,333,365,556,834,834,834,611,667,667,667,667,667,667,1000,722,667,667,667,667,278,278,' ||
        '278,278,722,722,778,778,778,778,778,584,778,722,722,722,722,667,667,611,556,556,556,556,556,556,889,' ||
        '500,556,556,556,556,278,278,278,278,556,556,556,556,556,556,556,549,611,556,556,556,556,500,556,500';
    c_widths_bold CONSTANT VARCHAR2(1000) :=
        '278,333,474,556,556,889,722,238,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,' ||
        '556,333,333,584,584,584,611,975,722,722,722,722,667,611,778,722,278,556,722,611,833,722,778,667,778,' ||
        '722,667,611,722,667,944,667,667,611,333,278,333,584,556,333,556,611,556,611,556,333,611,611,278,278,' ||
        '556,278,889,611,611,611,611,389,556,333,611,556,778,556,556,500,389,280,389,584,0,556,0,278,556,500,' ||
        '1000,556,556,333,1000,667,333,1000,0,611,0,0,278,278,500,500,350,556,1000,333,1000,556,333,944,0,' ||
        '500,667,278,333,556,556,556,556,280,556,333,737,370,556,584,333,737,552,400,549,333,333,333,576,556,' ||
        '333,333,333,365,556,834,834,834,611,722,722,722,722,722,722,1000,722,667,667,667,667,278,278,278,' ||
        '278,722,722,778,778,778,778,778,584,778,722,722,722,722,667,667,611,556,556,556,556,556,556,889,556,' ||
        '556,556,556,556,278,278,278,278,611,611,611,611,611,611,611,549,611,611,611,611,611,556,611,556';

    TYPE t_widths IS TABLE OF PLS_INTEGER INDEX BY PLS_INTEGER;
    TYPE t_blobs  IS TABLE OF BLOB INDEX BY PLS_INTEGER;

    g_w_regular   t_widths;
    g_w_bold      t_widths;
    g_pages       t_blobs;
    g_page_count  PLS_INTEGER := 0;
    g_current     PLS_INTEGER := 0;
    g_width_pt    NUMBER;
    g_height_pt   NUMBER;
    g_font        PLS_INTEGER := c_regular;
    g_size        NUMBER := 8;
    g_text_rgb    VARCHAR2(6) := '000000';

    -- ------------------------------------------------------------ utilidades
    PROCEDURE load_widths(p_list IN VARCHAR2, p_target IN OUT NOCOPY t_widths) IS
        l_code PLS_INTEGER := 32;
        l_pos  PLS_INTEGER := 1;
        l_next PLS_INTEGER;
    BEGIN
        LOOP
            l_next := INSTR(p_list, ',', l_pos);
            p_target(l_code) := TO_NUMBER(SUBSTR(p_list, l_pos,
                CASE WHEN l_next = 0 THEN LENGTH(p_list) + 1 ELSE l_next END - l_pos));
            EXIT WHEN l_next = 0;
            l_pos  := l_next + 1;
            l_code := l_code + 1;
        END LOOP;
    END load_widths;

    FUNCTION num(p_value IN NUMBER) RETURN VARCHAR2 IS
    BEGIN
        RETURN TO_CHAR(ROUND(p_value, 3), 'FM9999990.999', c_nls);
    END num;

    FUNCTION pt(p_mm IN NUMBER) RETURN VARCHAR2 IS
    BEGIN
        RETURN num(p_mm * c_pt_per_mm);
    END pt;

    FUNCTION y_pt(p_mm IN NUMBER) RETURN VARCHAR2 IS
    BEGIN
        RETURN num(g_height_pt - p_mm * c_pt_per_mm);
    END y_pt;

    FUNCTION rgb(p_hex IN VARCHAR2) RETURN VARCHAR2 IS
    BEGIN
        RETURN num(TO_NUMBER(SUBSTR(p_hex, 1, 2), 'XX') / 255) || ' ' ||
               num(TO_NUMBER(SUBSTR(p_hex, 3, 2), 'XX') / 255) || ' ' ||
               num(TO_NUMBER(SUBSTR(p_hex, 5, 2), 'XX') / 255);
    END rgb;

    FUNCTION to_ansi(p_text IN VARCHAR2) RETURN RAW IS
    BEGIN
        RETURN UTL_I18N.STRING_TO_RAW(p_text, 'WE8MSWIN1252');
    END to_ansi;

    PROCEDURE append(p_lob IN OUT NOCOPY BLOB, p_text IN VARCHAR2) IS
        l_raw RAW(32767) := to_ansi(p_text);
    BEGIN
        IF l_raw IS NOT NULL THEN
            DBMS_LOB.WRITEAPPEND(p_lob, UTL_RAW.LENGTH(l_raw), l_raw);
        END IF;
    END append;

    PROCEDURE out(p_text IN VARCHAR2) IS
    BEGIN
        IF g_current = 0 THEN
            RAISE_APPLICATION_ERROR(-20700, 'RPT_PDF: llame a NEW_DOCUMENT antes de dibujar.');
        END IF;
        append(g_pages(g_current), p_text || CHR(10));
    END out;

    FUNCTION escape_text(p_text IN VARCHAR2) RETURN VARCHAR2 IS
    BEGIN
        RETURN REPLACE(REPLACE(REPLACE(p_text, '\', '\\'), '(', '\('), ')', '\)');
    END escape_text;

    -- ------------------------------------------------------------ documento
    PROCEDURE new_document(p_landscape IN BOOLEAN DEFAULT FALSE) IS
    BEGIN
        FOR i IN 1 .. g_page_count LOOP
            IF DBMS_LOB.ISTEMPORARY(g_pages(i)) = 1 THEN
                DBMS_LOB.FREETEMPORARY(g_pages(i));
            END IF;
        END LOOP;
        g_pages.DELETE;
        g_page_count := 0;
        g_current    := 0;
        g_width_pt   := 210 * c_pt_per_mm;
        g_height_pt  := 297 * c_pt_per_mm;
        IF p_landscape THEN
            g_width_pt  := 297 * c_pt_per_mm;
            g_height_pt := 210 * c_pt_per_mm;
        END IF;
        IF g_w_regular.COUNT = 0 THEN
            load_widths(c_widths_regular, g_w_regular);
            load_widths(c_widths_bold, g_w_bold);
        END IF;
        new_page;
    END new_document;

    PROCEDURE new_page IS
        l_page BLOB;
    BEGIN
        DBMS_LOB.CREATETEMPORARY(l_page, TRUE);
        g_page_count := g_page_count + 1;
        g_pages(g_page_count) := l_page;
        g_current := g_page_count;
    END new_page;

    FUNCTION page_count RETURN PLS_INTEGER IS BEGIN RETURN g_page_count; END;
    FUNCTION current_page RETURN PLS_INTEGER IS BEGIN RETURN g_current; END;
    FUNCTION page_width_mm RETURN NUMBER IS BEGIN RETURN g_width_pt / c_pt_per_mm; END;
    FUNCTION page_height_mm RETURN NUMBER IS BEGIN RETURN g_height_pt / c_pt_per_mm; END;

    PROCEDURE set_page(p_page IN PLS_INTEGER) IS
    BEGIN
        IF p_page NOT BETWEEN 1 AND g_page_count THEN
            RAISE_APPLICATION_ERROR(-20701, 'RPT_PDF: página inexistente ' || p_page || '.');
        END IF;
        g_current := p_page;
    END set_page;

    -- ------------------------------------------------------------ dibujo
    PROCEDURE set_font(p_style IN PLS_INTEGER, p_size_pt IN NUMBER) IS
    BEGIN
        g_font := p_style;
        g_size := p_size_pt;
    END set_font;

    PROCEDURE set_text_color(p_rgb IN VARCHAR2) IS
    BEGIN
        g_text_rgb := UPPER(p_rgb);
    END set_text_color;

    FUNCTION text_width_mm(p_text IN VARCHAR2) RETURN NUMBER IS
        l_raw   RAW(32767);
        l_total PLS_INTEGER := 0;
        l_code  PLS_INTEGER;
    BEGIN
        IF p_text IS NULL THEN
            RETURN 0;
        END IF;
        l_raw := to_ansi(p_text);
        FOR i IN 1 .. UTL_RAW.LENGTH(l_raw) LOOP
            l_code := TO_NUMBER(RAWTOHEX(UTL_RAW.SUBSTR(l_raw, i, 1)), 'XX');
            IF l_code >= 32 THEN
                l_total := l_total + CASE
                    WHEN g_font IN (c_bold, c_bold_italic) THEN g_w_bold(l_code)
                    ELSE g_w_regular(l_code)
                END;
            END IF;
        END LOOP;
        RETURN l_total * g_size / 1000 / c_pt_per_mm;
    END text_width_mm;

    PROCEDURE text(
        p_x     IN NUMBER,
        p_y     IN NUMBER,
        p_text  IN VARCHAR2,
        p_align IN VARCHAR2 DEFAULT 'L',
        p_width IN NUMBER   DEFAULT 0
    ) IS
        l_x NUMBER := p_x;
    BEGIN
        IF p_text IS NULL THEN
            RETURN;
        END IF;
        IF p_align = 'R' THEN
            l_x := p_x + p_width - text_width_mm(p_text);
        ELSIF p_align = 'C' THEN
            l_x := p_x + (p_width - text_width_mm(p_text)) / 2;
        END IF;
        out('BT /F' || g_font || ' ' || num(g_size) || ' Tf ' || rgb(g_text_rgb) || ' rg '
            || pt(l_x) || ' ' || y_pt(p_y) || ' Td (' || escape_text(p_text) || ') Tj ET');
    END text;

    PROCEDURE line(
        p_x1 IN NUMBER, p_y1 IN NUMBER, p_x2 IN NUMBER, p_y2 IN NUMBER,
        p_width_pt IN NUMBER DEFAULT 0.5, p_rgb IN VARCHAR2 DEFAULT '000000'
    ) IS
    BEGIN
        out('q ' || num(p_width_pt) || ' w ' || rgb(p_rgb) || ' RG '
            || pt(p_x1) || ' ' || y_pt(p_y1) || ' m ' || pt(p_x2) || ' ' || y_pt(p_y2) || ' l S Q');
    END line;

    PROCEDURE rect(
        p_x IN NUMBER, p_y IN NUMBER, p_w IN NUMBER, p_h IN NUMBER,
        p_fill     IN VARCHAR2 DEFAULT NULL,
        p_stroke   IN VARCHAR2 DEFAULT '000000',
        p_width_pt IN NUMBER   DEFAULT 0.5
    ) IS
        l_op VARCHAR2(1);
    BEGIN
        l_op := CASE
            WHEN p_fill IS NOT NULL AND p_stroke IS NOT NULL THEN 'B'
            WHEN p_fill IS NOT NULL THEN 'f'
            WHEN p_stroke IS NOT NULL THEN 'S'
        END;
        IF l_op IS NULL THEN
            RETURN;
        END IF;
        out('q ' || num(p_width_pt) || ' w '
            || CASE WHEN p_fill IS NOT NULL THEN rgb(p_fill) || ' rg ' END
            || CASE WHEN p_stroke IS NOT NULL THEN rgb(p_stroke) || ' RG ' END
            || pt(p_x) || ' ' || y_pt(p_y + p_h) || ' ' || pt(p_w) || ' ' || pt(p_h) || ' re ' || l_op || ' Q');
    END rect;

    -- ------------------------------------------------------------ salida
    FUNCTION get_pdf RETURN BLOB IS
        TYPE t_offsets IS TABLE OF PLS_INTEGER INDEX BY PLS_INTEGER;
        l_pdf     BLOB;
        l_offsets t_offsets;
        l_objects PLS_INTEGER := 6 + g_page_count * 2;
        l_kids    VARCHAR2(32767);
        l_xref    PLS_INTEGER;
        l_fonts   CONSTANT SYS.ODCIVARCHAR2LIST := SYS.ODCIVARCHAR2LIST(
            'Helvetica', 'Helvetica-Bold', 'Helvetica-Oblique', 'Helvetica-BoldOblique');

        PROCEDURE begin_object(p_number IN PLS_INTEGER) IS
        BEGIN
            l_offsets(p_number) := DBMS_LOB.GETLENGTH(l_pdf);
            append(l_pdf, p_number || ' 0 obj' || CHR(10));
        END begin_object;
    BEGIN
        DBMS_LOB.CREATETEMPORARY(l_pdf, TRUE);
        DBMS_LOB.WRITEAPPEND(l_pdf, 15, HEXTORAW('255044462D312E340A25E2E3CFD30A'));  -- %PDF-1.4 + binario

        begin_object(1);
        append(l_pdf, '<< /Type /Catalog /Pages 2 0 R >>' || CHR(10) || 'endobj' || CHR(10));

        FOR i IN 1 .. g_page_count LOOP
            l_kids := l_kids || (5 + i * 2) || ' 0 R ';
        END LOOP;
        begin_object(2);
        append(l_pdf, '<< /Type /Pages /Kids [' || l_kids || '] /Count ' || g_page_count
            || ' >>' || CHR(10) || 'endobj' || CHR(10));

        FOR i IN 1 .. 4 LOOP
            begin_object(2 + i);
            append(l_pdf, '<< /Type /Font /Subtype /Type1 /BaseFont /' || l_fonts(i)
                || ' /Encoding /WinAnsiEncoding >>' || CHR(10) || 'endobj' || CHR(10));
        END LOOP;

        FOR i IN 1 .. g_page_count LOOP
            begin_object(5 + i * 2);
            append(l_pdf, '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ' || num(g_width_pt) || ' '
                || num(g_height_pt) || '] /Resources << /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R /F4 6 0 R >> >>'
                || ' /Contents ' || (6 + i * 2) || ' 0 R >>' || CHR(10) || 'endobj' || CHR(10));
            begin_object(6 + i * 2);
            append(l_pdf, '<< /Length ' || DBMS_LOB.GETLENGTH(g_pages(i)) || ' >>' || CHR(10)
                || 'stream' || CHR(10));
            IF DBMS_LOB.GETLENGTH(g_pages(i)) > 0 THEN
                DBMS_LOB.APPEND(l_pdf, g_pages(i));
            END IF;
            append(l_pdf, CHR(10) || 'endstream' || CHR(10) || 'endobj' || CHR(10));
        END LOOP;

        l_xref := DBMS_LOB.GETLENGTH(l_pdf);
        append(l_pdf, 'xref' || CHR(10) || '0 ' || (l_objects + 1) || CHR(10)
            || '0000000000 65535 f ' || CHR(10));
        FOR i IN 1 .. l_objects LOOP
            append(l_pdf, LPAD(l_offsets(i), 10, '0') || ' 00000 n ' || CHR(10));
        END LOOP;
        append(l_pdf, 'trailer' || CHR(10) || '<< /Size ' || (l_objects + 1) || ' /Root 1 0 R >>'
            || CHR(10) || 'startxref' || CHR(10) || l_xref || CHR(10) || '%%EOF' || CHR(10));
        RETURN l_pdf;
    END get_pdf;

END rpt_pdf;
/
