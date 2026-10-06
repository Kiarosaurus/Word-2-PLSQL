CREATE OR REPLACE PACKAGE BODY rpt_layout AS

    c_nls      CONSTANT VARCHAR2(40) := 'NLS_NUMERIC_CHARACTERS=''.,''';
    c_pt_to_mm CONSTANT NUMBER := 25.4 / 72;

    TYPE t_value IS RECORD (
        kind VARCHAR2(1),
        t    VARCHAR2(4000),
        n    NUMBER,
        d    DATE
    );
    TYPE t_value_map IS TABLE OF t_value INDEX BY VARCHAR2(261);
    TYPE t_flag_map  IS TABLE OF BOOLEAN INDEX BY VARCHAR2(128);

    g_queries  JSON_OBJECT_T;
    g_values   JSON_OBJECT_T;
    g_binds    t_binds;
    g_scalars  t_value_map;      -- "CONSULTA.COLUMNA", parámetros y constantes
    g_loaded   t_flag_map;       -- consultas escalares ya leídas
    g_row      t_value_map;      -- fila actual de una tabla de datos
    g_sums     t_value_map;      -- totales de la tabla de datos actual
    g_page     PLS_INTEGER := 1;
    g_pages    PLS_INTEGER := 1;
    g_now      DATE;
    g_top      NUMBER;
    g_limit    NUMBER;
    g_left     NUMBER;
    g_width    NUMBER;
    g_y        NUMBER;

    -- ------------------------------------------------------------ binds
    FUNCTION bind_text(p_name IN VARCHAR2, p_value IN VARCHAR2) RETURN t_bind IS
        l t_bind;
    BEGIN
        l.name := UPPER(p_name); l.kind := 'T'; l.v_text := p_value;
        RETURN l;
    END bind_text;

    FUNCTION bind_number(p_name IN VARCHAR2, p_value IN NUMBER) RETURN t_bind IS
        l t_bind;
    BEGIN
        l.name := UPPER(p_name); l.kind := 'N'; l.v_number := p_value;
        RETURN l;
    END bind_number;

    FUNCTION bind_date(p_name IN VARCHAR2, p_value IN DATE) RETURN t_bind IS
        l t_bind;
    BEGIN
        l.name := UPPER(p_name); l.kind := 'D'; l.v_date := p_value;
        RETURN l;
    END bind_date;

    -- ------------------------------------------------------------ valores
    FUNCTION fmt(p_value IN t_value, p_mask IN VARCHAR2) RETURN VARCHAR2 IS
    BEGIN
        IF p_value.kind = 'N' THEN
            IF p_value.n IS NULL THEN RETURN NULL; END IF;
            RETURN CASE WHEN p_mask IS NULL THEN TO_CHAR(p_value.n, 'TM9', c_nls)
                        ELSE TRIM(TO_CHAR(p_value.n, p_mask, c_nls)) END;
        ELSIF p_value.kind = 'D' THEN
            RETURN TO_CHAR(p_value.d, NVL(p_mask, 'DD/MM/YYYY'));
        END IF;
        RETURN p_value.t;
    END fmt;

    FUNCTION text_value(p_text IN VARCHAR2) RETURN t_value IS
        l t_value;
    BEGIN
        l.kind := 'T'; l.t := p_text;
        RETURN l;
    END text_value;

    FUNCTION open_query(p_name IN VARCHAR2) RETURN INTEGER IS
        l_query  JSON_OBJECT_T;
        l_names  JSON_ARRAY_T;
        l_cursor INTEGER;
        l_found  BOOLEAN;
        l_ignore INTEGER;
    BEGIN
        IF NOT g_queries.has(p_name) THEN
            RAISE_APPLICATION_ERROR(-20730, 'RPT_LAYOUT: la consulta ' || p_name || ' no está definida.');
        END IF;
        l_query  := g_queries.get_object(p_name);
        l_cursor := DBMS_SQL.OPEN_CURSOR;
        DBMS_SQL.PARSE(l_cursor, l_query.get_clob('sql'), DBMS_SQL.NATIVE);
        l_names := l_query.get_array('binds');
        FOR i IN 0 .. l_names.get_size - 1 LOOP
            l_found := FALSE;
            FOR j IN 1 .. g_binds.COUNT LOOP
                IF g_binds(j).name = UPPER(l_names.get_string(i)) THEN
                    l_found := TRUE;
                    IF g_binds(j).kind = 'N' THEN
                        DBMS_SQL.BIND_VARIABLE(l_cursor, ':' || g_binds(j).name, g_binds(j).v_number);
                    ELSIF g_binds(j).kind = 'D' THEN
                        DBMS_SQL.BIND_VARIABLE(l_cursor, ':' || g_binds(j).name, g_binds(j).v_date);
                    ELSE
                        DBMS_SQL.BIND_VARIABLE(l_cursor, ':' || g_binds(j).name, g_binds(j).v_text, 4000);
                    END IF;
                END IF;
            END LOOP;
            IF NOT l_found THEN
                DBMS_SQL.CLOSE_CURSOR(l_cursor);
                RAISE_APPLICATION_ERROR(-20731, 'RPT_LAYOUT: falta el parámetro ' || l_names.get_string(i)
                    || ' de la consulta ' || p_name || '.');
            END IF;
        END LOOP;
        l_ignore := DBMS_SQL.EXECUTE(l_cursor);
        RETURN l_cursor;
    END open_query;

    PROCEDURE define_columns(p_cursor IN INTEGER, p_desc OUT DBMS_SQL.DESC_TAB3, p_count OUT INTEGER) IS
        l_number NUMBER;
        l_date   DATE;
        l_text   VARCHAR2(4000);
    BEGIN
        DBMS_SQL.DESCRIBE_COLUMNS3(p_cursor, p_count, p_desc);
        FOR i IN 1 .. p_count LOOP
            IF p_desc(i).col_type IN (2, 100, 101) THEN
                DBMS_SQL.DEFINE_COLUMN(p_cursor, i, l_number);
            ELSIF p_desc(i).col_type IN (12, 180, 181, 231) THEN
                DBMS_SQL.DEFINE_COLUMN(p_cursor, i, l_date);
            ELSE
                DBMS_SQL.DEFINE_COLUMN(p_cursor, i, l_text, 4000);
            END IF;
        END LOOP;
    END define_columns;

    -- Lee la fila actual en p_target con claves p_prefix || COLUMNA.
    PROCEDURE read_row(
        p_cursor IN INTEGER, p_desc IN DBMS_SQL.DESC_TAB3, p_count IN INTEGER,
        p_prefix IN VARCHAR2, p_target IN OUT NOCOPY t_value_map
    ) IS
        l t_value;
    BEGIN
        FOR i IN 1 .. p_count LOOP
            l := NULL;
            IF p_desc(i).col_type IN (2, 100, 101) THEN
                l.kind := 'N'; DBMS_SQL.COLUMN_VALUE(p_cursor, i, l.n);
            ELSIF p_desc(i).col_type IN (12, 180, 181, 231) THEN
                l.kind := 'D'; DBMS_SQL.COLUMN_VALUE(p_cursor, i, l.d);
            ELSE
                l.kind := 'T'; DBMS_SQL.COLUMN_VALUE(p_cursor, i, l.t);
            END IF;
            p_target(p_prefix || UPPER(p_desc(i).col_name)) := l;
        END LOOP;
    END read_row;

    -- Consulta escalar (grupo maestro): se lee una vez, solo la primera fila.
    PROCEDURE load_scalar_query(p_name IN VARCHAR2) IS
        l_cursor INTEGER;
        l_desc   DBMS_SQL.DESC_TAB3;
        l_count  INTEGER;
    BEGIN
        IF g_loaded.EXISTS(p_name) THEN
            RETURN;
        END IF;
        g_loaded(p_name) := TRUE;
        l_cursor := open_query(p_name);
        define_columns(l_cursor, l_desc, l_count);
        IF DBMS_SQL.FETCH_ROWS(l_cursor) > 0 THEN
            read_row(l_cursor, l_desc, l_count, p_name || '.', g_scalars);
        ELSE
            FOR i IN 1 .. l_count LOOP
                g_scalars(p_name || '.' || UPPER(l_desc(i).col_name)) := text_value(NULL);
            END LOOP;
        END IF;
        DBMS_SQL.CLOSE_CURSOR(l_cursor);
    EXCEPTION
        WHEN OTHERS THEN
            IF DBMS_SQL.IS_OPEN(l_cursor) THEN DBMS_SQL.CLOSE_CURSOR(l_cursor); END IF;
            RAISE;
    END load_scalar_query;

    FUNCTION lookup(p_map IN t_value_map, p_key IN VARCHAR2, p_marker IN VARCHAR2) RETURN t_value IS
    BEGIN
        IF NOT p_map.EXISTS(p_key) THEN
            RAISE_APPLICATION_ERROR(-20732, 'RPT_LAYOUT: no hay valor para ' || p_marker
                || '; revise que la consulta devuelva esa columna.');
        END IF;
        RETURN p_map(p_key);
    END lookup;

    -- Sustituye las marcas {{...}} de un texto.
    FUNCTION resolve(p_text IN VARCHAR2) RETURN VARCHAR2 IS
        l_result VARCHAR2(32767);
        l_pos    PLS_INTEGER := 1;
        l_open   PLS_INTEGER;
        l_close  PLS_INTEGER;
        l_token  VARCHAR2(400);
        l_mask   VARCHAR2(200);
        l_bar    PLS_INTEGER;
        l_name   VARCHAR2(261);
        l_dot    PLS_INTEGER;
    BEGIN
        IF p_text IS NULL OR INSTR(p_text, '{{') = 0 THEN
            RETURN p_text;
        END IF;
        LOOP
            l_open := INSTR(p_text, '{{', l_pos);
            EXIT WHEN l_open = 0;
            l_close := INSTR(p_text, '}}', l_open);
            EXIT WHEN l_close = 0;
            l_result := l_result || SUBSTR(p_text, l_pos, l_open - l_pos);
            l_token := SUBSTR(p_text, l_open + 2, l_close - l_open - 2);
            l_bar := INSTR(l_token, '|');
            l_mask := NULL;
            IF l_bar > 0 THEN
                l_mask  := SUBSTR(l_token, l_bar + 1);
                l_token := SUBSTR(l_token, 1, l_bar - 1);
            END IF;
            l_token := UPPER(TRIM(l_token));
            IF l_token = 'PAGE' THEN
                l_result := l_result || g_page;
            ELSIF l_token = 'PAGES' THEN
                l_result := l_result || g_pages;
            ELSIF l_token = 'GENERATED_AT' THEN
                l_result := l_result || TO_CHAR(g_now, NVL(l_mask, 'DD/MM/YYYY HH24:MI'));
            ELSIF l_token IN ('REPORT_TITLE', 'APP_USER') THEN
                l_result := l_result || g_values.get_string(l_token);
            ELSIF l_token LIKE 'FIELD:%' THEN
                l_name := SUBSTR(l_token, 7);
                l_dot := INSTR(l_name, '.');
                IF l_dot > 0 THEN
                    load_scalar_query(SUBSTR(l_name, 1, l_dot - 1));
                END IF;
                l_result := l_result || fmt(lookup(g_scalars, l_name, '{{' || l_token || '}}'), l_mask);
            ELSIF l_token LIKE 'COLUMN:%' THEN
                l_name := SUBSTR(l_token, 8);
                l_result := l_result || fmt(lookup(g_row, SUBSTR(l_name, INSTR(l_name, '.') + 1),
                                                   '{{' || l_token || '}}'), l_mask);
            ELSIF l_token LIKE 'SUM:%' THEN
                l_name := SUBSTR(l_token, 5);
                l_result := l_result || fmt(lookup(g_sums, SUBSTR(l_name, INSTR(l_name, '.') + 1),
                                                   '{{' || l_token || '}}'), l_mask);
            ELSE
                RAISE_APPLICATION_ERROR(-20733, 'RPT_LAYOUT: marca desconocida {{' || l_token || '}}.');
            END IF;
            l_pos := l_close + 2;
        END LOOP;
        RETURN l_result || SUBSTR(p_text, l_pos);
    END resolve;

    -- ------------------------------------------------------------ medidas y dibujo
    FUNCTION line_height(p_line IN JSON_ARRAY_T, p_empty_size IN NUMBER) RETURN NUMBER IS
        l_max NUMBER := 0;
    BEGIN
        FOR i IN 0 .. p_line.get_size - 1 LOOP
            l_max := GREATEST(l_max, TREAT(p_line.get(i) AS JSON_OBJECT_T).get_number('s'));
        END LOOP;
        IF l_max = 0 THEN
            l_max := NVL(p_empty_size, 8);
        END IF;
        RETURN l_max * 1.2 * c_pt_to_mm;
    END line_height;

    -- Párrafo: devuelve su altura; dibuja si p_draw.
    FUNCTION paragraph(
        p_para IN JSON_OBJECT_T, p_x IN NUMBER, p_y IN NUMBER, p_w IN NUMBER, p_draw IN BOOLEAN
    ) RETURN NUMBER IS
        l_lines  JSON_ARRAY_T := p_para.get_array('ln');
        l_line   JSON_ARRAY_T;
        l_run    JSON_OBJECT_T;
        l_y      NUMBER := p_y + NVL(p_para.get_number('bf'), 0);
        l_h      NUMBER;
        l_total  NUMBER;
        l_x      NUMBER;
        TYPE t_texts IS TABLE OF VARCHAR2(32767) INDEX BY PLS_INTEGER;
        l_texts  t_texts;
    BEGIN
        FOR i IN 0 .. l_lines.get_size - 1 LOOP
            l_line := TREAT(l_lines.get(i) AS JSON_ARRAY_T);
            l_h := line_height(l_line, p_para.get_number('es'));
            IF p_draw AND l_line.get_size > 0 THEN
                l_total := 0;
                FOR j IN 0 .. l_line.get_size - 1 LOOP
                    l_run := TREAT(l_line.get(j) AS JSON_OBJECT_T);
                    l_texts(j) := resolve(l_run.get_string('t'));
                    rpt_pdf.set_font(l_run.get_number('f'), l_run.get_number('s'));
                    l_total := l_total + rpt_pdf.text_width_mm(l_texts(j));
                END LOOP;
                l_x := CASE p_para.get_string('al')
                    WHEN 'R' THEN p_x + p_w - l_total
                    WHEN 'C' THEN p_x + (p_w - l_total) / 2
                    ELSE p_x
                END;
                FOR j IN 0 .. l_line.get_size - 1 LOOP
                    l_run := TREAT(l_line.get(j) AS JSON_OBJECT_T);
                    rpt_pdf.set_font(l_run.get_number('f'), l_run.get_number('s'));
                    rpt_pdf.set_text_color(NVL(l_run.get_string('c'), '000000'));
                    rpt_pdf.text(l_x, l_y + l_h * 0.78, l_texts(j));
                    l_x := l_x + rpt_pdf.text_width_mm(l_texts(j));
                END LOOP;
            END IF;
            l_y := l_y + l_h;
        END LOOP;
        RETURN l_y + NVL(p_para.get_number('af'), 0) - p_y;
    END paragraph;

    FUNCTION paragraphs_height(p_paras IN JSON_ARRAY_T, p_w IN NUMBER) RETURN NUMBER IS
        l_total NUMBER := 0;
    BEGIN
        FOR i IN 0 .. p_paras.get_size - 1 LOOP
            l_total := l_total + paragraph(TREAT(p_paras.get(i) AS JSON_OBJECT_T), 0, 0, p_w, FALSE);
        END LOOP;
        RETURN l_total;
    END paragraphs_height;

    FUNCTION cell_x(p_table IN JSON_OBJECT_T, p_col IN PLS_INTEGER) RETURN NUMBER IS
        l_widths JSON_ARRAY_T := p_table.get_array('w');
        l_x      NUMBER := g_left + NVL(p_table.get_number('x'), 0);
    BEGIN
        FOR i IN 0 .. p_col - 1 LOOP
            l_x := l_x + l_widths.get_number(i);
        END LOOP;
        RETURN l_x;
    END cell_x;

    FUNCTION cell_w(p_table IN JSON_OBJECT_T, p_col IN PLS_INTEGER, p_span IN PLS_INTEGER) RETURN NUMBER IS
        l_widths JSON_ARRAY_T := p_table.get_array('w');
        l_w      NUMBER := 0;
    BEGIN
        FOR i IN p_col .. p_col + p_span - 1 LOOP
            l_w := l_w + l_widths.get_number(i);
        END LOOP;
        RETURN l_w;
    END cell_w;

    PROCEDURE border(p_spec IN JSON_ARRAY_T, p_x1 IN NUMBER, p_y1 IN NUMBER, p_x2 IN NUMBER, p_y2 IN NUMBER) IS
    BEGIN
        IF p_spec IS NOT NULL AND p_spec.get_size = 2 THEN
            rpt_pdf.line(p_x1, p_y1, p_x2, p_y2, p_spec.get_number(0), p_spec.get_string(1));
        END IF;
    END border;

    FUNCTION side(p_borders IN JSON_OBJECT_T, p_key IN VARCHAR2) RETURN JSON_ARRAY_T IS
    BEGIN
        IF p_borders IS NULL OR NOT p_borders.has(p_key) OR p_borders.get(p_key).is_null THEN
            RETURN NULL;
        END IF;
        RETURN p_borders.get_array(p_key);
    END side;

    -- Fila de una tabla (maquetación o datos): devuelve la altura; dibuja si p_draw.
    FUNCTION table_row(p_table IN JSON_OBJECT_T, p_row IN JSON_OBJECT_T, p_y IN NUMBER, p_draw IN BOOLEAN)
        RETURN NUMBER
    IS
        l_cells   JSON_ARRAY_T := p_row.get_array('cs');
        l_pad     JSON_ARRAY_T := p_table.get_array('pd');
        l_cell    JSON_OBJECT_T;
        l_borders JSON_OBJECT_T;
        l_h       NUMBER := 0;
        l_content NUMBER;
        l_x       NUMBER;
        l_w       NUMBER;
        l_y       NUMBER;
        l_rule    VARCHAR2(10) := NVL(p_row.get_string('hr'), 'auto');
    BEGIN
        FOR i IN 0 .. l_cells.get_size - 1 LOOP
            l_cell := TREAT(l_cells.get(i) AS JSON_OBJECT_T);
            l_w := cell_w(p_table, l_cell.get_number('c'), NVL(l_cell.get_number('sp'), 1))
                   - l_pad.get_number(0) - l_pad.get_number(2);
            l_h := GREATEST(l_h, paragraphs_height(l_cell.get_array('ps'), l_w)
                   + l_pad.get_number(1) + l_pad.get_number(3));
        END LOOP;
        IF l_rule = 'exact' THEN
            l_h := p_row.get_number('h');
        ELSIF l_rule = 'atLeast' THEN
            l_h := GREATEST(l_h, p_row.get_number('h'));
        END IF;
        IF NOT p_draw THEN
            RETURN l_h;
        END IF;

        FOR pass IN 1 .. 3 LOOP                 -- 1 rellenos, 2 textos, 3 bordes
            FOR i IN 0 .. l_cells.get_size - 1 LOOP
                l_cell := TREAT(l_cells.get(i) AS JSON_OBJECT_T);
                l_x := cell_x(p_table, l_cell.get_number('c'));
                l_w := cell_w(p_table, l_cell.get_number('c'), NVL(l_cell.get_number('sp'), 1));
                IF pass = 1 AND l_cell.has('fl') AND NOT l_cell.get('fl').is_null THEN
                    rpt_pdf.rect(l_x, p_y, l_w, l_h, l_cell.get_string('fl'), NULL);
                ELSIF pass = 2 THEN
                    l_content := paragraphs_height(l_cell.get_array('ps'), l_w);
                    l_y := p_y + l_pad.get_number(1);
                    IF l_cell.get_string('va') = 'C' THEN
                        l_y := p_y + (l_h - l_content) / 2;
                    ELSIF l_cell.get_string('va') = 'B' THEN
                        l_y := p_y + l_h - l_pad.get_number(3) - l_content;
                    END IF;
                    FOR k IN 0 .. l_cell.get_array('ps').get_size - 1 LOOP
                        l_y := l_y + paragraph(TREAT(l_cell.get_array('ps').get(k) AS JSON_OBJECT_T),
                            l_x + l_pad.get_number(0), l_y, l_w - l_pad.get_number(0) - l_pad.get_number(2), TRUE);
                    END LOOP;
                ELSIF pass = 3 AND l_cell.has('b') THEN
                    l_borders := l_cell.get_object('b');
                    border(side(l_borders, 't'), l_x, p_y, l_x + l_w, p_y);
                    border(side(l_borders, 'b'), l_x, p_y + l_h, l_x + l_w, p_y + l_h);
                    border(side(l_borders, 'l'), l_x, p_y, l_x, p_y + l_h);
                    border(side(l_borders, 'r'), l_x + l_w, p_y, l_x + l_w, p_y + l_h);
                END IF;
            END LOOP;
        END LOOP;
        RETURN l_h;
    END table_row;

    -- ------------------------------------------------------------ flujo del cuerpo
    PROCEDURE ensure(p_height IN NUMBER) IS
    BEGIN
        IF g_y + p_height > g_limit AND g_y > g_top + 0.01 THEN
            rpt_pdf.new_page;
            g_y := g_top;
        END IF;
    END ensure;

    PROCEDURE draw_rows(p_table IN JSON_OBJECT_T, p_rows IN JSON_ARRAY_T) IS
        l_row JSON_OBJECT_T;
        l_h   NUMBER;
    BEGIN
        IF p_rows IS NULL THEN RETURN; END IF;
        FOR i IN 0 .. p_rows.get_size - 1 LOOP
            l_row := TREAT(p_rows.get(i) AS JSON_OBJECT_T);
            l_h := table_row(p_table, l_row, g_y, FALSE);
            ensure(l_h);
            l_h := table_row(p_table, l_row, g_y, TRUE);
            g_y := g_y + l_h;
        END LOOP;
    END draw_rows;

    FUNCTION rows_height(p_table IN JSON_OBJECT_T, p_rows IN JSON_ARRAY_T) RETURN NUMBER IS
        l_total NUMBER := 0;
    BEGIN
        IF p_rows IS NULL THEN RETURN 0; END IF;
        FOR i IN 0 .. p_rows.get_size - 1 LOOP
            l_total := l_total + table_row(p_table, TREAT(p_rows.get(i) AS JSON_OBJECT_T), 0, FALSE);
        END LOOP;
        RETURN l_total;
    END rows_height;

    -- Tabla de datos: cabecera, una fila por registro de la consulta, totales.
    PROCEDURE data_table(p_table IN JSON_OBJECT_T) IS
        l_header JSON_ARRAY_T := p_table.get_array('hd');
        l_body   JSON_OBJECT_T := p_table.get_object('bd');
        l_footer JSON_ARRAY_T := p_table.get_array('ft');
        l_sums   JSON_ARRAY_T := p_table.get_array('sm');
        l_cursor INTEGER;
        l_desc   DBMS_SQL.DESC_TAB3;
        l_count  INTEGER;
        l_h      NUMBER;
        l_key    VARCHAR2(261);
        l_zero   t_value;
    BEGIN
        g_sums.DELETE;
        l_zero.kind := 'N'; l_zero.n := 0;
        FOR i IN 0 .. l_sums.get_size - 1 LOOP
            g_sums(UPPER(l_sums.get_string(i))) := l_zero;
        END LOOP;
        l_cursor := open_query(p_table.get_string('q'));
        define_columns(l_cursor, l_desc, l_count);

        g_row.DELETE;
        FOR i IN 1 .. l_count LOOP                -- fila vacía para medir antes de leer
            g_row(UPPER(l_desc(i).col_name)) := text_value(NULL);
        END LOOP;
        ensure(rows_height(p_table, l_header) + table_row(p_table, l_body, 0, FALSE));
        draw_rows(p_table, l_header);

        WHILE DBMS_SQL.FETCH_ROWS(l_cursor) > 0 LOOP
            g_row.DELETE;
            read_row(l_cursor, l_desc, l_count, NULL, g_row);
            FOR i IN 0 .. l_sums.get_size - 1 LOOP
                l_key := UPPER(l_sums.get_string(i));
                IF g_row.EXISTS(l_key) AND g_row(l_key).kind = 'N' THEN
                    g_sums(l_key).n := g_sums(l_key).n + NVL(g_row(l_key).n, 0);
                END IF;
            END LOOP;
            l_h := table_row(p_table, l_body, g_y, FALSE);
            IF g_y + l_h > g_limit THEN
                rpt_pdf.new_page;
                g_y := g_top;
                draw_rows(p_table, l_header);
            END IF;
            g_y := g_y + table_row(p_table, l_body, g_y, TRUE);
        END LOOP;
        DBMS_SQL.CLOSE_CURSOR(l_cursor);
        draw_rows(p_table, l_footer);
    EXCEPTION
        WHEN OTHERS THEN
            IF DBMS_SQL.IS_OPEN(l_cursor) THEN DBMS_SQL.CLOSE_CURSOR(l_cursor); END IF;
            RAISE;
    END data_table;

    PROCEDURE blocks(p_blocks IN JSON_ARRAY_T, p_flow IN BOOLEAN) IS
        l_block JSON_OBJECT_T;
        l_h     NUMBER;
    BEGIN
        FOR i IN 0 .. p_blocks.get_size - 1 LOOP
            l_block := TREAT(p_blocks.get(i) AS JSON_OBJECT_T);
            CASE l_block.get_string('k')
                WHEN 'P' THEN
                    l_h := paragraph(l_block, g_left, 0, g_width, FALSE);
                    IF p_flow THEN ensure(l_h); END IF;
                    g_y := g_y + paragraph(l_block, g_left, g_y, g_width, TRUE);
                WHEN 'G' THEN
                    IF p_flow THEN
                        draw_rows(l_block, l_block.get_array('rs'));
                    ELSE
                        FOR r IN 0 .. l_block.get_array('rs').get_size - 1 LOOP
                            g_y := g_y + table_row(l_block,
                                TREAT(l_block.get_array('rs').get(r) AS JSON_OBJECT_T), g_y, TRUE);
                        END LOOP;
                    END IF;
                WHEN 'D' THEN
                    data_table(l_block);
            END CASE;
        END LOOP;
    END blocks;

    FUNCTION blocks_height(p_blocks IN JSON_ARRAY_T) RETURN NUMBER IS
        l_block JSON_OBJECT_T;
        l_total NUMBER := 0;
    BEGIN
        FOR i IN 0 .. p_blocks.get_size - 1 LOOP
            l_block := TREAT(p_blocks.get(i) AS JSON_OBJECT_T);
            IF l_block.get_string('k') = 'P' THEN
                l_total := l_total + paragraph(l_block, 0, 0, g_width, FALSE);
            ELSIF l_block.get_string('k') = 'G' THEN
                l_total := l_total + rows_height(l_block, l_block.get_array('rs'));
            END IF;
        END LOOP;
        RETURN l_total;
    END blocks_height;

    -- ------------------------------------------------------------ entrada
    FUNCTION render(
        p_layout  IN CLOB,
        p_queries IN CLOB,
        p_values  IN CLOB,
        p_binds   IN t_binds
    ) RETURN BLOB IS
        l_layout JSON_OBJECT_T := JSON_OBJECT_T.parse(p_layout);
        l_page   JSON_OBJECT_T := l_layout.get_object('page');
        l_header JSON_ARRAY_T  := l_layout.get_array('header');
        l_footer JSON_ARRAY_T  := l_layout.get_array('footer');
        l_keys   JSON_KEY_LIST;
        l_header_h NUMBER;
        l_footer_h NUMBER;
        l_height   NUMBER := l_page.get_number('h');
    BEGIN
        g_queries := JSON_OBJECT_T.parse(p_queries);
        g_values  := JSON_OBJECT_T.parse(p_values);
        g_binds   := NVL(p_binds, t_binds());
        g_scalars.DELETE;
        g_loaded.DELETE;
        g_now := SYSDATE;
        l_keys := g_values.get_keys;
        FOR i IN 1 .. l_keys.COUNT LOOP
            g_scalars(UPPER(l_keys(i))) := text_value(g_values.get_string(l_keys(i)));
        END LOOP;
        FOR i IN 1 .. g_binds.COUNT LOOP        -- los parámetros también se pueden imprimir
            DECLARE
                l t_value;
            BEGIN
                l.kind := g_binds(i).kind;
                l.t := g_binds(i).v_text;
                l.n := g_binds(i).v_number;
                l.d := g_binds(i).v_date;
                g_scalars(g_binds(i).name) := l;
            END;
        END LOOP;

        g_left  := l_page.get_number('ml');
        g_width := l_page.get_number('w') - g_left - l_page.get_number('mr');
        l_header_h := blocks_height(l_header);
        l_footer_h := blocks_height(l_footer);
        g_top   := GREATEST(l_page.get_number('mt'), l_page.get_number('hd') + l_header_h);
        g_limit := LEAST(l_height - l_page.get_number('mb'), l_height - l_page.get_number('fd') - l_footer_h);

        rpt_pdf.new_document(p_landscape => l_page.get_number('w') > l_height);
        g_y := g_top;
        g_page := NULL; g_pages := NULL;          -- {{PAGE}} solo tiene sentido en encabezado y pie
        blocks(l_layout.get_array('body'), TRUE);

        g_pages := rpt_pdf.page_count;
        FOR p IN 1 .. g_pages LOOP
            rpt_pdf.set_page(p);
            g_page := p;
            g_y := l_page.get_number('hd');
            blocks(l_header, FALSE);
            g_y := l_height - l_page.get_number('fd') - l_footer_h;
            blocks(l_footer, FALSE);
        END LOOP;
        RETURN rpt_pdf.get_pdf;
    END render;

END rpt_layout;
/
