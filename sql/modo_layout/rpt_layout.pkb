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
    TYPE t_name_map  IS TABLE OF VARCHAR2(128) INDEX BY VARCHAR2(128);
    TYPE t_rows      IS TABLE OF t_value_map INDEX BY PLS_INTEGER;     -- filas leídas de una tabla de datos
    TYPE t_num_map   IS TABLE OF NUMBER INDEX BY VARCHAR2(4000);       -- subtotales por grupo
    TYPE t_keys      IS TABLE OF VARCHAR2(4000) INDEX BY PLS_INTEGER;
    -- Total acumulado fila a fila (numeración y acumulados de Reports).
    TYPE t_acc IS RECORD (
        name   VARCHAR2(128),
        source VARCHAR2(128),
        fn     VARCHAR2(30),
        keys   JSON_ARRAY_T,     -- columnas del grupo que lo vuelve a cero (vacío: nunca)
        each   BOOLEAN,          -- vuelve a cero en cada fila
        prev   VARCHAR2(4000),
        total  NUMBER,
        rows   PLS_INTEGER,
        result t_value
    );
    TYPE t_accs IS TABLE OF t_acc INDEX BY PLS_INTEGER;

    g_queries  JSON_OBJECT_T;
    g_values   JSON_OBJECT_T;
    g_binds    t_binds;
    g_scalars  t_value_map;      -- "CONSULTA.COLUMNA", parámetros y constantes
    g_loaded   t_flag_map;       -- consultas escalares ya leídas
    g_row      t_value_map;      -- fila actual de una tabla de datos
    g_sums     t_value_map;      -- totales de la tabla de datos actual
    -- Código convertido de Oracle Reports (p_model).
    g_package   VARCHAR2(130);
    g_formulas  JSON_OBJECT_T;
    g_holders   JSON_OBJECT_T;
    g_summaries JSON_OBJECT_T;
    g_ctx       t_value_map;     -- último valor de cada columna, fórmula y marcador de posición
    g_fresh     t_flag_map;      -- fórmulas ya calculadas para la fila actual de su consulta
    g_busy      t_flag_map;      -- fórmulas en cálculo (corta referencias circulares)
    g_sumcache  t_value_map;     -- totales de consultas sin Data Link, calculados una vez
    g_owner     t_name_map;      -- columna de Reports -> consulta
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

    FUNCTION value_of(p_name IN VARCHAR2) RETURN t_value;

    FUNCTION open_query(p_name IN VARCHAR2) RETURN INTEGER IS
        l_query  JSON_OBJECT_T;
        l_names  JSON_ARRAY_T;
        l_cursor INTEGER;
        l_found  BOOLEAN;
        l_ignore INTEGER;
        l_refs   JSON_ARRAY_T;
        l_value  t_value;
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
            IF NOT l_found AND l_query.has('refs') THEN
                -- :COLUMNA de otro grupo de Oracle Reports: valor de la fila actual de ese grupo.
                l_refs := l_query.get_array('refs');
                FOR j IN 0 .. l_refs.get_size - 1 LOOP
                    IF UPPER(l_refs.get_string(j)) = UPPER(l_names.get_string(i)) THEN
                        l_found := TRUE;
                        l_value := value_of(l_refs.get_string(j));
                        IF l_value.kind = 'N' THEN
                            DBMS_SQL.BIND_VARIABLE(l_cursor, ':' || l_names.get_string(i), l_value.n);
                        ELSIF l_value.kind = 'D' THEN
                            DBMS_SQL.BIND_VARIABLE(l_cursor, ':' || l_names.get_string(i), l_value.d);
                        ELSE
                            DBMS_SQL.BIND_VARIABLE(l_cursor, ':' || l_names.get_string(i), l_value.t, 4000);
                        END IF;
                    END IF;
                END LOOP;
            END IF;
            IF NOT l_found THEN
                DBMS_SQL.CLOSE_CURSOR(l_cursor);
                RAISE_APPLICATION_ERROR(-20731, 'RPT_LAYOUT: falta el parámetro ' || l_names.get_string(i)
                    || ' de la consulta ' || p_name || '.');
            END IF;
        END LOOP;
        l_ignore := DBMS_SQL.EXECUTE(l_cursor);
        RETURN l_cursor;
    END open_query;

    -- Describe y define las columnas. Si la consulta trae "columns" (nombres del Modelo
    -- de Datos de Oracle Reports, en el orden del SELECT), esos nombres reemplazan a
    -- los alias por posición; las columnas sobrantes o con nombre null conservan su alias.
    PROCEDURE define_columns(
        p_query IN VARCHAR2, p_cursor IN INTEGER, p_desc OUT DBMS_SQL.DESC_TAB3, p_count OUT INTEGER
    ) IS
        l_number NUMBER;
        l_date   DATE;
        l_text   VARCHAR2(4000);
        l_query  JSON_OBJECT_T := g_queries.get_object(p_query);
        l_names  JSON_ARRAY_T;
    BEGIN
        DBMS_SQL.DESCRIBE_COLUMNS3(p_cursor, p_count, p_desc);
        IF l_query.has('columns') THEN
            l_names := l_query.get_array('columns');
            FOR i IN 1 .. LEAST(p_count, l_names.get_size) LOOP
                IF NOT l_names.get(i - 1).is_null THEN
                    p_desc(i).col_name := l_names.get_string(i - 1);
                END IF;
            END LOOP;
        END IF;
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

    -- ------------------------------------------------------------ modelo de Oracle Reports

    FUNCTION has_key(p_object IN JSON_OBJECT_T, p_key IN VARCHAR2) RETURN BOOLEAN IS
    BEGIN
        RETURN p_object IS NOT NULL AND p_object.has(p_key);
    END has_key;

    FUNCTION query_list(p_query IN VARCHAR2, p_key IN VARCHAR2) RETURN JSON_ARRAY_T IS
        l_query JSON_OBJECT_T;
    BEGIN
        IF NOT g_queries.has(p_query) THEN
            RETURN NULL;
        END IF;
        l_query := g_queries.get_object(p_query);
        IF NOT l_query.has(p_key) OR l_query.get(p_key).is_null THEN
            RETURN NULL;
        END IF;
        RETURN l_query.get_array(p_key);
    END query_list;

    FUNCTION as_number(p_value IN t_value) RETURN NUMBER IS
    BEGIN
        IF p_value.kind = 'N' THEN
            RETURN p_value.n;
        ELSIF p_value.kind = 'D' THEN
            RETURN NULL;
        END IF;
        RETURN TO_NUMBER(p_value.t);
    EXCEPTION
        WHEN VALUE_ERROR OR INVALID_NUMBER THEN
            RETURN NULL;
    END as_number;

    FUNCTION as_text(p_value IN t_value) RETURN VARCHAR2 IS
    BEGIN
        RETURN CASE p_value.kind WHEN 'N' THEN TO_CHAR(p_value.n) WHEN 'D' THEN TO_CHAR(p_value.d) ELSE p_value.t END;
    END as_text;

    FUNCTION as_date(p_value IN t_value) RETURN DATE IS
    BEGIN
        IF p_value.kind = 'D' THEN
            RETURN p_value.d;
        ELSIF p_value.kind = 'N' THEN
            RETURN NULL;
        END IF;
        RETURN TO_DATE(p_value.t);
    EXCEPTION
        WHEN OTHERS THEN
            RETURN NULL;
    END as_date;

    -- Comparación de un Data Link (como en SQL: con NULL nunca se cumple).
    FUNCTION compare(p_left IN t_value, p_right IN t_value, p_op IN VARCHAR2) RETURN BOOLEAN IS
        l_cmp   PLS_INTEGER;
        l_left  VARCHAR2(4000);
        l_right VARCHAR2(4000);
    BEGIN
        IF p_left.kind = 'D' OR p_right.kind = 'D' THEN
            IF as_date(p_left) IS NULL OR as_date(p_right) IS NULL THEN RETURN FALSE; END IF;
            l_cmp := SIGN(as_date(p_left) - as_date(p_right));
        ELSIF (p_left.kind = 'N' OR p_right.kind = 'N') AND as_number(p_left) IS NOT NULL
              AND as_number(p_right) IS NOT NULL THEN
            l_cmp := SIGN(as_number(p_left) - as_number(p_right));
        ELSE
            l_left := as_text(p_left);
            l_right := as_text(p_right);
            IF l_left IS NULL OR l_right IS NULL THEN RETURN FALSE; END IF;
            l_cmp := CASE WHEN l_left < l_right THEN -1 WHEN l_left > l_right THEN 1 ELSE 0 END;
        END IF;
        RETURN CASE p_op
            WHEN 'eq' THEN l_cmp = 0  WHEN 'ne' THEN l_cmp <> 0
            WHEN 'lt' THEN l_cmp < 0  WHEN 'le' THEN l_cmp <= 0
            WHEN 'gt' THEN l_cmp > 0  WHEN 'ge' THEN l_cmp >= 0
            ELSE FALSE END;
    END compare;

    -- Data Links: la fila de la consulta hija debe coincidir con la fila actual del grupo padre.
    FUNCTION matches(p_query IN VARCHAR2, p_row IN t_value_map) RETURN BOOLEAN IS
        l_filters JSON_ARRAY_T := query_list(p_query, 'filters');
        l_filter  JSON_OBJECT_T;
        l_child   t_value;
    BEGIN
        IF l_filters IS NULL THEN
            RETURN TRUE;
        END IF;
        FOR i IN 0 .. l_filters.get_size - 1 LOOP
            l_filter := TREAT(l_filters.get(i) AS JSON_OBJECT_T);
            l_child := CASE WHEN p_row.EXISTS(l_filter.get_string('c')) THEN p_row(l_filter.get_string('c'))
                            ELSE text_value(NULL) END;
            IF NOT compare(l_child, value_of(l_filter.get_string('p')), l_filter.get_string('op')) THEN
                RETURN FALSE;
            END IF;
        END LOOP;
        RETURN TRUE;
    END matches;

    -- La fila pasa a ser la actual de su consulta: sus fórmulas deben recalcularse.
    PROCEDURE remember(p_query IN VARCHAR2, p_row IN t_value_map) IS
        l_key      VARCHAR2(261) := p_row.FIRST;
        l_formulas JSON_ARRAY_T := query_list(p_query, 'formulas');
    BEGIN
        WHILE l_key IS NOT NULL LOOP
            g_ctx(l_key) := p_row(l_key);
            l_key := p_row.NEXT(l_key);
        END LOOP;
        IF l_formulas IS NOT NULL THEN
            FOR i IN 0 .. l_formulas.get_size - 1 LOOP
                g_fresh.DELETE(l_formulas.get_string(i));
            END LOOP;
        END IF;
    END remember;

    -- Como Reports: al leer una fila se calculan las fórmulas de su grupo (en ese orden).
    PROCEDURE on_row(p_query IN VARCHAR2, p_row IN OUT NOCOPY t_value_map) IS
        l_formulas JSON_ARRAY_T := query_list(p_query, 'formulas');
    BEGIN
        remember(p_query, p_row);
        IF l_formulas IS NOT NULL THEN
            FOR i IN 0 .. l_formulas.get_size - 1 LOOP
                p_row(l_formulas.get_string(i)) := value_of(l_formulas.get_string(i));
            END LOOP;
        END IF;
    END on_row;

    -- Consulta escalar (grupo maestro): primera fila (que cumpla sus Data Links). Sin Data
    -- Links ni :COLUMNAS de otro grupo se lee una sola vez; si no, cada vez, porque depende
    -- de la fila padre.
    PROCEDURE load_scalar_query(p_name IN VARCHAR2) IS
        l_cursor INTEGER;
        l_desc   DBMS_SQL.DESC_TAB3;
        l_count  INTEGER;
        l_row    t_value_map;
        l_found  BOOLEAN := FALSE;
        l_key    VARCHAR2(261);
    BEGIN
        IF g_loaded.EXISTS(p_name) THEN
            RETURN;
        END IF;
        IF query_list(p_name, 'filters') IS NULL AND query_list(p_name, 'refs') IS NULL THEN
            g_loaded(p_name) := TRUE;
        END IF;
        l_cursor := open_query(p_name);
        define_columns(p_name, l_cursor, l_desc, l_count);
        WHILE NOT l_found AND DBMS_SQL.FETCH_ROWS(l_cursor) > 0 LOOP
            l_row.DELETE;
            read_row(l_cursor, l_desc, l_count, NULL, l_row);
            l_found := matches(p_name, l_row);
        END LOOP;
        DBMS_SQL.CLOSE_CURSOR(l_cursor);
        IF l_found THEN
            on_row(p_name, l_row);
        ELSE
            l_row.DELETE;
            FOR i IN 1 .. l_count LOOP
                l_row(UPPER(l_desc(i).col_name)) := text_value(NULL);
            END LOOP;
            remember(p_name, l_row);
        END IF;
        l_key := l_row.FIRST;
        WHILE l_key IS NOT NULL LOOP
            g_scalars(p_name || '.' || l_key) := l_row(l_key);
            l_key := l_row.NEXT(l_key);
        END LOOP;
    EXCEPTION
        WHEN OTHERS THEN
            IF DBMS_SQL.IS_OPEN(l_cursor) THEN DBMS_SQL.CLOSE_CURSOR(l_cursor); END IF;
            RAISE;
    END load_scalar_query;

    -- Fórmula convertida: función pública del package del reporte.
    FUNCTION eval_formula(p_name IN VARCHAR2) RETURN t_value IS
        l_formula JSON_OBJECT_T := g_formulas.get_object(p_name);
        l_call    VARCHAR2(400);
        l         t_value;
        l_n       NUMBER;
        l_d       DATE;
        l_t       VARCHAR2(4000);
    BEGIN
        l.kind := NVL(l_formula.get_string('t'), 'T');
        IF g_busy.EXISTS(p_name) THEN
            RETURN l;                            -- referencia circular: NULL, como un valor aún no calculado
        END IF;
        l_call := 'BEGIN :r := ' || g_package || '.'
                  || DBMS_ASSERT.SIMPLE_SQL_NAME(l_formula.get_string('f')) || '; END;';
        g_busy(p_name) := TRUE;
        BEGIN
            IF l.kind = 'N' THEN
                EXECUTE IMMEDIATE l_call USING OUT l_n;
            ELSIF l.kind = 'D' THEN
                EXECUTE IMMEDIATE l_call USING OUT l_d;
            ELSE
                EXECUTE IMMEDIATE l_call USING OUT l_t;
            END IF;
        EXCEPTION
            WHEN OTHERS THEN
                g_busy.DELETE(p_name);
                IF SQLCODE = -20736 THEN
                    RAISE;
                END IF;
                RAISE_APPLICATION_ERROR(-20736, 'RPT_LAYOUT: error en la fórmula ' || p_name || ': '
                    || SUBSTR(SQLERRM, 1, 1800));
        END;
        g_busy.DELETE(p_name);
        l.n := l_n; l.d := l_d; l.t := l_t;
        g_ctx(p_name) := l;
        g_fresh(p_name) := TRUE;
        RETURN l;
    END eval_formula;

    -- Acumula un valor según la función del total de Reports.
    PROCEDURE accumulate(
        p_fn IN VARCHAR2, p_value IN t_value,
        p_total IN OUT NUMBER, p_rows IN OUT PLS_INTEGER, p_result IN OUT t_value
    ) IS
    BEGIN
        IF p_fn = 'last' OR (p_fn = 'first' AND p_rows = 0) THEN
            p_result := p_value;
        ELSIF p_fn IN ('minimum', 'maximum') AND as_text(p_value) IS NOT NULL THEN
            IF as_text(p_result) IS NULL
               OR compare(p_value, p_result, CASE p_fn WHEN 'minimum' THEN 'lt' ELSE 'gt' END) THEN
                p_result := p_value;
            END IF;
        ELSIF as_text(p_value) IS NOT NULL THEN
            p_total := p_total + CASE WHEN p_fn = 'count' THEN 1 ELSE NVL(as_number(p_value), 0) END;
        END IF;
        IF as_text(p_value) IS NOT NULL OR p_fn IN ('first', 'last') THEN
            p_rows := p_rows + 1;
        END IF;
    END accumulate;

    FUNCTION accumulated(p_fn IN VARCHAR2, p_total IN NUMBER, p_rows IN PLS_INTEGER, p_result IN t_value)
        RETURN t_value
    IS
        l t_value := p_result;
    BEGIN
        IF p_fn IN ('sum', 'count', 'average') THEN
            l := NULL;
            l.kind := 'N';
            l.n := CASE WHEN p_fn <> 'average' THEN p_total WHEN p_rows > 0 THEN p_total / p_rows END;
        END IF;
        l.kind := NVL(l.kind, 'N');
        RETURN l;
    END accumulated;

    -- Total de Reports (CS_): recorre su consulta (respetando los Data Links).
    FUNCTION summary(p_name IN VARCHAR2) RETURN t_value IS
        l_def    JSON_OBJECT_T := g_summaries.get_object(p_name);
        l_query  VARCHAR2(128) := l_def.get_string('q');
        l_source VARCHAR2(128) := l_def.get_string('s');
        l_fn     VARCHAR2(30)  := l_def.get_string('f');
        l_cache  BOOLEAN := query_list(l_query, 'filters') IS NULL AND query_list(l_query, 'refs') IS NULL;
        l_row    t_value_map := g_row;
        l_ctx    t_value_map := g_ctx;
        l_fresh  t_flag_map  := g_fresh;
        l_cursor INTEGER;
        l_desc   DBMS_SQL.DESC_TAB3;
        l_count  INTEGER;
        l_value  t_value;
        l_result t_value;
        l_total  NUMBER := 0;
        l_rows   PLS_INTEGER := 0;
    BEGIN
        IF l_cache AND g_sumcache.EXISTS(p_name) THEN
            RETURN g_sumcache(p_name);
        END IF;
        l_cursor := open_query(l_query);
        define_columns(l_query, l_cursor, l_desc, l_count);
        WHILE DBMS_SQL.FETCH_ROWS(l_cursor) > 0 LOOP
            g_row.DELETE;
            read_row(l_cursor, l_desc, l_count, NULL, g_row);
            IF matches(l_query, g_row) THEN
                remember(l_query, g_row);
                l_value := value_of(l_source);
                accumulate(l_fn, l_value, l_total, l_rows, l_result);
            END IF;
        END LOOP;
        DBMS_SQL.CLOSE_CURSOR(l_cursor);
        g_row := l_row; g_ctx := l_ctx; g_fresh := l_fresh;
        l_result := accumulated(l_fn, l_total, l_rows, l_result);
        IF l_cache THEN
            g_sumcache(p_name) := l_result;
        END IF;
        RETURN l_result;
    EXCEPTION
        WHEN OTHERS THEN
            IF DBMS_SQL.IS_OPEN(l_cursor) THEN DBMS_SQL.CLOSE_CURSOR(l_cursor); END IF;
            g_row := l_row; g_ctx := l_ctx; g_fresh := l_fresh;
            RAISE;
    END summary;

    -- Valor de un nombre del reporte: fila actual, fórmula, total, marcador de posición,
    -- parámetro o constante, o columna de otra consulta (se lee su primera fila).
    FUNCTION value_of(p_name IN VARCHAR2) RETURN t_value IS
        l_name   VARCHAR2(261) := UPPER(p_name);
        l        t_value;
        l_by     JSON_ARRAY_T;
        l_ignore t_value;
    BEGIN
        IF g_row.EXISTS(l_name) THEN
            RETURN g_row(l_name);
        END IF;
        IF has_key(g_formulas, l_name) THEN
            IF g_fresh.EXISTS(l_name) AND g_ctx.EXISTS(l_name) THEN
                RETURN g_ctx(l_name);
            END IF;
            RETURN eval_formula(l_name);
        END IF;
        IF has_key(g_summaries, l_name) THEN
            RETURN summary(l_name);
        END IF;
        IF has_key(g_holders, l_name) THEN
            IF NOT g_ctx.EXISTS(l_name) THEN       -- aún sin asignar: se calculan las fórmulas que lo asignan
                l_by := g_holders.get_object(l_name).get_array('by');
                FOR i IN 0 .. l_by.get_size - 1 LOOP
                    l_ignore := value_of(l_by.get_string(i));
                END LOOP;
            END IF;
            IF g_ctx.EXISTS(l_name) THEN
                RETURN g_ctx(l_name);
            END IF;
            l.kind := NVL(g_holders.get_object(l_name).get_string('t'), 'T');
            RETURN l;
        END IF;
        IF g_scalars.EXISTS(l_name) THEN
            RETURN g_scalars(l_name);
        END IF;
        IF NOT g_ctx.EXISTS(l_name) AND g_owner.EXISTS(l_name) THEN
            load_scalar_query(g_owner(l_name));
        END IF;
        IF g_ctx.EXISTS(l_name) THEN
            RETURN g_ctx(l_name);
        END IF;
        RAISE_APPLICATION_ERROR(-20732, 'RPT_LAYOUT: no hay valor para ' || l_name
            || '; revise que la consulta devuelva esa columna.');
    END value_of;

    FUNCTION num(p_name IN VARCHAR2) RETURN NUMBER IS
    BEGIN
        RETURN as_number(value_of(p_name));
    END num;

    FUNCTION txt(p_name IN VARCHAR2) RETURN VARCHAR2 IS
    BEGIN
        RETURN as_text(value_of(p_name));
    END txt;

    FUNCTION dat(p_name IN VARCHAR2) RETURN DATE IS
    BEGIN
        RETURN as_date(value_of(p_name));
    END dat;

    PROCEDURE set_num(p_name IN VARCHAR2, p_value IN NUMBER) IS
        l t_value;
    BEGIN
        l.kind := 'N'; l.n := p_value;
        g_ctx(UPPER(p_name)) := l;
    END set_num;

    PROCEDURE set_txt(p_name IN VARCHAR2, p_value IN VARCHAR2) IS
    BEGIN
        g_ctx(UPPER(p_name)) := text_value(p_value);
    END set_txt;

    PROCEDURE set_dat(p_name IN VARCHAR2, p_value IN DATE) IS
        l t_value;
    BEGIN
        l.kind := 'D'; l.d := p_value;
        g_ctx(UPPER(p_name)) := l;
    END set_dat;

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
                -- Sin la columna en la consulta: fórmula, total o marcador de posición de Reports.
                l_result := l_result || fmt(CASE WHEN g_scalars.EXISTS(l_name) THEN g_scalars(l_name)
                                                 ELSE value_of(SUBSTR(l_name, l_dot + 1)) END, l_mask);
            ELSIF l_token LIKE 'COLUMN:%' THEN
                l_name := SUBSTR(l_token, 8);
                l_result := l_result || fmt(value_of(SUBSTR(l_name, INSTR(l_name, '.') + 1)), l_mask);
            ELSIF l_token LIKE 'SUM:%' THEN
                l_name := SUBSTR(SUBSTR(l_token, 5), INSTR(SUBSTR(l_token, 5), '.') + 1);
                l_result := l_result || fmt(CASE WHEN g_sums.EXISTS(l_name) THEN g_sums(l_name)
                                                 ELSE value_of(l_name) END, l_mask);
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
        l_w      NUMBER;
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
                    l_w := rpt_pdf.text_width_mm(l_texts(j));
                    IF l_w > 0 AND l_run.has('u') THEN      -- subrayado, bajo la línea base
                        rpt_pdf.line(l_x, l_y + l_h * 0.78 + l_run.get_number('s') * 0.12 * c_pt_to_mm,
                                     l_x + l_w, l_y + l_h * 0.78 + l_run.get_number('s') * 0.12 * c_pt_to_mm,
                                     l_run.get_number('s') / 16, NVL(l_run.get_string('c'), '000000'));
                    END IF;
                    IF l_w > 0 AND l_run.has('x') THEN      -- tachado, a media altura de las minúsculas
                        rpt_pdf.line(l_x, l_y + l_h * 0.78 - l_run.get_number('s') * 0.3 * c_pt_to_mm,
                                     l_x + l_w, l_y + l_h * 0.78 - l_run.get_number('s') * 0.3 * c_pt_to_mm,
                                     l_run.get_number('s') / 16, NVL(l_run.get_string('c'), '000000'));
                    END IF;
                    l_x := l_x + l_w;
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

    -- Clave de un grupo de ruptura: valores de sus columnas en la fila.
    FUNCTION group_key(p_row IN t_value_map, p_cols IN JSON_ARRAY_T) RETURN VARCHAR2 IS
        l_key VARCHAR2(4000);
    BEGIN
        IF p_cols IS NULL THEN
            RETURN NULL;
        END IF;
        FOR i IN 0 .. p_cols.get_size - 1 LOOP
            l_key := SUBSTR(l_key || CASE WHEN p_row.EXISTS(p_cols.get_string(i))
                                          THEN as_text(p_row(p_cols.get_string(i))) END || CHR(1), 1, 3800);
        END LOOP;
        RETURN l_key;
    END group_key;

    -- Fila que se repite: si no cabe, salta de página y repite los encabezados.
    PROCEDURE data_row(p_table IN JSON_OBJECT_T, p_row IN JSON_OBJECT_T, p_header IN JSON_ARRAY_T,
                       p_next IN JSON_OBJECT_T DEFAULT NULL) IS
        l_h NUMBER := table_row(p_table, p_row, g_y, FALSE);
    BEGIN
        IF p_next IS NOT NULL THEN                -- una fila de grupo no queda sola al pie de la página
            l_h := l_h + table_row(p_table, p_next, g_y, FALSE);
        END IF;
        IF g_y + l_h > g_limit AND g_y > g_top + 0.01 THEN
            rpt_pdf.new_page;
            g_y := g_top;
            draw_rows(p_table, p_header);
        END IF;
        g_y := g_y + table_row(p_table, p_row, g_y, TRUE);
    END data_row;

    -- Tabla de datos: encabezados, filas de grupo, una fila por registro de la consulta, totales.
    -- Las filas se leen primero en memoria cuando hacen falta antes de dibujar: totales arriba de
    -- las filas ("hs"), filas de grupo con subtotal ("g") u ocultar la tabla si no hay filas ("he").
    PROCEDURE data_table(p_table IN JSON_OBJECT_T) IS
        l_all    JSON_ARRAY_T  := p_table.get_array('hd');
        l_header JSON_ARRAY_T  := JSON_ARRAY_T();   -- encabezados (se repiten en cada página)
        l_groups JSON_ARRAY_T  := JSON_ARRAY_T();   -- filas de grupo, en orden
        l_body   JSON_OBJECT_T := p_table.get_object('bd');
        l_footer JSON_ARRAY_T  := p_table.get_array('ft');
        l_sums   JSON_ARRAY_T  := p_table.get_array('sm');
        l_query  VARCHAR2(128) := p_table.get_string('q');
        l_hide   BOOLEAN := p_table.has('he');
        l_buffer BOOLEAN;
        l_cursor INTEGER;
        l_desc   DBMS_SQL.DESC_TAB3;
        l_count  INTEGER;
        l_key    VARCHAR2(261);
        l_zero   t_value;
        l_rows   t_rows;
        l_n      PLS_INTEGER := 0;
        l_totals t_value_map;
        l_gsums  t_num_map;
        l_prev   t_keys;
        l_gkey   VARCHAR2(4000);
        l_accs   t_accs;
        l_names  JSON_KEY_LIST;
        l_def    JSON_OBJECT_T;
        l_new    BOOLEAN;
        l_row    JSON_OBJECT_T;

        -- Datos de la fila recién leída (en g_row): fórmulas, acumulados y totales.
        PROCEDURE take_row IS
            l_holder JSON_KEY_LIST;
        BEGIN
            on_row(l_query, g_row);
            IF g_holders IS NOT NULL THEN            -- marcadores de posición asignados en esta fila
                l_holder := g_holders.get_keys;
                FOR i IN 1 .. l_holder.COUNT LOOP
                    IF g_ctx.EXISTS(l_holder(i)) AND NOT g_row.EXISTS(l_holder(i)) THEN
                        g_row(l_holder(i)) := g_ctx(l_holder(i));
                    END IF;
                END LOOP;
            END IF;
            FOR i IN 1 .. l_accs.COUNT LOOP
                l_gkey := group_key(g_row, l_accs(i).keys);
                IF l_accs(i).each OR l_accs(i).rows IS NULL OR NVL(l_gkey, CHR(0)) <> NVL(l_accs(i).prev, CHR(0)) THEN
                    l_accs(i).total := 0; l_accs(i).rows := 0; l_accs(i).result := NULL;
                END IF;
                l_accs(i).prev := l_gkey;
                accumulate(l_accs(i).fn, value_of(l_accs(i).source), l_accs(i).total, l_accs(i).rows, l_accs(i).result);
                g_row(l_accs(i).name) := accumulated(l_accs(i).fn, l_accs(i).total, l_accs(i).rows, l_accs(i).result);
            END LOOP;
            FOR i IN 0 .. l_sums.get_size - 1 LOOP
                l_key := UPPER(l_sums.get_string(i));
                IF g_sums.EXISTS(l_key) AND g_row.EXISTS(l_key) AND g_row(l_key).kind = 'N' THEN
                    g_sums(l_key).n := g_sums(l_key).n + NVL(g_row(l_key).n, 0);
                    IF l_buffer THEN
                        FOR j IN 0 .. l_groups.get_size - 1 LOOP
                            l_gkey := j || CHR(2) || l_key || CHR(2)
                                      || group_key(g_row, TREAT(l_groups.get(j) AS JSON_OBJECT_T).get_array('g'));
                            l_gsums(l_gkey) := CASE WHEN l_gsums.EXISTS(l_gkey) THEN l_gsums(l_gkey) ELSE 0 END
                                               + NVL(g_row(l_key).n, 0);
                        END LOOP;
                    END IF;
                END IF;
            END LOOP;
        END take_row;
    BEGIN
        FOR i IN 0 .. l_all.get_size - 1 LOOP
            l_row := TREAT(l_all.get(i) AS JSON_OBJECT_T);
            IF l_row.has('g') THEN l_groups.append(l_row); ELSE l_header.append(l_row); END IF;
        END LOOP;
        l_buffer := l_hide OR p_table.has('hs') OR l_groups.get_size > 0;
        g_sums.DELETE;
        l_zero.kind := 'N'; l_zero.n := 0;
        FOR i IN 0 .. l_sums.get_size - 1 LOOP
            IF NOT has_key(g_summaries, UPPER(l_sums.get_string(i))) THEN   -- los CS_ se calculan aparte
                g_sums(UPPER(l_sums.get_string(i))) := l_zero;
            END IF;
        END LOOP;
        IF g_summaries IS NOT NULL THEN               -- acumulados fila a fila de esta consulta
            l_names := g_summaries.get_keys;
            FOR i IN 1 .. l_names.COUNT LOOP
                l_def := g_summaries.get_object(l_names(i));
                IF l_def.get_string('q') = l_query AND l_def.has('run') THEN
                    l_accs(l_accs.COUNT + 1).name := l_names(i);
                    l_accs(l_accs.COUNT).source := l_def.get_string('s');
                    l_accs(l_accs.COUNT).fn := l_def.get_string('f');
                    l_accs(l_accs.COUNT).keys := CASE WHEN l_def.has('rk') THEN l_def.get_array('rk') END;
                    l_accs(l_accs.COUNT).each := l_def.has('each');
                END IF;
            END LOOP;
        END IF;
        l_cursor := open_query(l_query);
        define_columns(l_query, l_cursor, l_desc, l_count);

        IF NOT l_buffer THEN                          -- directo: se dibuja mientras se lee
            g_row.DELETE;
            FOR i IN 1 .. l_count LOOP                -- fila vacía para medir antes de leer
                g_row(UPPER(l_desc(i).col_name)) := text_value(NULL);
            END LOOP;
            ensure(rows_height(p_table, l_header) + table_row(p_table, l_body, 0, FALSE));
            draw_rows(p_table, l_header);
        END IF;
        WHILE DBMS_SQL.FETCH_ROWS(l_cursor) > 0 LOOP
            g_row.DELETE;
            read_row(l_cursor, l_desc, l_count, NULL, g_row);
            CONTINUE WHEN NOT matches(l_query, g_row);
            take_row;
            IF l_buffer THEN
                l_n := l_n + 1;
                l_rows(l_n) := g_row;
            ELSE
                data_row(p_table, l_body, l_header);
            END IF;
        END LOOP;
        DBMS_SQL.CLOSE_CURSOR(l_cursor);

        IF l_buffer THEN
            IF l_n = 0 AND l_hide THEN
                RETURN;                               -- sin filas: la sección no se imprime
            END IF;
            l_totals := g_sums;
            g_row := CASE WHEN l_n > 0 THEN l_rows(1) ELSE g_row END;
            ensure(rows_height(p_table, l_header) + rows_height(p_table, l_groups) + table_row(p_table, l_body, 0, FALSE));
            draw_rows(p_table, l_header);
            FOR r IN 1 .. l_n LOOP
                g_row := l_rows(r);
                l_new := FALSE;
                FOR j IN 0 .. l_groups.get_size - 1 LOOP
                    l_row := TREAT(l_groups.get(j) AS JSON_OBJECT_T);
                    l_gkey := group_key(g_row, l_row.get_array('g'));
                    IF l_new OR NOT l_prev.EXISTS(j) OR l_prev(j) <> l_gkey THEN
                        l_new := TRUE;                -- cambia este grupo y los de adentro
                        l_prev(j) := l_gkey;
                        FOR i IN 0 .. l_sums.get_size - 1 LOOP   -- subtotales del grupo
                            l_key := UPPER(l_sums.get_string(i));
                            IF l_totals.EXISTS(l_key) THEN
                                g_sums(l_key).n := CASE WHEN l_gsums.EXISTS(j || CHR(2) || l_key || CHR(2) || l_gkey)
                                                        THEN l_gsums(j || CHR(2) || l_key || CHR(2) || l_gkey) ELSE 0 END;
                            END IF;
                        END LOOP;
                        data_row(p_table, l_row, l_header, l_body);
                        g_sums := l_totals;
                    END IF;
                END LOOP;
                data_row(p_table, l_body, l_header);
            END LOOP;
        END IF;
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
        p_binds   IN t_binds,
        p_model   IN CLOB DEFAULT NULL
    ) RETURN BLOB IS
        l_layout JSON_OBJECT_T := JSON_OBJECT_T.parse(p_layout);
        l_model  JSON_OBJECT_T := CASE WHEN p_model IS NOT NULL THEN JSON_OBJECT_T.parse(p_model) END;
        l_names  JSON_ARRAY_T;
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
        g_row.DELETE;
        g_ctx.DELETE;
        g_fresh.DELETE;
        g_busy.DELETE;
        g_sumcache.DELETE;
        g_owner.DELETE;
        g_package := NULL;
        g_formulas := NULL; g_holders := NULL; g_summaries := NULL;
        IF l_model IS NOT NULL THEN
            g_package   := DBMS_ASSERT.SIMPLE_SQL_NAME(l_model.get_string('package'));
            g_formulas  := l_model.get_object('formulas');
            g_holders   := l_model.get_object('placeholders');
            g_summaries := l_model.get_object('summaries');
        END IF;
        l_keys := g_queries.get_keys;               -- columna de Reports -> su consulta
        FOR i IN 1 .. l_keys.COUNT LOOP
            l_names := query_list(l_keys(i), 'columns');
            IF l_names IS NOT NULL THEN
                FOR j IN 0 .. l_names.get_size - 1 LOOP
                    IF NOT l_names.get(j).is_null AND NOT g_owner.EXISTS(l_names.get_string(j)) THEN
                        g_owner(l_names.get_string(j)) := l_keys(i);
                    END IF;
                END LOOP;
            END IF;
        END LOOP;
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
