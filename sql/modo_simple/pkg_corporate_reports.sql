-- Versión 8: conserva el flujo IG v7 y añade reportes SQL declarativos.
CREATE OR REPLACE PACKAGE pkg_corporate_reports AUTHID CURRENT_USER AS

    /**
     * Descarga como PDF o XLSX los datos actualmente filtrados de un
     * Interactive Grid, respetando el orden y la selección de columnas
     * recibidos desde el navegador.
     *
     * p_application_id       ID de la aplicación APEX actual.
     * p_page_id              Página que contiene el Interactive Grid.
     * p_region_static_id     Static ID de la región Interactive Grid.
     * p_visible_columns_json JSON generado por el código JavaScript común.
     * p_title                Título mostrado en el encabezado del PDF.
     * p_file_name            Nombre base del archivo, sin extensión.
     * p_generated_by         Usuario que solicita la descarga.
     * p_format               PDF o XLSX.
     * p_orientation          AUTO, PORTRAIT o LANDSCAPE.
     * p_max_rows             Se conserva solo por compatibilidad de la firma y
     *                        se ignora: nunca se limita la cantidad de registros.
     * p_excluded_columns_json
     *                        Array JSON con columnas que nunca deben
     *                        exportarse, aunque estén visibles.
     * p_column_spans_json    Array JSON de pesos relativos, aplicado a las
     *                        columnas finales de izquierda a derecha. El
     *                        valor especial "*" deja esa columna adaptable.
     * p_display_columns_json Parámetro obsoleto, conservado únicamente para
     *                        no romper llamadas existentes. Desde esta
     *                        versión, APEX_REGION.EXPORT_DATA proporciona
     *                        directamente los valores de presentación LOV.
     */
    PROCEDURE download_ig(
        p_application_id       IN NUMBER,
        p_page_id              IN NUMBER,
        p_region_static_id     IN VARCHAR2,
        p_visible_columns_json IN CLOB,
        p_title                IN VARCHAR2,
        p_file_name            IN VARCHAR2,
        p_generated_by         IN VARCHAR2 DEFAULT NULL,
        p_format               IN VARCHAR2 DEFAULT 'PDF',
        p_orientation          IN VARCHAR2 DEFAULT 'AUTO',
        p_max_rows             IN PLS_INTEGER DEFAULT 1000,
        p_excluded_columns_json IN CLOB DEFAULT NULL,
        p_column_spans_json    IN CLOB DEFAULT NULL,
        p_display_columns_json IN CLOB DEFAULT NULL
    );

    /* Compatibilidad con las páginas que ya llaman al procedimiento PDF. */
    PROCEDURE download_ig_pdf(
        p_application_id       IN NUMBER,
        p_page_id              IN NUMBER,
        p_region_static_id     IN VARCHAR2,
        p_visible_columns_json IN CLOB,
        p_title                IN VARCHAR2,
        p_file_name            IN VARCHAR2,
        p_generated_by         IN VARCHAR2 DEFAULT NULL,
        p_orientation          IN VARCHAR2 DEFAULT 'AUTO',
        p_max_rows             IN PLS_INTEGER DEFAULT 1000,
        p_excluded_columns_json IN CLOB DEFAULT NULL,
        p_column_spans_json    IN CLOB DEFAULT NULL,
        p_display_columns_json IN CLOB DEFAULT NULL
    );

    /**
     * Descarga un reporte tabular basado en una única consulta SQL SELECT/WITH.
     * El SQL, las columnas y los estilos deben ser configuración del proceso
     * APEX; nunca deben recibirse desde el navegador.
     *
     * Los binds son lógicos y se resuelven contra Page Items autorizados.
     * La consulta debe devolver directamente los display values que deban
     * imprimirse (por ejemplo, mediante JOIN para columnas LOV).
     */
    PROCEDURE download_query(
        p_sql_query              IN VARCHAR2,
        p_columns_json           IN CLOB,
        p_bindings_json          IN CLOB DEFAULT NULL,
        p_fields_json            IN CLOB DEFAULT NULL,
        p_style_json             IN CLOB DEFAULT NULL,
        p_header_template        IN VARCHAR2 DEFAULT '{{REPORT_TITLE}}',
        p_footer_template        IN VARCHAR2 DEFAULT NULL,
        p_title                  IN VARCHAR2 DEFAULT 'Reporte',
        p_file_name              IN VARCHAR2 DEFAULT 'reporte',
        p_format                 IN VARCHAR2 DEFAULT 'PDF',
        p_orientation            IN VARCHAR2 DEFAULT 'AUTO',
        p_excluded_columns_json  IN CLOB DEFAULT NULL,
        p_column_widths_json     IN CLOB DEFAULT NULL
    );

END pkg_corporate_reports;
/

CREATE OR REPLACE PACKAGE BODY pkg_corporate_reports AS

    c_query_max_columns CONSTANT PLS_INTEGER := 50;
    c_width_budget      CONSTANT NUMBER := 99;
    c_auto_reserve_pct  CONSTANT NUMBER := 20;

    TYPE t_name_set IS TABLE OF BOOLEAN
        INDEX BY VARCHAR2(4000);

    TYPE t_varchar_map IS TABLE OF VARCHAR2(32767)
        INDEX BY VARCHAR2(4000);

    TYPE t_number_name_map IS TABLE OF NUMBER
        INDEX BY VARCHAR2(4000);

    TYPE t_style_config IS RECORD (
        title_font_color       apex_data_export.t_color,
        title_font_family      apex_data_export.t_font_family,
        title_font_weight      apex_data_export.t_font_weight,
        title_font_size        NUMBER,
        title_alignment        apex_data_export.t_alignment,
        footer_font_color      apex_data_export.t_color,
        footer_font_family     apex_data_export.t_font_family,
        footer_font_weight     apex_data_export.t_font_weight,
        footer_font_size       NUMBER,
        footer_alignment       apex_data_export.t_alignment,
        header_bg_color        apex_data_export.t_color,
        header_font_color      apex_data_export.t_color,
        header_font_family     apex_data_export.t_font_family,
        header_font_weight     apex_data_export.t_font_weight,
        header_font_size       NUMBER,
        header_alignment       apex_data_export.t_alignment,
        body_bg_color          apex_data_export.t_color,
        body_font_color        apex_data_export.t_color,
        body_font_family       apex_data_export.t_font_family,
        body_font_weight       apex_data_export.t_font_weight,
        body_font_size         NUMBER,
        body_alignment         apex_data_export.t_alignment,
        border_width           NUMBER,
        border_color           apex_data_export.t_color
    );

    FUNCTION is_valid_json(p_json IN CLOB) RETURN BOOLEAN
    IS
        l_count PLS_INTEGER;
    BEGIN
        IF p_json IS NULL THEN
            RETURN TRUE;
        END IF;

        SELECT COUNT(*)
          INTO l_count
          FROM dual
         WHERE p_json IS JSON;

        RETURN l_count = 1;
    END is_valid_json;


    PROCEDURE assert_json(
        p_json  IN CLOB,
        p_label IN VARCHAR2,
        p_code  IN PLS_INTEGER
    )
    IS
    BEGIN
        IF p_json IS NOT NULL AND NOT is_valid_json(p_json) THEN
            raise_application_error(
                p_code,
                'La configuración JSON de ' || p_label || ' no es válida.'
            );
        END IF;
    END assert_json;


    FUNCTION normalize_name(
        p_name  IN VARCHAR2,
        p_label IN VARCHAR2,
        p_code  IN PLS_INTEGER
    ) RETURN VARCHAR2
    IS
        l_name VARCHAR2(128) := UPPER(TRIM(p_name));
    BEGIN
        IF l_name IS NULL OR
           NOT REGEXP_LIKE(l_name, '^[A-Z][A-Z0-9_$#]{0,29}$')
        THEN
            raise_application_error(
                p_code,
                p_label || ' contiene un identificador no permitido.'
            );
        END IF;

        RETURN l_name;
    END normalize_name;


    FUNCTION normalize_item_name(
        p_name  IN VARCHAR2,
        p_label IN VARCHAR2,
        p_code  IN PLS_INTEGER
    ) RETURN VARCHAR2
    IS
        l_name VARCHAR2(128) := UPPER(TRIM(p_name));
    BEGIN
        IF l_name IS NULL OR
           NOT REGEXP_LIKE(l_name, '^P[0-9]+_[A-Z][A-Z0-9_]{0,119}$')
        THEN
            raise_application_error(
                p_code,
                p_label ||
                ' debe cumplir la convención P<page>_NOMBRE.'
            );
        END IF;

        RETURN l_name;
    END normalize_item_name;


    PROCEDURE collect_sql_bind_names(
        p_sql_query IN VARCHAR2,
        p_names     IN OUT NOCOPY t_name_set
    )
    IS
        l_index       PLS_INTEGER := 1;
        l_length      PLS_INTEGER := LENGTH(p_sql_query);
        l_start       PLS_INTEGER;
        l_character   VARCHAR2(1);
        l_next        VARCHAR2(1);
        l_delimiter   VARCHAR2(1);
        l_closing     VARCHAR2(1);
        l_name        VARCHAR2(128);
    BEGIN
        /*
         * Escáner deliberadamente pequeño: reconoce binds fuera de strings,
         * identificadores citados y comentarios. También ignora literales
         * alternativos q'[...]', necesarios en consultas Oracle reales.
         */
        WHILE l_index <= l_length
        LOOP
            l_character := SUBSTR(p_sql_query, l_index, 1);
            l_next := SUBSTR(p_sql_query, l_index + 1, 1);

            IF l_character = '-' AND l_next = '-' THEN
                l_index := INSTR(p_sql_query, CHR(10), l_index + 2);
                IF l_index = 0 THEN
                    EXIT;
                END IF;

            ELSIF l_character = '/' AND l_next = '*' THEN
                l_index := INSTR(p_sql_query, '*/', l_index + 2);
                IF l_index = 0 THEN
                    raise_application_error(
                        -20103,
                        'La consulta contiene un comentario sin cerrar.'
                    );
                END IF;
                l_index := l_index + 2;

            ELSIF UPPER(l_character) = 'Q' AND l_next = '''' AND
                  l_index + 2 <= l_length
            THEN
                l_delimiter := SUBSTR(p_sql_query, l_index + 2, 1);
                l_closing := CASE l_delimiter
                    WHEN '[' THEN ']'
                    WHEN '{' THEN '}'
                    WHEN '(' THEN ')'
                    WHEN '<' THEN '>'
                    ELSE l_delimiter
                END;
                l_index := INSTR(
                    p_sql_query,
                    l_closing || '''',
                    l_index + 3
                );
                IF l_index = 0 THEN
                    raise_application_error(
                        -20164,
                        'La consulta contiene un literal q sin cerrar.'
                    );
                END IF;
                l_index := l_index + 2;

            ELSIF l_character = '''' THEN
                l_index := l_index + 1;
                LOOP
                    l_index := INSTR(p_sql_query, '''', l_index);
                    IF l_index = 0 THEN
                        raise_application_error(
                            -20165,
                            'La consulta contiene un literal sin cerrar.'
                        );
                    END IF;
                    IF SUBSTR(p_sql_query, l_index + 1, 1) = '''' THEN
                        l_index := l_index + 2;
                    ELSE
                        l_index := l_index + 1;
                        EXIT;
                    END IF;
                END LOOP;

            ELSIF l_character = '"' THEN
                l_index := l_index + 1;
                LOOP
                    l_index := INSTR(p_sql_query, '"', l_index);
                    IF l_index = 0 THEN
                        raise_application_error(
                            -20166,
                            'La consulta contiene un identificador citado ' ||
                            'sin cerrar.'
                        );
                    END IF;
                    IF SUBSTR(p_sql_query, l_index + 1, 1) = '"' THEN
                        l_index := l_index + 2;
                    ELSE
                        l_index := l_index + 1;
                        EXIT;
                    END IF;
                END LOOP;

            ELSIF l_character = ':' AND
                  REGEXP_LIKE(l_next, '[A-Za-z]') AND
                  NVL(SUBSTR(p_sql_query, l_index - 1, 1), ' ') <> ':'
            THEN
                l_start := l_index + 1;
                l_index := l_start + 1;
                WHILE l_index <= l_length AND
                      REGEXP_LIKE(
                          SUBSTR(p_sql_query, l_index, 1),
                          '[A-Za-z0-9_$#]'
                      )
                LOOP
                    l_index := l_index + 1;
                END LOOP;

                l_name := normalize_name(
                    SUBSTR(p_sql_query, l_start, l_index - l_start),
                    'La consulta SQL',
                    -20167
                );
                p_names(l_name) := TRUE;

            ELSIF l_character = ':' AND
                  NVL(SUBSTR(p_sql_query, l_index - 1, 1), ' ') <> ':' AND
                  l_next <> ':'
            THEN
                raise_application_error(
                    -20168,
                    'Solo se admiten binds con nombre que comience por una ' ||
                    'letra, por ejemplo :P_CLIENTE.'
                );

            ELSIF l_character = ';' THEN
                /*
                 * Solo se rechazan separadores fuera de strings,
                 * identificadores citados y comentarios. Un punto y coma
                 * dentro de texto no convierte la consulta en dos sentencias.
                 */
                raise_application_error(
                    -20105,
                    'La consulta no debe contener punto y coma.'
                );

            ELSE
                l_index := l_index + 1;
            END IF;
        END LOOP;
    END collect_sql_bind_names;


    FUNCTION normalize_color(
        p_value   IN VARCHAR2,
        p_default IN VARCHAR2,
        p_label   IN VARCHAR2
    ) RETURN VARCHAR2
    IS
        l_value VARCHAR2(20) := UPPER(NVL(TRIM(p_value), p_default));
    BEGIN
        IF NOT REGEXP_LIKE(l_value, '^#[0-9A-F]{6}$') THEN
            raise_application_error(
                -20140,
                'El color de ' || p_label || ' debe usar el formato #RRGGBB.'
            );
        END IF;

        RETURN l_value;
    END normalize_color;


    FUNCTION normalize_font_family(
        p_value   IN VARCHAR2,
        p_default IN apex_data_export.t_font_family,
        p_label   IN VARCHAR2
    ) RETURN apex_data_export.t_font_family
    IS
    BEGIN
        CASE UPPER(NVL(TRIM(p_value), p_default))
            WHEN 'HELVETICA' THEN
                RETURN apex_data_export.c_font_family_helvetica;
            WHEN 'TIMES' THEN
                RETURN apex_data_export.c_font_family_times;
            WHEN 'COURIER' THEN
                RETURN apex_data_export.c_font_family_courier;
            ELSE
                raise_application_error(
                    -20141,
                    'La fuente de ' || p_label ||
                    ' debe ser HELVETICA, TIMES o COURIER.'
                );
        END CASE;
    END normalize_font_family;


    FUNCTION normalize_font_weight(
        p_value   IN VARCHAR2,
        p_default IN apex_data_export.t_font_weight,
        p_label   IN VARCHAR2
    ) RETURN apex_data_export.t_font_weight
    IS
    BEGIN
        CASE UPPER(NVL(TRIM(p_value), p_default))
            WHEN 'NORMAL' THEN
                RETURN apex_data_export.c_font_weight_normal;
            WHEN 'BOLD' THEN
                RETURN apex_data_export.c_font_weight_bold;
            ELSE
                raise_application_error(
                    -20142,
                    'El peso de fuente de ' || p_label ||
                    ' debe ser NORMAL o BOLD.'
                );
        END CASE;
    END normalize_font_weight;


    FUNCTION normalize_alignment(
        p_value   IN VARCHAR2,
        p_default IN apex_data_export.t_alignment,
        p_label   IN VARCHAR2
    ) RETURN apex_data_export.t_alignment
    IS
    BEGIN
        CASE UPPER(NVL(TRIM(p_value), p_default))
            WHEN 'START' THEN
                RETURN apex_data_export.c_align_start;
            WHEN 'LEFT' THEN
                RETURN apex_data_export.c_align_start;
            WHEN 'CENTER' THEN
                RETURN apex_data_export.c_align_center;
            WHEN 'END' THEN
                RETURN apex_data_export.c_align_end;
            WHEN 'RIGHT' THEN
                RETURN apex_data_export.c_align_end;
            ELSE
                raise_application_error(
                    -20143,
                    'La alineación de ' || p_label ||
                    ' debe ser START, CENTER o END.'
                );
        END CASE;
    END normalize_alignment;


    FUNCTION bounded_number(
        p_value   IN NUMBER,
        p_default IN NUMBER,
        p_minimum IN NUMBER,
        p_maximum IN NUMBER,
        p_label   IN VARCHAR2
    ) RETURN NUMBER
    IS
        l_value NUMBER := NVL(p_value, p_default);
    BEGIN
        IF l_value < p_minimum OR l_value > p_maximum THEN
            raise_application_error(
                -20144,
                p_label || ' debe estar entre ' || p_minimum ||
                ' y ' || p_maximum || '.'
            );
        END IF;

        RETURN l_value;
    END bounded_number;


    PROCEDURE assert_allowed_keys(
        p_object       IN json_object_t,
        p_allowed_keys IN VARCHAR2,
        p_section      IN VARCHAR2
    )
    IS
        l_keys json_key_list;
        l_key  VARCHAR2(4000);
    BEGIN
        IF p_object IS NULL THEN
            RETURN;
        END IF;

        l_keys := p_object.get_keys;

        FOR key_index IN 1 .. l_keys.COUNT
        LOOP
            l_key := UPPER(l_keys(key_index));

            IF INSTR(
                   ',' || UPPER(p_allowed_keys) || ',',
                   ',' || l_key || ','
               ) = 0
            THEN
                raise_application_error(
                    -20145,
                    'La propiedad "' || SUBSTR(l_keys(key_index), 1, 100) ||
                    '" no está permitida en el estilo ' || p_section || '.'
                );
            END IF;
        END LOOP;
    END assert_allowed_keys;


    FUNCTION child_object(
        p_parent  IN json_object_t,
        p_key     IN VARCHAR2,
        p_section IN VARCHAR2
    ) RETURN json_object_t
    IS
        l_element json_element_t;
    BEGIN
        IF p_parent IS NULL OR NOT p_parent.has(p_key) THEN
            RETURN NULL;
        END IF;

        l_element := p_parent.get(p_key);

        IF l_element IS NULL OR NOT l_element.is_object THEN
            raise_application_error(
                -20146,
                'La sección de estilo ' || p_section ||
                ' debe ser un objeto JSON.'
            );
        END IF;

        RETURN TREAT(l_element AS json_object_t);
    END child_object;


    FUNCTION json_string(
        p_object  IN json_object_t,
        p_key     IN VARCHAR2,
        p_default IN VARCHAR2
    ) RETURN VARCHAR2
    IS
    BEGIN
        IF p_object IS NULL OR NOT p_object.has(p_key) THEN
            RETURN p_default;
        END IF;

        RETURN p_object.get_string(p_key);
    END json_string;


    FUNCTION json_number(
        p_object  IN json_object_t,
        p_key     IN VARCHAR2,
        p_default IN NUMBER
    ) RETURN NUMBER
    IS
    BEGIN
        IF p_object IS NULL OR NOT p_object.has(p_key) THEN
            RETURN p_default;
        END IF;

        RETURN p_object.get_number(p_key);
    END json_number;


    FUNCTION parse_style(p_style_json IN CLOB) RETURN t_style_config
    IS
        l_style       t_style_config;
        l_root        json_object_t;
        l_title       json_object_t;
        l_header      json_object_t;
        l_body        json_object_t;
        l_border      json_object_t;
        l_footer      json_object_t;
    BEGIN
        l_style.title_font_color := '#2F343A';
        l_style.title_font_family :=
            apex_data_export.c_font_family_helvetica;
        l_style.title_font_weight :=
            apex_data_export.c_font_weight_bold;
        l_style.title_font_size := 15;
        l_style.title_alignment := apex_data_export.c_align_center;

        l_style.footer_font_color := '#666666';
        l_style.footer_font_family :=
            apex_data_export.c_font_family_helvetica;
        l_style.footer_font_weight :=
            apex_data_export.c_font_weight_normal;
        l_style.footer_font_size := 8;
        l_style.footer_alignment := apex_data_export.c_align_center;

        l_style.header_bg_color := '#4A4F55';
        l_style.header_font_color := '#FFFFFF';
        l_style.header_font_family :=
            apex_data_export.c_font_family_helvetica;
        l_style.header_font_weight :=
            apex_data_export.c_font_weight_bold;
        l_style.header_font_size := 9;
        l_style.header_alignment := apex_data_export.c_align_center;

        l_style.body_bg_color := '#FFFFFF';
        l_style.body_font_color := '#25282B';
        l_style.body_font_family :=
            apex_data_export.c_font_family_helvetica;
        l_style.body_font_weight :=
            apex_data_export.c_font_weight_normal;
        l_style.body_font_size := 8.5;
        l_style.body_alignment := apex_data_export.c_align_start;

        l_style.border_width := 0.5;
        l_style.border_color := '#BFC3C7';

        IF p_style_json IS NULL THEN
            RETURN l_style;
        END IF;

        assert_json(p_style_json, 'estilos', -20147);

        l_root := json_object_t.parse(p_style_json);

        assert_allowed_keys(
            l_root,
            'TITLE,TABLE_HEADER,TABLE_BODY,BORDER,FOOTER',
            'raíz'
        );

        l_title := child_object(l_root, 'title', 'title');
        l_header := child_object(l_root, 'table_header', 'table_header');
        l_body := child_object(l_root, 'table_body', 'table_body');
        l_border := child_object(l_root, 'border', 'border');
        l_footer := child_object(l_root, 'footer', 'footer');

        assert_allowed_keys(
            l_title,
            'FONT_FAMILY,FONT_SIZE,FONT_WEIGHT,FONT_COLOR,ALIGNMENT',
            'title'
        );
        assert_allowed_keys(
            l_header,
            'BACKGROUND_COLOR,FONT_FAMILY,FONT_SIZE,FONT_WEIGHT,' ||
            'FONT_COLOR,ALIGNMENT',
            'table_header'
        );
        assert_allowed_keys(
            l_body,
            'BACKGROUND_COLOR,FONT_FAMILY,FONT_SIZE,FONT_WEIGHT,' ||
            'FONT_COLOR,ALIGNMENT',
            'table_body'
        );
        assert_allowed_keys(l_border, 'WIDTH,COLOR', 'border');
        assert_allowed_keys(
            l_footer,
            'FONT_FAMILY,FONT_SIZE,FONT_WEIGHT,FONT_COLOR,ALIGNMENT',
            'footer'
        );

        l_style.title_font_color := normalize_color(
            json_string(l_title, 'font_color', l_style.title_font_color),
            l_style.title_font_color,
            'title'
        );
        l_style.title_font_family := normalize_font_family(
            json_string(l_title, 'font_family', l_style.title_font_family),
            l_style.title_font_family,
            'title'
        );
        l_style.title_font_weight := normalize_font_weight(
            json_string(l_title, 'font_weight', l_style.title_font_weight),
            l_style.title_font_weight,
            'title'
        );
        l_style.title_font_size := bounded_number(
            json_number(l_title, 'font_size', l_style.title_font_size),
            l_style.title_font_size,
            8,
            24,
            'El tamaño de fuente del título'
        );
        l_style.title_alignment := normalize_alignment(
            json_string(l_title, 'alignment', 'CENTER'),
            apex_data_export.c_align_center,
            'title'
        );

        IF l_style.title_alignment <>
           apex_data_export.c_align_center
        THEN
            raise_application_error(
                -20148,
                'El título debe estar alineado al centro.'
            );
        END IF;

        l_style.header_bg_color := normalize_color(
            json_string(
                l_header,
                'background_color',
                l_style.header_bg_color
            ),
            l_style.header_bg_color,
            'table_header.background_color'
        );
        l_style.header_font_color := normalize_color(
            json_string(l_header, 'font_color', l_style.header_font_color),
            l_style.header_font_color,
            'table_header.font_color'
        );
        l_style.header_font_family := normalize_font_family(
            json_string(l_header, 'font_family', l_style.header_font_family),
            l_style.header_font_family,
            'table_header'
        );
        l_style.header_font_weight := normalize_font_weight(
            json_string(l_header, 'font_weight', l_style.header_font_weight),
            l_style.header_font_weight,
            'table_header'
        );
        l_style.header_font_size := bounded_number(
            json_number(l_header, 'font_size', l_style.header_font_size),
            l_style.header_font_size,
            6,
            16,
            'El tamaño de fuente del encabezado de tabla'
        );
        l_style.header_alignment := normalize_alignment(
            json_string(l_header, 'alignment', 'CENTER'),
            apex_data_export.c_align_center,
            'table_header'
        );

        l_style.body_bg_color := normalize_color(
            json_string(l_body, 'background_color', l_style.body_bg_color),
            l_style.body_bg_color,
            'table_body.background_color'
        );
        l_style.body_font_color := normalize_color(
            json_string(l_body, 'font_color', l_style.body_font_color),
            l_style.body_font_color,
            'table_body.font_color'
        );
        l_style.body_font_family := normalize_font_family(
            json_string(l_body, 'font_family', l_style.body_font_family),
            l_style.body_font_family,
            'table_body'
        );
        l_style.body_font_weight := normalize_font_weight(
            json_string(l_body, 'font_weight', l_style.body_font_weight),
            l_style.body_font_weight,
            'table_body'
        );
        l_style.body_font_size := bounded_number(
            json_number(l_body, 'font_size', l_style.body_font_size),
            l_style.body_font_size,
            6,
            14,
            'El tamaño de fuente del cuerpo de tabla'
        );
        l_style.body_alignment := normalize_alignment(
            json_string(l_body, 'alignment', 'START'),
            apex_data_export.c_align_start,
            'table_body'
        );

        l_style.border_width := bounded_number(
            json_number(l_border, 'width', l_style.border_width),
            l_style.border_width,
            0,
            5,
            'El grosor del borde'
        );
        l_style.border_color := normalize_color(
            json_string(l_border, 'color', l_style.border_color),
            l_style.border_color,
            'border.color'
        );

        l_style.footer_font_color := normalize_color(
            json_string(l_footer, 'font_color', l_style.footer_font_color),
            l_style.footer_font_color,
            'footer.font_color'
        );
        l_style.footer_font_family := normalize_font_family(
            json_string(l_footer, 'font_family', l_style.footer_font_family),
            l_style.footer_font_family,
            'footer'
        );
        l_style.footer_font_weight := normalize_font_weight(
            json_string(l_footer, 'font_weight', l_style.footer_font_weight),
            l_style.footer_font_weight,
            'footer'
        );
        l_style.footer_font_size := bounded_number(
            json_number(l_footer, 'font_size', l_style.footer_font_size),
            l_style.footer_font_size,
            6,
            12,
            'El tamaño de fuente del pie'
        );
        l_style.footer_alignment := normalize_alignment(
            json_string(l_footer, 'alignment', 'CENTER'),
            apex_data_export.c_align_center,
            'footer'
        );

        RETURN l_style;
    EXCEPTION
        WHEN OTHERS THEN
            IF SQLCODE BETWEEN -20999 AND -20000 THEN
                RAISE;
            END IF;

            raise_application_error(
                -20149,
                'No se pudo interpretar la configuración de estilos.'
            );
    END parse_style;


    PROCEDURE assert_read_only_query(p_sql_query IN VARCHAR2)
    IS
        /* Espacio, tabulador, LF y CR: igual que el compilador local. */
        c_blank CONSTANT VARCHAR2(4) := ' ' || CHR(9) || CHR(10) || CHR(13);
        l_check VARCHAR2(32767);
        l_end   PLS_INTEGER;
    BEGIN
        IF LENGTHB(p_sql_query) > 32767 THEN
            raise_application_error(
                -20102,
                'La consulta SQL supera el límite de 32767 bytes.'
            );
        END IF;

        l_check := LTRIM(p_sql_query, c_blank);

        IF l_check IS NULL THEN
            raise_application_error(-20101, 'No se recibió la consulta SQL.');
        END IF;

        IF INSTR(p_sql_query, CHR(0)) > 0 THEN
            raise_application_error(
                -20172,
                'La consulta contiene un carácter NUL no permitido.'
            );
        END IF;

        /* Se permiten comentarios únicamente antes del primer token. */
        LOOP
            IF SUBSTR(l_check, 1, 2) = '--' THEN
                l_end := INSTR(l_check, CHR(10));
                IF l_end = 0 THEN
                    l_check := NULL;
                ELSE
                    l_check := LTRIM(SUBSTR(l_check, l_end + 1), c_blank);
                END IF;
            ELSIF SUBSTR(l_check, 1, 2) = '/*' THEN
                l_end := INSTR(l_check, '*/', 3);
                IF l_end = 0 THEN
                    raise_application_error(
                        -20103,
                        'La consulta contiene un comentario sin cerrar.'
                    );
                END IF;
                l_check := LTRIM(SUBSTR(l_check, l_end + 2), c_blank);
            ELSE
                EXIT;
            END IF;
        END LOOP;

        IF l_check IS NULL OR NOT REGEXP_LIKE(
                   l_check,
                   '^(SELECT|WITH)([^A-Z0-9_$#]|$)',
                   'i'
               )
        THEN
            raise_application_error(
                -20104,
                'El reporte solo admite una consulta SELECT o WITH.'
            );
        END IF;

        /*
         * collect_sql_bind_names comprueba después los separadores fuera de
         * literales y comentarios mientras obtiene los binds autorizados.
         */
    END assert_read_only_query;


    FUNCTION template_uses_field(
        p_template   IN VARCHAR2,
        p_field_name IN VARCHAR2
    ) RETURN BOOLEAN
    IS
        l_scan_pos   PLS_INTEGER := 1;
        l_open_pos   PLS_INTEGER;
        l_close_pos  PLS_INTEGER;
        l_token      VARCHAR2(4000);
        l_name       VARCHAR2(128);
    BEGIN
        IF p_template IS NULL THEN
            RETURN FALSE;
        END IF;

        LOOP
            l_open_pos := INSTR(p_template, '{{', l_scan_pos);
            EXIT WHEN l_open_pos = 0;

            l_close_pos := INSTR(p_template, '}}', l_open_pos + 2);
            EXIT WHEN l_close_pos = 0;

            l_token := UPPER(
                TRIM(
                    SUBSTR(
                        p_template,
                        l_open_pos + 2,
                        l_close_pos - l_open_pos - 2
                    )
                )
            );

            IF REGEXP_LIKE(l_token, '^FIELD[[:space:]]*:') THEN
                l_name := normalize_name(
                    REGEXP_REPLACE(
                        l_token,
                        '^FIELD[[:space:]]*:[[:space:]]*',
                        ''
                    ),
                    'La plantilla',
                    -20151
                );
                IF l_name = p_field_name THEN
                    RETURN TRUE;
                END IF;
            END IF;

            l_scan_pos := l_close_pos + 2;
        END LOOP;

        RETURN FALSE;
    END template_uses_field;


    FUNCTION resolve_text_template(
        p_template     IN VARCHAR2,
        p_title        IN VARCHAR2,
        p_fields       IN t_varchar_map,
        p_generated_at IN VARCHAR2,
        p_label        IN VARCHAR2
    ) RETURN VARCHAR2
    IS
        l_result      VARCHAR2(32767);
        l_scan_pos    PLS_INTEGER := 1;
        l_open_pos    PLS_INTEGER;
        l_close_pos   PLS_INTEGER;
        l_stray_close PLS_INTEGER;
        l_token       VARCHAR2(4000);
        l_field_name  VARCHAR2(4000);
        l_replacement VARCHAR2(32767);
        l_app_user    VARCHAR2(4000) := NVL(
            apex_application.g_user,
            'Usuario APEX'
        );
    BEGIN
        IF p_template IS NULL THEN
            RETURN NULL;
        END IF;

        LOOP
            l_open_pos := INSTR(p_template, '{{', l_scan_pos);
            l_stray_close := INSTR(p_template, '}}', l_scan_pos);

            IF l_open_pos = 0 THEN
                IF l_stray_close > 0 THEN
                    raise_application_error(
                        -20150,
                        'La plantilla de ' || p_label ||
                        ' contiene un marcador incompleto.'
                    );
                END IF;
                l_result := l_result || SUBSTR(p_template, l_scan_pos);
                EXIT;
            END IF;

            IF l_stray_close > 0 AND l_stray_close < l_open_pos THEN
                raise_application_error(
                    -20150,
                    'La plantilla de ' || p_label ||
                    ' contiene un marcador incompleto.'
                );
            END IF;

            l_result := l_result || SUBSTR(
                p_template,
                l_scan_pos,
                l_open_pos - l_scan_pos
            );

            l_close_pos := INSTR(p_template, '}}', l_open_pos + 2);

            IF l_close_pos = 0 THEN
                raise_application_error(
                    -20150,
                    'La plantilla de ' || p_label ||
                    ' contiene un marcador sin cerrar.'
                );
            END IF;

            l_token := UPPER(
                    TRIM(
                        SUBSTR(
                        p_template,
                        l_open_pos + 2,
                        l_close_pos - l_open_pos - 2
                    )
                )
            );

            IF l_token IN ('REPORT_TITLE', 'APP_USER', 'GENERATED_AT') THEN
                CASE l_token
                    WHEN 'REPORT_TITLE' THEN
                        l_replacement := NVL(TRIM(p_title), 'Reporte');
                    WHEN 'APP_USER' THEN
                        l_replacement := l_app_user;
                    WHEN 'GENERATED_AT' THEN
                        l_replacement := p_generated_at;
                END CASE;
            ELSIF REGEXP_LIKE(l_token, '^FIELD[[:space:]]*:') THEN
                l_field_name := normalize_name(
                    REGEXP_REPLACE(
                        l_token,
                        '^FIELD[[:space:]]*:[[:space:]]*',
                        ''
                    ),
                    'La plantilla de ' || p_label,
                    -20151
                );

                IF NOT p_fields.EXISTS(l_field_name) THEN
                    raise_application_error(
                        -20152,
                        'No existe un mapeo para {{FIELD:' ||
                        l_field_name || '}}.'
                    );
                END IF;
                l_replacement := p_fields(l_field_name);
            ELSE
                raise_application_error(
                    -20153,
                    'La plantilla de ' || p_label ||
                    ' contiene el marcador no permitido {{' ||
                    SUBSTR(l_token, 1, 200) || '}}.'
                );
            END IF;

            l_result := l_result || l_replacement;
            l_scan_pos := l_close_pos + 2;

            IF LENGTHB(l_result) > 4000 THEN
                raise_application_error(
                    -20154,
                    'El texto final de ' || p_label ||
                    ' supera el límite de 4000 bytes.'
                );
            END IF;
        END LOOP;

        IF LENGTHB(l_result) > 4000 THEN
            raise_application_error(
                -20154,
                'El texto final de ' || p_label ||
                ' supera el límite de 4000 bytes.'
            );
        END IF;

        RETURN l_result;
    END resolve_text_template;

    FUNCTION get_alignment(
        p_alignment IN VARCHAR2,
        p_data_type IN apex_exec.t_data_type
    ) RETURN apex_data_export.t_alignment
    IS
    BEGIN
        CASE LOWER(NVL(p_alignment, ''))
            WHEN 'end' THEN
                RETURN apex_data_export.c_align_end;
            WHEN 'right' THEN
                RETURN apex_data_export.c_align_end;
            WHEN 'center' THEN
                RETURN apex_data_export.c_align_center;
            WHEN 'start' THEN
                RETURN apex_data_export.c_align_start;
            WHEN 'left' THEN
                RETURN apex_data_export.c_align_start;
            ELSE
                IF p_data_type IN (
                    apex_exec.c_data_type_number,
                    apex_exec.c_data_type_binary_number
                ) THEN
                    RETURN apex_data_export.c_align_end;
                ELSE
                    RETURN apex_data_export.c_align_start;
                END IF;
        END CASE;
    END get_alignment;


    PROCEDURE download_ig(
        p_application_id       IN NUMBER,
        p_page_id              IN NUMBER,
        p_region_static_id     IN VARCHAR2,
        p_visible_columns_json IN CLOB,
        p_title                IN VARCHAR2,
        p_file_name            IN VARCHAR2,
        p_generated_by         IN VARCHAR2,
        p_format               IN VARCHAR2,
        p_orientation          IN VARCHAR2,
        p_max_rows             IN PLS_INTEGER,
        p_excluded_columns_json IN CLOB,
        p_column_spans_json    IN CLOB,
        p_display_columns_json IN CLOB
    )
    IS
        l_region_id          NUMBER;
        l_report_id          NUMBER;
        l_context            apex_exec.t_context;
        l_context_is_open    BOOLEAN := FALSE;

        l_export             apex_data_export.t_export;
        l_native_export      apex_data_export.t_export;
        l_print_config       apex_data_export.t_print_config;
        l_export_columns     apex_data_export.t_columns;
        l_export_format      apex_data_export.t_format;
        l_orientation        apex_data_export.t_orientation;

        l_column_position    PLS_INTEGER;
        l_column_metadata    apex_exec.t_column;
        l_column_count       PLS_INTEGER := 0;

        l_effective_title    VARCHAR2(255);
        l_base_file_name     VARCHAR2(180);
        l_export_file_name   VARCHAR2(255);
        l_generated_by       VARCHAR2(255);
        l_page_header        VARCHAR2(4000);
        l_page_footer        VARCHAR2(4000);
        l_native_json        CLOB;
        l_sql_query          VARCHAR2(32767);
        l_sql_fragment       VARCHAR2(32767);
        l_collection_name    VARCHAR2(255);
        l_collection_exists  BOOLEAN := FALSE;

        l_span_count         PLS_INTEGER := 0;
        l_total_span         NUMBER := 0;
        l_column_width       NUMBER;
        l_use_custom_spans   BOOLEAN := FALSE;
        l_auto_column_index  PLS_INTEGER;

        TYPE t_name_set IS TABLE OF BOOLEAN
            INDEX BY VARCHAR2(4000);

        TYPE t_number_map IS TABLE OF NUMBER
            INDEX BY PLS_INTEGER;

        l_excluded_columns t_name_set;
        l_column_spans     t_number_map;

        TYPE t_visible_column IS RECORD (
            column_name      VARCHAR2(4000),
            source_name      VARCHAR2(4000),
            column_label     VARCHAR2(4000),
            display_position PLS_INTEGER,
            alignment        VARCHAR2(20),
            data_type        apex_exec.t_data_type,
            format_mask      VARCHAR2(4000)
        );

        TYPE t_visible_columns IS TABLE OF t_visible_column
            INDEX BY PLS_INTEGER;

        l_visible_columns t_visible_columns;

    BEGIN
        IF p_application_id IS NULL OR p_page_id IS NULL THEN
            raise_application_error(
                -20001,
                'No se recibió la aplicación o página del reporte.'
            );
        END IF;

        IF TRIM(p_region_static_id) IS NULL THEN
            raise_application_error(
                -20002,
                'No se recibió el Static ID del Interactive Grid.'
            );
        END IF;

        IF p_visible_columns_json IS NULL THEN
            raise_application_error(
                -20003,
                'No se recibieron las columnas visibles del reporte.'
            );
        END IF;

        CASE UPPER(NVL(TRIM(p_format), 'PDF'))
            WHEN 'PDF' THEN
                l_export_format := apex_data_export.c_format_pdf;
            WHEN 'XLSX' THEN
                l_export_format := apex_data_export.c_format_xlsx;
            ELSE
                raise_application_error(
                    -20009,
                    'El formato debe ser PDF o XLSX.'
                );
        END CASE;

        l_effective_title := SUBSTR(
            NVL(TRIM(p_title), 'Reporte'),
            1,
            255
        );

        l_base_file_name := SUBSTR(
            REGEXP_REPLACE(
                NVL(TRIM(p_file_name), 'reporte'),
                '[^A-Za-z0-9_-]+',
                '_'
            ),
            1,
            180
        );

        IF l_base_file_name IS NULL THEN
            l_base_file_name := 'reporte';
        END IF;

        l_generated_by := SUBSTR(
            NVL(TRIM(p_generated_by), 'Usuario APEX'),
            1,
            255
        );

        l_page_header := l_effective_title;

        l_page_footer := SUBSTR(
            'Usuario: ' || l_generated_by ||
            '  |  Fecha: ' ||
            TO_CHAR(SYSTIMESTAMP, 'DD/MM/YYYY HH24:MI'),
            1,
            4000
        );

        /*
         * Las exclusiones son configuradas por el desarrollador.
         * Las comparaciones no distinguen mayúsculas.
         */
        IF p_excluded_columns_json IS NOT NULL THEN
            FOR excluded_record IN (
                SELECT column_name
                  FROM JSON_TABLE(
                           p_excluded_columns_json,
                           '$[*]'
                           COLUMNS (
                               column_name VARCHAR2(4000)
                                   PATH '$'
                                   NULL ON ERROR
                           )
                       )
            )
            LOOP
                IF TRIM(excluded_record.column_name) IS NOT NULL THEN
                    l_excluded_columns(
                        UPPER(TRIM(excluded_record.column_name))
                    ) := TRUE;
                END IF;
            END LOOP;
        END IF;

        BEGIN
            SELECT region_id
              INTO l_region_id
              FROM apex_application_page_regions
             WHERE application_id = p_application_id
               AND page_id = p_page_id
               AND static_id = p_region_static_id;
        EXCEPTION
            WHEN NO_DATA_FOUND THEN
                raise_application_error(
                    -20005,
                    'No existe una región con Static ID "' ||
                    SUBSTR(p_region_static_id, 1, 200) ||
                    '" en la página indicada.'
                );
            WHEN TOO_MANY_ROWS THEN
                raise_application_error(
                    -20006,
                    'El Static ID de la región no es único en la página.'
                );
        END;

        l_report_id := apex_ig.get_last_viewed_report_id(
            p_page_id   => p_page_id,
            p_region_id => l_region_id
        );

        /*
         * El navegador solo define el orden y las columnas visibles.
         * Los datos nunca se aceptan desde el cliente.
         */
        FOR column_record IN (
            SELECT column_name,
                   column_label,
                   display_position,
                   alignment
              FROM JSON_TABLE(
                       p_visible_columns_json,
                       '$[*]'
                       COLUMNS (
                           column_name      VARCHAR2(4000)
                               PATH '$.name',
                           column_label     VARCHAR2(4000)
                               PATH '$.label',
                           display_position NUMBER
                               PATH '$.position',
                           alignment        VARCHAR2(20)
                               PATH '$.alignment'
                       )
                   )
             ORDER BY display_position
        )
        LOOP
            IF l_excluded_columns.EXISTS(
                   UPPER(TRIM(column_record.column_name))
               )
            THEN
                CONTINUE;
            END IF;

            IF TRIM(column_record.column_name) IS NULL OR
               NOT REGEXP_LIKE(
                   TRIM(column_record.column_name),
                   '^[A-Za-z][A-Za-z0-9_$#]*$'
               )
            THEN
                raise_application_error(
                    -20010,
                    'El reporte contiene un nombre de columna no permitido.'
                );
            END IF;

            l_column_count := l_column_count + 1;

            l_visible_columns(l_column_count).column_name :=
                UPPER(TRIM(column_record.column_name));

            l_visible_columns(l_column_count).source_name :=
                UPPER(TRIM(column_record.column_name));

            l_visible_columns(l_column_count).column_label :=
                SUBSTR(
                    NVL(
                        column_record.column_label,
                        column_record.column_name
                    ),
                    1,
                    4000
                );

            l_visible_columns(l_column_count).display_position :=
                column_record.display_position;

            l_visible_columns(l_column_count).alignment :=
                column_record.alignment;

        END LOOP;

        IF l_column_count = 0 THEN
            raise_application_error(
                -20007,
                'Debe existir al menos una columna visible.'
            );
        END IF;

        /*
         * APEX_REGION.EXPORT_DATA aplica el reporte actual del IG y, a
         * diferencia de OPEN_QUERY_CONTEXT, serializa el display value de
         * las columnas LOV. El JSON nativo es la única fuente de datos para
         * el PDF corporativo.
         */
        l_native_export := apex_region.export_data(
            p_format       => apex_data_export.c_format_json,
            p_page_id      => p_page_id,
            p_region_id    => l_region_id,
            p_component_id => l_report_id,
            p_as_clob      => TRUE,
            p_file_name    => 'corporate_report_source',
            p_data_only    => TRUE
        );

        l_native_json := l_native_export.content_clob;

        IF l_native_json IS NULL THEN
            raise_application_error(
                -20014,
                'APEX no devolvió datos JSON para el reporte.'
            );
        END IF;

        l_collection_name :=
            'CORP_RPT_' || SUBSTR(RAWTOHEX(SYS_GUID()), 1, 24);

        apex_collection.create_or_truncate_collection(
            p_collection_name => l_collection_name
        );
        l_collection_exists := TRUE;

        apex_collection.add_member(
            p_collection_name => l_collection_name,
            p_clob001         => l_native_json
        );

        /*
         * Se proyectan exclusivamente las columnas finales. Los nombres
         * JSON nativos de APEX se generan en minúsculas, como "cod_dpto".
         * Cada valor se lee como texto para conservar exactamente su valor
         * de presentación (incluidos los LOV y formatos del IG).
         */
        l_sql_query := 'SELECT ';

        FOR column_index IN 1 .. l_column_count
        LOOP
            IF column_index > 1 THEN
                l_sql_query := l_sql_query || ', ';
            END IF;

            l_sql_query := l_sql_query ||
                'jt.' ||
                DBMS_ASSERT.ENQUOTE_NAME(
                    l_visible_columns(column_index).source_name,
                    FALSE
                );
        END LOOP;

        l_sql_query := l_sql_query ||
            ' FROM apex_collections c, JSON_TABLE(' ||
            'c.clob001, ''$.items[*]'' COLUMNS (';

        FOR column_index IN 1 .. l_column_count
        LOOP
            IF column_index > 1 THEN
                l_sql_query := l_sql_query || ', ';
            END IF;

            l_sql_fragment :=
                DBMS_ASSERT.ENQUOTE_NAME(
                    l_visible_columns(column_index).source_name,
                    FALSE
                ) ||
                ' VARCHAR2(4000) PATH ''$."' ||
                LOWER(l_visible_columns(column_index).column_name) ||
                '"'' NULL ON ERROR';

            IF LENGTH(l_sql_query) + LENGTH(l_sql_fragment) > 32000 THEN
                raise_application_error(
                    -20015,
                    'La definición de columnas supera el límite permitido.'
                );
            END IF;

            l_sql_query := l_sql_query || l_sql_fragment;
        END LOOP;

        l_sql_query := l_sql_query ||
            ')) jt WHERE c.collection_name = ''' ||
            l_collection_name || ''' AND c.seq_id = 1';

        l_context := apex_exec.open_query_context(
            p_location  => apex_exec.c_location_local_db,
            p_sql_query => l_sql_query
        );
        l_context_is_open := TRUE;

        FOR column_index IN 1 .. l_column_count
        LOOP
            l_column_position := apex_exec.get_column_position(
                p_context         => l_context,
                p_column_name     =>
                    l_visible_columns(column_index).source_name,
                p_attribute_label =>
                    l_visible_columns(column_index).column_label,
                p_is_required     => TRUE,
                p_data_type       => NULL
            );

            l_column_metadata := apex_exec.get_column(
                p_context    => l_context,
                p_column_idx => l_column_position
            );

            l_visible_columns(column_index).data_type :=
                l_column_metadata.data_type;

            l_visible_columns(column_index).format_mask := NULL;
        END LOOP;

        /*
         * Los spans son pesos relativos. Si faltan, las columnas restantes
         * reciben peso 1. Los valores sobrantes se ignoran.
         * Si no se proporciona ningún span, se conserva el ajuste nativo de
         * APEX_DATA_EXPORT dejando p_width en NULL.
         */
        IF p_column_spans_json IS NOT NULL THEN
            FOR span_record IN (
                SELECT span_position,
                       span_token,
                       span_value
                  FROM JSON_TABLE(
                           p_column_spans_json,
                           '$[*]'
                           COLUMNS (
                               span_position FOR ORDINALITY,
                               span_token VARCHAR2(100) PATH '$'
                                   NULL ON ERROR,
                               span_value NUMBER PATH '$'
                                   NULL ON ERROR
                           )
                       )
                 WHERE span_position <= l_column_count
                 ORDER BY span_position
            )
            LOOP
                /*
                 * "*" marca la columna que APEX debe dimensionar con el
                 * espacio restante. Solo puede existir una marca aplicable.
                 */
                IF TRIM(span_record.span_token) = '*' THEN
                    IF l_auto_column_index IS NOT NULL THEN
                        raise_application_error(
                            -20013,
                            'Solo se permite un span adaptable "*".'
                        );
                    END IF;

                    l_auto_column_index := span_record.span_position;
                    l_column_spans(span_record.span_position) := 1;

                ELSIF span_record.span_value IS NULL OR
                      span_record.span_value <= 0
                THEN
                    raise_application_error(
                        -20011,
                        'Cada span debe ser un número mayor que cero o "*".'
                    );

                ELSE
                    l_column_spans(span_record.span_position) :=
                        span_record.span_value;
                END IF;

                l_span_count := l_span_count + 1;
            END LOOP;

            l_use_custom_spans := l_span_count > 0;
        END IF;

        IF l_use_custom_spans THEN
            FOR column_index IN 1 .. l_column_count
            LOOP
                IF NOT l_column_spans.EXISTS(column_index) THEN
                    l_column_spans(column_index) := 1;
                END IF;

                l_total_span :=
                    l_total_span + l_column_spans(column_index);
            END LOOP;

            IF l_total_span <= 0 THEN
                raise_application_error(
                    -20012,
                    'La suma de spans debe ser mayor que cero.'
                );
            END IF;

            /*
             * Sin "*" explícito, se mantiene la compatibilidad anterior:
             * la última columna queda adaptable.
             */
            IF l_auto_column_index IS NULL THEN
                l_auto_column_index := l_column_count;
            END IF;
        END IF;

        CASE LOWER(NVL(TRIM(p_orientation), 'auto'))
            WHEN 'portrait' THEN
                l_orientation :=
                    apex_data_export.c_orientation_portrait;

            WHEN 'landscape' THEN
                l_orientation :=
                    apex_data_export.c_orientation_landscape;

            WHEN 'auto' THEN
                IF l_column_count <= 3 THEN
                    l_orientation :=
                        apex_data_export.c_orientation_portrait;
                ELSE
                    l_orientation :=
                        apex_data_export.c_orientation_landscape;
                END IF;

            ELSE
                raise_application_error(
                    -20008,
                    'La orientación debe ser AUTO, PORTRAIT o LANDSCAPE.'
                );
        END CASE;

        FOR column_index IN 1 .. l_column_count
        LOOP
            /*
             * La columna marcada con "*" queda automática para que APEX
             * absorba el espacio restante. Si no hay marca, se utiliza la
             * última. El factor 99 reserva además un margen técnico.
             */
            IF l_use_custom_spans AND
               column_index <> l_auto_column_index
            THEN
                l_column_width := ROUND(
                    l_column_spans(column_index) /
                    l_total_span * 99,
                    4
                );
            ELSE
                l_column_width := NULL;
            END IF;

            IF l_column_width IS NULL THEN
                apex_data_export.add_column(
                    p_columns           => l_export_columns,
                    p_name              =>
                        l_visible_columns(column_index).source_name,
                    p_heading           =>
                        l_visible_columns(column_index).column_label,
                    p_format_mask       =>
                        l_visible_columns(column_index).format_mask,
                    p_heading_alignment =>
                        apex_data_export.c_align_center,
                    p_value_alignment   =>
                        get_alignment(
                            l_visible_columns(column_index).alignment,
                            l_visible_columns(column_index).data_type
                        )
                );
            ELSE
                apex_data_export.add_column(
                    p_columns           => l_export_columns,
                    p_name              =>
                        l_visible_columns(column_index).source_name,
                    p_heading           =>
                        l_visible_columns(column_index).column_label,
                    p_format_mask       =>
                        l_visible_columns(column_index).format_mask,
                    p_heading_alignment =>
                        apex_data_export.c_align_center,
                    p_value_alignment   =>
                        get_alignment(
                            l_visible_columns(column_index).alignment,
                            l_visible_columns(column_index).data_type
                        ),
                    p_width             => l_column_width
                );
            END IF;
        END LOOP;

        l_print_config := apex_data_export.get_print_config(
            p_paper_size              =>
                apex_data_export.c_size_a4,

            p_width_units             =>
                apex_data_export.c_width_unit_percentage,

            p_orientation             =>
                l_orientation,

            p_page_header             =>
                l_page_header,

            p_page_header_font_color  =>
                '#2F343A',

            p_page_header_font_family =>
                apex_data_export.c_font_family_helvetica,

            p_page_header_font_weight =>
                apex_data_export.c_font_weight_bold,

            p_page_header_font_size   =>
                15,

            p_page_header_alignment   =>
                apex_data_export.c_align_center,

            p_page_footer             =>
                l_page_footer,

            p_page_footer_font_color  =>
                '#666666',

            p_page_footer_font_family =>
                apex_data_export.c_font_family_helvetica,

            p_page_footer_font_size   =>
                8,

            p_page_footer_alignment   =>
                apex_data_export.c_align_center,

            p_header_bg_color         =>
                '#4A4F55',

            p_header_font_color       =>
                '#FFFFFF',

            p_header_font_family      =>
                apex_data_export.c_font_family_helvetica,

            p_header_font_weight      =>
                apex_data_export.c_font_weight_bold,

            p_header_font_size        =>
                9,

            p_body_bg_color           =>
                '#FFFFFF',

            p_body_font_color         =>
                '#25282B',

            p_body_font_family        =>
                apex_data_export.c_font_family_helvetica,

            p_body_font_size          =>
                8.5,

            p_border_width            =>
                0.5,

            p_border_color            =>
                '#BFC3C7'
        );

        l_export_file_name :=
            l_base_file_name || '_' ||
            TO_CHAR(SYSDATE, 'YYYYMMDD_HH24MISS');

        IF l_export_format = apex_data_export.c_format_pdf THEN
            l_export := apex_data_export.export(
                p_context        => l_context,
                p_format         => l_export_format,
                p_columns        => l_export_columns,
                p_file_name      => l_export_file_name,
                p_print_config   => l_print_config,
                p_pdf_accessible => TRUE
            );
        ELSE
            l_export := apex_data_export.export(
                p_context      => l_context,
                p_format       => l_export_format,
                p_columns      => l_export_columns,
                p_file_name    => l_export_file_name,
                p_print_config => l_print_config
            );
        END IF;

        apex_exec.close(l_context);
        l_context_is_open := FALSE;

        apex_collection.delete_collection(
            p_collection_name => l_collection_name
        );
        l_collection_exists := FALSE;

        apex_data_export.download(
            p_export              => l_export,
            p_content_disposition =>
                apex_data_export.c_attachment,
            p_add_file_extension  => TRUE
        );

    EXCEPTION
        WHEN apex_application.e_stop_apex_engine THEN
            IF l_context_is_open THEN
                apex_exec.close(l_context);
                l_context_is_open := FALSE;
            END IF;

            IF l_collection_exists THEN
                apex_collection.delete_collection(
                    p_collection_name => l_collection_name
                );
                l_collection_exists := FALSE;
            END IF;

            RAISE;

        WHEN OTHERS THEN
            IF l_context_is_open THEN
                apex_exec.close(l_context);
            END IF;

            IF l_collection_exists THEN
                apex_collection.delete_collection(
                    p_collection_name => l_collection_name
                );
            END IF;

            RAISE;
    END download_ig;


    PROCEDURE download_query(
        p_sql_query              IN VARCHAR2,
        p_columns_json           IN CLOB,
        p_bindings_json          IN CLOB,
        p_fields_json            IN CLOB,
        p_style_json             IN CLOB,
        p_header_template        IN VARCHAR2,
        p_footer_template        IN VARCHAR2,
        p_title                  IN VARCHAR2,
        p_file_name              IN VARCHAR2,
        p_format                 IN VARCHAR2,
        p_orientation            IN VARCHAR2,
        p_excluded_columns_json  IN CLOB,
        p_column_widths_json     IN CLOB
    )
    IS
        l_context            apex_exec.t_context;
        l_context_is_open    BOOLEAN := FALSE;
        l_parameters         apex_exec.t_parameters;
        l_export             apex_data_export.t_export;
        l_export_columns     apex_data_export.t_columns;
        l_print_config       apex_data_export.t_print_config;
        l_export_format      apex_data_export.t_format;
        l_paper_size         apex_data_export.t_size;
        l_orientation        apex_data_export.t_orientation;
        l_style              t_style_config;

        l_effective_title    VARCHAR2(255);
        l_base_file_name     VARCHAR2(180);
        l_export_file_name   VARCHAR2(255);
        l_page_header        VARCHAR2(4000);
        l_page_footer        VARCHAR2(4000);
        l_generated_at       VARCHAR2(100);

        l_bind_names         t_name_set;
        l_sql_bind_names     t_name_set;
        l_field_names        t_name_set;
        l_field_values       t_varchar_map;
        l_excluded_columns   t_name_set;
        l_all_columns        t_name_set;
        l_width_names        t_name_set;
        l_width_modes        t_varchar_map;
        l_width_values       t_number_name_map;

        l_varchar_value      VARCHAR2(32767);
        l_number_value       NUMBER;
        l_date_value         DATE;
        l_timestamp_value    TIMESTAMP;
        l_field_value        VARCHAR2(32767);
        l_required           BOOLEAN;
        l_data_type          VARCHAR2(30);
        l_source             VARCHAR2(30);
        l_context_key        VARCHAR2(128);
        l_name               VARCHAR2(128);
        l_item_name          VARCHAR2(128);
        l_format_mask        VARCHAR2(4000);
        l_key                VARCHAR2(4000);
        l_template_text      VARCHAR2(32767);

        l_column_position    PLS_INTEGER;
        l_column_metadata    apex_exec.t_column;
        l_column_count       PLS_INTEGER := 0;
        l_template_count     PLS_INTEGER := 0;
        l_auto_count         PLS_INTEGER := 0;
        l_fixed_total        NUMBER := 0;
        l_weight_total       NUMBER := 0;
        l_weight_denominator NUMBER := 0;
        l_remaining_width    NUMBER;
        l_weight_budget      NUMBER;
        l_column_width       NUMBER;
        l_width_mode         VARCHAR2(30);
        l_width_value        NUMBER;

        TYPE t_query_column IS RECORD (
            column_name       VARCHAR2(128),
            column_heading    VARCHAR2(255),
            value_alignment   VARCHAR2(20),
            heading_alignment VARCHAR2(20),
            format_mask       VARCHAR2(4000),
            data_type         apex_exec.t_data_type
        );

        TYPE t_query_columns IS TABLE OF t_query_column
            INDEX BY PLS_INTEGER;

        l_columns t_query_columns;

        FUNCTION is_true(p_value IN VARCHAR2) RETURN BOOLEAN
        IS
            l_value VARCHAR2(20) := UPPER(NVL(TRIM(p_value), 'FALSE'));
        BEGIN
            IF l_value IN ('TRUE', '1', 'Y', 'YES', 'SI', 'SÍ') THEN
                RETURN TRUE;
            ELSIF l_value IN ('FALSE', '0', 'N', 'NO') THEN
                RETURN FALSE;
            ELSE
                raise_application_error(
                    -20110,
                    'La propiedad required de un bind debe ser true o false.'
                );
            END IF;
        END is_true;

    BEGIN
        assert_read_only_query(p_sql_query);
        collect_sql_bind_names(p_sql_query, l_sql_bind_names);

        IF p_columns_json IS NULL THEN
            raise_application_error(
                -20106,
                'No se recibió la definición de columnas del reporte.'
            );
        END IF;

        assert_json(p_columns_json, 'columnas', -20108);
        assert_json(p_bindings_json, 'binds', -20109);
        assert_json(p_fields_json, 'campos', -20111);
        assert_json(p_excluded_columns_json, 'exclusiones', -20112);
        assert_json(p_column_widths_json, 'anchos', -20113);

        CASE UPPER(NVL(TRIM(p_format), 'PDF'))
            WHEN 'PDF' THEN
                l_export_format := apex_data_export.c_format_pdf;
            WHEN 'XLSX' THEN
                l_export_format := apex_data_export.c_format_xlsx;
            ELSE
                raise_application_error(
                    -20114,
                    'El formato debe ser PDF o XLSX.'
                );
        END CASE;

        /* El contrato congelado utiliza siempre papel A4. */
        l_paper_size := apex_data_export.c_size_a4;

        l_effective_title := SUBSTRB(
            NVL(TRIM(p_title), 'Reporte'),
            1,
            255
        );

        /* Lista explícita: los rangos A-Z dependen de NLS_SORT. */
        l_base_file_name := SUBSTRB(
            REGEXP_REPLACE(
                NVL(TRIM(p_file_name), 'reporte'),
                '[^ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-]+',
                '_'
            ),
            1,
            180
        );

        IF l_base_file_name IS NULL THEN
            l_base_file_name := 'reporte';
        END IF;

        l_style := parse_style(p_style_json);

        /*
         * Los nombres de bind y de Page Item son configuración protegida.
         * p_auto_bind_items=FALSE impide que APEX resuelva otros items por
         * coincidencia accidental de nombres.
         */
        IF p_bindings_json IS NOT NULL THEN
            FOR bind_record IN (
                SELECT bind_name,
                       item_name,
                       data_type,
                       required_flag,
                       format_mask
                  FROM JSON_TABLE(
                           p_bindings_json,
                           '$[*]'
                           COLUMNS (
                               bind_name VARCHAR2(128)
                                   PATH '$.bind' ERROR ON ERROR,
                               item_name VARCHAR2(128)
                                   PATH '$.item' ERROR ON ERROR,
                               data_type VARCHAR2(30)
                                   PATH '$.type' NULL ON ERROR,
                               required_flag VARCHAR2(20)
                                   PATH '$.required' NULL ON ERROR,
                               format_mask VARCHAR2(4000)
                                   PATH '$.format_mask' NULL ON ERROR
                           )
                       )
            )
            LOOP
                l_name := normalize_name(
                    bind_record.bind_name,
                    'El mapeo de binds',
                    -20115
                );
                l_item_name := normalize_item_name(
                    bind_record.item_name,
                    'El mapeo de Page Items',
                    -20116
                );

                IF l_name IN (
                       'APP_USER', 'APP_ID', 'APP_PAGE_ID', 'APP_SESSION',
                       'APP_ALIAS', 'APP_PAGE_ALIAS', 'APP_BUILDER_SESSION',
                       'SESSION', 'REQUEST', 'DEBUG', 'WORKSPACE_ID',
                       'APP_REQUEST_DATA_HASH', 'APP_SESSION_VISIBLE'
                   )
                THEN
                    raise_application_error(
                        -20173,
                        'El bind "' || l_name ||
                        '" es un nombre reservado de APEX.'
                    );
                END IF;

                IF l_bind_names.EXISTS(l_name) THEN
                    raise_application_error(
                        -20118,
                        'El bind "' || l_name || '" está repetido.'
                    );
                END IF;
                l_bind_names(l_name) := TRUE;

                IF NOT l_sql_bind_names.EXISTS(l_name) THEN
                    raise_application_error(
                        -20117,
                        'El mapeo declara el bind "' || l_name ||
                        '", pero la consulta no lo utiliza.'
                    );
                END IF;

                l_data_type := UPPER(
                    NVL(TRIM(bind_record.data_type), 'VARCHAR2')
                );
                l_required := is_true(bind_record.required_flag);
                l_format_mask := bind_record.format_mask;

                BEGIN
                    CASE l_data_type
                        WHEN 'VARCHAR2' THEN
                            l_varchar_value :=
                                apex_session_state.get_varchar2(l_item_name);

                            IF l_required AND l_varchar_value IS NULL THEN
                                raise_application_error(
                                    -20119,
                                    'El parámetro "' || l_name ||
                                    '" es obligatorio.'
                                );
                            END IF;

                            apex_exec.add_parameter(
                                p_parameters => l_parameters,
                                p_name       => l_name,
                                p_value      => l_varchar_value
                            );

                        WHEN 'NUMBER' THEN
                            l_number_value :=
                                apex_session_state.get_number(l_item_name);

                            IF l_required AND l_number_value IS NULL THEN
                                raise_application_error(
                                    -20119,
                                    'El parámetro "' || l_name ||
                                    '" es obligatorio.'
                                );
                            END IF;

                            apex_exec.add_parameter(
                                p_parameters => l_parameters,
                                p_name       => l_name,
                                p_value      => l_number_value
                            );

                        WHEN 'TIMESTAMP' THEN
                            IF l_format_mask IS NOT NULL THEN
                                l_varchar_value :=
                                    apex_session_state.get_varchar2(
                                        l_item_name
                                    );
                                IF l_varchar_value IS NULL THEN
                                    l_timestamp_value := NULL;
                                ELSE
                                    l_timestamp_value := TO_TIMESTAMP(
                                        l_varchar_value,
                                        l_format_mask
                                    );
                                END IF;
                            ELSE
                                l_timestamp_value :=
                                    apex_session_state.get_timestamp(
                                        l_item_name
                                    );
                            END IF;

                            IF l_required AND l_timestamp_value IS NULL THEN
                                raise_application_error(
                                    -20119,
                                    'El parámetro "' || l_name ||
                                    '" es obligatorio.'
                                );
                            END IF;

                            apex_exec.add_parameter(
                                p_parameters => l_parameters,
                                p_name       => l_name,
                                p_value      => l_timestamp_value
                            );

                        WHEN 'DATE' THEN
                            IF l_format_mask IS NOT NULL THEN
                                l_varchar_value :=
                                    apex_session_state.get_varchar2(
                                        l_item_name
                                    );
                                IF l_varchar_value IS NULL THEN
                                    l_date_value := NULL;
                                ELSE
                                    l_date_value := TO_DATE(
                                        l_varchar_value,
                                        l_format_mask
                                    );
                                END IF;
                            ELSE
                                l_timestamp_value :=
                                    apex_session_state.get_timestamp(
                                        l_item_name
                                    );
                                l_date_value := CAST(
                                    l_timestamp_value AS DATE
                                );
                            END IF;

                            IF l_required AND l_date_value IS NULL THEN
                                raise_application_error(
                                    -20119,
                                    'El parámetro "' || l_name ||
                                    '" es obligatorio.'
                                );
                            END IF;

                            apex_exec.add_parameter(
                                p_parameters => l_parameters,
                                p_name       => l_name,
                                p_value      => l_date_value
                            );

                        ELSE
                            raise_application_error(
                                -20120,
                                'El tipo del bind "' || l_name ||
                                '" debe ser VARCHAR2, NUMBER, DATE o TIMESTAMP.'
                            );
                    END CASE;
                EXCEPTION
                    WHEN OTHERS THEN
                        /*
                         * Solo se propagan los errores propios de este
                         * bloque. Un error de APEX_SESSION_STATE podría
                         * incluir el valor recibido en su mensaje.
                         */
                        IF SQLCODE IN (-20119, -20120) THEN
                            RAISE;
                        END IF;

                        raise_application_error(
                            -20121,
                            'No se pudo convertir el parámetro "' ||
                            l_name || '" al tipo ' || l_data_type || '.'
                        );
                END;
            END LOOP;
        END IF;

        l_key := l_sql_bind_names.FIRST;
        WHILE l_key IS NOT NULL
        LOOP
            IF NOT l_bind_names.EXISTS(l_key) THEN
                raise_application_error(
                    -20168,
                    'La consulta usa el bind "' || l_key ||
                    '", pero no existe un mapeo autorizado.'
                );
            END IF;
            l_key := l_sql_bind_names.NEXT(l_key);
        END LOOP;

        /* Campos escalares restringidos para header y footer. */
        IF p_fields_json IS NOT NULL THEN
            FOR field_record IN (
                SELECT field_name,
                       field_source,
                       item_name,
                       data_type,
                       context_key,
                       constant_value,
                       COALESCE(format_mask_value, format_value) AS format_mask
                  FROM JSON_TABLE(
                           p_fields_json,
                           '$[*]'
                           COLUMNS (
                               field_name VARCHAR2(128)
                                   PATH '$.name' ERROR ON ERROR,
                               field_source VARCHAR2(30)
                                   PATH '$.source' ERROR ON ERROR,
                               item_name VARCHAR2(128)
                                   PATH '$.item' NULL ON ERROR,
                               data_type VARCHAR2(30)
                                   PATH '$.type' NULL ON ERROR,
                               context_key VARCHAR2(128)
                                   PATH '$.key' NULL ON ERROR,
                               constant_value VARCHAR2(4000)
                                   PATH '$.value' NULL ON ERROR,
                               format_mask_value VARCHAR2(4000)
                                   PATH '$.format_mask' NULL ON ERROR,
                               format_value VARCHAR2(4000)
                                   PATH '$.format' NULL ON ERROR
                           )
                       )
            )
            LOOP
                l_name := normalize_name(
                    field_record.field_name,
                    'El mapeo de campos',
                    -20122
                );

                IF l_name IN ('REPORT_TITLE', 'APP_USER', 'GENERATED_AT') THEN
                    raise_application_error(
                        -20123,
                        'El campo "' || l_name || '" usa un nombre reservado.'
                    );
                END IF;

                IF l_field_names.EXISTS(l_name) THEN
                    raise_application_error(
                        -20124,
                        'El campo "' || l_name || '" está repetido.'
                    );
                END IF;
                l_field_names(l_name) := TRUE;

                l_source := UPPER(TRIM(field_record.field_source));
                l_data_type := UPPER(
                    NVL(TRIM(field_record.data_type), 'VARCHAR2')
                );
                l_format_mask := field_record.format_mask;

                BEGIN
                    CASE l_source
                        WHEN 'ITEM' THEN
                            l_item_name := normalize_item_name(
                                field_record.item_name,
                                'El campo "' || l_name || '"',
                                -20125
                            );

                            CASE l_data_type
                                WHEN 'VARCHAR2' THEN
                                    l_field_value :=
                                        apex_session_state.get_varchar2(
                                            l_item_name
                                        );
                                WHEN 'NUMBER' THEN
                                    l_number_value :=
                                        apex_session_state.get_number(
                                            l_item_name
                                        );
                                    IF l_number_value IS NULL THEN
                                        l_field_value := NULL;
                                    ELSIF l_format_mask IS NULL THEN
                                        l_field_value := TO_CHAR(l_number_value);
                                    ELSE
                                        l_field_value := TO_CHAR(
                                            l_number_value,
                                            l_format_mask
                                        );
                                    END IF;
                                WHEN 'DATE' THEN
                                    l_timestamp_value :=
                                        apex_session_state.get_timestamp(
                                            l_item_name
                                        );
                                    IF l_timestamp_value IS NULL THEN
                                        l_field_value := NULL;
                                    ELSE
                                        l_field_value := TO_CHAR(
                                            CAST(l_timestamp_value AS DATE),
                                            NVL(l_format_mask, 'DD/MM/YYYY')
                                        );
                                    END IF;
                                WHEN 'TIMESTAMP' THEN
                                    l_timestamp_value :=
                                        apex_session_state.get_timestamp(
                                            l_item_name
                                        );
                                    IF l_timestamp_value IS NULL THEN
                                        l_field_value := NULL;
                                    ELSE
                                        l_field_value := TO_CHAR(
                                            l_timestamp_value,
                                            NVL(
                                                l_format_mask,
                                                'DD/MM/YYYY HH24:MI'
                                            )
                                        );
                                    END IF;
                                ELSE
                                    raise_application_error(
                                        -20127,
                                        'El tipo del campo "' || l_name ||
                                        '" debe ser VARCHAR2, NUMBER, DATE ' ||
                                        'o TIMESTAMP.'
                                    );
                            END CASE;

                        WHEN 'CONSTANT' THEN
                            l_field_value := field_record.constant_value;

                        WHEN 'CONTEXT' THEN
                            l_context_key := UPPER(
                                NVL(
                                    TRIM(field_record.context_key),
                                    l_name
                                )
                            );
                            CASE l_context_key
                                WHEN 'APP_USER' THEN
                                    l_field_value := apex_application.g_user;
                                WHEN 'APP_ID' THEN
                                    l_field_value := TO_CHAR(
                                        apex_application.g_flow_id
                                    );
                                WHEN 'APP_PAGE_ID' THEN
                                    l_field_value := TO_CHAR(
                                        apex_application.g_flow_step_id
                                    );
                                ELSE
                                    raise_application_error(
                                        -20128,
                                        'La clave CONTEXT del campo "' ||
                                        l_name || '" no está permitida.'
                                    );
                            END CASE;

                        WHEN 'SYSTEM' THEN
                            l_context_key := UPPER(
                                NVL(
                                    TRIM(field_record.context_key),
                                    l_name
                                )
                            );
                            IF l_context_key <> 'GENERATED_AT' THEN
                                raise_application_error(
                                    -20129,
                                    'La clave SYSTEM del campo "' || l_name ||
                                    '" no está permitida.'
                                );
                            END IF;
                            l_field_value := TO_CHAR(
                                SYSTIMESTAMP,
                                NVL(
                                    l_format_mask,
                                    'DD/MM/YYYY HH24:MI'
                                )
                            );

                        WHEN 'BUILTIN' THEN
                            /* Compatibilidad con manifiestos preliminares. */
                            l_context_key := UPPER(
                                NVL(
                                    TRIM(field_record.context_key),
                                    l_name
                                )
                            );
                            CASE l_context_key
                                WHEN 'APP_USER' THEN
                                    l_field_value := apex_application.g_user;
                                WHEN 'APP_ID' THEN
                                    l_field_value := TO_CHAR(
                                        apex_application.g_flow_id
                                    );
                                WHEN 'APP_PAGE_ID' THEN
                                    l_field_value := TO_CHAR(
                                        apex_application.g_flow_step_id
                                    );
                                WHEN 'GENERATED_AT' THEN
                                    l_field_value := TO_CHAR(
                                        SYSTIMESTAMP,
                                        NVL(
                                            l_format_mask,
                                            'DD/MM/YYYY HH24:MI'
                                        )
                                    );
                                ELSE
                                    raise_application_error(
                                        -20130,
                                        'La clave BUILTIN del campo "' ||
                                        l_name || '" no está permitida.'
                                    );
                            END CASE;

                        ELSE
                            raise_application_error(
                                -20131,
                                'El origen del campo "' || l_name ||
                                '" debe ser ITEM, CONSTANT, CONTEXT o SYSTEM.'
                            );
                    END CASE;
                EXCEPTION
                    WHEN OTHERS THEN
                        /* Errores propios de este bloque; el resto se sanea. */
                        IF SQLCODE IN (
                               -20125, -20127, -20128,
                               -20129, -20130, -20131
                           )
                        THEN
                            RAISE;
                        END IF;

                        raise_application_error(
                            -20132,
                            'No se pudo resolver el campo "' || l_name || '".'
                        );
                END;

                l_field_values(l_name) := SUBSTR(l_field_value, 1, 4000);
            END LOOP;
        END IF;

        /* Un mapeo escalar sobrante suele ser un error de mantenimiento. */
        l_template_text := UPPER(
            NVL(p_header_template, '{{REPORT_TITLE}}') || CHR(10) ||
            NVL(p_footer_template, '')
        );
        l_key := l_field_names.FIRST;
        WHILE l_key IS NOT NULL
        LOOP
            IF NOT template_uses_field(l_template_text, l_key)
            THEN
                raise_application_error(
                    -20169,
                    'El campo "' || l_key ||
                    '" está mapeado, pero no se usa en encabezado o pie.'
                );
            END IF;
            l_key := l_field_names.NEXT(l_key);
        END LOOP;

        l_generated_at := TO_CHAR(
            SYSTIMESTAMP,
            'DD/MM/YYYY HH24:MI'
        );

        l_page_header := resolve_text_template(
            p_template     => NVL(p_header_template, '{{REPORT_TITLE}}'),
            p_title        => l_effective_title,
            p_fields       => l_field_values,
            p_generated_at => l_generated_at,
            p_label        => 'encabezado'
        );

        l_page_footer := resolve_text_template(
            p_template     => p_footer_template,
            p_title        => l_effective_title,
            p_fields       => l_field_values,
            p_generated_at => l_generated_at,
            p_label        => 'pie de página'
        );

        IF p_excluded_columns_json IS NOT NULL THEN
            FOR excluded_record IN (
                SELECT column_name
                  FROM JSON_TABLE(
                           p_excluded_columns_json,
                           '$[*]'
                           COLUMNS (
                               column_name VARCHAR2(128)
                                   PATH '$' ERROR ON ERROR
                           )
                       )
            )
            LOOP
                l_name := normalize_name(
                    excluded_record.column_name,
                    'La lista de exclusiones',
                    -20133
                );
                l_excluded_columns(l_name) := TRUE;
            END LOOP;
        END IF;

        l_context := apex_exec.open_query_context(
            p_location        => apex_exec.c_location_local_db,
            p_sql_query       => p_sql_query,
            p_sql_parameters  => l_parameters,
            p_auto_bind_items => FALSE,
            p_first_row       => 1
        );                                -- sin p_max_rows: se exportan todas las filas
        l_context_is_open := TRUE;

        FOR column_record IN (
            SELECT column_name,
                   column_heading,
                   value_alignment,
                   heading_alignment,
                   legacy_alignment,
                   format_mask
              FROM JSON_TABLE(
                       p_columns_json,
                       '$[*]'
                       COLUMNS (
                           column_name VARCHAR2(128)
                               PATH '$.name' ERROR ON ERROR,
                           column_heading VARCHAR2(4000)
                               PATH '$.heading' NULL ON ERROR,
                           value_alignment VARCHAR2(20)
                               PATH '$.value_alignment' NULL ON ERROR,
                           heading_alignment VARCHAR2(20)
                               PATH '$.heading_alignment' NULL ON ERROR,
                           legacy_alignment VARCHAR2(20)
                               PATH '$.alignment' NULL ON ERROR,
                           format_mask VARCHAR2(4000)
                               PATH '$.format_mask' NULL ON ERROR
                       )
                   )
        )
        LOOP
            l_template_count := l_template_count + 1;

            IF l_template_count > c_query_max_columns THEN
                raise_application_error(
                    -20134,
                    'La plantilla supera el máximo de ' ||
                    c_query_max_columns || ' columnas.'
                );
            END IF;

            l_name := normalize_name(
                column_record.column_name,
                'La definición de columnas',
                -20135
            );

            IF l_all_columns.EXISTS(l_name) THEN
                raise_application_error(
                    -20136,
                    'La columna "' || l_name || '" está repetida.'
                );
            END IF;
            l_all_columns(l_name) := TRUE;

            IF LENGTHB(column_record.column_heading) > 255 THEN
                raise_application_error(
                    -20137,
                    'El encabezado de la columna "' || l_name ||
                    '" supera 255 bytes.'
                );
            END IF;

            /*
             * El contrato se valida antes de excluir: una plantilla no puede
             * ocultar accidentalmente un alias inexistente en la consulta.
             */
            l_column_position := apex_exec.get_column_position(
                p_context         => l_context,
                p_column_name     => l_name,
                p_attribute_label =>
                    SUBSTRB(NVL(column_record.column_heading, l_name), 1, 255),
                p_is_required     => TRUE,
                p_data_type       => NULL
            );

            l_column_metadata := apex_exec.get_column(
                p_context    => l_context,
                p_column_idx => l_column_position
            );

            IF l_excluded_columns.EXISTS(l_name) THEN
                CONTINUE;
            END IF;

            l_column_count := l_column_count + 1;
            l_columns(l_column_count).column_name := l_name;
            l_columns(l_column_count).column_heading := SUBSTRB(
                NVL(column_record.column_heading, l_name),
                1,
                255
            );
            l_columns(l_column_count).value_alignment := NVL(
                column_record.value_alignment,
                column_record.legacy_alignment
            );
            l_columns(l_column_count).heading_alignment :=
                column_record.heading_alignment;
            l_columns(l_column_count).format_mask :=
                column_record.format_mask;
            l_columns(l_column_count).data_type :=
                l_column_metadata.data_type;
        END LOOP;

        IF l_template_count = 0 THEN
            raise_application_error(
                -20138,
                'La plantilla no define columnas.'
            );
        END IF;

        IF l_column_count = 0 THEN
            raise_application_error(
                -20139,
                'Todas las columnas del reporte fueron excluidas.'
            );
        END IF;

        /* Default por nombre: todas las columnas activas tienen peso 1. */
        FOR column_index IN 1 .. l_column_count
        LOOP
            l_width_modes(l_columns(column_index).column_name) := 'WEIGHT';
            l_width_values(l_columns(column_index).column_name) := 1;
        END LOOP;

        IF p_column_widths_json IS NOT NULL THEN
            FOR width_record IN (
                SELECT column_name,
                       width_mode,
                       width_value
                  FROM JSON_TABLE(
                           p_column_widths_json,
                           '$[*]'
                           COLUMNS (
                               column_name VARCHAR2(128)
                                   PATH '$.column' ERROR ON ERROR,
                               width_mode VARCHAR2(30)
                                   PATH '$.mode' ERROR ON ERROR,
                               width_value NUMBER
                                   PATH '$.value' NULL ON ERROR
                           )
                       )
            )
            LOOP
                l_name := normalize_name(
                    width_record.column_name,
                    'La configuración de anchos',
                    -20155
                );

                IF NOT l_all_columns.EXISTS(l_name) THEN
                    raise_application_error(
                        -20156,
                        'El ancho hace referencia a la columna desconocida "' ||
                        l_name || '".'
                    );
                END IF;

                IF l_width_names.EXISTS(l_name) THEN
                    raise_application_error(
                        -20157,
                        'El ancho de la columna "' || l_name ||
                        '" está repetido.'
                    );
                END IF;
                l_width_names(l_name) := TRUE;

                l_width_mode := UPPER(TRIM(width_record.width_mode));
                IF l_width_mode = 'FIXED' THEN
                    /* Alias preliminar aceptado para migrar proyectos beta. */
                    l_width_mode := 'FIXED_PERCENT';
                END IF;

                IF l_width_mode = 'FIXED_PERCENT' THEN
                    IF width_record.width_value IS NULL OR
                       width_record.width_value < 1 OR
                       width_record.width_value > 95
                    THEN
                        raise_application_error(
                            -20158,
                            'El ancho fijo de "' || l_name ||
                            '" debe estar entre 1 y 95.'
                        );
                    END IF;
                ELSIF l_width_mode = 'WEIGHT' THEN
                    IF width_record.width_value IS NULL OR
                       width_record.width_value <= 0
                    THEN
                        raise_application_error(
                            -20159,
                            'El peso de "' || l_name ||
                            '" debe ser mayor que cero.'
                        );
                    END IF;
                ELSIF l_width_mode = 'AUTO' THEN
                    NULL;
                ELSE
                    raise_application_error(
                        -20161,
                        'El modo de ancho de "' || l_name ||
                        '" debe ser FIXED_PERCENT, WEIGHT o AUTO.'
                    );
                END IF;

                /* Una columna excluida puede conservar su configuración. */
                IF NOT l_excluded_columns.EXISTS(l_name) THEN
                    l_width_modes(l_name) := l_width_mode;
                    IF l_width_mode <> 'AUTO' THEN
                        l_width_values(l_name) :=
                            width_record.width_value;
                    END IF;
                END IF;
            END LOOP;
        END IF;

        /* Recontar AUTO solo entre columnas finales. */
        l_auto_count := 0;
        FOR column_index IN 1 .. l_column_count
        LOOP
            l_name := l_columns(column_index).column_name;
            l_width_mode := l_width_modes(l_name);

            IF l_width_mode = 'FIXED_PERCENT' THEN
                l_fixed_total := l_fixed_total + l_width_values(l_name);
            ELSIF l_width_mode = 'WEIGHT' THEN
                l_weight_total := l_weight_total + l_width_values(l_name);
            ELSIF l_width_mode = 'AUTO' THEN
                l_auto_count := l_auto_count + 1;
            END IF;
        END LOOP;

        IF l_auto_count > 1 THEN
            raise_application_error(-20160, 'Solo se permite una columna AUTO.');
        END IF;

        IF l_fixed_total >= c_width_budget THEN
            raise_application_error(
                -20162,
                'La suma de anchos fijos debe ser menor que 99.'
            );
        END IF;

        l_remaining_width := c_width_budget - l_fixed_total;
        l_weight_denominator := l_weight_total;
        l_weight_budget := l_remaining_width;

        /*
         * AUTO queda sin p_width. La constante interna reserva una fracción del
         * espacio que queda después de los anchos fijos. La reserva es una
         * heurística: APEX conserva la decisión final para la columna AUTO.
         * Aplicarla sobre el espacio restante mantiene siempre un presupuesto
         * positivo y evita divisiones inválidas incluso con anchos fijos altos.
         */
        IF l_auto_count = 1 AND l_weight_total > 0 THEN
            l_weight_budget :=
                l_remaining_width * (100 - c_auto_reserve_pct) / 100;
        END IF;

        CASE UPPER(NVL(TRIM(p_orientation), 'AUTO'))
            WHEN 'PORTRAIT' THEN
                l_orientation := apex_data_export.c_orientation_portrait;
            WHEN 'LANDSCAPE' THEN
                l_orientation := apex_data_export.c_orientation_landscape;
            WHEN 'AUTO' THEN
                IF l_column_count <= 3 THEN
                    l_orientation := apex_data_export.c_orientation_portrait;
                ELSE
                    l_orientation := apex_data_export.c_orientation_landscape;
                END IF;
            ELSE
                raise_application_error(
                    -20163,
                    'La orientación debe ser AUTO, PORTRAIT o LANDSCAPE.'
                );
        END CASE;

        FOR column_index IN 1 .. l_column_count
        LOOP
            l_name := l_columns(column_index).column_name;
            l_width_mode := l_width_modes(l_name);

            IF l_width_mode = 'FIXED_PERCENT' THEN
                l_column_width := l_width_values(l_name);
            ELSIF l_width_mode = 'WEIGHT' AND
                  l_weight_denominator > 0
            THEN
                l_column_width := ROUND(
                    l_weight_budget * l_width_values(l_name) /
                    l_weight_denominator,
                    4
                );
            ELSE
                l_column_width := NULL;
            END IF;

            IF l_column_width IS NULL THEN
                apex_data_export.add_column(
                    p_columns           => l_export_columns,
                    p_name              => l_name,
                    p_heading           =>
                        l_columns(column_index).column_heading,
                    p_format_mask       =>
                        l_columns(column_index).format_mask,
                    p_heading_alignment => normalize_alignment(
                        l_columns(column_index).heading_alignment,
                        l_style.header_alignment,
                        'la columna ' || l_name
                    ),
                    p_value_alignment   => CASE
                        WHEN l_columns(column_index).value_alignment IS NULL
                        THEN get_alignment(
                            NULL,
                            l_columns(column_index).data_type
                        )
                        ELSE normalize_alignment(
                            l_columns(column_index).value_alignment,
                            l_style.body_alignment,
                            'la columna ' || l_name
                        )
                    END
                );
            ELSE
                apex_data_export.add_column(
                    p_columns           => l_export_columns,
                    p_name              => l_name,
                    p_heading           =>
                        l_columns(column_index).column_heading,
                    p_format_mask       =>
                        l_columns(column_index).format_mask,
                    p_heading_alignment => normalize_alignment(
                        l_columns(column_index).heading_alignment,
                        l_style.header_alignment,
                        'la columna ' || l_name
                    ),
                    p_value_alignment   => CASE
                        WHEN l_columns(column_index).value_alignment IS NULL
                        THEN get_alignment(
                            NULL,
                            l_columns(column_index).data_type
                        )
                        ELSE normalize_alignment(
                            l_columns(column_index).value_alignment,
                            l_style.body_alignment,
                            'la columna ' || l_name
                        )
                    END,
                    p_width             => l_column_width
                );
            END IF;
        END LOOP;

        -- El XLSX lleva solo la fila de títulos y los datos, listo para filtrar
        -- y analizar: título, campos, usuario, fecha y pie son solo del PDF.
        l_print_config := apex_data_export.get_print_config(
            p_paper_size              => l_paper_size,
            p_width_units             =>
                apex_data_export.c_width_unit_percentage,
            p_orientation             => l_orientation,
            p_page_header             => CASE
                WHEN l_export_format = apex_data_export.c_format_xlsx THEN NULL
                ELSE l_page_header
            END,
            p_page_header_font_color  => l_style.title_font_color,
            p_page_header_font_family => l_style.title_font_family,
            p_page_header_font_weight => l_style.title_font_weight,
            p_page_header_font_size   => l_style.title_font_size,
            p_page_header_alignment   => l_style.title_alignment,
            p_page_footer             => CASE
                WHEN l_export_format = apex_data_export.c_format_xlsx THEN NULL
                ELSE l_page_footer
            END,
            p_page_footer_font_color  => l_style.footer_font_color,
            p_page_footer_font_family => l_style.footer_font_family,
            p_page_footer_font_weight => l_style.footer_font_weight,
            p_page_footer_font_size   => l_style.footer_font_size,
            p_page_footer_alignment   => l_style.footer_alignment,
            p_header_bg_color         => l_style.header_bg_color,
            p_header_font_color       => l_style.header_font_color,
            p_header_font_family      => l_style.header_font_family,
            p_header_font_weight      => l_style.header_font_weight,
            p_header_font_size        => l_style.header_font_size,
            p_body_bg_color           => l_style.body_bg_color,
            p_body_font_color         => l_style.body_font_color,
            p_body_font_family        => l_style.body_font_family,
            p_body_font_weight        => l_style.body_font_weight,
            p_body_font_size          => l_style.body_font_size,
            p_border_width            => l_style.border_width,
            p_border_color            => l_style.border_color
        );

        l_export_file_name :=
            l_base_file_name || '_' ||
            TO_CHAR(SYSDATE, 'YYYYMMDD_HH24MISS');

        IF l_export_format = apex_data_export.c_format_pdf THEN
            l_export := apex_data_export.export(
                p_context        => l_context,
                p_format         => l_export_format,
                p_columns        => l_export_columns,
                p_file_name      => l_export_file_name,
                p_print_config   => l_print_config,
                p_pdf_accessible => TRUE
            );
        ELSE
            l_export := apex_data_export.export(
                p_context      => l_context,
                p_format       => l_export_format,
                p_columns      => l_export_columns,
                p_file_name    => l_export_file_name,
                p_print_config => l_print_config
            );
        END IF;

        apex_exec.close(l_context);
        l_context_is_open := FALSE;

        apex_data_export.download(
            p_export              => l_export,
            p_content_disposition => apex_data_export.c_attachment,
            p_add_file_extension  => TRUE
        );

    EXCEPTION
        WHEN apex_application.e_stop_apex_engine THEN
            IF l_context_is_open THEN
                apex_exec.close(l_context);
                l_context_is_open := FALSE;
            END IF;
            RAISE;
        WHEN OTHERS THEN
            IF l_context_is_open THEN
                apex_exec.close(l_context);
            END IF;
            RAISE;
    END download_query;


    PROCEDURE download_ig_pdf(
        p_application_id       IN NUMBER,
        p_page_id              IN NUMBER,
        p_region_static_id     IN VARCHAR2,
        p_visible_columns_json IN CLOB,
        p_title                IN VARCHAR2,
        p_file_name            IN VARCHAR2,
        p_generated_by         IN VARCHAR2,
        p_orientation          IN VARCHAR2,
        p_max_rows             IN PLS_INTEGER,
        p_excluded_columns_json IN CLOB,
        p_column_spans_json    IN CLOB,
        p_display_columns_json IN CLOB
    )
    IS
    BEGIN
        download_ig(
            p_application_id       => p_application_id,
            p_page_id              => p_page_id,
            p_region_static_id     => p_region_static_id,
            p_visible_columns_json => p_visible_columns_json,
            p_title                => p_title,
            p_file_name            => p_file_name,
            p_generated_by         => p_generated_by,
            p_format               => 'PDF',
            p_orientation          => p_orientation,
            p_max_rows             => p_max_rows,
            p_excluded_columns_json => p_excluded_columns_json,
            p_column_spans_json    => p_column_spans_json,
            p_display_columns_json => p_display_columns_json
        );
    END download_ig_pdf;

END pkg_corporate_reports;
/
