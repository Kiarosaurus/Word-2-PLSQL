-- GENERADO por apex-word-report-compiler 1.0.0.
-- Revise el SQL y los nombres de Page Items antes de copiar este bloque.
-- Los valores de los items se enlazan en PKG_CORPORATE_REPORTS; no se concatenan.
declare
    l_sql                  varchar2(32767) :=
        q'~select e.vdni          as vdni,
       e.vnom          as vnom,
       e.vdirec_actual as vdirec_actual,
       d.descripcion   as departamento,
       di.descripcion  as distrito,
       e.vnro_tlf1     as vnro_tlf1
  from pruebap_entidad e
  left join pruebap_departamento d
    on d.id = e.departamento_id
  left join pruebap_distrito di
    on di.id = e.distrito_id
 where (:DNI is null or e.vdni = :DNI)
   and (:DEPARTAMENTO_ID is null or e.departamento_id = :DEPARTAMENTO_ID)
   and (:FECHA_DESDE is null or e.fecha_registro >= :FECHA_DESDE)
 order by e.vnom~';
    l_bindings_json        clob := to_clob(q'~[{"bind":"DNI","item":"P70_DNI","type":"VARCHAR2","required":false},{"bind":"DEPARTAMENTO_ID","item":"P70_DEPARTAMENTO_ID","type":"NUMBER","required":false},{"bind":"FECHA_DESDE","item":"P70_FECHA_DESDE","type":"DATE","required":false,"format_mask":"DD/MM/YYYY"}]~');
    l_fields_json          clob := to_clob(q'~[{"name":"UNIDAD","source":"CONSTANT","value":"Oficina de Informática"}]~');
    l_columns_json         clob := to_clob(q'~[{"name":"VDNI","heading":"DNI","alignment":"CENTER","format_mask":null},{"name":"VNOM","heading":"Nombre completo","alignment":"START","format_mask":null},{"name":"VDIREC_ACTUAL","heading":"Dirección actual","alignment":"START","format_mask":null},{"name":"DEPARTAMENTO","heading":"Departamento","alignment":"START","format_mask":null},{"name":"DISTRITO","heading":"Distrito","alignment":"START","format_mask":null},{"name":"VNRO_TLF1","heading":"Teléfono","alignment":"CENTER","format_mask":null}]~');
    l_style_json           clob := to_clob(q'~{"title":{"font_family":"HELVETICA","font_size":15.0,"font_weight":"BOLD","font_color":"#2F343A","alignment":"CENTER"},"table_header":{"font_family":"HELVETICA","font_size":9.0,"font_weight":"BOLD","font_color":"#FFFFFF","alignment":"CENTER","background_color":"#4A4F55"},"table_body":{"font_family":"HELVETICA","font_size":8.0,"font_weight":"NORMAL","font_color":"#25282B","background_color":"#FFFFFF"},"border":{"width":0.5,"color":"#BFC3C7"},"footer":{"font_family":"HELVETICA","font_size":8.0,"font_weight":"NORMAL","font_color":"#666666","alignment":"CENTER"}}~');
    l_excluded_columns     clob := NULL;
    l_column_widths        clob := to_clob(q'~[{"column":"VDNI","mode":"FIXED_PERCENT","value":12.0},{"column":"VNOM","mode":"WEIGHT","value":2.0},{"column":"VDIREC_ACTUAL","mode":"AUTO"},{"column":"DEPARTAMENTO","mode":"WEIGHT","value":1.0},{"column":"DISTRITO","mode":"WEIGHT","value":1.0},{"column":"VNRO_TLF1","mode":"FIXED_PERCENT","value":12.0}]~');
begin
    pkg_corporate_reports.download_query(
        p_sql_query             => l_sql,
        p_bindings_json         => l_bindings_json,
        p_fields_json           => l_fields_json,
        p_columns_json          => l_columns_json,
        p_style_json            => l_style_json,
        p_header_template       => q'~{{REPORT_TITLE}}
Unidad: {{FIELD:UNIDAD}}~',
        p_title                 => q'~Relación de entidades~',
        p_footer_template       => q'~Usuario: {{APP_USER}} | Fecha: {{GENERATED_AT}}~',
        p_file_name             => q'~reporte_entidades~',
        p_format                => coalesce(:P0_REPORT_FORMAT, 'PDF'),
        p_orientation           => coalesce(:P0_REPORT_ORIENTATION, q'~LANDSCAPE~'),
        p_excluded_columns_json => l_excluded_columns,
        p_column_widths_json    => l_column_widths
    );
end;
