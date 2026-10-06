-- Proceso APEX generado por apex-word-report-compiler 1.0.0 para ESTADO_CUENTA.
-- Requiere RPT_PDF, RPT_LAYOUT y el package del reporte instalados en el parsing schema.
declare
    l_pdf blob;
begin
    l_pdf := rpt_estado_cuenta.build(
        p_cod_alumno => :P70_COD_ALUMNO,
        p_app_user => :APP_USER
    );
    sys.htp.init;
    sys.owa_util.mime_header('application/pdf', false);
    sys.htp.p('Content-Length: ' || dbms_lob.getlength(l_pdf));
    sys.htp.p('Content-Disposition: inline; filename="estado_cuenta.pdf"');
    sys.owa_util.http_header_close;
    sys.wpg_docload.download_file(l_pdf);
    apex_application.stop_apex_engine;
end;
