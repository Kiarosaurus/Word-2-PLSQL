-- Tablas de prueba del ejemplo ESTADO_CUENTA (convención PRUEBAP_, página 70).
-- Datos ficticios: academia, alumno, montos y fechas son inventados.
-- Idempotente: crea cada tabla solo si no existe e inserta solo si está vacía.
-- No borra ni modifica ninguna tabla existente.
DECLARE
    PROCEDURE create_if_missing(p_name IN VARCHAR2, p_ddl IN VARCHAR2) IS
        l_count PLS_INTEGER;
    BEGIN
        SELECT COUNT(*) INTO l_count FROM user_tables WHERE table_name = p_name;
        IF l_count = 0 THEN
            EXECUTE IMMEDIATE p_ddl;
        END IF;
    END create_if_missing;
BEGIN
    create_if_missing('PRUEBAP_ALUMNO', q'~
        CREATE TABLE pruebap_alumno (
            cod_alumno        VARCHAR2(10) PRIMARY KEY,
            carrera           VARCHAR2(10),
            nombre            VARCHAR2(100),
            fec_matricula     DATE,
            modalidad         VARCHAR2(30),
            sede              VARCHAR2(30),
            costo_programa    NUMBER(12,2),
            saldo_anterior    NUMBER(12,2),
            seguro            NUMBER(12,2),
            beca              NUMBER(12,2),
            cuota_pension     NUMBER(12,2),
            cuota_servicios   NUMBER(12,2),
            nro_cuotas        NUMBER(4),
            servicios_total   NUMBER(12,2),
            recargo_mensual   NUMBER(6,2)
        )~');
    create_if_missing('PRUEBAP_ALUMNO_CUOTA', q'~
        CREATE TABLE pruebap_alumno_cuota (
            cod_alumno   VARCHAR2(10) REFERENCES pruebap_alumno,
            nro_cuota    NUMBER(4),
            anio         NUMBER(4),
            mes          NUMBER(2),
            caja         VARCHAR2(5),
            pension      NUMBER(12,2),
            servicios    NUMBER(12,2),
            descuento    NUMBER(12,2),
            mora         NUMBER(12,2),
            boleta       VARCHAR2(12),
            estado       VARCHAR2(1),   -- C: pagada en caja, P: por pagar
            PRIMARY KEY (cod_alumno, nro_cuota)
        )~');
    create_if_missing('PRUEBAP_ALUMNO_TRANSFERENCIA', q'~
        CREATE TABLE pruebap_alumno_transferencia (
            cod_alumno    VARCHAR2(10) REFERENCES pruebap_alumno,
            nro           NUMBER(4),
            fecha         DATE,
            nro_operacion VARCHAR2(20),
            nro_cuota     NUMBER(4),
            pension       NUMBER(12,2),
            servicios     NUMBER(12,2),
            descuento     NUMBER(12,2),
            mora          NUMBER(12,2),
            PRIMARY KEY (cod_alumno, nro)
        )~');
    create_if_missing('PRUEBAP_ALUMNO_ATRASO', q'~
        CREATE TABLE pruebap_alumno_atraso (
            cod_alumno   VARCHAR2(10) REFERENCES pruebap_alumno,
            nro          NUMBER(4),
            anio         NUMBER(4),
            mes          NUMBER(2),
            importe      NUMBER(12,2),
            PRIMARY KEY (cod_alumno, nro)
        )~');
END;
/

INSERT INTO pruebap_alumno
SELECT 'A20240157', 'ING-SIS', 'CARLA ANDREA MARTINEZ ROJAS', DATE '2024-03-04',
       'Presencial', 'Campus Norte', 9000, NULL, 120, 450, 187.50, 37.50, 48, 1800, 1.25
  FROM dual
 WHERE NOT EXISTS (SELECT 1 FROM pruebap_alumno WHERE cod_alumno = 'A20240157');

-- 48 cuotas desde abril de 2024: las 4 primeras pagadas en caja; las 8 últimas con descuento.
INSERT INTO pruebap_alumno_cuota
SELECT 'A20240157',
       LEVEL,
       EXTRACT(YEAR FROM ADD_MONTHS(DATE '2024-04-01', LEVEL - 1)),
       EXTRACT(MONTH FROM ADD_MONTHS(DATE '2024-04-01', LEVEL - 1)),
       CASE WHEN LEVEL <= 4 THEN 'C01' END,
       CASE WHEN LEVEL <= 40 THEN 187.50 ELSE 175.00 END,
       CASE WHEN LEVEL <= 40 THEN 37.50 ELSE 35.00 END,
       NULL,
       CASE WHEN LEVEL <= 4 THEN 0 END,
       CASE WHEN LEVEL <= 4 THEN 'B001-' || LPAD(1200 + LEVEL, 6, '0') END,
       CASE WHEN LEVEL <= 4 THEN 'C' ELSE 'P' END
  FROM dual
 WHERE NOT EXISTS (SELECT 1 FROM pruebap_alumno_cuota WHERE cod_alumno = 'A20240157')
CONNECT BY LEVEL <= 48;

INSERT INTO pruebap_alumno_transferencia
SELECT 'A20240157', 1, DATE '2025-02-14', 'OP-7781520', NULL, 1125.00, 225.00, 0, 0
  FROM dual
 WHERE NOT EXISTS (SELECT 1 FROM pruebap_alumno_transferencia WHERE cod_alumno = 'A20240157');

-- 15 meses con pago atrasado: 2024-08..2024-12 y 2025-06..2026-03
INSERT INTO pruebap_alumno_atraso
SELECT 'A20240157', ROWNUM, EXTRACT(YEAR FROM periodo), EXTRACT(MONTH FROM periodo), 225.00
  FROM (SELECT ADD_MONTHS(DATE '2024-08-01', LEVEL - 1) periodo FROM dual CONNECT BY LEVEL <= 5
        UNION ALL
        SELECT ADD_MONTHS(DATE '2025-06-01', LEVEL - 1) FROM dual CONNECT BY LEVEL <= 10
        ORDER BY 1)
 WHERE NOT EXISTS (SELECT 1 FROM pruebap_alumno_atraso WHERE cod_alumno = 'A20240157');

COMMIT;
