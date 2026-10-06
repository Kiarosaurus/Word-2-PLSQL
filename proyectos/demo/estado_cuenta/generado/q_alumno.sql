-- Q_ALUMNO: grupo maestro (una fila). Las columnas de fórmula (importe neto,
-- cuota mensual, deuda total) van como expresiones del SELECT.
select a.cod_alumno,
       a.carrera,
       a.nombre,
       a.fec_matricula,
       a.modalidad,
       a.sede,
       a.costo_programa,
       a.saldo_anterior,
       a.seguro,
       a.beca,
       nvl(a.costo_programa, 0) + nvl(a.saldo_anterior, 0)
         + nvl(a.seguro, 0) - nvl(a.beca, 0)            as importe_neto,
       a.cuota_pension,
       a.cuota_servicios,
       a.cuota_pension + a.cuota_servicios              as cuota_mensual,
       a.nro_cuotas,
       a.servicios_total,
       a.costo_programa + a.servicios_total             as deuda_total,
       a.recargo_mensual
  from pruebap_alumno a
 where a.cod_alumno = :P_COD_ALUMNO
