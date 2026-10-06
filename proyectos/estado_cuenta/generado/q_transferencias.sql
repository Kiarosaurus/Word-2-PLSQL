-- Q_TRANSFERENCIAS: pagos por transferencia bancaria.
select t.nro,
       t.fecha,
       t.nro_operacion,
       t.nro_cuota,
       t.pension,
       t.servicios,
       t.descuento,
       t.mora,
       t.pension + t.servicios + nvl(t.mora, 0)  as total
  from pruebap_alumno_transferencia t
 where t.cod_alumno = :P_COD_ALUMNO
 order by t.nro
