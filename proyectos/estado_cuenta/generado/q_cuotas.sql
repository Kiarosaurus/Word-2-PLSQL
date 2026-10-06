-- Q_CUOTAS: cuotas por pagar.
select c.anio,
       lpad(c.mes, 2, '0')        as mes,
       c.caja,
       c.nro_cuota,
       c.pension,
       c.servicios,
       c.pension + c.servicios    as total
  from pruebap_alumno_cuota c
 where c.cod_alumno = :P_COD_ALUMNO
   and c.estado = 'P'
 order by c.nro_cuota
