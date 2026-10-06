-- Q_CAJA: cuotas pagadas en caja.
select c.anio,
       lpad(c.mes, 2, '0')                        as mes,
       c.caja,
       c.nro_cuota,
       c.pension,
       c.servicios,
       c.descuento,
       c.mora,
       c.pension + c.servicios + nvl(c.mora, 0)   as total,
       c.boleta,
       c.pension + c.servicios                    as importe
  from pruebap_alumno_cuota c
 where c.cod_alumno = :P_COD_ALUMNO
   and c.estado = 'C'
 order by c.nro_cuota
