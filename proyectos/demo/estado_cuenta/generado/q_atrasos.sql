-- Q_ATRASOS: meses con pago atrasado.
select a.anio,
       lpad(a.mes, 2, '0')  as mes,
       a.nro,
       a.importe
  from pruebap_alumno_atraso a
 where a.cod_alumno = :P_COD_ALUMNO
 order by a.nro
