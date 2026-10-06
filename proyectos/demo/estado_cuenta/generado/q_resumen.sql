-- Q_RESUMEN: columnas de resumen calculadas en SQL.
select r.*,
       r.costo_total - r.pagos_caja - r.becas
         - r.pagos_transferencia - nvl(r.otros_abonos, 0) + r.devoluciones   as saldo
  from (select a.costo_programa + a.servicios_total                    as costo_total,
               (select nvl(sum(c.pension + c.servicios), 0)
                  from pruebap_alumno_cuota c
                 where c.cod_alumno = a.cod_alumno
                   and c.estado = 'C')                                 as pagos_caja,
               nvl(a.beca, 0)                                          as becas,
               (select nvl(sum(t.pension + t.servicios + nvl(t.mora, 0)), 0)
                  from pruebap_alumno_transferencia t
                 where t.cod_alumno = a.cod_alumno)                    as pagos_transferencia,
               cast(null as number)                                    as otros_abonos,
               0                                                       as devoluciones
          from pruebap_alumno a
         where a.cod_alumno = :P_COD_ALUMNO) r
