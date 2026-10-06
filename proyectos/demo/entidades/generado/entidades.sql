select e.vdni          as vdni,
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
 order by e.vnom
