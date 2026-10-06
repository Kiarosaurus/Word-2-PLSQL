-- Tablas de prueba del ejemplo ENTIDADES (convención PRUEBAP_, página 70).
-- Copia del bloque «Tablas de prueba» de README.docx. Ejecútelo una vez en el
-- parsing schema: CREATE TABLE falla si las tablas ya existen (no borra nada).

create table pruebap_departamento (
    id          number primary key,
    descripcion varchar2(100) not null
);

create table pruebap_distrito (
    id          number primary key,
    descripcion varchar2(100) not null
);

create table pruebap_entidad (
    vdni            varchar2(8) primary key,
    vnom            varchar2(200) not null,
    vdirec_actual   varchar2(300),
    departamento_id number references pruebap_departamento(id),
    distrito_id     number references pruebap_distrito(id),
    vnro_tlf1       varchar2(20),
    fecha_registro  date default sysdate
);

insert into pruebap_departamento values (1, 'Lima');
insert into pruebap_departamento values (2, 'Arequipa');
insert into pruebap_departamento values (3, 'Cusco');
insert into pruebap_distrito values (1, 'Miraflores');
insert into pruebap_distrito values (2, 'Cayma');
insert into pruebap_distrito values (3, 'Wanchaq');

insert into pruebap_entidad values ('12345678', 'Ana Torres Ríos', 'Av. Larco 123', 1, 1, '987654321', date '2025-01-15');
insert into pruebap_entidad values ('23456789', 'Luis Pérez Soto', 'Calle Mercaderes 45', 2, 2, '976543210', date '2025-06-01');
insert into pruebap_entidad values ('34567890', 'María Quispe Huamán', 'Av. de la Cultura 800', 3, 3, null, date '2026-02-10');
-- dirección larga para la columna AUTO
insert into pruebap_entidad values ('45678901', 'Carlos Ñique', rpad('Jr. Muy Largo ', 250, 'x'), 1, null, '014445555', date '2026-09-01');

-- 2500 filas para probar un volumen grande (se exportan todas)
insert into pruebap_entidad (vdni, vnom, vdirec_actual, departamento_id, distrito_id, vnro_tlf1, fecha_registro)
select lpad(to_char(50000000 + level), 8, '0'), 'Persona ' || level, 'Dirección ' || level,
       mod(level, 3) + 1, mod(level, 3) + 1, '9' || lpad(level, 8, '0'),
       date '2024-01-01' + mod(level, 900)
  from dual connect by level <= 2500;
commit;
