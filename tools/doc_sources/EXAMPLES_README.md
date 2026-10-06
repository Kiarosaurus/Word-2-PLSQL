# Ejemplo ejecutable ENTIDADES

La carpeta `proyectos/entidades/` contiene un proyecto completo que el compilador puede validar y compilar sin mover archivos. Sigue la misma organización que tendrá cualquier reporte nuevo; es el único proyecto que además incluye esta documentación (`README.docx`). Sus tablas tienen el prefijo `pruebap_` y están pensadas para una página de prueba (página 70); la sección «Tablas de prueba» incluye el DDL y los datos para crearlas.

## Archivos

```
proyectos/entidades/
  entidades.docx             <- plantilla Word (A4 horizontal, seis columnas, cinco tipos de marcador)
  apex_process.sql           <- lo único que se pega en APEX (generado al compilar)
  README.docx                <- esta documentación (solo en el proyecto de ejemplo)
  generado/                  <- material de trabajo que no se sube a APEX
    entidades.sql            <- consulta fuente; su texto queda incrustado en apex_process.sql
    entidades.report.json    <- binds, campos, anchos y estilos
    template.json            <- definición compilada (automática)
    validation.json          <- diagnósticos de la compilación (automático)
```

`entidades.sql` no se sube a APEX como archivo: al compilar, su consulta se copia dentro de `apex_process.sql` (`l_sql clob := to_clob(q'~select …~')`). Por eso vive en `generado/` junto con el resto del material de trabajo.

El DOCX se utiliza solamente durante la compilación local. En APEX se instala el paquete una vez y se pega el bloque `apex_process.sql`. Las tablas del SQL (`pruebap_entidad`, `pruebap_departamento`, `pruebap_distrito`) son tablas de prueba: créelas en el parsing schema (sección «Tablas de prueba») o adapte el SQL a objetos reales y recompile.

## Tablas de prueba

Ejecútelo conectado como el *parsing schema*. Los tipos coinciden con los binds: `vdni` es texto, `departamento_id` es número y `fecha_registro` es fecha.

```
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

-- 2500 filas para probar el límite max_rows = 2000
insert into pruebap_entidad (vdni, vnom, vdirec_actual, departamento_id, distrito_id, vnro_tlf1, fecha_registro)
select lpad(to_char(50000000 + level), 8, '0'), 'Persona ' || level, 'Dirección ' || level,
       mod(level, 3) + 1, mod(level, 3) + 1, '9' || lpad(level, 8, '0'),
       date '2024-01-01' + mod(level, 900)
  from dual connect by level <= 2500;
commit;
```

## Contrato de la plantilla

El párrafo de título contiene el título integrado y un campo escalar, separados por un salto de línea manual (Mayús+Intro) dentro del mismo párrafo:

```
{{REPORT_TITLE}}            <- se reemplaza por "title": "Relación de entidades"
Unidad: {{FIELD:UNIDAD}}    <- se reemplaza por el campo UNIDAD de "fields"
```

El pie de página incluye los valores integrados:

```
Usuario: {{APP_USER}} | Fecha: {{GENERATED_AT}}   <- usuario APEX y fecha DD/MM/YYYY HH24:MI
```

La tabla contiene una fila de etiquetas y otra de marcadores. Los seis aliases coinciden con el `SELECT`:

| Encabezado | Marcador Word | Alias SQL |
|---|---|---|
| DNI | {{COLUMN:VDNI}} | VDNI |
| Nombre completo | {{COLUMN:VNOM}} | VNOM |
| Dirección actual | {{COLUMN:VDIREC_ACTUAL}} | VDIREC_ACTUAL |
| Departamento | {{COLUMN:DEPARTAMENTO}} | DEPARTAMENTO |
| Distrito | {{COLUMN:DISTRITO}} | DISTRITO |
| Teléfono | {{COLUMN:VNRO_TLF1}} | VNRO_TLF1 |

El campo `{{FIELD:UNIDAD}}` se declara en `fields` como constante:

```
"fields": [
  {"name": "UNIDAD", "source": "CONSTANT", "value": "Oficina de Informática"}
]
```

Para que el valor provenga de la página en lugar de una constante, cámbielo a `{"name": "UNIDAD", "source": "ITEM", "item": "P70_UNIDAD", "type": "VARCHAR2"}` y cree ese Page Item.

## Binds lógicos y Page Items

| Bind SQL | Page Item | Tipo |
|---|---|---|
| :DNI | P70_DNI | VARCHAR2 |
| :DEPARTAMENTO_ID | P70_DEPARTAMENTO_ID | NUMBER |
| :FECHA_DESDE | P70_FECHA_DESDE | DATE, máscara DD/MM/YYYY |

Cada entrada de `bindings` relaciona un `:BIND` del SQL con un Page Item mediante las claves `bind` e `item`. El proceso generado no concatena estos valores en el SQL: `PKG_CORPORATE_REPORTS` los obtiene del estado de sesión y los añade como parámetros tipados de `APEX_EXEC`.

Antes de la descarga, los tres Page Items deben estar en el estado de sesión. El botón de descarga usa *Submit Page*, que los envía automáticamente; no use una Dynamic Action Ajax para descargar.

El proyecto declara además `P0_REPORT_FORMAT` y `P0_REPORT_ORIENTATION` en la página global (página 0), con valores de retorno `PDF`/`XLSX` y `AUTO`/`PORTRAIT`/`LANDSCAPE`. Si no desea que el usuario elija, elimine `format_item` y `orientation_item` del proyecto y recompile.

## LOV: imprimir display, no return

`P70_DEPARTAMENTO_ID` puede almacenar el identificador numérico del LOV. Ese valor sirve solo como filtro. El reporte imprime `d.descripcion AS departamento`, por lo que el PDF recibe el texto visible y no el ID.

La misma regla debe aplicarse a cualquier otro LOV: la consulta del reporte debe hacer el `JOIN` correspondiente y seleccionar el display con el alias declarado en Word.

## Anchos y exclusiones

- `VDNI` y `VNRO_TLF1` usan `FIXED_PERCENT`.
- `VNOM`, `DEPARTAMENTO` y `DISTRITO` distribuyen el espacio restante mediante `WEIGHT`.
- `VDIREC_ACTUAL` usa `AUTO`; solo puede existir una columna con ese modo.
- `excluded_columns` está presente y vacío. Para una exclusión institucional se agrega allí uno de los seis aliases.

Los estilos avanzados están en `style_overrides`, con las zonas `title`, `table_header`, `table_body`, `border` y `footer`. Sus valores prevalecen sobre los que el compilador extrae automáticamente del DOCX; en este ejemplo coinciden con la plantilla y sirven como muestra para editarlos directamente.

## Validar y compilar

Desde la raíz del proyecto, con el entorno virtual activado:

```
# 1. Validar: no escribe archivos
python -m corporate_report_compiler validate `
  --project .\proyectos\entidades\generado\entidades.report.json

# 2. Compilar: sin --output, apex_process.sql va a proyectos\entidades\
#    y template.json / validation.json a proyectos\entidades\generado\
python -m corporate_report_compiler compile `
  --project .\proyectos\entidades\generado\entidades.report.json
```

Los archivos compilados que acompañan al ejemplo son también la referencia que detecta regresiones: tras compilar, `git status` no debe mostrar cambios en `proyectos\entidades`. La prueba automática `tests\test_workspace.py` compila una copia y la compara byte a byte.

En Linux o macOS:

```
PYTHONPATH=src python -m corporate_report_compiler compile \
  --project proyectos/entidades/generado/entidades.report.json
git status --short proyectos/entidades     # sin salida = idéntico a la referencia
```

Una validación correcta no escribe archivos. La compilación genera automáticamente, reemplazando solo estos tres archivos:

- `generado/template.json`: contrato normalizado y estilos extraídos del DOCX;
- `generado/validation.json`: diagnóstico reproducible de la compilación;
- `apex_process.sql`: llamada mínima a `PKG_CORPORATE_REPORTS.DOWNLOAD_QUERY`, junto al DOCX.

## Cargar de nuevo un DOCX con el mismo nombre

«Nuevo proyecto desde DOCX» (o `apex-report-compiler new --docx ...`) copia el Word a `proyectos/<nombre>/`. Si esa carpeta ya existe, pide confirmación y reemplaza únicamente sus archivos conocidos: el DOCX, `apex_process.sql` y los cuatro de `generado/`. Antes de reemplazar cada uno guarda una copia `generado/<archivo>.bak`; la siguiente carga sobrescribe esa copia. Cualquier otro archivo que haya puesto en la carpeta (documentación, notas, capturas) no se toca.

```
# Primera carga desde cualquier carpeta
apex-report-compiler new --docx C:\Descargas\ventas.docx --page 42

# Nueva versión del Word con el mismo nombre: sin --replace se detiene (GUIDE-002)
apex-report-compiler new --docx C:\Descargas\ventas.docx --page 42 --replace
#   proyectos\ventas\generado\ventas.sql.bak         <- su SQL editado anterior
#   proyectos\ventas\generado\ventas.report.json.bak <- su proyecto anterior
```

Después de reemplazar, recupere desde los `.bak` lo que había editado a mano (tablas reales, filtros, anchos) y vuelva a compilar.

## Adaptación a una página real

```flujo
# Figura 1. De este ejemplo a un reporte real
Informática | Copiar entidades.docx con otro nombre y editarlo en Word | etiquetas, estilos y columnas; mantener los marcadores {{...}}
Compilador | Nuevo proyecto desde DOCX | copia el Word a proyectos\<nombre>\ y crea automáticamente generado\<nombre>.sql y .report.json
Informática | Adaptar SQL y proyecto en generado\ | tablas reales, aliases del DOCX, Page Items, título, anchos y max_rows
Compilador | Validar y compilar | apex_process.sql (junto al DOCX) y generado\template.json se regeneran automáticamente
APEX | Pegar proyectos\<nombre>\apex_process.sql y probar | Manual de uso, sección 9
```

Alternativa: copiar también `generado\entidades.sql` y `generado\entidades.report.json` a la carpeta `generado` del nuevo proyecto, renombrarlos y editarlos directamente (incluidas las claves `template` y `query_file`).

1. Cambiar las tablas y joins del SQL manteniendo los aliases del DOCX.
2. Cambiar los Page Items de `bindings` por los de la página.
3. Si se añade o quita una columna en Word, añadir o quitar su alias en el `SELECT` y, si figura, en `column_widths`. Las etiquetas, estilos y el orden se toman automáticamente del DOCX al recompilar.
4. Ajustar título, nombre de archivo, límite de filas, campos, exclusiones y anchos directamente en el `.report.json`.
5. Validar y compilar localmente.
6. Revisar `template.json`, `validation.json` y `apex_process.sql`.
7. Copiar el proceso generado a APEX (Manual de uso, sección 9) y probarlo allí. Una validación local correcta no equivale a aceptación en APEX.

El SQL, el JSON y el proceso son artefactos administrados por Informática. No deben recibirse desde el navegador ni ser editables por el usuario final.
