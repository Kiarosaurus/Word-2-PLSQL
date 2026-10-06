# Proyectos de demostración

La carpeta `proyectos/demo/` contiene dos proyectos completos que el compilador puede validar y compilar sin mover archivos, uno por cada modo. Siguen la misma organización que tendrá cualquier reporte nuevo y usan tablas de prueba con el prefijo `pruebap_`. Son los únicos proyectos que incluyen esta documentación.

| | entidades | estado_cuenta |
|---|---|---|
| Modo | Simple (APEX_DATA_EXPORT) | Layout (motor PDF en PL/SQL) |
| Página APEX de prueba | 70 | 71 |
| Qué demuestra | Una tabla con filtros tipados, LOV, anchos, PDF y Excel, formato y orientación elegidos por el usuario | Cabecera en cuadrícula con bordes visibles, blancos y sin borde, cuatro tablas de datos con totales, recuadro de resumen y «X DE Y» |
| Consultas | Una (entidades.sql) | Seis (q_alumno.sql, q_caja.sql…) |
| Salida | PDF y XLSX | PDF |

```
proyectos/demo/
  README.docx                    <- esta documentación
  entidades/                     <- demo del modo simple (página 70)
    entidades.docx
    apex_process.sql             <- se pega en el proceso de la página 70
    datos_prueba.sql             <- tablas PRUEBAP_ de esta demo
    generado/
      entidades.sql, entidades.report.json, template.json, validation.json
  estado_cuenta/                 <- demo del modo layout (página 71)
    estado_cuenta.docx
    apex_process.sql             <- se pega en el proceso de la página 71
    rpt_estado_cuenta.sql        <- package de este reporte
    datos_prueba.sql             <- tablas PRUEBAP_ALUMNO* (datos ficticios)
    ejemplo_estado_cuenta.pdf    <- resultado esperado
    generado/
      q_*.sql (una por consulta), estado_cuenta.report.json, layout.json, validation.json
```

Los motores comunes no están aquí: viven en `sql/modo_simple/` y `sql/modo_layout/` y se instalan una sola vez por esquema (ver `sql/README.docx`).

## Preparar la base de datos

Conectado como el *parsing schema* (SQL Workshop > SQL Scripts, Upload y Run):

| Orden | Archivo | Para |
|---|---|---|
| 1 | sql/modo_simple/pkg_corporate_reports.sql | Demo entidades (y todo reporte simple) |
| 2 | proyectos/demo/entidades/datos_prueba.sql | Demo entidades |
| 3 | sql/modo_layout/rpt_pdf.pks, rpt_pdf.pkb, rpt_layout.pks, rpt_layout.pkb | Demo estado_cuenta (y todo reporte layout) |
| 4 | proyectos/demo/estado_cuenta/datos_prueba.sql | Demo estado_cuenta |
| 5 | proyectos/demo/estado_cuenta/rpt_estado_cuenta.sql | Demo estado_cuenta |

`estado_cuenta/datos_prueba.sql` es idempotente: crea las tablas solo si no existen y nunca borra datos. `entidades/datos_prueba.sql` usa `CREATE TABLE` directo: ejecútelo una sola vez.

## Demo 1: entidades (modo simple, página 70)

### Archivos

`entidades.sql` no se sube a APEX como archivo: al compilar, su consulta se copia dentro de `apex_process.sql` (`l_sql clob := to_clob(q'~select …~')`). Por eso vive en `generado/` junto con el resto del material de trabajo.

El DOCX se utiliza solamente durante la compilación local. En APEX se instala el paquete una vez y se pega el bloque `apex_process.sql`. Las tablas del SQL (`pruebap_entidad`, `pruebap_departamento`, `pruebap_distrito`) son tablas de prueba: créelas con `datos_prueba.sql` o adapte el SQL a objetos reales y recompile.

### Tablas de prueba

El script es `entidades/datos_prueba.sql`. Los tipos coinciden con los binds: `vdni` es texto, `departamento_id` es número y `fecha_registro` es fecha.

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

-- 2500 filas para probar un volumen grande (se exportan todas)
insert into pruebap_entidad (vdni, vnom, vdirec_actual, departamento_id, distrito_id, vnro_tlf1, fecha_registro)
select lpad(to_char(50000000 + level), 8, '0'), 'Persona ' || level, 'Dirección ' || level,
       mod(level, 3) + 1, mod(level, 3) + 1, '9' || lpad(level, 8, '0'),
       date '2024-01-01' + mod(level, 900)
  from dual connect by level <= 2500;
commit;
```

### Contrato de la plantilla

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

### Binds lógicos y Page Items

| Bind SQL | Page Item | Tipo |
|---|---|---|
| :DNI | P70_DNI | VARCHAR2 |
| :DEPARTAMENTO_ID | P70_DEPARTAMENTO_ID | NUMBER |
| :FECHA_DESDE | P70_FECHA_DESDE | DATE, máscara DD/MM/YYYY |

Cada entrada de `bindings` relaciona un `:BIND` del SQL con un Page Item mediante las claves `bind` e `item`. El proceso generado no concatena estos valores en el SQL: `PKG_CORPORATE_REPORTS` los obtiene del estado de sesión y los añade como parámetros tipados de `APEX_EXEC`.

Antes de la descarga, los tres Page Items deben estar en el estado de sesión. El botón de descarga usa *Submit Page*, que los envía automáticamente; no use una Dynamic Action Ajax para descargar.

El proyecto declara además `P0_REPORT_FORMAT` y `P0_REPORT_ORIENTATION` en la página global (página 0), con valores de retorno `PDF`/`XLSX` y `AUTO`/`PORTRAIT`/`LANDSCAPE`. Si no desea que el usuario elija, elimine `format_item` y `orientation_item` del proyecto y recompile. El XLSX lleva solo la fila de títulos y los datos, sin encabezado ni pie.

### LOV: imprimir display, no return

`P70_DEPARTAMENTO_ID` puede almacenar el identificador numérico del LOV. Ese valor sirve solo como filtro. El reporte imprime `d.descripcion AS departamento`, por lo que el PDF recibe el texto visible y no el ID.

La misma regla debe aplicarse a cualquier otro LOV: la consulta del reporte debe hacer el `JOIN` correspondiente y seleccionar el display con el alias declarado en Word.

### Anchos y exclusiones

- `VDNI` y `VNRO_TLF1` usan `FIXED_PERCENT`.
- `VNOM`, `DEPARTAMENTO` y `DISTRITO` distribuyen el espacio restante mediante `WEIGHT`.
- `VDIREC_ACTUAL` usa `AUTO`; solo puede existir una columna con ese modo.
- `excluded_columns` está presente y vacío. Para una exclusión institucional se agrega allí uno de los seis aliases.

Los estilos avanzados están en `style_overrides`, con las zonas `title`, `table_header`, `table_body`, `border` y `footer`. Sus valores prevalecen sobre los que el compilador extrae automáticamente del DOCX; en este ejemplo coinciden con la plantilla y sirven como muestra para editarlos directamente.

### Página 70 en APEX

1. Page Items `P70_DNI` (Text Field), `P70_DEPARTAMENTO_ID` (Select List con LOV de `pruebap_departamento`, return = id) y `P70_FECHA_DESDE` (Date Picker, máscara DD/MM/YYYY).
2. En la página 0: `P0_REPORT_FORMAT` (PDF/XLSX) y `P0_REPORT_ORIENTATION` (AUTO/PORTRAIT/LANDSCAPE).
3. Botón con *Action: Submit Page*; en la página, *Reload on Submit: Always*.
4. Proceso *Execute Code* en *Processing* con el contenido de `entidades/apex_process.sql` y *When Button Pressed* = el botón.

## Demo 2: estado_cuenta (modo layout, página 71)

Estado de cuenta ficticio de un alumno de una academia de demostración, con la apariencia típica de un reporte de Oracle Reports. Todos los nombres, códigos y montos son inventados.

### Tablas de prueba

`estado_cuenta/datos_prueba.sql` crea `PRUEBAP_ALUMNO`, `PRUEBAP_ALUMNO_CUOTA`, `PRUEBAP_ALUMNO_TRANSFERENCIA` y `PRUEBAP_ALUMNO_ATRASO` con un alumno (código `A20240157`), 48 cuotas, una transferencia y 15 meses atrasados. Es idempotente: se puede ejecutar varias veces.

### Qué demuestra la plantilla

| Zona del Word | Contenido | Qué demuestra |
|---|---|---|
| Encabezado de Word | Tabla invisible de 3 celdas y {{REPORT_TITLE}} | Tablas sin borde; se repite en cada página; constantes {{FIELD:SISTEMA}} y fechas con máscara |
| Cabecera del alumno | Cuadrícula de 8 columnas con banda gris combinada | {{FIELD:ALUMNO.COLUMNA}}; bordes negros, BLANCOS y sin borde; recuadros sobre un solo valor |
| Pagos en caja, Cuotas por pagar, Transferencias, Atrasos | Cuatro tablas de datos | {{COLUMN:...}} por registro, banda y títulos repetidos al saltar de página, totales {{SUM:...}}, columnas vacías como separación |
| Resumen | Tabla centrada con borde exterior | Valores de una consulta de resumen y una raya antes del saldo |
| Pie de Word | Usuario y {{PAGE}} DE {{PAGES}} en casillas | Número de página |

### Consultas y parámetros

Cada consulta es un archivo de `generado/`, como una Query del Data Model de Oracle Reports, y todas usan el mismo parámetro `:P_COD_ALUMNO`, enlazado al Page Item `P71_COD_ALUMNO`.

| Consulta | Archivo | Uso en el Word |
|---|---|---|
| ALUMNO | q_alumno.sql | Cabecera (una fila; incluye columnas de fórmula) |
| CAJA | q_caja.sql | Tabla «Pagos registrados en caja» |
| CUOTAS | q_cuotas.sql | Tabla «Cuotas por pagar» |
| TRANSFERENCIAS | q_transferencias.sql | Tabla «Pagos por transferencia bancaria» |
| ATRASOS | q_atrasos.sql | Tabla «Meses con pago atrasado» |
| RESUMEN | q_resumen.sql | Recuadro de resumen y saldo |

Las constantes `SISTEMA`, `MODULO` y `CODIGO` del encabezado están en `constants` de `estado_cuenta.report.json`.

### Página 71 en APEX

1. Instale los objetos de la sección «Preparar la base de datos» (pasos 3 a 5).
2. *Create Page > Blank Page*, número 71, con una región *Static Content*.
3. Page Item `P71_COD_ALUMNO` (Text Field, valor por defecto `A20240157`).
4. Botón `DESCARGAR` con *Action: Submit Page*; en la página, *Reload on Submit: Always*.
5. Proceso *Execute Code* en *Processing* con el contenido de `estado_cuenta/apex_process.sql` y *When Button Pressed* = DESCARGAR.
6. Descargue: debe obtener un PDF de 2 páginas igual a `ejemplo_estado_cuenta.pdf` (saldo pendiente 8,100.00).

La plantilla se regenera con `python tools\generate_layout_template.py`; deja una copia en `templates\estado_cuenta_layout.docx`.

## Validar y compilar las demos

Desde la raíz del proyecto, con el entorno virtual activado:

```
# 1. Validar: no escribe archivos
python -m corporate_report_compiler validate `
  --project .\proyectos\demo\entidades\generado\entidades.report.json

# 2. Compilar: sin --output, lo que se sube a APEX queda junto al DOCX
#    y el material de trabajo en generado\
python -m corporate_report_compiler compile `
  --project .\proyectos\demo\entidades\generado\entidades.report.json
python -m corporate_report_compiler compile `
  --project .\proyectos\demo\estado_cuenta\generado\estado_cuenta.report.json
```

Los archivos compilados que acompañan a las demos son también la referencia que detecta regresiones: tras compilar, `git status` no debe mostrar cambios en `proyectos\demo`. Las pruebas automáticas `tests\test_workspace.py` y `tests\test_layout.py` compilan una copia de cada demo y la comparan byte a byte.

En Linux o macOS:

```
PYTHONPATH=src python -m corporate_report_compiler compile \
  --project proyectos/demo/entidades/generado/entidades.report.json
git status --short proyectos/demo     # sin salida = idéntico a la referencia
```

| Modo | Lo que se sube a APEX (junto al DOCX) | Material de trabajo (generado/) |
|---|---|---|
| Simple | apex_process.sql | template.json, validation.json |
| Layout | apex_process.sql, rpt_<nombre>.sql | layout.json, validation.json |

## Cargar de nuevo un DOCX con el mismo nombre

«Nuevo proyecto desde DOCX» (o `apex-report-compiler new --docx ...`) copia el Word a `proyectos/<nombre>/`. Si esa carpeta ya existe, pide confirmación y reemplaza únicamente sus archivos conocidos: el DOCX, lo que se sube a APEX y los archivos de `generado/`. Antes de reemplazar cada uno guarda una copia `generado/<archivo>.bak`; la siguiente carga sobrescribe esa copia. Cualquier otro archivo que haya puesto en la carpeta (documentación, notas, capturas, datos de prueba) no se toca.

```
# Primera carga desde cualquier carpeta
apex-report-compiler new --docx C:\Descargas\ventas.docx --page 42

# Nueva versión del Word con el mismo nombre: sin --replace se detiene (GUIDE-002)
apex-report-compiler new --docx C:\Descargas\ventas.docx --page 42 --replace
#   proyectos\ventas\generado\ventas.sql.bak         <- su SQL editado anterior
#   proyectos\ventas\generado\ventas.report.json.bak <- su proyecto anterior
```

Después de reemplazar, recupere desde los `.bak` lo que había editado a mano (tablas reales, filtros, anchos) y vuelva a compilar.

## De una demo a un reporte real

```flujo
# Figura 1. Modo simple: de entidades a un reporte real
Informática | Copiar entidades.docx con otro nombre y editarlo en Word | etiquetas, estilos y columnas; mantener los marcadores {{...}}
Compilador | Nuevo proyecto desde DOCX | copia el Word a proyectos\<nombre>\ y crea automáticamente generado\<nombre>.sql y .report.json
Informática | Adaptar SQL y proyecto en generado\ | tablas reales, aliases del DOCX, Page Items, título y anchos
Compilador | Validar y compilar | apex_process.sql (junto al DOCX) y generado\template.json se regeneran automáticamente
APEX | Pegar proyectos\<nombre>\apex_process.sql y probar | Manual de uso, sección 9
```

```flujo
# Figura 2. Modo layout: de estado_cuenta a un reporte real
Informática | Diseñar el Word copiando estado_cuenta.docx o con el prompt de IA | docs\PROMPT_IA_PLANTILLA.txt genera la primera versión
Compilador | Nuevo proyecto desde DOCX | crea generado\q_<consulta>.sql por cada consulta y el .report.json
Informática | Pegar en cada q_<consulta>.sql la Query del reporte original | declarar los parámetros :P_NOMBRE en 'parameters'
Compilador | Validar y compilar | rpt_<nombre>.sql y apex_process.sql quedan junto al DOCX
APEX | Instalar rpt_<nombre>.sql y pegar apex_process.sql | Manual de uso, sección 13
```

1. Cambiar las tablas y joins del SQL manteniendo los aliases del DOCX.
2. Cambiar los Page Items (`bindings` en modo simple, `parameters` en modo layout) por los de la página.
3. Si se añade o quita una columna en Word, añadir o quitar su alias en el `SELECT` y, si figura, en `column_widths`. Las etiquetas, estilos y el orden se toman automáticamente del DOCX al recompilar.
4. Ajustar título, nombre de archivo, campos, exclusiones y anchos directamente en el `.report.json`.
5. Validar y compilar localmente.
6. Revisar los archivos generados.
7. Publicar en APEX y probarlo allí. Una validación local correcta no equivale a aceptación en APEX.

El SQL, el JSON y el proceso son artefactos administrados por Informática. No deben recibirse desde el navegador ni ser editables por el usuario final.
