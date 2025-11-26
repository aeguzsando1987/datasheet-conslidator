# Consolidador de datos xlxs, xls, csv
## Descripcion general
Este programa recorre mutiples subdirectorios, detectando todos los archivos Excel cuyo nombre inicie con el patron “REPORTE SEMANAL”.  Extrae la información de sus tablas internas y consolida todos los datos en un archivo maestro de unificación con estructura estándar o una base de datos dedicada.
## ESTRUCTURA DEL PROYECTO
1.	**Directorio raíz (Nombrar: planeacion_vendedores/)**
### Contiene:
a. El archivo maestro de unificación (un xlsx).
b. Varios subdirectorios, uno por región.
c. Ejemplos de subdirectorios:
   Region_Leon/
   Region_Apodaca/
   Region_Guadalajara/
   .../
   **NOTA**: Se pueden generar cuantos subdirectorios y archivos como se quiera y cualquier otro que se agregue dinámicamente (ejemplo: Region_Sureste/).
2.	**Archivos dentro de los subdirectorios**
Cada subdirectorio puede contener uno o varios archivos cuyo nombre siempre inicia con "REPORTE SEMANAL".
**Ejemplos válidos:**
    REPORTE SEMANAL.xlsx
    REPORTE SEMANAL semana 45.xlsm
    REPORTE SEMANAL - Zona Norte.xls
**NOTA**: Todos deben teber la misma estructura estándar. Se agrega al repositorio un archivo de prueba.
## ESTRUCTURA DE CADA ARCHIVO “REPORTE SEMANAL…”
Cada archivo contiene tres hojas, todas empezando sus encabezados en la fila 5.
**Hoja 1 — BASE DE DATOS DE PROSPECTOS LOCALIZADOS**
**Tabla: BaseDatos**
**Columnas:**
    NUM
    RESPONSABLE
    REGION
    SEMANA
    FECHA
    CLASIFICACION
    NOMBRE DE LA EMPRESA
    GIRO
    SECTOR
    ESTADO
    CIUDAD
    DOMICILIO
    CONTACTO
    PUESTO
    E-MAIL
    TELEFONO
    WHATSAPP
    FUENTE DE INFORMACION
    B2B
    FECHA DE VISITA
    OPORTUNIDAD
**Hoja 2 — Planeación**
**Tabla: Planeacion**
**Columnas:**
    NUM
    RESPONSABLE
    REGION
    SEMANA
    FECHA
    NOMBRE DE LA EMPRESA
    FUENTE DE INFORMACION
    COMENTARIOS
**Hoja 3 — Reporte semanal**
**Tabla: Reporte**
**Columnas:**
    NUM
    RESPONSABLE
    REGION
    SEMANA
    FECHA
    NOMBRE DE LA EMPRESA
    FUENTE DE INFORMACION
    ACTIVIDAD PROGRAMADA
    COMENTARIOS
    NECESIDAD DETECTADA

## ARCHIVO MAESTRO DE UNIFICACIÓN
**Caracteristicas**:
 - El archivo maestro tiene la misma estructura que los archivos de en subdirectorios.
 - Para cada tabla (BaseDatos, Planeacion y Reporte), los datos de todas las regiones se consolidan en este archivo maestro.
 - La columna NUM no se copia, sino que se renumera globalmente desde 1 o ultimo registro numerado (por ejemplo 1001) hasta el total de filas unificadas (1002 hasta N).

## PROCESO DEL PROGRAMA
1. El programa se ejecuta a partri de una CLI. El proceso puede automatizarse para refrescar datos cada N horas.
2. Identifica todos los subdirectorios dentro del directorio raíz.
3. Recorre dinámicamente los subdirectorios, sin depender de sus nombres.
3. Detecta cualquier archivo Excel cuyo nombre inicie con “REPORTE SEMANAL”.
4. Lee todas las hojas correspondientes a las tablas estándar.
5. Extrae los datos de las tablas. (Nota: en el archivo, en cada hoja, las tablas incian con sus encabezados en la fila 5)
6. Consolida la información en el archivo maestro.
7. Guarda la base unificada.
8. Borra o refresca los datos de las tablas en los archivos origen a partir de la segunda fila con datos, pues la primera fila con datos contiene formulas (importante: no borra los archivos ni los renombra).
