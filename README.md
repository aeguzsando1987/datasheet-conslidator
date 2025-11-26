# CONSOLIDADOR DE REPORTES v1.0

## CONTENIDO

1.  Introducción\
2.  Requisitos del Sistema\
3.  Instalación\
4.  Configuración\
5.  Uso del Programa\
6.  Características Principales\
7.  Solución de Problemas\
8.  Preguntas Frecuentes (FAQ)

------------------------------------------------------------------------

## 1. INTRODUCCIÓN

El **Consolidador de Reportes v1.0** es un sistema automático
que:

-   Busca archivos Excel con nombre **"REPORTE SEMANAL\*"** en
    subdirectorios\
-   Extrae datos de 3 hojas: **BASE DE DATOS**, **PLANEACION**,
    **REPORTE**\
-   Consolida todos los datos en un archivo maestro único\
-   Renumera incrementalmente (continúa desde último **NUM** del
    maestro)\
-   Preserva fórmulas en archivos origen al limpiar datos\
-   Crea backups automáticos antes de modificar archivos

------------------------------------------------------------------------

## 2. REQUISITOS DEL SISTEMA

-   **Sistema Operativo:** Windows 10/11 (64-bit)\
-   **Memoria RAM:** Mínimo 4 GB\
-   **Espacio en Disco:** 500 MB libres\
-   **NO requiere Excel instalado**\
-   **NO requiere .NET instalado** (el ejecutable es autónomo)

------------------------------------------------------------------------

## 3. INSTALACIÓN

### PASO 0: Compilar y generar ejecutable

-   Abrir la consola de comandos
-   Navegar a la carpeta del proyecto
-   Compilar el proyecto escribiendo `dotnet build`
-   Crear el ejecutable escribiendo: `dotnet publish -c Release -r win-x64 --self-contained`
-   El ejecutable se genera en la carpeta `bin\Release\net8.0-windows\win-x64\publish`

### PASO 1: Crear Carpeta de Trabajo

-   Copiar el ejecutable `ConsolidadorReportes.exe` a la carpeta de trabajo\
-   Crea una carpeta donde procesarás los reportes\
    Ejemplo: `C:\Reportes_Semanales\`

### PASO 2: Copiar Archivos del Programa

-   Copia `ConsolidadorReportes.exe` a la carpeta de trabajo\
-   Copia `appsettings.json` a la misma carpeta\
-   Las subcarpetas `logs/` y `backups/` se crean automáticamente

### PASO 3: Verificar Estructura

    C:\Reportes_Semanales├── ConsolidadorReportes.exe
    ├── appsettings.json
    ├── Region_Centro (carpeta de ejmplo)│   └── REPORTE SEMANAL centro.xlsx
    ├── Region_Norte (carpeta de ejmplo)│   └── REPORTE SEMANAL norte.xlsx
    └── ... (otros subdirectorios con reportes)

------------------------------------------------------------------------

## 4. CONFIGURACIÓN

Edita el archivo `appsettings.json`:

### 4.1. RUTAS (Paths)

-   `DirectorioRaiz`: `"."`\
-   `ArchivoMaestro`: `"consolidado_maestro.xlsx"`\
-   `DirectorioLogs`: `"logs"`\
-   `DirectorioBackups`: `"backups"`

### 4.2. OPCIONES (Options)

-   `CrearBackup`: `true`\
-   `LimpiarArchivosOrigen`: `false`\
-   `ValidarDuplicados`: `false`\
-   `IncluirSubdirectorios`: `true`

### 4.3. TABLAS (Tablas)

-   `NombreHoja1`: `"BASE DE DATOS"`\
-   `NombreHoja2`: `"PLANEACION"`\
-   `NombreHoja3`: `"REPORTE"`\
-   `FilaEncabezados`: `5`\
-   `FilaPrimeraConFormulas`: `6`\
-   `FilaInicioDatosABorrar`: `7`

------------------------------------------------------------------------

## 5. USO DEL PROGRAMA

### EJECUCIÓN BÁSICA

1.  Abre **cmd** o PowerShell\
2.  Navega a la carpeta:\
    `cd C:\Reportes_Semanales`\
3.  Ejecuta:\
    `ConsolidadorReportes.exe`\
4.  Revisa `consolidado_maestro.xlsx`

### EJECUCIÓN POR DOBLE CLIC

-   Doble clic en `ConsolidadorReportes.exe`\
-   La ventana muestra progreso\
-   Revisa los logs en `logs\`

### FLUJO DEL PROCESO

1.  Escaneo del directorio raíz\
2.  Búsqueda de archivos "REPORTE SEMANAL"\
3.  Lectura del último NUM\
4.  Backup del maestro\
5.  Consolidación\
6.  Renumeración\
7.  Escritura en modo **APPEND**\
8.  Limpieza opcional de archivos origen\
9.  Resumen y estadísticas

------------------------------------------------------------------------

## 6. CARACTERÍSTICAS PRINCIPALES

### 6.1. RENUMERACIÓN INCREMENTAL

-   No reinicia en 1\
-   Continúa desde el último NUM\
-   Asegura trazabilidad

### 6.2. MODO APPEND

-   Agrega nuevos datos al final\
-   No sobrescribe historial

### 6.3. PRESERVACIÓN DE FÓRMULAS

-   Fila 5 intacta\
-   Fila 6 con fórmulas preservada\
-   Fila 7+ borrada solo si `LimpiarArchivosOrigen=true`

### 6.4. BACKUPS AUTOMÁTICOS

-   Backups del maestro y de origen\
-   Se retienen por 30 días

### 6.5. LOGGING DETALLADO

-   Archivos en `/logs`\
-   Incluye errores, números asignados, tiempos

------------------------------------------------------------------------

## 7. SOLUCIÓN DE PROBLEMAS

### No se encontraron archivos

-   Revisa `DirectorioRaiz`\
-   Verifica nombres "REPORTE SEMANAL..."

### Archivo en uso

-   Cierra Excel\
-   Intenta de nuevo

### Error al leer hoja

-   Verifica nombres exactos\
-   Revisa fila de encabezados

### NUM duplicados

-   No debería ocurrir\
-   Revisa el log

------------------------------------------------------------------------

## 8. PREGUNTAS FRECUENTES (FAQ)

**¿Puedo ejecutar varias veces al día?**\
Sí. Sin embargo, lo ideal es ejecutarlo con una tarea automatizada

**¿Reprocesa archivos ya consolidados?**\
Sí, esto generará duplicados. Sin mebargo en las configuraciones del programa (appsetting.json) se puede activar la opcion "LimpiarArchivosOrigen": true. Esto permitira que el archivo de orgne sea borrado requiriendo que se suba uno nuevo. Queda a criterio de los usuarios verificar posibles duplicados

**¿Necesito Excel instalado?**\
No. El programa genera un archivo con extencion .xlxs que puede ser abiereto con el programa de hojas e calculo de preferencia

**¿Cómo sé si fue exitoso?**\
- Revisar consola\
- Revisar Log\
- Verificar queel archivo maestro está actualizado

------------------------------------------------------------------------

## SOPORTE

Para reportar problemas:\
- Ver logs\
- Incluir mensaje de error\
- Indicar qué operación realizabas\
- Versión v1.0

------------------------------------------------------------------------

## LICENCIA Y CRÉDITOS

Desarrolado por: E. Guzmán
Consolidador de Reportes Semanales v1.0\
Desarrollado con: **C# .NET 8.0, ClosedXML, Serilog, Spectre.Console**\
Fecha: **Noviembre 2025**

------------------------------------------------------------------------

## FIN DEL MANUAL
