# Documentacion de `external_files.json`

Este documento explica de forma detallada el archivo de configuracion `external_files.json`, su proposito dentro del bot, su alcance y el significado de cada llave.

## 1. Proposito

`external_files.json` es el punto central de configuracion de rutas externas del proyecto.

Su objetivo es:

- Eliminar rutas hardcodeadas(adsolutas) dentro del codigo Python.
- Permitir cambios operativos sin modificar `course_duplication.py` ni `helpers.py`.
- Validar prerequisitos antes de ejecutar la duplicacion de cursos.
- Controlar archivos de entrada, festivos y ubicacion de logs.

En resumen: si una ruta cambia en el entorno, se actualiza el JSON y no la logica del bot.

## 2. Alcance

Este JSON aplica al flujo completo del proceso de duplicacion y se carga al inicio de la ejecucion en `course_duplication.py`.

Afecta:

- Carga del archivo de cursos a procesar.
- Carga de unidades organizacionales (logica QM / NO QM).
- Carga de Banner (docentes - primera hoja [0]) (NRC, periodo y fechas de curso).
- Carga de festivos de Colombia para calculo de dia habil anterior.
- Carpeta donde se escribe el log de ejecucion.

No cubre:

- Credenciales de acceso a plataformas.
- Parametros de Selenium distintos a rutas.
- Reglas de negocio de columnas internas del Excel (esas se validan en codigo).

## 3. Estructura actual del JSON

```json
{
  "path": "courses.xlsx",
  "unidades_path": "C:\\Users\\...\\UnidadesOrganización.xlsx",
  "BDBANNER": "C:\\Users\\...\\ListadoDocentesEstudiantes.xlsx",
  "HOLIDAYS_JSON": "colombia_holidays.json",
  "DIRLOGS": "C:\\Users\\...\\logs"
}
```

## 4. Reglas generales de uso

- Todas las llaves son obligatorias.
- `path`, `unidades_path`, `BDBANNER` y `HOLIDAYS_JSON` deben existir como archivo.
- `DIRLOGS` debe ser una carpeta valida; si no existe, el sistema la crea.
- Se aceptan rutas absolutas y relativas.
- Las rutas relativas se resuelven desde la carpeta del script (`course_duplication.py`).

## 5. Detalle de cada llave

## `path`

**Tipo:** `string`  
**Obligatoria:** si  
**Que representa:** archivo Excel maestro con los cursos que se van a duplicar.  
**Uso en el sistema:** se lee al inicio para iterar curso por curso.

### Recomendaciones

- Debe contener la estructura de columnas esperada por el bot.
- Si el archivo esta vacio o con formato invalido, la ejecucion se detiene.

## `unidades_path`

**Tipo:** `string`  
**Obligatoria:** si  
**Que representa:** archivo Excel de unidades organizacionales (fuente de referencia para identificar QM).  
**Uso en el sistema:** determina si cada maestro es QM y define comportamiento de configuracion posterior.

### Recomendaciones

- Validar que contenga las columnas requeridas (`Code`, `Name`).
- Mantener actualizada esta fuente para evitar clasificaciones incorrectas.

## `BDBANNER`

**Tipo:** `string`  
**Obligatoria:** si  
**Que representa:** archivo Excel de Banner con datos academicos (NRC, periodo, fecha inicio, fecha fin).  
**Uso en el sistema:** se usa para obtener fechas del curso y aplicarlas en el LMS.

### Recomendaciones

- Verificar la consistencia de `LISTA_CRUZADA`, `PERIODO`, `FECHA_INICIO_CURSO`, `FECHA_FIN_CURSO`.
- Evitar celdas vacias en fechas para no bloquear la iteracion.

## `HOLIDAYS_JSON`

**Tipo:** `string`  
**Obligatoria:** si  
**Que representa:** archivo JSON con festivos de Colombia organizado por `anio -> mes -> fechas`.  
**Uso en el sistema:** permite calcular la fecha de inicio para LMS como **1 dia habil anterior** a la fecha de inicio de Banner.

### Recomendaciones

- Mantener años actualizados (ejemplo: ampliar periodicamente el rango de años).
- Usar formato de fecha ISO `YYYY-MM-DD`.

## `DIRLOGS`

**Tipo:** `string`  
**Obligatoria:** si  
**Que representa:** carpeta destino de los logs de ejecucion.  
**Uso en el sistema:** se crea un archivo `log_YYYYMMDDHHMMSS.log` por ejecucion.

### Recomendaciones

- Usar una carpeta con permisos de escritura.
- Evitar rutas de red inestables para no perder trazabilidad.

## 6. Validaciones y fallas comunes

## Fallas frecuentes

- Ruta mal escrita o archivo movido.
- `HOLIDAYS_JSON` inexistente.
- `BDBANNER` con fechas vacias.
- `DIRLOGS` sin permisos de escritura.

## Comportamiento esperado ante error

- El error se registra en log.
- El error se muestra en pantalla.
- La iteracion afectada se detiene para evitar datos inconsistentes.

## 7. Buenas practicas operativas

- Versionar `external_files.json` junto al codigo.
- Evitar nombres de llave diferentes a los definidos (son sensibles al nombre exacto).
- Si cambia el entorno (rutas), actualizar primero el JSON y luego ejecutar.
- Revisar el log generado en `DIRLOGS` despues de cada corrida.

## 8. Mantenimiento recomendado

- Revisar anualmente `HOLIDAYS_JSON`.
- Verificar periodicidad de actualizacion de `BDBANNER` y `unidades_path`.
- Mantener copias de respaldo de archivos de entrada antes de ejecuciones masivas.

