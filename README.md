# Calculadora de precios de la GIOEE

Aplicación local para la estimación de precios de Cartera, con base en el
`ModeloCartera-2025` y en el procedimiento operativo disponible. Los insumos
originales dentro de `Data_Informacion` se leen sin modificarse.

## Ejecutar

En Windows, desde este directorio:

Configura la conexión a la base de datos en la sesión de PowerShell (no guardes
la contraseña en el código):

```powershell
.\.venv\Scripts\python -m pip install -r requirements.txt
$env:DATABASE_URL = "postgresql://USUARIO:CONTRASENA@127.0.0.1:5432/calculo_precios"
.\.venv\Scripts\python setup_database.py
.\.venv\Scripts\python app.py
```

La aplicación inicia en `http://127.0.0.1:4001`. PostgreSQL conserva el
historial de versiones y cuál set está vigente; cada versión guarda propiedades,
catálogo y marcadores normalizados en archivos Parquet bajo
`Resultados\Cartera\Insumos`. El panel izquierdo organiza los tipos de cálculo;
por ahora, solo Precios Cartera está implementado.

En el primer cálculo, si aún no hay una versión vigente, se leen los insumos
2025 disponibles y se guardan como Parquet. En cálculos posteriores se reutiliza
esa versión sin volver a abrir los Excel. Si se carga un nuevo archivo de
propiedades o catálogo y pasa la validación, se crea una nueva versión completa
y se convierte en el set vigente. Una carga inválida no cambia la versión
vigente. Los marcadores de los escenarios Alto, Medio y Bajo se incluyen en el
set inicial.

## Resultados

Los cálculos son bajo demanda. Cada ejecución conserva un archivo Parquet y
actualiza `Resultados\Cartera\bitacora.parquet`; ambos registran la versión de
insumos utilizada. Desde la interfaz, el reporte se descarga en Excel con el
nombre `precios_cartera_ESCENARIO_YYYYMMDD_version_NN.xlsx` (por ejemplo,
`precios_cartera_medio_20261115_version_03.xlsx`), incluyendo la tabla anual y
una hoja de metadatos. Las versiones anteriores de insumos se conservan para
trazabilidad. El reporte conserva los años 2025–2124 del entregable de
referencia; desde 2051 mantiene el último marcador anual disponible, conforme
al comportamiento del modelo 2025.

La interfaz muestra gráficas para revisar las series de marcadores y la
distribución de API/poder calorífico. La gráfica de aceite incluye las
referencias Homol (37.1° API), Ket (15.37° API) y Kayab (8.6° API), calculadas
por año con los marcadores del escenario activo. No se asignan umbrales
estadísticos de anomalía que no estén definidos en el procedimiento.

Se puede elegir entre validación de precios por campo (fórmula normal) y reporte
final de proceso. En validación, el condensado con `Calculado_como_Aceite = 13`
usa la curva de aceite; en el reporte final Nejo usa directamente su marcador
de condensado. El gas conserva la conversión por poder calorífico en ambos
modos. En el set vigente v4, la propiedad de condensado de Nejo no tiene
bandera `Calculado_como_Aceite`, y el catálogo activo no incluye Nejo; por ello
Nejo no aparece actualmente en los resultados, sin omitir la validación del
catálogo. Además, la propiedad vigente asigna a su gas el marcador `GH Dulce
Norte`, mientras el procedimiento menciona `Gas seco de campos`; la equivalencia
debe confirmarse antes de cambiar esa asignación. El modo queda registrado en
el Parquet y en la bitácora.

## Pruebas

```powershell
.\.venv\Scripts\python -m unittest discover -s tests -v
```

La prueba de regresión compara los tres escenarios contra los entregables 2025.
Los renglones históricos Nejo / Condensado y Nejo / Gas corresponden a la
validación por campo y no son el valor del reporte final: para el reporte final
se aplican los marcadores actuales. La regresión verifica estos dos renglones
contra los marcadores y compara los demás con los entregables 2025.
Los procesos PDEs, Deterioro, Cartera Exploración y Programa Operativo quedan
pendientes de su definición.
