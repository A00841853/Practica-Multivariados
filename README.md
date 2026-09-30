# Practica-Multivariados

## Analisis Monte Carlo de AMD

Requiere Python 3.10 o posterior. Desde la carpeta del repositorio:

```powershell
git pull --ff-only
py -m pip install -r requirements.txt
py amd_monte_carlo.py --s0 604.72
```

604.72 es un ejemplo proporcionado por el usuario, no una cotizacion actual
verificada. Sustituyelo por tu precio inicial. En macOS/Linux usa python3.

Cada ejecucion crea en la carpeta actual:
- **resumen.csv**: las 20 columnas originales, con precision completa.
- **analisis_amd.xlsx**: hojas Resumen y Datos completos, filtros, porcentajes y
  precios formateados y cuatro graficas editables: media/mediana, percentiles
  P05/mediana/P95 en el mayor plazo seleccionado, probabilidad de superar USD 700
  y probabilidad de perder mas de 20% por plazo.

Para abrirlo en Windows:
```powershell
start excel analisis_amd.xlsx
```

Puedes cambiar los supuestos y los nombres de salida:
```powershell
py amd_monte_carlo.py --s0 604.72 --volatility 0.60 --csv resumen_60.csv --xlsx analisis_60.xlsx
py amd_monte_carlo.py --s0 604.72 --returns -0.10 0.10 0.25 0.40 --months 1 3 6 12 --simulations 100000 --seed 42
```

0.60 significa 60% de volatilidad anual. El 57.5% predeterminado es un supuesto,
no una estimacion actual. Las graficas admiten rendimientos y plazos personalizados.
No se asigna mas credibilidad a un escenario: el rendimiento anual esperado es
un supuesto de entrada, no una prediccion.

Los resultados se calculan en Python. Excel guarda esos resultados y vincula el
resumen y las graficas a las celdas de datos. **Editar precio o volatilidad en
Excel no vuelve a simular**: para cambiar supuestos, ejecuta de nuevo el programa.

El CSV usa probabilidades de 0 a 1; Excel las muestra como porcentajes.
P05-P95 describe el 90% central de precios simulados, no un intervalo de confianza
de la media. Perdida >20% significa ST < 0.8*S0 al final del horizonte, no una
caida maxima durante el periodo. theoretical_mean es la media exacta GBM.
Los archivos contienen resumenes, no los 100000 precios individuales.

Se conservan los escenarios -10%, 10%, 25%, 40%, los plazos 1, 3, 6, 12 meses,
100000 simulaciones por combinacion y semilla 42. El modelo supone volatilidad
constante y precios lognormales. No es recomendacion financiera.

Se reemplazan los archivos con los nombres elegidos. Cierralos en Excel antes
de ejecutar; si estan bloqueados se informa del problema. Para conservar una
corrida anterior, utiliza otros nombres. Las carpetas de destino deben existir.
Si falta XlsxWriter, ejecuta el comando de instalacion anterior.

Pruebas:
```powershell
py -m unittest discover -s tests
```
