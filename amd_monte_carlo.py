#!/usr/bin/env python3
"""Simulacion Monte Carlo de AMD con movimiento browniano geometrico (GBM).

Python >= 3.10. Instalar exportacion Excel: python -m pip install -r requirements.txt
Cada ejecucion genera resumen.csv y analisis_amd.xlsx en la carpeta actual.
Ejemplo (600 USD es un supuesto ilustrativo, NO una cotizacion actual):
    python amd_monte_carlo.py --s0 600
    python amd_monte_carlo.py --s0 600 --volatility 0.60 --seed 42 --csv resumen.csv
    python amd_monte_carlo.py --s0 600 --returns -0.10 0.10 0.25 0.40

Los rendimientos son expectativas anuales simples: -0.10 equivale a -10%.
Se usa mu = log(1 + rendimiento), de modo que E[S_1]/S0 - 1
coincide exactamente con el rendimiento anual indicado. La volatilidad
es anualizada (0.575 = 57.5%) y el tiempo se mide como meses / 12.

Modelo: ST = S0 * exp((mu - sigma**2 / 2)*T + sigma*sqrt(T)*Z),
con Z normal estandar. Se simula directamente la distribucion terminal:
no hay error de discretizacion temporal ni trayectorias intraperiodo.
No se utiliza la tasa libre de riesgo: son escenarios de rendimiento,
no una valoracion neutral al riesgo de opciones Black-Scholes.

LIMITACIONES: mu y sigma constantes; sin saltos, dividendos, comisiones,
impuestos ni cambios de regimen. No descarga ni verifica datos de mercado.
Las probabilidades dependen de los supuestos y no son pronosticos.
Este ejercicio educativo NO es una recomendacion financiera.
"""
import argparse
import math
import random
import sys
from pathlib import Path

from amd_report import export_reports

DEFAULT_RETURNS = (-0.10, 0.10, 0.25, 0.40)
DEFAULT_MONTHS = (1, 3, 6, 12)
PERCENTILES = (1, 5, 25, 50, 75, 95, 99)


def quantile(sorted_values, probability):
    """Percentil por interpolacion lineal entre observaciones ordenadas."""
    index = (len(sorted_values) - 1) * probability
    lower = math.floor(index)
    upper = math.ceil(index)
    weight = index - lower
    return sorted_values[lower] * (1 - weight) + sorted_values[upper] * weight


def validate(s0, volatility, returns, months, simulations, seed):
    if not math.isfinite(s0) or s0 <= 0:
        raise ValueError("s0 debe ser finito y mayor que cero.")
    if not math.isfinite(volatility) or volatility < 0:
        raise ValueError("volatility debe ser finita y no negativa.")
    if not returns or any(not math.isfinite(r) or r <= -1 for r in returns):
        raise ValueError("Cada rendimiento debe ser finito y mayor que -1.")
    if not months or any(not isinstance(m, int) or m <= 0 for m in months):
        raise ValueError("Los horizontes deben ser meses enteros positivos.")
    if not isinstance(simulations, int) or simulations < 1:
        raise ValueError("simulations debe ser un entero positivo.")
    if not isinstance(seed, int):
        raise ValueError("seed debe ser entero.")


def simulate(s0, volatility=0.575, returns=DEFAULT_RETURNS,
             months=DEFAULT_MONTHS, simulations=100_000, seed=42):
    """Devuelve una fila por escenario y horizonte; probabilidades en [0, 1]."""
    validate(s0, volatility, returns, months, simulations, seed)
    rng = random.Random(seed)
    # Numeros aleatorios comunes facilitan comparar escenarios. Cada fila
    # contiene simulations sorteos independientes; las filas estan correlacionadas
    # entre si y NO representan una trayectoria conjunta entre horizontes.
    normals = [rng.normalvariate(0.0, 1.0) for _ in range(simulations)]
    rows = []
    for annual_return in returns:
        mu = math.log1p(annual_return)
        for horizon in months:
            years = horizon / 12.0
            drift = (mu - volatility**2 / 2.0) * years
            diffusion = volatility * math.sqrt(years)
            try:
                terminal = sorted(
                    s0 * math.exp(drift + diffusion * z) for z in normals
                )
                theoretical_mean = s0 * math.exp(mu * years)
            except OverflowError as exc:
                raise ValueError("Parametros demasiado extremos para el modelo.") from exc
            if any(not math.isfinite(price) or price <= 0 for price in terminal):
                raise ValueError("Los parametros producen desbordamiento numerico.")
            row = {
                "s0": s0,
                "annual_return": annual_return,
                "volatility": volatility,
                "months": horizon,
                "simulations": simulations,
                "seed": seed,
                "mean": math.fsum(price / simulations for price in terminal),
                "median": quantile(terminal, 0.5),
                "theoretical_mean": theoretical_mean,
            }
            for percentile in PERCENTILES:
                row[f"p{percentile:02d}"] = quantile(terminal, percentile / 100)
            row.update({
                "prob_st_gt_s0": sum(p > s0 for p in terminal) / simulations,
                "prob_st_gt_700": sum(p > 700 for p in terminal) / simulations,
                "prob_st_lt_500": sum(p < 500 for p in terminal) / simulations,
                # Perdida terminal >20% equivale a ST < 0.8*S0.
                # NO mide caidas maximas ni tocar ese nivel antes del vencimiento.
                "prob_loss_gt_20pct": sum(p < 0.8 * s0 for p in terminal) / simulations,
            })
            rows.append(row)
    return rows


def main(argv=None):
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("--s0", type=float, required=True, help="Precio inicial en USD.")
    parser.add_argument("--volatility", type=float, default=0.575,
                        help="Volatilidad anual decimal; predeterminado: 0.575.")
    parser.add_argument("--returns", type=float, nargs="+", default=DEFAULT_RETURNS,
                        help="Rendimientos anuales simples en decimales.")
    parser.add_argument("--months", type=int, nargs="+", default=DEFAULT_MONTHS)
    parser.add_argument("--simulations", type=int, default=100_000,
                        help="Simulaciones por escenario y horizonte.")
    parser.add_argument("--seed", type=int, default=42,
                        help="Semilla fija para reproducibilidad.")
    parser.add_argument("--csv", type=Path, default=Path("resumen.csv"),
                        help="Ruta CSV; predeterminado: resumen.csv.")
    parser.add_argument("--xlsx", type=Path, default=Path("analisis_amd.xlsx"),
                        help="Ruta Excel; predeterminado: analisis_amd.xlsx.")
    args = parser.parse_args(argv)
    if args.csv.suffix.lower() != ".csv" or args.xlsx.suffix.lower() != ".xlsx":
        parser.error("Usa una ruta .csv y otra .xlsx.")
    if args.csv.resolve() == args.xlsx.resolve():
        parser.error("Las rutas CSV y Excel deben ser distintas.")
    try:
        rows = simulate(args.s0, args.volatility, args.returns, args.months,
                        args.simulations, args.seed)
    except ValueError as exc:
        parser.error(str(exc))

    print("AMD | Monte Carlo GBM | Ejercicio educativo; no es recomendacion financiera.")
    print(f"S0={args.s0:.2f} USD | sigma={args.volatility:.2%} | "
          f"N={args.simulations:,} por fila | semilla={args.seed}")
    print("Rendimientos anuales simples; probabilidades de precio al final del horizonte.")
    print("Precios y percentiles en USD; probabilidades mostradas en porcentaje.\n")
    print(f"{'R anual':>8} {'Meses':>5} {'Media':>10} {'Mediana':>10} "
          f"{'P05':>10} {'P95':>10} {'P(>S0)':>9} {'P(>700)':>9} "
          f"{'P(<500)':>9} {'P(perd>20%)':>12}")
    for row in rows:
        print(f"{row['annual_return']:>8.0%} {row['months']:>5} "
              f"{row['mean']:>10.2f} {row['median']:>10.2f} "
              f"{row['p05']:>10.2f} {row['p95']:>10.2f} "
              f"{row['prob_st_gt_s0']:>9.2%} {row['prob_st_gt_700']:>9.2%} "
              f"{row['prob_st_lt_500']:>9.2%} {row['prob_loss_gt_20pct']:>12.2%}")
    print("\nPercentiles adicionales (USD):")
    for row in rows:
        values = " | ".join(f"P{p:02d}={row[f'p{p:02d}']:.2f}" for p in PERCENTILES)
        print(f"r={row['annual_return']:+.0%}, {row['months']:2d} meses | {values}")
    print("\nError estandar Monte Carlo de cada probabilidad: sqrt(p*(1-p)/N).")
    print(f"Maximo por fila: {0.5 / math.sqrt(args.simulations):.4%} "
          "(no incluye incertidumbre de los supuestos).")

    try:
        export_reports(rows, args.csv, args.xlsx)
    except (OSError, RuntimeError) as exc:
        parser.exit(1, f"No se pudieron guardar ambos archivos: {exc}\n")
    print(f"CSV guardado en {args.csv.resolve()} (probabilidades entre 0 y 1).")
    print(f"Excel guardado en {args.xlsx.resolve()} (tablas y graficas).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
