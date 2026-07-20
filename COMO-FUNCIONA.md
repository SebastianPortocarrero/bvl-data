# bvl-data — base de datos de EEFF oficiales de SMV (Perú)

Cada `{TICKER}.json` contiene los estados financieros trimestrales oficiales de la SMV
ya procesados (montos en MILLONES; los EEFF de SMV vienen en miles). `index.json` lista
lo disponible. Los lee el GPT "Analista Fundamental BVL" vía Action.

## Incorporar un aporte nuevo (sirve en la nube, sin depender de ninguna Mac)

Un aporte es un `.zip` con Excel oficiales de SMV nombrados `SMV_{EMPRESA}_{AÑO}_TRIMESTRE{N}.xlsx`
(los genera el script `descargar_smv.py` que usa el profesor).

```bash
pip install openpyxl
python3 procesar_aporte.py APORTE.zip            # el ticker se resuelve solo
python3 procesar_aporte.py APORTE.zip --ticker XXXXXC1   # o se fuerza
git add -A && git commit -m "aporte: XXXXXC1" && git push
```

`procesar_aporte.py` parsea los Excel respetando las trampas de los EEFF de SMV
(deuda financiera en DOS líneas, utilidad neta duplicada en el flujo de efectivo,
cifras atribuibles a la controladora, etiquetas distintas en bancos), **fusiona** con
los trimestres que ya existían para ese ticker (nunca los pierde) y regenera `index.json`.

Profundidad objetivo: 20 trimestres (5 años). El script avisa si un ticker queda por debajo.
