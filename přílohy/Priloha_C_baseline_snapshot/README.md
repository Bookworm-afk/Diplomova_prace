# Příloha C

## Skript pro tvorbu baseline snapshotu

Tato příloha obsahuje skript pro vytvoření reprodukovatelného baseline snapshotu, který slouží jako opora pro vyhodnocení modelů a souhrnné porovnání dostupných modelových artefaktů.

## Hlavní soubor

- [generate_baseline_snapshot.py](./generate_baseline_snapshot.py)

## Co příloha obsahuje

- výpočet statistik segmentů pro heating a cooling
- sestavení inventáře dostupných modelů
- výpočet globálních a pozičních evaluačních metrik
- uložení výstupů do reportovací složky `NN/reports/baseline_*`

## Význam v práci

Tato příloha reprezentuje skript, z něhož vznikají baseline reporty používané v praktické části práce při interpretaci přesnosti a stability modelů.
