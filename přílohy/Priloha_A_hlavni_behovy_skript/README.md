# Příloha A

## Hlavní běhový skript pro produkční stanici

Tato příloha obsahuje hlavní produkční skript, který představuje jádro rozhodovací logiky pro řízení natápění linky POL1 v provozním prostředí.

## Hlavní soubor

- [comm_decision_RELIANCE_PROD_WIN7.py](./comm_decision_RELIANCE_PROD_WIN7.py)

## Co příloha obsahuje

- načítání provozních dat a stavových informací
- vyhodnocení podmínek pro start a stop natápění
- výpočet rozhodnutí nad provozními vstupy
- zápis výsledných povelů do komunikační fronty
- implementaci přizpůsobenou prostředí Windows 7 a Python 3.4

## Význam v práci

Tato příloha dokumentuje finální běhovou podobu navrženého řešení v produkčním nasazení a ukazuje, jak byla predikční a rozhodovací logika integrována do reálného provozu.
