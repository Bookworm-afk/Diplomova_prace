# Přílohy k diplomové práci

Tento repozitář slouží jako přehledně členěný archiv hlavních příloh k diplomové práci zaměřené na prediktivní řízení předehřevu linky POL1. Obsah je rozdělen do samostatných složek podle označení příloh A-H, aby byl na GitHubu snadno čitelný a přímo navazoval na seznam příloh uvedený v textu práce.

## Jak se orientovat

- Každá příloha má vlastní složku a vlastní `README.md` se stručným představením obsahu.
- V každé složce je uložen hlavní soubor přílohy.
- Kořenový `README.md` funguje jako rozcestník pro rychlou orientaci.

## Přehled příloh

| Příloha | Obsah | Hlavní soubor |
| --- | --- | --- |
| [A](./Priloha_A_hlavni_behovy_skript) | Hlavní běhový skript produkční stanice s rozhodovací logikou natápění | [`comm_decision_RELIANCE_PROD_WIN7.py`](./Priloha_A_hlavni_behovy_skript/comm_decision_RELIANCE_PROD_WIN7.py) |
| [B](./Priloha_B_komunikacni_skript_reliance) | Komunikační VBScript pro prostředí Reliance a práci s provozními tagy | [`Reliance_VBscript.vbs`](./Priloha_B_komunikacni_skript_reliance/Reliance_VBscript.vbs) |
| [C](./Priloha_C_baseline_snapshot) | Skript pro vytvoření baseline snapshotu a souvisejících evaluačních přehledů | [`generate_baseline_snapshot.py`](./Priloha_C_baseline_snapshot/generate_baseline_snapshot.py) |
| [D](./Priloha_D_manualni_anotace_segmentu) | Nástroj pro manuální anotaci segmentů při přípravě trénovacích dat | [`manual_labeling.py`](./Priloha_D_manualni_anotace_segmentu/manual_labeling.py) |
| [E](./Priloha_E_shadow_log) | Ukázka provozního logu ze shadow režimu při testování rozhodovací logiky | [`Shadow_log.txt`](./Priloha_E_shadow_log/Shadow_log.txt) |
| [F](./Priloha_F_eval_metrics) | Agregované metriky baseline evaluace modelů v režimech heating a cooling | [`eval_metrics.json`](./Priloha_F_eval_metrics/eval_metrics.json) |
| [G](./Priloha_G_manifest_segmentu) | Manifest rozdělení datových segmentů pro augmented variantu se sin/cos příznaky | [`segment_split_manifest_augmented_cos_sin.json`](./Priloha_G_manifest_segmentu/segment_split_manifest_augmented_cos_sin.json) |
| [H](./Priloha_H_midnight_backtest) | Export offline midnight backtestu s porovnáním predikce a skutečného dosažení cíle | [`midnight_backtest_performance (3).xlsx`](./Priloha_H_midnight_backtest/midnight_backtest_performance%20(3).xlsx) |

## Poznámka

Struktura repozitáře je přizpůsobena prezentaci příloh na GitHubu. Jednotlivé podsložky odpovídají označení příloh použitých v diplomové práci.
