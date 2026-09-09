# Calendario Partite (Serie A, UCL, UEL) con Emittenti TV

Questo repository aggiorna automaticamente ogni giorno alle 08:00 italiane il file **`calendario_partite.csv`** con la programmazione aggiornata delle partite e delle relative emittenti TV per:
- **Serie A** (DAZN, DAZN / SKY)
- **UEFA Champions League (UCL)** (Sky, Amazon Prime Video)
- **UEFA Europa League (UEL)** (Sky)

## Formato del file CSV

Il file è delimitato da virgola (`,`) con codifica UTF-8 BOM e presenta le seguenti colonne:
1. `Number`: Competizione (`Serie A`, `UCL`, `UEL`)
2. `Planned Start Date`: Data e ora inizio (`GG/MM/AAAA HH:mm`)
3. `Planned End Date`: Data e ora fine (`GG/MM/AAAA HH:mm`, +2 ore)
4. `Short Description`: Partita (`Squadra Casa - Squadra Ospite`)
5. `State`: vuoto (null)
6. `Type`: vuoto (null)
7. `Impatto`: vuoto (null)
8. `Note`: Emittente TV (`Sky`, `DAZN`, `DAZN / SKY`, `Amazon Prime Video`)

## Come collegarlo ad Excel (Aggiornamento automatico nel Cloud)

1. Apri **Microsoft Excel**.
2. Vai nella scheda **Dati** -> **Da Web** (o *Recupera dati* -> *Da altre origini* -> *Da Web*).
3. Incolla l'URL Raw del file CSV:
   `https://raw.githubusercontent.com/francolacioppa/calendario-partite/main/calendario_partite.csv`
4. Seleziona il delimitatore **Virgola** (`,`) e la codifica **65001: Unicode (UTF-8)**.
5. Clicca su **Carica**.
