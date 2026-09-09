# Calendario Partite (Serie A, UCL, UEL) con Emittenti TV

Questo repository aggiorna automaticamente ogni giorno alle 08:00 italiane il file **`calendario_partite.csv`** con la programmazione aggiornata delle partite e delle relative emittenti TV per:
- **Serie A** (DAZN, DAZN / SKY)
- **UEFA Champions League (UCL)** (Sky, Amazon Prime Video)
- **UEFA Europa League (UEL)** (Sky)

## Come collegarlo ad Excel (Aggiornamento automatico nel Cloud)

Puoi collegare questo file direttamente ad Excel per visualizzarlo sempre aggiornato senza dover lanciare alcuno script sul tuo PC:

1. Apri **Microsoft Excel**.
2. Vai nella scheda **Dati** -> **Da Web** (o *Recupera dati* -> *Da altre origini* -> *Da Web*).
3. Incolla l'URL Raw del file CSV:
   `https://raw.githubusercontent.com/francolacioppa/calendario-partite/main/calendario_partite.csv`
4. Seleziona il delimitatore **Punto e virgola** (`;`) e la codifica **65001: Unicode (UTF-8)**.
5. Clicca su **Carica**.

Da questo momento, ogni volta che apri la cartella di lavoro Excel o clicchi su **Dati -> Aggiorna tutti**, Excel scaricherà l'ultima versione aggiornata da GitHub.

