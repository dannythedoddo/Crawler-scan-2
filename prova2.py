import os
import re
import gzip
import time
import requests
import openpyxl
from datetime import datetime, timezone
from openpyxl import Workbook
from warcio.archiveiterator import ArchiveIterator
from concurrent.futures import ThreadPoolExecutor, as_completed

# ---------------------------------------------------------
# CONFIGURAZIONE
# ---------------------------------------------------------
BASE = "https://data.commoncrawl.org/"
PRIMO_CRAWL = "CC-MAIN-2013-20"  # primo crawl in formato WARC/WET standard

N_WORKER = 16  # quanti file WET scaricare/analizzare in parallelo

MAX_SEQ_LEN = 24  # tetto massimo di parole consecutive per una catena

# Cartella dove leggere/scrivere vocabolario, tolleranza, checkpoint, risultati.
# In locale resta "." (cartella corrente). Su GitHub Actions viene impostata
# tramite variabile d'ambiente DATA_DIR per puntare al repository dati privato.
DATA_DIR = os.environ.get("DATA_DIR", ".")
os.makedirs(DATA_DIR, exist_ok=True)

# Se impostato (in minuti) tramite variabile d'ambiente, lo script si ferma da solo
# poco prima di questo limite, salva tutto e chiude senza attendere gli altri file.
# 0 = nessun limite (uso normale su PC di casa).
MAX_RUNTIME_MINUTES = int(os.environ.get("MAX_RUNTIME_MINUTES", "0"))

# Data del giorno corrente (UTC, coerente con l'orario dei workflow schedulati)
OGGI = datetime.now(timezone.utc).strftime("%m_%d")

FILE_VOCABOLARIO = os.path.join(DATA_DIR, "vocabolario2.txt")
FILE_TOLLERANZA = os.path.join(DATA_DIR, "tolleranza2.txt")
FILE_CRAWL_ATTUALE = os.path.join(DATA_DIR, "crawl_attuale2.txt")  # quale crawl stiamo scansionando ora
FILE_STATO_OGGI = os.path.join(DATA_DIR, "stato_oggi2.txt")        # riepilogo per il workflow (email, ecc.)

# Questi tre dipendono dal crawl in corso: vengono assegnati in MAIN
# dopo aver determinato automaticamente CRAWL_ID.
FILE_LISTA_PATH = None
FILE_CHECKPOINT = None
FILE_RISULTATI = None

# regex per tokenizzare: cattura solo lettere (accentate incluse), non numeri/simboli
TOKEN_REGEX = re.compile(r"[^\W\d_]+", re.UNICODE)


# ---------------------------------------------------------
# Caricamento vocabolario da .txt
# ---------------------------------------------------------
def carica_vocabolario(path):
    parole = set()
    with open(path, "r", encoding="utf-8") as f:
        for riga in f:
            pulita = riga.strip().strip(",").strip('"').strip().lower()
            if pulita:
                parole.add(pulita)
    return parole


# ---------------------------------------------------------
# PASSO 3: scarica la lista di TUTTI i file WET del crawl
# ---------------------------------------------------------
def scarica_lista_path(crawl_id):
    url_lista = f"{BASE}crawl-data/{crawl_id}/wet.paths.gz"
    r = requests.get(url_lista)
    r.raise_for_status()
    testo = gzip.decompress(r.content).decode("utf-8")
    paths = [BASE + riga.strip() for riga in testo.splitlines() if riga.strip()]
    return paths


# ---------------------------------------------------------
# Sequenza automatica dei crawl: dal primo compatibile al più recente
# ---------------------------------------------------------
PATTERN_CRAWL_ID = re.compile(r"^CC-MAIN-(\d{4})-(\d{2})$")


def carica_lista_crawl():
    """Scarica l'elenco ufficiale dei crawl e restituisce solo quelli in formato
    standard (anno-settimana), ordinati cronologicamente, da PRIMO_CRAWL in poi."""
    r = requests.get("https://index.commoncrawl.org/collinfo.json")
    r.raise_for_status()
    dati = r.json()

    validi = []
    for voce in dati:
        m = PATTERN_CRAWL_ID.match(voce["id"])
        if m:
            anno, settimana = int(m.group(1)), int(m.group(2))
            validi.append((anno, settimana, voce["id"]))
    validi.sort()  # ordine cronologico crescente
    lista_id = [v[2] for v in validi]

    if PRIMO_CRAWL in lista_id:
        lista_id = lista_id[lista_id.index(PRIMO_CRAWL):]
    return lista_id


def leggi_crawl_attuale(lista_crawl):
    try:
        with open(FILE_CRAWL_ATTUALE, "r") as f:
            valore = f.read().strip()
        if valore in lista_crawl:
            return valore
    except FileNotFoundError:
        pass
    # non esiste ancora o non è valido: iniziamo dal primo della sequenza
    primo = lista_crawl[0]
    scrivi_crawl_attuale(primo)
    return primo


def scrivi_crawl_attuale(crawl_id):
    with open(FILE_CRAWL_ATTUALE, "w") as f:
        f.write(crawl_id)


# ---------------------------------------------------------
# Tokenizzazione + ricerca catene di parole consecutive nel vocabolario
# ---------------------------------------------------------
MAX_RIPETIZIONI_PAROLA = 2  # una stessa parola non può comparire più di N volte nella stessa catena


def trova_catene(testo, vocabolario, tolleranza, lunghezza_minima, lunghezza_massima=MAX_SEQ_LEN):
    tokens = TOKEN_REGEX.findall(testo.lower())

    catene_trovate = []
    catena_corrente = []
    conta_parole = {}      # occorrenze di ogni parola nella catena corrente
    conta_tolleranza = 0   # quante parole della catena corrente sono "a tolleranza"

    for parola in tokens:
        if parola in vocabolario:
            occorrenze_future = conta_parole.get(parola, 0) + 1
            lunghezza_futura = len(catena_corrente) + 1
            tolleranza_futura = conta_tolleranza + (1 if parola in tolleranza else 0)
            tetto_tolleranza = lunghezza_futura // 4  # un quarto della lunghezza, arrotondato per difetto

            supera_ripetizioni = occorrenze_future > MAX_RIPETIZIONI_PAROLA
            supera_tolleranza = tolleranza_futura > tetto_tolleranza

            if supera_ripetizioni or supera_tolleranza:
                # la parola violerebbe un vincolo: chiudo la catena qui
                # e ricomincio una nuova catena partendo da questa stessa parola
                if len(catena_corrente) >= lunghezza_minima:
                    catene_trovate.append(" ".join(catena_corrente))
                catena_corrente = [parola]
                conta_parole = {parola: 1}
                conta_tolleranza = 1 if parola in tolleranza else 0
                continue

            catena_corrente.append(parola)
            conta_parole[parola] = occorrenze_future
            conta_tolleranza = tolleranza_futura

            # tetto massimo di lunghezza raggiunto: chiudo la catena qui, non la estendo oltre
            if len(catena_corrente) == lunghezza_massima:
                if len(catena_corrente) >= lunghezza_minima:
                    catene_trovate.append(" ".join(catena_corrente))
                catena_corrente = []
                conta_parole = {}
                conta_tolleranza = 0
        else:
            if len(catena_corrente) >= lunghezza_minima:
                catene_trovate.append(" ".join(catena_corrente))
            catena_corrente = []
            conta_parole = {}
            conta_tolleranza = 0

    # controlla l'ultima catena rimasta a fine testo
    if len(catena_corrente) >= lunghezza_minima:
        catene_trovate.append(" ".join(catena_corrente))

    return catene_trovate


# ---------------------------------------------------------
# PASSO 4: cerca catene di parole all'interno di UN file WET
# ---------------------------------------------------------
def cerca_in_wet(wet_url, vocabolario, tolleranza, lunghezza_minima):
    trovati = []  # lista di tuple (url_pagina, catena_testo)
    n_record = 0
    with requests.get(wet_url, stream=True, timeout=60) as r:
        r.raise_for_status()
        for record in ArchiveIterator(r.raw):
            if record.rec_type != 'conversion':
                continue
            n_record += 1
            testo = record.content_stream().read().decode('utf-8', errors='ignore')
            catene = trova_catene(testo, vocabolario, tolleranza, lunghezza_minima)
            if catene:
                url_pagina = record.rec_headers.get_header('WARC-Target-URI')
                for catena in catene:
                    trovati.append((url_pagina, catena))
    return trovati, n_record


# ---------------------------------------------------------
# Gestione checkpoint (path già completati)
# ---------------------------------------------------------
def carica_checkpoint():
    try:
        with open(FILE_CHECKPOINT, "r") as f:
            return set(riga.strip() for riga in f if riga.strip())
    except FileNotFoundError:
        return set()


def segna_come_completato(path):
    with open(FILE_CHECKPOINT, "a") as f:
        f.write(path + "\n")


# ---------------------------------------------------------
# Gestione Excel incrementale
# ---------------------------------------------------------
def apri_o_crea_excel():
    try:
        wb = openpyxl.load_workbook(FILE_RISULTATI)
        ws = wb.active
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.title = "Risultati"
        ws.append(["Sequenza trovata", "URL"])
        wb.save(FILE_RISULTATI)
    return wb, ws


def salva_risultati_excel(wb, ws, trovati):
    if not trovati:
        return
    for url_pagina, catena in trovati:
        ws.append([catena, url_pagina])
    wb.save(FILE_RISULTATI)


# ---------------------------------------------------------
# Worker per un singolo file (usato dai thread paralleli)
# ---------------------------------------------------------
def worker(path, vocabolario, tolleranza, lunghezza_minima):
    t0 = time.time()
    try:
        trovati, n_record = cerca_in_wet(path, vocabolario, tolleranza, lunghezza_minima)
        dt = time.time() - t0
        return {
            "path": path,
            "ok": True,
            "n_record": n_record,
            "trovati": trovati,
            "tempo": dt,
        }
    except Exception as e:
        return {
            "path": path,
            "ok": False,
            "errore": str(e),
        }


# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
if __name__ == "__main__":
    # 0. carica vocabolario e tolleranza, chiede la lunghezza minima della sequenza
    vocabolario = carica_vocabolario(FILE_VOCABOLARIO)
    print(f"Vocabolario caricato: {len(vocabolario)} parole")

    tolleranza = carica_vocabolario(FILE_TOLLERANZA)
    parole_non_in_vocabolario = tolleranza - vocabolario
    if parole_non_in_vocabolario:
        print(
            f"ATTENZIONE: {len(parole_non_in_vocabolario)} parole in {FILE_TOLLERANZA} "
            f"non sono presenti in {FILE_VOCABOLARIO} e verranno ignorate: "
            f"{sorted(parole_non_in_vocabolario)}"
        )
        tolleranza = tolleranza & vocabolario
    print(f"Tolleranza caricata: {len(tolleranza)} parole")

    valore_ambiente = os.environ.get("LUNGHEZZA_MINIMA")
    if valore_ambiente is not None:
        lunghezza_minima = int(valore_ambiente)
        print(f"Lunghezza minima della sequenza (da variabile d'ambiente): {lunghezza_minima}")
    else:
        lunghezza_minima = int(input("Lunghezza minima della sequenza da cercare (numero di parole): "))

    # 0b. determina automaticamente su quale crawl lavorare oggi
    print("Scarico l'elenco ufficiale dei crawl disponibili...")
    lista_crawl = carica_lista_crawl()
    print(f"Crawl compatibili disponibili (da {PRIMO_CRAWL} al più recente): {len(lista_crawl)}")

    crawl_forzato = os.environ.get("CRAWL_ID_FORZATO")
    if crawl_forzato:
        CRAWL_ID = crawl_forzato
        print(f"Crawl forzato manualmente (via CRAWL_ID_FORZATO): {CRAWL_ID}")
    else:
        CRAWL_ID = leggi_crawl_attuale(lista_crawl)
        print(f"Crawl attuale nella sequenza automatica: {CRAWL_ID}")

    FILE_LISTA_PATH = os.path.join(DATA_DIR, f"wet_paths_{CRAWL_ID}.txt")
    FILE_CHECKPOINT = os.path.join(DATA_DIR, f"checkpoint2_{CRAWL_ID}.txt")
    FILE_RISULTATI = os.path.join(DATA_DIR, f"risultati2_{OGGI}_{CRAWL_ID}.xlsx")

    # 1. lista completa dei file WET (la scarica solo se non l'abbiamo già salvata)
    try:
        with open(FILE_LISTA_PATH, "r") as f:
            paths = [riga.strip() for riga in f if riga.strip()]
        print(f"Lista file WET già presente su disco: {len(paths)} file.")
    except FileNotFoundError:
        print(f"Scarico la lista dei file WET per il crawl {CRAWL_ID}...")
        paths = scarica_lista_path(CRAWL_ID)
        with open(FILE_LISTA_PATH, "w") as f:
            f.write("\n".join(paths))
        print(f"Numero totale di file WET nel crawl: {len(paths)}")

    # 2. rimuovo dalla lista i path già completati in run precedenti
    completati = carica_checkpoint()
    da_fare = [p for p in paths if p not in completati]
    print(f"File già completati: {len(completati)} | File da processare ora: {len(da_fare)}")

    # 3. apre (o crea) il file Excel del giorno, sempre, anche se non resta nulla da fare
    wb, ws = apri_o_crea_excel()

    if not da_fare:
        print("Tutti i file di questo crawl sono già stati processati in run precedenti.")
        n_completati_ora = 0
        n_match_totali = 0
        t_inizio = time.time()
    else:
        # 4. elaborazione parallela con salvataggio incrementale e budget di tempo opzionale
        n_completati_ora = 0
        n_match_totali = 0
        t_inizio = time.time()
        tempo_scaduto = False

        executor = ThreadPoolExecutor(max_workers=N_WORKER)
        future_to_path = {
            executor.submit(worker, p, vocabolario, tolleranza, lunghezza_minima): p
            for p in da_fare
        }

        for future in as_completed(future_to_path):
            risultato = future.result()
            path = risultato["path"]

            if not risultato["ok"]:
                print(f"[ERRORE] {path.split('/')[-1]}: {risultato['errore']}")
                # NON segniamo come completato: verrà ritentato al prossimo avvio
            else:
                n_completati_ora += 1
                n_match_totali += len(risultato["trovati"])

                # salvo risultati (Excel) e checkpoint SUBITO, non alla fine
                salva_risultati_excel(wb, ws, risultato["trovati"])
                segna_come_completato(path)

                print(
                    f"[{n_completati_ora}/{len(da_fare)}] "
                    f"{path.split('/')[-1]} -> "
                    f"record: {risultato['n_record']}, "
                    f"catene trovate: {len(risultato['trovati'])}, "
                    f"tempo: {risultato['tempo']:.1f}s"
                )

            # controllo budget di tempo (solo se impostato tramite variabile d'ambiente)
            if MAX_RUNTIME_MINUTES and (time.time() - t_inizio) > MAX_RUNTIME_MINUTES * 60:
                tempo_scaduto = True
                break

        if tempo_scaduto:
            print(
                f"\nBudget di tempo di {MAX_RUNTIME_MINUTES} minuti raggiunto. "
                f"Annullo i file non ancora avviati e chiudo (quelli in corso terminano regolarmente)..."
            )
            executor.shutdown(wait=True, cancel_futures=True)
        else:
            executor.shutdown(wait=True)

        dt_totale = time.time() - t_inizio
        print(f"\nFatto. File processati in questa sessione: {n_completati_ora}")
        print(f"Catene totali trovate in questa sessione: {n_match_totali}")
        print(f"Tempo totale sessione: {dt_totale/60:.1f} minuti")
        print(f"Risultati salvati in: {FILE_RISULTATI}")
        print(f"Checkpoint salvato in: {FILE_CHECKPOINT}")

    # 5. valuta se il crawl è stato completato del tutto, ed eventualmente avanza alla prossimo
    completati_finali = carica_checkpoint()
    crawl_completato = len(completati_finali) >= len(paths)
    crawl_successivo = ""

    if crawl_completato:
        print(f"\nIl crawl {CRAWL_ID} risulta completato per intero ({len(completati_finali)}/{len(paths)} file).")
        if not crawl_forzato:
            try:
                idx = lista_crawl.index(CRAWL_ID)
                crawl_successivo = lista_crawl[idx + 1]
                scrivi_crawl_attuale(crawl_successivo)
                print(f"Passo automaticamente al prossimo crawl: {crawl_successivo}")
            except (ValueError, IndexError):
                crawl_successivo = ""
                print("Questo era l'ultimo crawl disponibile: in attesa che Common Crawl ne pubblichi uno nuovo.")

    # 6. scrive un riepilogo dello stato del giorno, usato dal workflow per l'email
    with open(FILE_STATO_OGGI, "w") as f:
        f.write(f"CRAWL_ID_USATO={CRAWL_ID}\n")
        f.write(f"CRAWL_COMPLETATO={'si' if crawl_completato else 'no'}\n")
        f.write(f"CRAWL_SUCCESSIVO={crawl_successivo}\n")
