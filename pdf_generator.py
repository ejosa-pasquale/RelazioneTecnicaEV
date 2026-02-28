from __future__ import annotations

"""PDF generator.

Obiettivi della versione:
- Cover page in stile "relazione tecnico-specialistica" (3 riquadri: titolo/indice/firma)
- Pagina "Elenco delle revisioni" (subito dopo la cover)
- Header/Footer e numerazione "Pagina X di Y" su tutte le pagine successive
- Wording e struttura più legali/rigorosi (capitoli allineati al template: 1..6)
- Contenuti condizionali: stampa solo sezioni significative

Nota: la cover riprende l'impostazione a tre riquadri del PDF campione.
"""

from io import BytesIO
from typing import Dict, List, Any, Optional

from xml.sax.saxutils import escape

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    Flowable,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- Template contenuti (derivato dal DOCX campione) ---
TEMPLATE_SECTIONS_DEFAULT = [{'default': 'L’area di intervento consiste nell’installazione di tipo Outdoor (esterno) di punto di ricarica.',
  'num': 9,
  'title': 'AREA DI INTERVENTO E TIPO DI ATTIVITÀ.'},
 {'default': 'Trattasi di impianto elettrico, utilizzatore di Ia categoria (50 V < Vn <1000 V), con alimentazione da '
             'rete privata di bassa tensione tramite un unico punto di consegna (POD) dell’Ente Distributore.',
  'num': 10,
  'title': 'TIPO DI IMPIANTO.'},
 {'default': 'Il punto di origine dell’impianto dista circa 60 metri con il punto di consegna da parte dell’Ente '
             'Distributore.\n'
             '\n'
             'Per maggiori dettagli si rimanda ai disegni e agli elaborati tecnici allegati a questa relazione.',
  'num': 11,
  'title': 'PUNTO DI ORIGINE.'},
 {'default': 'La fornitura sarà di tipo a corrente alternata monofase in bassa tensione 230 V, a frequenza nominale 50 '
             'Hz.',
  'num': 12,
  'title': 'SISTEMA DI FORNITURA.'},
 {'default': 'Gli impianti elettrici presenti avranno le seguenti tensioni:\n'
             '\n'
             'Circuiti elettrici di tipo monofase a 230 V;',
  'num': 13,
  'title': 'TENSIONE NOMINALE.'},
 {'default': 'Si tratta di un impianto di tipo TT, con impianto di terra comune a tutte le sezioni dell’impianto.',
  'num': 14,
  'title': 'SISTEMA DI DISTRIBUZIONE.'},
 {'default': 'Per la fornitura delle parti comuni, la corrente di corto circuito presunta per guasto trifase nel punto '
             'di fornitura è pari a 10 kA. (salvo diversa indicazione del Distributore), mentre per la corrente di '
             'corto circuito presunta per guasto monofase nel punto di fornitura è pari a 6 kA.\n'
             '\n'
             'In ogni caso il dispositivo generale avrà un potere d’interruzione nominale monofase 230V maggiore della '
             'corrente di corto circuito monofase presunta nel punto di installazione.',
  'num': 15,
  'title': 'CORRENTE DI CORTO CIRCUITO.'},
 {'default': 'La potenza impegnata tiene conto di tutti i fattori di contemporaneità e di utilizzazione delle varie '
             'utenze presenti nell’area d’intervento ed è pari a quella contrattualmente richiesta all’Ente Fornitore, '
             'più precisamente: 4 kW.',
  'num': 16,
  'title': 'POTENZA IMPEGNATA.'},
 {'default': 'Per gli impianti di 1ª categoria la tensione misurata tra il quadro principale immediatamente a valle '
             "del punto di consegna dell’energia elettrica ed un qualsiasi punto dell'impianto utilizzatore (ivi "
             'compresi apparecchi d’illuminazione, prese a spina, ecc.), quando sono inseriti e funzionanti al '
             'rispettivo carico nominale non deve superare il 4% (a fondo linea).',
  'num': 17,
  'title': 'CADUTA DI TENSIONE.'},
 {'default': 'Ai fini della determinazione delle correnti d’impiego sono state fatte le seguenti considerazioni: le '
             'linee asservite alle utenze (WallBox) sono state dimensionate per il massimo carico previsto: per il '
             'dimensionamento della linea tra punto di consegna e il quadro di distribuzione si ipotizzato un fattore '
             'di contemporaneità ed utilizzazione dei carichi pari al 100% della somma dei carichi (c.a. 7,4 kW).\n'
             '\n'
             'La portata delle condutture è ricavata dalle tabelle CEI-UNEL vigenti ed applicando i coefficienti di '
             'riduzione relativi alle condizioni di posa ed alle temperature ambiente. La portata delle singole '
             'condutture (o lz) è valutata secondo le tabelle CEI UNEL 35024 e CEI UNEL 35026 con fattore di '
             'correzione in funzione del tipo posa e presenza di più circuiti elettrici.',
  'num': 18,
  'title': 'CORRENTI DI IMPIEGO E PORTATE DEI CAVI.'},
 {'default': 'Le sezioni dovranno essere tali da soddisfare le più restrittive prescrizioni in proposito dettate dalle '
             'norme CEI e delle disposizioni di legge vigenti in materia antinfortunistica.\n'
             '\n'
             'La sezione dei cavi sarà determinata anche in funzione dei seguenti parametri:\n'
             '\n'
             'carico installato;\n'
             '\n'
             "temperatura ambiente di 30°C per installazione all'interno, 40°C per posa nei percorsi\n"
             '\n'
             "all'esterno su canaletta;\n"
             '\n'
             'coefficiente di riduzione relativo alle condizioni di posa nella situazione più restrittiva nello '
             'sviluppo della linea;\n'
             '\n'
             'caduta di tensione massima ammissibile come meglio precisato nel paragrafo 12.',
  'num': 19,
  'title': 'SEZIONE MINIMA DEI CONDUTTORI DI FASE.'},
 {'default': 'Per i conduttori di neutro la sezione dovrà essere la stessa del conduttore di fase nei circuiti:\n'
             '\n'
             'monofase a due fili;\n'
             '\n'
             'polifase, quando la sezione del conduttore di fase sia inferiore o uguale a 16 mm2 se in rame e 25 mm2 '
             'se in alluminio.\n'
             '\n'
             'per i circuiti nei quali la dimensione del conduttore di fase è maggiore di quelle sopra citate, è '
             'ammesso l’uso di un conduttore di neutro avente sezione inferiore a quella di fase se la corrente che '
             'percorre il neutro, durante il servizio ordinario, non sia maggiore della corrente sopportabile dal cavo '
             'e che la sezione del neutro sia almeno uguale a 16 mm2 se in rame e 25 mm2 se in alluminio.',
  'num': 20,
  'title': 'SEZIONE MINIMA DEI CONDUTTORI DI NEUTRO.'},
 {'default': 'Si dovranno rispettare le sezioni precisate dalla tabella 54F della norma CEI 64-8 art. 543.1.2\n'
             '\n'
             'La sezione del conduttore di protezione dovrà essere scelta fra le seguenti possibilità (CEI 64-8 art. '
             '543.1):\n'
             '\n'
             '•\tnon inferiore al valore determinato dalla formula seguente:\n'
             '\n'
             'Sp = \uf0d6\uf020I2 t / K\n'
             '\n'
             'dove\n'
             '\n'
             'Sp è la sezione del conduttore di protezione in mm2\n'
             '\n'
             'I è la corrente di guasto che può percorrere il conduttore di protezione\n'
             '\n'
             't è il tempo di intervento delle protezioni in secondi\n'
             '\n'
             'K è il fattore che dipende dal materiale del conduttore di protezione (PVC = 115);\n'
             '\n'
             'secondo la seguente tabella (Tab. 54 F della norma CEI 64-8):\n'
             '\n'
             'La sezione del conduttore di protezione non facente parte della conduttura di alimentazione non dovrà '
             'essere inferiore a:\n'
             '\n'
             '•\t2,5 mm2 se protetto meccanicamente;\n'
             '\n'
             '•\t4 mm2 se non protetto meccanicamente.\n'
             '\n'
             'Quando un conduttore di protezione è comune a più circuiti, dovrà essere proporzionato alla sezione del '
             'conduttore di fase avente sezione maggiore.',
  'num': 21,
  'title': 'SEZIONE MINIMA DEI CONDUTTORI DI PROTEZIONE (PE).'},
 {'default': 'Con riferimento all’art. 542.3 ed alla tabella 54A della norma CEI 64-8 le sezioni minime dei conduttori '
             'di terra si distinguono in relazione alla loro protezione meccanica e alla loro protezione contro la '
             'corrosione. Si avranno quindi:\n'
             '\n'
             'conduttori privi di protezione contro la corrosione: sezione 25 mm2 se in rame e 50 mm2 se in ferro '
             'zincato;\n'
             '\n'
             'conduttori protetti contro la corrosione e protetti meccanicamente: sezione in accordo con l’art. '
             '543.1;\n'
             '\n'
             'conduttori protetti contro la corrosione, ma non protetti meccanicamente: sezione 16 mm2 se in rame e 16 '
             'mm2 se ferro zincato.',
  'num': 22,
  'title': 'SEZIONE MINIMA DEL CONDUTTORE DI TERRA.'},
 {'default': 'In accordo con art. 514.31 della CEI 64-8/5 e della CEI 16-4 i colori da utilizzare per '
             "l'identificazione dei vari conduttori saranno unicamente i seguenti:\n"
             '\n'
             '•\tconduttori di fase: marrone, grigio e nero;\n'
             '\n'
             '•\tconduttore di neutro: blu chiaro;\n'
             '\n'
             '•\tconduttori di protezione: giallo verde;\n'
             '\n'
             '•\tritorni ed interrotte: rosso;\n'
             '\n'
             '•\tbassissima tensione: bianco, arancione, violetto.',
  'num': 23,
  'title': 'COLORI DI IDENTIFICAZIONE.'},
 {'default': 'Di seguito si riportano le caratteristiche delle apparecchiature per il comando ed il sezionamento dei '
             'circuiti elettrici.',
  'num': 24,
  'title': 'SEZIONAMENTO E COMANDO.'},
 {'default': 'Ogni circuito dovrà poter essere sezionato dall’alimentazione, in particolare il sezionamento dovrà '
             'avvenire su tutti i conduttori attivi.\n'
             '\n'
             'Dovrà in ogni modo essere possibile sezionare diversi circuiti con un solo dispositivo purché le '
             'condizioni di esercizio lo consentano.\n'
             '\n'
             'Quando un componente elettrico, oppure un involucro, contenga parti attive collegate a più di '
             'un’alimentazione, una scritta od una segnalazione dovrà essere posta in posizione tale che qualsiasi '
             'persona che acceda alle parti attive sia avvertita della necessità di sezionare dette parti dalle '
             'proprie alimentazioni nel caso non sia presente un interblocco tale da assicurare che tutti i conduttori '
             'attivi siano sezionati.',
  'num': 25,
  'title': 'SEZIONAMENTO.'},
 {'default': 'Quando la manutenzione non elettrica può comportare rischi per le persone, si dovranno provvedere '
             'dispositivi di interruzione dell’alimentazione.\n'
             '\n'
             'Dovranno essere presi adatti provvedimenti per evitare che le apparecchiature meccaniche alimentate '
             'elettricamente siano riattivate accidentalmente durante la manutenzione non elettrica, salvo che i '
             'dispositivi di interruzione non siano continuamente sotto il controllo dell’operatore.\n'
             '\n'
             'Dovranno quindi utilizzare dispositivi di sezionamento in grado di interrompere la corrente di pieno '
             'carico.',
  'num': 26,
  'title': 'INTERRUZIONE PER MANUTENZIONE NON ELETTRICA.'},
 {'default': 'Gli apparecchi di comando funzionale non dovranno necessariamente interrompere tutti i conduttori attivi '
             'di un circuito. In ogni caso un dispositivo di comando unipolare non dovrà essere inserito sul '
             'conduttore di neutro.\n'
             '\n'
             'Le prese a spina possono essere utilizzate come comando funzionale se la loro portata non è superiore a '
             '16 A.\n'
             '\n'
             'Il comando funzionale potrà essere realizzato mediante:\n'
             '\n'
             'interruttori di manovra;\n'
             '\n'
             'interruttori automatici;\n'
             '\n'
             'contattori;\n'
             '\n'
             'relè ausiliari;\n'
             '\n'
             'prese a spina fino a 16 A compresi.',
  'num': 27,
  'title': 'COMANDO FUNZIONALE.'},
 {'default': 'Le parti attive risultano ricoperte tramite un isolamento che può essere rimosso solo mediante '
             'distruzione, Art. 412.1 CEI 64-8/4.\n'
             '\n'
             'Ne consegue che i componenti in tensione e le parti attive dovranno essere segregati, mediante posa '
             'entro involucri o dietro barriere, in modo da assicurare un grado di protezione IPXXB (CEI 64-8 art. '
             '412.2.1).\n'
             '\n'
             'Per le superfici superiori orizzontali degli involucri e delle barriere a portata di mano si dovrà '
             'garantire un grado di protezione IPXXD (CEI 64-8 art. 412.2.2).\n'
             '\n'
             'Nei luoghi soggetti a normativa specifica o con ambienti ed applicazioni particolari, il grado di '
             'protezione dovrà essere adeguato ai singoli casi, considerati in dettaglio nei capitoli specifici.\n'
             '\n'
             'Le barriere e/o gli involucri di protezione dovranno essere fissati saldamente in modo da garantire '
             'stabilità e durata nel tempo e dovranno poter essere rimossi esclusivamente:\n'
             '\n'
             "mediante l'uso di chiave o attrezzo;\n"
             '\n'
             "se l'alimentazione, dopo l'interruzione a seguito della rimozione degli involucri di protezione, sia "
             'ripristinabile solo con la richiusura degli stessi;\n'
             '\n'
             "se esiste una barriera intermedia, con grado di protezione minimo IPXXB, rimovibile solo con l'uso di "
             'chiave od attrezzo.\n'
             '\n'
             'Sono possibili altri sistemi di protezione dai contatti diretti (ostacoli, distanziamento ecc.) che '
             'dovranno in ogni modo essere analizzati ed applicati solo in casi particolari e specifici (CEI 64-8 art. '
             '412.2.4).',
  'num': 28,
  'title': 'PROTEZIONE CONTRO I CONTATTI DIRETTI.'},
 {'default': 'Per la protezione dai contatti indiretti dovrà essere garantito il coordinamento dell’impianto di terra '
             "con i dispositivi di protezione (CEI 64-8/4 art. 413.1.4.2) in modo da assicurare l'interruzione "
             "automatica dell'alimentazione nei tempi richiesti.\n"
             '\n'
             'Il coordinamento sarà soddisfatto dalla relazione:\n'
             '\n'
             'Ra * Ia < 50\n'
             '\n'
             'dove:\n'
             '\n'
             'Ra = somma della resistenza del dispersore e dei conduttori di protezione\n'
             '\n'
             'Ia = corrente che provoca il funzionamento automatico del dispositivo di protezione (Idn se il '
             'dispositivo è differenziale).\n'
             '\n'
             'Nel caso di dispositivo con caratteristica di funzionamento a tempo inverso (interruttore '
             'magnetotermico) si dovrà garantire che tra una parte attiva e una massa (o un conduttore di protezione) '
             'non possa permanere una tensione di contatto superiore a 50 V (in corrente alternata) per un tempo '
             'superiore a 5s (CEI 64-8 art. 413.1.4.2).\n'
             '\n'
             "Nell'utilizzo di dispositivi differenziali, che dovranno rispettare le prescrizioni della Norma CEI "
             "23-18, l'intervento dovrà essere istantaneo.\n"
             '\n'
             'Se si usano dispositivi differenziali di tipo selettivo (S) o ritardati, posti in serie a dispositivi '
             'differenziali di tipo generale, il tempo di intervento non dovrà essere superiore a 1s.',
  'num': 29,
  'title': 'PROTEZIONE CONTRO I CONTATTI INDIRETTI – SISTEMA TT.'},
 {'default': 'Le parti attive poste entro involucri o barriere devono assicurare almeno il grado di protezione IP XXB '
             '(IP 20) Art. 412.2.1 CEI 64-8/4.\n'
             '\n'
             'Le superfici orizzontali degli involucri o barriere poste a portata di mano (sotto i m. 2,5 dal '
             'calpestio) devono assicurare almeno il grado di protezione IP XXD (IP 40) Art. 412.2.2 CEI 64-8/4.\n'
             '\n'
             'Per un più esplicito riferimento si rimanda a quanto segue:',
  'num': 30,
  'title': 'PROTEZIONE DA PARTI IN TENSIONE POSTE ALL’INTERNO DELL’INVOLUCRO.'},
 {'default': "La protezione da contatti indiretti può essere realizzata anche con l'utilizzo di componenti in classe "
             'II. Sono da considerare tali le condutture elettriche costituite da:\n'
             '\n'
             'cavi con guaina non metallica aventi tensione nominale maggiore di un gradino rispetto a quella '
             'necessaria per il sistema elettrico servito e che non comprendano un rivestimento metallico;\n'
             '\n'
             'cavi unipolari senza guaina installati in tubo protettivo o canale isolante e rispondente alle '
             'rispettive Norme;\n'
             '\n'
             'cavi con guaina metallica aventi isolamento idoneo per la tensione nominale del sistema elettrico '
             "servito, tra la parte attiva e la guaina metallica e tra questa e l'esterno.",
  'num': 31,
  'title': 'COMPONENTI ELETTRICI IN CLASSE II O CON ISOLAMENTO EQUIVALENTE.'},
 {'default': 'Di seguito si riportano quanto previsto dalle Norme e Leggi relativamente alle protezioni delle '
             'condutture.',
  'num': 32,
  'title': 'PROTEZIONE DELLE CONDUTTURE CONTRO LE SOVRACORRENTI.'},
 {'default': 'Tutte le condutture saranno protette dai sovraccarichi, con la sola esclusione dei circuiti la cui '
             'interruzione potrebbe dar luogo a pericolo per le persone. Le protezioni dai sovraccarichi saranno '
             'realizzate con interruttori automatici, rispondenti alle norme CEI 17-5 e CEI 23-3.\n'
             '\n'
             'Per proteggere le linee contro i sovraccarichi saranno soddisfatte le seguenti condizioni:\n'
             '\n'
             'Ib \uf0a3\uf020In \uf0a3\uf020Iz\n'
             '\n'
             'e\n'
             '\n'
             'If \uf0a3\uf0201,45 Iz\n'
             '\n'
             'dove:\n'
             '\n'
             'In è la corrente nominale dell’interruttore o la sua taratura termica;\n'
             '\n'
             'If è la corrente convenzionale di funzionamento dell’interruttore;\n'
             '\n'
             'Ib è la corrente d’impiego;\n'
             '\n'
             'Iz è la portata della linea.\n'
             '\n'
             'Per quanto riguarda il soddisfacimento della seconda condizione, si terrà presente che:\n'
             '\n'
             'gli interruttori per uso domestico o similare (norma CEI 23-3 e 23-18) hanno una corrente di '
             'funzionamento If \uf0a3 1,45 x In;\n'
             '\n'
             'gli interruttori conformi alla norma CEI 17-5 hanno una corrente di funzionamento If = 1,35 x In, per '
             'correnti nominali fino a 63 A e If = 1,25 x In, per valori della corrente nominale superiori a 63 A.\n'
             '\n'
             'Quando la protezione dalle sovracorrenti sarà effettuata con fusibili si terranno presenti le seguenti '
             'relazioni:\n'
             '\n'
             'a)\t4 A \uf0a3 In \uf0a3 10 A\t\tIf =1,9\te quindi\tIb \uf0a3 In \uf0a3 0,763 IZ\n'
             '\n'
             'b)\t10 A \uf0a3 In \uf0a3 25 A    \tIf =1,75\te quindi\tIb \uf0a3 In \uf0a3 0,828 Iz\n'
             '\n'
             'c)\t25 A \uf0a3 In\t\tIf =1,6\te quindi\tIb \uf0a3 In \uf0a3 0,6 Iz',
  'num': 33,
  'title': 'PROTEZIONE CONTRO I SOVRACCARICHI.'},
 {'default': 'Per la protezione da corto circuito (CEI 64-8 art. 434.3), affinché la temperatura dei conduttori non '
             'superi il valore massimo ammissibile, si dovrà tener conto della relazione seguente:\n'
             '\n'
             '(I²* t) \uf0a3 K²* S²\n'
             '\n'
             'dove:\n'
             '\n'
             'I = corrente di corto circuito in Ampere;\n'
             '\n'
             't = durata del corto circuito in secondi;\n'
             '\n'
             "K = fattore relativo alla natura dell'isolante\n"
             '\n'
             '115 per cavo in rame con guaina esterna in PVC;\n'
             '\n'
             '135 per cavi in rame isolati con gomma ordinaria o gomma butilica;\n'
             '\n'
             '143 per cavi in rame isolati con gomma etilenpropilenica e propilene reticolato.\n'
             '\n'
             'S = sezione del conduttore in mm.',
  'num': 34,
  'title': 'PROTEZIONE CONTRO I CORTO CIRCUITI.'},
 {'default': 'Gli impianti saranno realizzati in modo tale da assicurare la massima selettività possibile onde evitare '
             'che, in caso di guasto su un circuito a valle, intervengano anche le protezioni generali installate a '
             'monte.',
  'num': 35,
  'title': 'SELETTIVITÀ.'},
 {'default': 'Saranno forniti al manutentore i documenti di disposizione topografica dell’impianto elettrico, '
             'unitamente a rapporti di verifica, disegni, schemi e relative modifiche, così come istruzioni per '
             'l’esercizio e la manutenzione.',
  'num': 36,
  'title': 'SCHEMI E DOCUMENTAZIONE.'},
 {'default': 'Relativamente a quanto descritto in questo capitolo, si precisa che per maggiori dettagli si rimanda a '
             'tutta la documentazione allegata alla presente.',
  'num': 37,
  'title': 'DESCRIZIONE DEGLI IMPIANTI'},
 {'default': 'Tutti i quadri e i dispositivi di protezione scelti (fusibili, interruttori) saranno di primaria marca: '
             'e dovranno essere del tipo in materiale termoplastico salvo diversa indicazione.\n'
             '\n'
             'I quadri saranno completi di telai e pannellature idonee per il montaggio di apparecchi modulari e '
             'scatolati, e dovranno essere corredati di appositi cartellini fissati in modo imperdibile che '
             'indicheranno chiaramente le funzioni svolte dalle varie apparecchiature installate.\n'
             '\n'
             'Sui quadri troveranno posto le protezioni magnetotermiche differenziali necessarie per attuare la '
             'protezione, il sezionamento e la suddivisione dei circuiti previsti con riferimento alla vigente '
             'normativa e in considerazione delle esigenze di sicurezza, continuità del servizio e praticità di '
             'manutenzione.\n'
             '\n'
             "L'ingresso delle condutture (cavi provenienti dal contatore) sarà realizzato nella parte inferiore dello "
             'stesso; analogamente l’uscita delle condutture (cavi verso le utenze) sarà realizzata nella parte '
             'inferiore dello stesso.\n'
             '\n'
             'Si raccomanda, nell’ingresso delle condutture al quadro, il mantenimento del grado di protezione '
             'iniziale dello stesso, mediante l’utilizzo di appositi pressa-cavi o guarnizioni. Le dimensioni dei '
             'quadri e le caratteristiche tecniche delle apparecchiature in essi installate sono specificate negli '
             'schemi elettrici allegati.',
  'num': 38,
  'title': 'QUADRI ELETTRICI.'},
 {'default': 'Quanto segue è valido solo nel caso in cui il POD, e quindi il contatore fiscale ed il DG.\n'
             '\n'
             'Gli ingressi delle condutture dall’Ente Distributore e delle condutture del fornitore dei servizi sono '
             'convogliati alla conchiglia dei quadri a mezzo di tubazione interrata all’interno della proprietà.\n'
             '\n'
             'Le condutture della distribuzione saranno unicamente cavi unipolari tipo FG16R16/FG16M16, che saranno '
             'usati per sezioni di conduttore superiore a 25 mm2, mentre per sezioni inferiore sarà ammesso uso di '
             'cavi multipolare del tipo FG16(O)R16 o FRG17, che saranno utilizzati anche per le alimentazioni di '
             'utenze o i collegamenti di segnale nei locali tecnologici.',
  'num': 39,
  'title': 'DISTRIBUZIONE PRINCIPALE.'},
 {'default': 'La stazione di ricarica è classificabile come “Ambienti ed applicazioni particolari, Alimentazione di '
             'veicoli elettrici” della Norma CEI 64-8:2012:06, variante V1:2013:07, Sezione 722.\n'
             '\n'
             'Il modo di ricarica sarà tipo 3 e 4, mentre il modo di connessione sarà di tipo B e C (Norma CEI EN '
             '61851-1:2012-05).\n'
             '\n'
             'Attualmente la norma, che riporta le prescrizioni necessarie per la ricarica dei veicoli elettrici, è la '
             'Norma CEI EN 61851-1:2012-05 “Sistema di ricarica conduttiva dei veicoli elettrici – Parte 1: '
             'Prescrizioni generali\n'
             '\n'
             '“…con riferimento ai modi di carica in corrente alternata adottati in Italia, al fine di garantire la '
             'necessaria sicurezza durante la carica conduttiva dei veicoli elettrici, quando questa viene eseguita in '
             'ambienti aperti a terzi deve essere adottato il Modo di carica 3”.\n'
             '\n'
             'Sulla base delle classificazioni realizzate da Cives ed Eurelectric, il Piano Nazionale individua le '
             'seguenti classi di infrastrutture di ricarica sulla base della capacità di erogazione dell’energia:\n'
             '\n'
             'Normal power (Slow charging) - fino a 3,7 kW\n'
             '\n'
             'Medium power (Quick charging) - da 3,7 fino a 22 kW\n'
             '\n'
             'High power (Fast charging) - superiore a 22 kW\n'
             '\n'
             'Lo specifico progetto prevede la posa di Medium power.\n'
             '\n'
             'Architettura EVC.\n'
             '\n'
             'punto di ricarica prevedrà:\n'
             '\n'
             '•\tuna WallBox del tipo a ricarica fino a 7,4kW dotata, nel caso in oggetto, di una presa di ricarica '
             'opportunamente modulata con apposito softwere;',
  'num': 40,
  'title': 'COLONNINA.'},
 {'default': 'Per la valutazione delle verifiche da eseguire sui quadri elettrici bisogna distinguerli per categoria '
             'in base alle loro caratteristiche.',
  'num': 41,
  'title': 'VERIFICHE.'},
 {'default': 'Le apparecchiature installate nei quadri di comando e negli armadi dovranno essere del tipo modulare e '
             'componibile con fissaggio a scatto su profilato normalizzato EN 50022, ad eccezione di eventuali '
             'interruttori automatici superiori a 125 A che si fisseranno a mezzo di bulloni sulla piastra di '
             "cablaggio, mentre per il fissaggio di relè contattori all'interno del quadro si adotterà il sistema di "
             'fissaggio e cablaggio su piastra.\n'
             '\n'
             'Gli interruttori di tipo magnetotermico, magnetotermico differenziale e differenziale puro dovranno '
             'avere potere di interruzione adeguato alla corrente di corto circuito calcolato nel punto di '
             'installazione. La corrente di soglia di intervento differenziale potrà essere da 0,5 A - 0,3 A - 0,03 A, '
             'a seconda della selettività che si vuole conseguire.',
  'num': 42,
  'title': 'APPARECCHIATURE MODULARI.'},
 {'default': 'Ogni quadro sarà dotato di un interruttore generale provvisto di comando manuale che consenta di '
             'interrompere simultaneamente la continuità metallica di tutti i conduttori.\n'
             '\n'
             'Esso dovrà portare una chiara indicazione della posizione di aperto o chiuso in corrispondenza '
             "dell'organo di manovra.",
  'num': 43,
  'title': 'INTERRUTTORE GENERALE.'},
 {'default': 'Si è fatto uso di interruttori automatici magnetotermici aventi meccanica di tipo autoportante '
             'svincolata dall’involucro isolante, di comando a leva nera piombabile in posizione ON-OFF.\n'
             '\n'
             'I morsetti di collegamento saranno predisposti per il collegamento di cavi e barrette di collegamento e '
             'l’alimentazione sarà possibile sia dai morsetti superiori che inferiori.\n'
             '\n'
             'Si riportano di seguito le caratteristiche generali:\n'
             '\n'
             'Tensione nominale di funzionamento in corrente alternata: 230/400 V;\n'
             '\n'
             'Frequenza di esercizio: 50-60 Hz;\n'
             '\n'
             'Nr. poli: (1+N; 1; 2; 3; 4);\n'
             '\n'
             'Potere di inter. (CEI 23.3): 15 kA o inferiore se è previsto protezione di back up a monte;\n'
             '\n'
             'Corrente nominale ininterrotta:\n'
             '\n'
             '(caratteristica B): (6…63) A;\n'
             '\n'
             '(caratteristiche C, D, K): (0.5…63) A;\n'
             '\n'
             'Caratteristica di intervento: B-C-D-K;\n'
             '\n'
             'Tenuta alla tensione a frequenza industriale: 3 kV;\n'
             '\n'
             'Numero di manovre meccaniche: 20.000;\n'
             '\n'
             'Numero di manovre elettriche a Ue e In: 10.000;\n'
             '\n'
             'Tensione di isolamento 500 V;\n'
             '\n'
             'Grado di inquinamento 2;\n'
             '\n'
             'Gruppo materiale II, idoneo al sezionamento.\n'
             '\n'
             'Il potere di interruzione dovrà essere adeguato alla corrente di corto circuito simmetrica trifase '
             'presunta nel punto di installazione o in alternativa dovrà essere previsto un interruttore generale con '
             'opportuno coordinamento avente potere di interruzione comunque non inferiore a 15 kA.',
  'num': 44,
  'title': 'INTERRUTTORI MAGNETOTERMICI MODULARI.'},
 {'default': 'I blocchi differenziali avranno meccanica di tipo autoportante svincolata dall’involucro isolante, di '
             'comando a leva piombabile in posizione ON-OFF.\n'
             '\n'
             'Il dispositivo differenziale sarà idoneo al funzionamento in presenza di correnti alternate sinusoidali '
             'e  immune agli scatti intempestivi dovuti alle sovratensioni pari a 250A di picco con onda 8/20 µs.\n'
             '\n'
             'Tensione nominale di funzionamento in corrente alternata: 230/400 V;\n'
             '\n'
             'Frequenza di esercizio: 50-60 Hz;\n'
             '\n'
             'Potere di interruzione in corto circuito pari a quello dell’interruttore automatico a cui è accoppiato '
             '(se non già un differenziale puro);\n'
             '\n'
             'Taglia: 25, 40, 63, 100 A;\n'
             '\n'
             'Nr. poli: (2-3-4);\n'
             '\n'
             'Sensibilità nominale differenziale: 0.03 – 0,1 – 0,3 – 0,5 – 1 – 2;\n'
             '\n'
             'Numero di manovre meccaniche: 20.000;\n'
             '\n'
             'Numero di manovre elettriche a Ue e In: 10.000; 20.000 (taglia 100 A).',
  'num': 45,
  'title': 'INTERRUTTORI DIFFERENZIALI MODULARI.'},
 {'default': 'Se necessari si farà uso di contattori accessoriabili, rispondenti alle normative EN60947-1 e 947-4-1 e '
             'saranno idonei per montaggio su barra DIN o piastra di fondo, in versioni a 3 o 4 poli con morsetti a '
             'vite e grado di protezione IP20 ed avranno circuito magnetico (bobina) in corrente alternata e contatto '
             'ausiliario (n° 1 normalmente aperto NA o n° 1 normalmente chiuso NC) integrato nelle versioni '
             'tripolari.\n'
             '\n'
             'In generale avranno le seguenti caratteristiche tecniche:\n'
             '\n'
             'Tensione nominale d’isolamento Ui: 1.000 V;\n'
             '\n'
             'Tensione nominale di impulso Uimp: 8 kV;\n'
             '\n'
             'Tensione nominale di impiego Ue: 690 Vca;\n'
             '\n'
             'Temperatura ambiente:\n'
             '\n'
             'immagazzinaggio: da –60 °C a +80 °C;\n'
             '\n'
             'in funzionamento da –40 °C a +55 °C;\n'
             '\n'
             'in funzionamento con relè termico da –25 °C a +55 °C\n'
             '\n'
             'Durata meccanica: 10 milioni di manovre;\n'
             '\n'
             'Durata elettrica (AC3 - 400 V - 16 A): 2.000.000 manovre.',
  'num': 46,
  'title': 'CONTATTORI DI POTENZA E AUSILIARI.'},
 {'default': 'Qualora previsti, la grandezza per i contattori di potenza dovrà essere scelta tenendo conto di una '
             'corrente minima di 9 A in classe di funzionamento AC3, inoltre dovranno avere almeno due contatti '
             'ausiliari (n° 1 NA e n° 1 NC) in più di quelli previsti dallo schema di quadro.\n'
             '\n'
             'Si riportano di seguito le caratteristiche minime generali a cui devono rispondere:\n'
             '\n'
             'Contatti ausiliari frontali e/o laterali;\n'
             '\n'
             'Interblocchi meccanici ed elettro-meccanici;\n'
             '\n'
             'Temporizzatori pneumatici ed elettronici;\n'
             '\n'
             'Limitatori di sovratensioni;\n'
             '\n'
             'Barrette di collegamento;\n'
             '\n'
             'Bobine di ricambio;\n'
             '\n'
             'Relè termici.',
  'num': 47,
  'title': 'ACCESSORI.'},
 {'default': 'Quanto segue è valido solo nel caso in cui il POD, e quindi il contatore fiscale ed il DG non sia '
             'all’interno di EVC, come nel caso oggetto della presente relazione tecnica.\n'
             '\n'
             'Tutti i cavi impiegati nella realizzazione degli impianti descritti nel presente progetto risponderanno '
             "all'unificazione UNEL ed alle Norme costruttive stabilite dal Comitato Elettrotecnico Italiano.\n"
             '\n'
             'In particolare, tutti i cavi citati nel presente documento si intendono del tipo non propagante '
             'l’incendio ed a bassissima emissione di fumi e gas tossici, secondo le norme vigenti quali CEI 20-22 '
             'III, CEI 20-35, CEI 20-37 e CEI 20-38 o s.m.i.\n'
             '\n'
             'Per la distribuzione dell’energia dovranno essere utilizzati i cavi unipolari isolati in XLPE, qualità '
             'R2 non propaganti l’incendio, con corde flessibili in rame, rispondenti alla norma CEI 20-22, ed avranno '
             'una tensione di isolamento minimo, superiore di un gradino alla tensione di impiego (Uo/U = 0,6/1 kV).\n'
             '\n'
             'I cavi saranno contrassegnati in modo da individuare prontamente il servizio a cui appartengono; il '
             'transito di cavi attraverso la struttura di canali portacavi, cassette di derivazione etc., sarà '
             "effettuato con l'ausilio di pressacavi del tipo con bullone a stringere.\n"
             '\n'
             'I conduttori previsti saranno dimensionati secondo i dati della tabella CEI-UNEL 35024/1 e 35024/2 '
             'tenendo conto di una temperatura iniziale di 30°C, di una temperatura massima di esercizio e di una '
             "temperatura massima di corto circuito adeguati al tipo dell'isolante (CEI 64-8 tabella 52 D); per la "
             'posa interrata si farà riferimento alle tabelle CEI-UNEL 35026.\n'
             '\n'
             'Nel caso siano posati nella stessa conduttura conduttori di sistemi a tensione diversa (cavi per '
             'energia, impianto rivelazione incendio, impianti trasmissione dati, ecc.), tutti i conduttori dovranno '
             'essere isolati per la tensione più elevata (CEI 64-8 art. 521.6).',
  'num': 48,
  'title': 'CAVI.'},
 {'default': 'I conduttori impiegati nell’esecuzione degli impianti dovranno essere contraddistinti dalla colorazione '
             'prevista dalle vigenti tabelle di unificazione CEI - UNEL 00722 e 00712.\n'
             '\n'
             'Per quanto riguarda i conduttori di fase dovranno essere contraddistinti in modo univoco per tutto '
             "l'impianto dai colori: nero, grigio e marrone. Nella scelta del colore dei conduttori, il bicolore "
             'giallo-verde sarà tassativamente riservato ai conduttori di protezione ed equipotenziali ed il colore '
             'blu chiaro sarà destinato esclusivamente al conduttore di neutro (CEI 64-8 art. 514.3.1).',
  'num': 49,
  'title': 'COLORI DEI CAVI.'},
 {'default': 'In accordo con la Tabella 52A della Norma CEI 64-8, si potranno utilizzare, ad esempio, i seguenti tipi '
             'di cavo:\n'
             '\n'
             'posa all’interno e all’esterno non interrata: H07V-K, FS17, FG17 – 450/750 V;\n'
             '\n'
             'posa all’interno e all’esterno anche interrata: FG16OR16-0,6/1 kV, FG16R16-0,6/1 kV, N1VV-K.\n'
             '\n'
             'Per gli ambienti trattati nella Sezione 751 della Norma CEI 64-8:',
  'num': 50,
  'title': 'CAVI PER LA DISTRIBUZIONE DELL’ENERGIA.'},
 {'default': 'Quanto segue è valido solo nel caso in cui il POD, e quindi il contatore fiscale ed il DG non sia '
             'all’interno di EVC, come nel caso oggetto della presente relazione tecnica.\n'
             '\n'
             'Le condutture dovranno essere realizzate in modo da ridurre al minimo la probabilità di innesco e '
             'propagazione dell’incendio nelle condizioni di posa. Per soddisfare questi requisiti le condutture '
             'dovranno rispondere alle prescrizioni della Sezione 751 della Norma CEI 64-8/7.\n'
             '\n'
             "Per conduttura si dovrà intendere l'insieme costituito da uno o più conduttori elettrici e dagli "
             'elementi che assicurano il loro isolamento, il loro supporto, il loro fissaggio e la loro eventuale '
             'protezione meccanica (CEI 64-8/2 art. 26.1).\n'
             '\n'
             'I conduttori dovranno essere sempre protetti meccanicamente. Dette protezioni saranno realizzate '
             'mediante tubazioni anche interrate, canalette portacavi, passerelle, condotti o cunicoli, eventualmente '
             'ricavati nella struttura edile ecc.\n'
             '\n'
             "I tubi protettivi, le cassette e le scatole per l'impianto di energia, per trasmissione dati, di "
             'allarme, di controllo e di segnalazione, dovranno essere dedicate e distinte fra loro (CEI 64-8/5 art. '
             '528.1.1).\n'
             '\n'
             'Le condutture elettriche dovranno essere opportunamente distanziate da tubazioni che producano calore, '
             'fumi o vapori. Se ciò non fosse possibile si dovranno utilizzare opportuni accorgimenti onde evitare '
             'eventuali effetti dannosi.',
  'num': 51,
  'title': 'CONDUTTURE.'},
 {'default': 'In considerazione delle diverse tipologie impiantistiche si potranno utilizzare, oltre a quelli già '
             'esistenti, i tubi e le guaine di seguito descritte:\n'
             '\n'
             'tubo rigido autoestinguente in PVC serie pesante conforme alla Norma CEI 23-8 e varianti ed alle '
             'relative tabelle UNEL 37118-37119-37120 e s.m.i.;\n'
             '\n'
             'tubo flessibile autoestinguente in PVC serie pesante conforme alla Norma CEI 23-14 e varianti;\n'
             '\n'
             'guaine in PVC flessibile autoestinguente, serie pesante, complete di accessori di giunzione e '
             'derivazione, conformi alle relative tabelle UNEL 37118-37119-37120 e s.m.i.\n'
             '\n'
             'Il diametro dei tubi non dovrà essere inferiore a 16 mm. Tutte le curve eseguite senza l’impiego di '
             'pezzi speciali dovranno essere di raggio proporzionato al diametro del tubo e tale da non diminuirne in '
             'corrispondenza delle stesse la sezione libera di passaggio.\n'
             '\n'
             'I tubi di nuova installazione dovranno essere dimensionati in modo che il loro diametro sia pari ad '
             'almeno 1,3 volte il diametro del cerchio circoscritto al fascio dei conduttori in essi contenuti.\n'
             '\n'
             "Tale accorgimento renderà possibile un'eventuale aggiunta di conduttori senza arrecare deterioramento "
             "all'isolamento degli esistenti e permetterà di non apportare pregiudizio alla sfilabilità dei cavi.\n"
             '\n'
             'Tutte le tubazioni, qualunque sia il tipo di posa, dovranno avere andamento prevalentemente rettilineo, '
             'si potranno seguire percorsi non rigorosamente rettilinei solamente in corrispondenza di eventuali '
             'ostacoli (canali, tubazioni di altri impianti).',
  'num': 52,
  'title': 'TUBI E GUAINE.'},
 {'default': 'Le condutture elettriche dovranno essere opportunamente distanziate da tubazioni che producano calore, '
             'fumi o vapori. Se ciò non fosse possibile si dovranno utilizzare opportuni accorgimenti onde evitare '
             'eventuali effetti dannosi.\n'
             '\n'
             'I tubi protettivi installati sotto traccia dovranno avere un percorso orizzontale, verticale o parallelo '
             'allo spigolo della parete, ad esclusione dei percorsi nei soffitti e nei pavimenti ove il percorso potrà '
             'essere omnidirezionale.\n'
             '\n'
             'I tipi di posa delle condutture in funzione dei tipi di cavi utilizzati dovranno essere in accordo con '
             'la Tabella 52A della norma CEI 64-8 sotto riportata.\n'
             '\n'
             'Nei cavi con guaina sono compresi i cavi provvisti di armatura e quelli con isolamento minerale.\n'
             '\n'
             'Legenda:\n'
             '\n'
             '+ permesso\n'
             '\n'
             '- non permesso\n'
             '\n'
             '0 non applicabile o non usato in generale nella pratica.\n'
             '\n'
             'I tipi di posa delle condutture in funzione delle varie condizioni di utilizzo dovranno essere in '
             'accordo con la Tabella 52B della norma CEI 64-8 di seguito riportata.',
  'num': 53,
  'title': 'TIPI DI POSA.'},
 {'default': 'La rete di distribuzione per questa tipologia di impianto, qualora previsto, si svilupperà all’interno '
             'di canalizzazioni predisposte allo scopo, separata dagli impianti di energia. La tipologia di cavo '
             'utilizzata sarà conforme alle specifiche del fornitore/costruttore dell’EVC o, più in generale, del '
             'fornitore del sistema di trasmissione dati adottato. I cavi impiegati saranno adeguati alla modalità di '
             'posa prevista.',
  'num': 54,
  'title': 'IMPIANTO TRASMISSIONE DATI.'},
 {'default': 'Ad impianto ultimato si dovrà provvedere alle seguenti verifiche di collaudo.\n'
             '\n'
             'Rispondenza alle disposizioni di legge.\n'
             '\n'
             'Rispondenza alle prescrizioni particolari concordate in progetto e in sede di offerta.\n'
             '\n'
             'Rispondenza alle norme CEI relative al tipo di impianto, come meglio descritto sulla Norma CEI 64-8 '
             'Cap.61 "Verifiche iniziali" e s.m.i.\n'
             '\n'
             'Entrando più in dettaglio, l’esame dell’impianto elettrico consiste in un controllo di rispondenza '
             'dell’opera realizzata ai dati di progetto e a regola d’arte e dovrà essere effettuata prendendo tutte le '
             'precauzioni possibili per la sicurezza del personale e per evitare danni ai beni ed ai componenti '
             'elettrici.\n'
             '\n'
             'I tipi di verifica si distinguono in iniziale, periodica o straordinaria:\n'
             '\n'
             'iniziale: effettuata prima della messa in servizio dell’impianto elettrico;\n'
             '\n'
             'periodica: effettuata ad intervalli di tempo solitamente stabiliti;\n'
             '\n'
             'straordinaria: effettuata dopo aver modificato o ampliato l’impianto elettrico.\n'
             '\n'
             'Dovranno essere registrate le date ed i risultati delle prove e delle misure di ciascuna verifica, la '
             'quale dovrà essere effettuata da un tecnico qualificato.',
  'num': 55,
  'title': 'VERIFICHE'},
 {'default': 'Durante la realizzazione e prima della messa in servizio, l’impianto elettrico dovrà essere esaminato a '
             'vista e provato per verificare che le prescrizioni richiamate dalla Guida CEI 64-52 siano state '
             'rispettate.\n'
             '\n'
             'A tale scopo dovranno essere eseguite tutte le verifiche prescritte dalle norme impiantistiche ed in '
             'particolare quelle del Capitolo 61 della Norma CEI 64-8.\n'
             '\n'
             'Le verifiche dovranno essere effettuate prima della messa in servizio iniziale e, dopo modifiche o '
             'riparazioni, prima della nuova messa in servizio.',
  'num': 56,
  'title': 'VERIFICHE INIZIALI.'},
 {'default': 'In questo caso la verifica di un impianto elettrico è:\n'
             '\n'
             'di tipo ordinario accertando tutti quei difetti evidenti allo sguardo ad esempio: involucri rotti, '
             'connessioni interrotte, mancanza di ancoraggi, ecc.\n'
             '\n'
             'di tipo approfondito ispezionando, per mezzo di attrezzi ed utensili, i componenti elettrici per '
             'identificarne i difetti di installazione ad esempio connessioni lente, ecc.',
  'num': 57,
  'title': 'ESAME A VISTA.'},
 {'default': 'Si intende l’effettuazione di misure o di altre operazioni sull’impianto elettrico per mezzo di '
             'strumenti appropriati al fine di accertare che i valori risultanti siano in accordo con le Norme CEI.\n'
             '\n'
             'Le prove da effettuare, ovviamente in funzione di quanto effettivamente installato, sono:\n'
             '\n'
             'prove della protezione contro i contatti diretti:\n'
             '\n'
             'prova del grado di protezione;\n'
             '\n'
             'prove della protezione contro i contatti indiretti:\n'
             '\n'
             'prova della continuità dei conduttori di terra, di protezione ed equipotenziali (se previsti);\n'
             '\n'
             'prova del funzionamento dei dispositivi differenziali;\n'
             '\n'
             'misura della resistenza di terra;\n'
             '\n'
             'misura dell’impedenza dell’anello di guasto.\n'
             '\n'
             'prove per la verifica della corretta scelta dei componenti elettrici e loro corretta installazione:\n'
             '\n'
             'prova di tensione applicata;\n'
             '\n'
             'prova di funzionamento.\n'
             '\n'
             'prove delle condutture e connessioni:\n'
             '\n'
             'misura della resistenza di isolamento dell’impianto elettrico.\n'
             '\n'
             'Per l’effettuazione delle sopracitate prove dovranno essere utilizzati i seguenti strumenti:\n'
             '\n'
             'apparecchio per la prova di continuità dei conduttori di protezione ed equipotenziali (se previsti);\n'
             '\n'
             'misuratore della resistenza di isolamento;\n'
             '\n'
             'misuratori della resistenza di terra con metodo volt-amperometrico;\n'
             '\n'
             'apparecchio per il controllo di funzionalità degli interruttori differenziali;\n'
             '\n'
             'dito e filo di prova.\n'
             '\n'
             'Le verifiche ed i loro risultati dovranno essere riportati su di un registro corredato da timbro e firma '
             'del tecnico esecutore e dalla data di verifica.\n'
             '\n'
             'Busnago, 06/02/2024\n'
             '\n'
             'Il Progettista\n'
             '\n'
             '………………………………………………',
  'num': 58,
  'title': 'PROVE.'}]



def _p(text: str, style):
    safe = escape(text or "").replace("\n", "<br/>")
    return Paragraph(safe, style)


def _meaningful(value: Any) -> bool:
    if value is None:
        return False
    s = str(value).strip()
    if not s:
        return False
    low = s.lower()
    bad = {
        "non pertinente",
        "non applicabile",
        "n/a",
        "na",
        "—",
        "-",
        "nessuna",
        "nessuna / non applicabile",
    }
    if low in bad:
        return False
    if "xxxx" in low or "inserire" in low or "da compil" in low or "da defin" in low:
        return False
    return True


def _kv_table(rows: List[list], col_widths):
    tbl = Table(rows, colWidths=col_widths, hAlign="LEFT")
    tbl.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    return tbl


def _first_nonempty_line(text: str) -> str:
    for ln in (text or "").splitlines():
        ln = ln.strip()
        if ln:
            return ln
    return ""


class _NumberedCanvas(canvas.Canvas):
    """Canvas con numerazione 'Pagina X di Y' senza duplicare le pagine."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states: List[dict] = []

    def showPage(self):
        # Salva lo stato della pagina corrente e prepara la successiva
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._draw_page_number(page_count)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def _draw_page_number(self, page_count: int):
        self.saveState()
        self.setFont("Helvetica", 9)
        self.setFillColor(colors.grey)
        self.drawRightString(200 * mm, 10 * mm, f"Pagina {self._pageNumber} di {page_count}")
        self.setFillColor(colors.black)
        self.restoreState()



def _build_indice_items(_: Dict[str, Any]) -> List[str]:
    # Struttura capitoli "editoriale". L'indice è volutamente sintetico:
    # i sotto-paragrafi vengono resi nel testo solo se significativi.
    return [
        "CAPITOLO 1: Premessa",
        "CAPITOLO 2: Riferimenti legislativi e normativi",
        "CAPITOLO 3: Criteri di progetto degli impianti",
        "CAPITOLO 4: Soluzione progettuale adottata (dati tecnici, EVSE, quadri e linee)",
        "CAPITOLO 5: Sicurezza, verifiche e manutenzione",
        "CAPITOLO 6: Allegati e documentazione",
    ]


def _evse_table(data: Dict[str, Any], styles) -> Optional[Table]:
    evse = data.get("evse") or []
    # drop righe vuote (Marca/Modello entrambi non significativi)
    filtered = []
    for r in evse:
        marca = str(r.get("Marca", "")).strip()
        modello = str(r.get("Modello", "")).strip()
        if not _meaningful(marca) and not _meaningful(modello):
            continue
        filtered.append(r)
    if not filtered:
        return None

    th = ParagraphStyle("evth", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8, leading=9)
    tc = ParagraphStyle("evtc", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=9)

    tdata = [[
        _p("Tipo", th),
        _p("Marca/Modello", th),
        _p("Punti", th),
        _p("Potenza", th),
        _p("Alim.", th),
        _p("Connettore/Modo", th),
        _p("IP/IK", th),
        _p("RCD", th),
        _p("Note", th),
    ]]
    for r in filtered:
        marca = str(r.get('Marca', '')).strip()
        modello = str(r.get('Modello', '')).strip()
        marca_modello = (f"{marca} {modello}").strip()
        tdata.append([
            _p(str(r.get("Tipo", "")), tc),
            _p(marca_modello, tc),
            _p(str(r.get("N. punti", "")), tc),
            _p(f"{r.get('Potenza (kW)','')} kW", tc),
            _p(str(r.get("Alimentazione", "")), tc),
            _p(f"{r.get('Connettore','')} / {r.get('Modo ricarica','')}", tc),
            _p(str(r.get("IP/IK", "")), tc),
            _p(str(r.get("RCD richiesto", "")), tc),
            _p(str(r.get("Note", "")), tc),
        ])

    tbl = Table(
        tdata,
        colWidths=[14 * mm, 35 * mm, 10 * mm, 16 * mm, 16 * mm, 26 * mm, 16 * mm, 28 * mm, 23 * mm],
        repeatRows=1,
        hAlign="LEFT",
    )
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return tbl


def _photo_flowables(data: Dict[str, Any]) -> List[Flowable]:
    photos = data.get("foto") or []
    # massimo 6 per non gonfiare il PDF
    photos = [p for p in photos if isinstance(p, dict) and p.get("bytes")][:6]
    if not photos:
        return []

    flows: List[Flowable] = []
    flows.append(Spacer(1, 4))
    # tabella 2 colonne
    cells = []
    for p in photos:
        try:
            img = ImageReader(BytesIO(p["bytes"]))
            iw, ih = img.getSize()
            # fit in box
            box_w = 80 * mm
            box_h = 55 * mm
            scale = min(box_w / iw, box_h / ih)
            w = iw * scale
            h = ih * scale
            # canvas Image can't be directly inserted; use reportlab.platypus Image
            from reportlab.platypus import Image as RLImage

            im = RLImage(BytesIO(p["bytes"]), width=w, height=h)
            caption = Paragraph(escape(p.get("name", "")), getSampleStyleSheet()["BodyText"])
            cells.append([im, caption])
        except Exception:
            continue

    # build rows
    row = []
    rows = []
    for item in cells:
        row.append(item)
        if len(row) == 2:
            rows.append(row)
            row = []
    if row:
        # pad
        row.append([Paragraph("", getSampleStyleSheet()["BodyText"])])
        rows.append(row)

    # convert nested lists into Table with inner tables per cell
    cell_flow_rows = []
    for r in rows:
        out_row = []
        for cell in r:
            if isinstance(cell, list) and len(cell) == 2:
                inner = Table([[cell[0]], [cell[1]]], colWidths=[80 * mm])
                inner.setStyle(TableStyle([
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ]))
                out_row.append(inner)
            else:
                out_row.append(Paragraph("", getSampleStyleSheet()["BodyText"]))
        cell_flow_rows.append(out_row)

    tbl = Table(cell_flow_rows, colWidths=[86 * mm, 86 * mm], hAlign="LEFT")
    tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    flows.append(tbl)
    flows.append(Spacer(1, 6))
    return flows



class EngineeringCoverPage(Flowable):
    """Cover page tipica per documenti di ingegneria con title-block e spazio timbro.

    Layout:
    - Titolo documento al centro-alto
    - Nome progetto in evidenza
    - Title-block tecnico in basso a destra (Cod. progetto / N. doc / Rev / Data / Progettista / Committente)
    - Riquadro timbro/firma in basso a sinistra (con immagine opzionale PNG)
    """

    def __init__(self, data: Dict[str, Any]):
        super().__init__()
        self.data = data

    def wrap(self, availWidth, availHeight):
        return availWidth, availHeight

    def _first_line(self, s: Any) -> str:
        if not s:
            return ""
        if isinstance(s, str):
            return s.strip().split("\n")[0].strip()
        return str(s).strip()

    def draw(self):
        c = self.canv
        # Nota: questo Flowable viene disegnato all'interno del frame (origine = margini).
        # Per ottenere una cover "a pagina intera" su A4 verticale, trasliamo
        # l'origine al vero (0,0) della pagina.
        width, height = A4

        # Deve essere coerente con i margini del SimpleDocTemplate (vedi genera_pdf_relazione_bytes)
        left_margin = 18 * mm
        bottom_margin = 18 * mm
        c.saveState()
        c.translate(-left_margin, -bottom_margin)

        margin_x = 18 * mm
        margin_top = 18 * mm
        margin_bottom = 18 * mm

        # ---- dati
        titolo = self.data.get("titolo_cover") or "RELAZIONE TECNICO-SPECIALISTICA"
        nome_progetto = self.data.get("nome_progetto") or self.data.get("oggetto_intervento") or "Progetto: ________"
        sottotitolo = self.data.get("sottotitolo_cover") or "IMPIANTO ELETTRICO"
        committente = self.data.get("committente_nome") or ""
        indirizzo = self.data.get("impianto_indirizzo") or self.data.get("luogo_intervento") or ""
        cod_progetto = self.data.get("cod_progetto") or ""
        n_doc = self.data.get("n_documento") or self.data.get("n_doc") or ""
        rev = self.data.get("revisione") or self.data.get("rev") or ""
        data_doc = self.data.get("data_documento") or self.data.get("data_doc") or ""
        progettista = (self.data.get("progettista_nome") or self._first_line(self.data.get("progettista_blocco")) 
                       or self._first_line(self.data.get("firma")))

        # ---- stile base
        c.setStrokeColor(colors.black)
        c.setFillColor(colors.black)

        # Titoli (centro pagina, stile "engineering")
        c.setFont("Times-Bold", 22)
        c.drawCentredString(width / 2, height - margin_top - 18 * mm, titolo)

        c.setFont("Times-Bold", 18)
        c.drawCentredString(width / 2, height - margin_top - 32 * mm, nome_progetto)

        c.setFont("Times-Roman", 12)
        c.drawCentredString(width / 2, height - margin_top - 42 * mm, sottotitolo)

        # Riga info (committente / indirizzo)
        info_y = height - margin_top - 58 * mm
        c.setFont("Times-Bold", 10)
        c.drawString(margin_x, info_y, "Committente:")
        c.setFont("Times-Roman", 10)
        c.drawString(margin_x + 26*mm, info_y, committente[:95])

        c.setFont("Times-Bold", 10)
        c.drawString(margin_x, info_y - 6*mm, "Luogo:")
        c.setFont("Times-Roman", 10)
        c.drawString(margin_x + 26*mm, info_y - 6*mm, indirizzo[:95])

        # Linea di separazione
        c.setLineWidth(1)
        c.line(margin_x, info_y - 14*mm, width - margin_x, info_y - 14*mm)

        # ---- title-block (basso)
        block_h = 52 * mm
        block_w = 92 * mm
        block_x = width - margin_x - block_w
        block_y = margin_bottom

        # Riquadro timbro (basso sinistra)
        stamp_w = 82 * mm
        stamp_h = 52 * mm
        stamp_x = margin_x
        stamp_y = margin_bottom
        c.setLineWidth(1)
        c.rect(stamp_x, stamp_y, stamp_w, stamp_h)

        c.setFont("Times-Bold", 9)
        c.drawString(stamp_x + 3*mm, stamp_y + stamp_h - 5*mm, "Spazio timbro / firma")

        timbro_bytes = self.data.get("timbro_bytes") or self.data.get("timbro_image_bytes") or None
        if timbro_bytes:
            try:
                img = ImageReader(BytesIO(timbro_bytes))
                # area immagine con margini
                c.drawImage(img, stamp_x + 3*mm, stamp_y + 3*mm, stamp_w - 6*mm, stamp_h - 12*mm, preserveAspectRatio=True, anchor='c')
            except Exception:
                pass

        # Title-block tecnico
        c.setLineWidth(1)
        c.rect(block_x, block_y, block_w, block_h)

        # griglia interna: 6 righe
        rows = 6
        row_h = block_h / rows
        for i in range(1, rows):
            y = block_y + i * row_h
            c.line(block_x, y, block_x + block_w, y)

        # 2 colonne (label / value)
        split = block_x + 28 * mm
        c.line(split, block_y, split, block_y + block_h)

        labels = ["Cod. Progetto", "N. Documento", "Revisione", "Data", "Progettista", "Committente"]
        values = [cod_progetto, n_doc, rev, data_doc, progettista, committente]

        c.setFont("Times-Bold", 8.5)
        c.setFillColor(colors.black)
        for i, (lab, val) in enumerate(zip(labels, values)):
            y_text = block_y + block_h - (i + 0.7) * row_h
            c.drawString(block_x + 2*mm, y_text, lab)
            c.setFont("Times-Roman", 8.5)
            c.drawString(split + 2*mm, y_text, (val or "")[:40])
            c.setFont("Times-Bold", 8.5)

        # Nota legale minima
        note = self.data.get("disclaimer_cover") or "Documento emesso a supporto della DiCo ex D.M. 37/08; eventuali aggiornamenti normativi successivi non sono inclusi."
        c.setFont("Times-Roman", 8)
        c.drawString(margin_x, block_y + block_h + 6*mm, note[:120])

        c.restoreState()


class LegacyCoverPage(Flowable):
    """Cover a riquadri (titolo / indice / firma) - legacy."""

    def __init__(self, data: Dict[str, Any]):
        super().__init__()
        self.data = data

    def wrap(self, availWidth, availHeight):
        return availWidth, availHeight

    def draw(self):
        c = self.canv
        width, height = A4

        # Anche questa cover è un Flowable: riportiamo l'origine al (0,0) pagina
        left_margin = 18 * mm
        bottom_margin = 18 * mm
        c.saveState()
        c.translate(-left_margin, -bottom_margin)

        left = 20 * mm
        right = 20 * mm
        top = height - 22 * mm
        bottom = 22 * mm
        w = width - left - right

        box1_h = 92 * mm
        box2_h = 45 * mm
        box3_h = (top - bottom) - box1_h - box2_h - 12 * mm

        y1_top = top
        y1_bot = y1_top - box1_h
        y2_top = y1_bot - 6 * mm
        y2_bot = y2_top - box2_h
        y3_top = y2_bot - 6 * mm
        y3_bot = bottom

        c.setLineWidth(1)
        c.rect(left, y1_bot, w, box1_h)
        c.rect(left, y2_bot, w, box2_h)
        c.rect(left, y3_bot, w, box3_h)

        titolo_grande = (self.data.get("titolo_cover") or "RELAZIONE TECNICO-SPECIALISTICA").upper()
        sottotitolo = self.data.get("sottotitolo_cover") or self.data.get("oggetto_intervento") or ""
        committente = self.data.get("committente_nome") or ""
        luogo = self.data.get("impianto_indirizzo") or ""

        # Titolo grande (Times-Bold, centrato, spezzato su più righe)
        c.setFont("Times-Bold", 20)
        words = titolo_grande.split()
        lines: List[str] = []
        cur = ""
        for w0 in words:
            test = (cur + " " + w0).strip()
            if len(test) > 42 and cur:
                lines.append(cur)
                cur = w0
            else:
                cur = test
        if cur:
            lines.append(cur)

        y = y1_top - 12 * mm
        for ln in lines[:6]:
            c.drawCentredString(left + w / 2, y, ln)
            y -= 9 * mm

        c.setFont("Times-Roman", 14)
        y -= 2 * mm
        if _meaningful(sottotitolo):
            c.drawCentredString(left + w / 2, y, str(sottotitolo))
            y -= 8 * mm

        c.setFont("Times-Roman", 13)
        c.drawCentredString(left + w / 2, y, "IMPIANTO ELETTRICO")
        y -= 6 * mm
        c.setFont("Times-Roman", 12)
        c.drawCentredString(left + w / 2, y, "RELAZIONE TECNICO-SPECIALISTICA")
        y -= 10 * mm

        c.setFont("Times-Roman", 12)
        if _meaningful(committente):
            c.drawCentredString(left + w / 2, y, f"Committente: {committente}")
            y -= 6 * mm
        if _meaningful(luogo):
            luogo_txt = str(luogo)
            max_len = 70
            luogo_lines = [luogo_txt[i : i + max_len] for i in range(0, len(luogo_txt), max_len)]
            for ll in luogo_lines[:2]:
                c.drawCentredString(left + w / 2, y, ll)
                y -= 6 * mm

        # Indice con checkbox
        c.setFont("Times-Roman", 13)
        idx = _build_indice_items(self.data)
        x0 = left + 10 * mm
        y = y2_top - 10 * mm
        for it in idx:
            c.rect(x0, y - 3 * mm, 3.5 * mm, 3.5 * mm)
            c.drawString(x0 + 7 * mm, y - 2 * mm, it)
            y -= 7.5 * mm

        # Firma
        progettista = (
            self.data.get("progettista_nome")
            or _first_nonempty_line(self.data.get("progettista_blocco", ""))
            or self.data.get("firma")
            or ""
        )
        c.setFont("Times-Roman", 14)
        c.drawCentredString(left + w / 2, y3_top - 18 * mm, "Il progettista:")
        c.drawCentredString(left + w / 2, y3_top - 28 * mm, str(progettista))

        # Timbro/firma: se fornito PNG, lo disegna; altrimenti placeholder
        stamp_w = 70 * mm
        stamp_h = 45 * mm
        sx = left + (w - stamp_w) / 2
        sy = y3_bot + 16 * mm
        png_bytes = self.data.get("timbro_png")
        if png_bytes:
            try:
                img = ImageReader(BytesIO(png_bytes))
                c.drawImage(img, sx, sy, width=stamp_w, height=stamp_h, preserveAspectRatio=True, mask='auto')
            except Exception:
                png_bytes = None
        if not png_bytes:
            c.setLineWidth(0.5)
            c.rect(sx, sy, stamp_w, stamp_h)
            c.setFont("Helvetica-Oblique", 9)
            c.setFillColor(colors.grey)
            c.drawCentredString(left + w / 2, sy + stamp_h / 2, "Spazio firma/timbro")
            c.setFillColor(colors.black)

        # Metadati (riga bassa)
        cod = self.data.get("cod_progetto", "")
        rev = self.data.get("rev", "")
        data_doc = self.data.get("data", "")
        meta = " · ".join(
            [
                x
                for x in [
                    f"Cod. {cod}" if _meaningful(cod) else "",
                    f"Rev. {rev}" if _meaningful(rev) else "",
                    f"Data {data_doc}" if _meaningful(data_doc) else "",
                ]
                if x
            ]
        )
        if meta:
            c.setFont("Helvetica", 9)
            c.setFillColor(colors.grey)
            c.drawCentredString(left + w / 2, y3_bot + 8 * mm, meta)
            c.setFillColor(colors.black)

        c.restoreState()


def _draw_header_footer(c: canvas.Canvas, doc, data: Dict[str, Any]):
    """Header/Footer per pagine successive alla cover."""
    width, height = A4
    left = doc.leftMargin
    right = width - doc.rightMargin

    titolo = data.get("header_titolo") or "Relazione Tecnico-Specialistica"
    cod = data.get("cod_progetto")
    rev = data.get("rev")
    data_doc = data.get("data")

    c.saveState()
    c.setFont("Helvetica", 9)
    c.setFillColor(colors.grey)

    header = titolo
    if _meaningful(cod):
        header = f"{titolo} · Cod. {cod}"
    c.drawString(left, height - 12 * mm, header)

    c.setStrokeColor(colors.lightgrey)
    c.setLineWidth(0.3)
    c.line(left, 15 * mm, right, 15 * mm)

    meta = " · ".join(
        [x for x in [f"Rev. {rev}" if _meaningful(rev) else "", f"Data {data_doc}" if _meaningful(data_doc) else ""] if x]
    )
    if meta:
        c.drawString(left, 10 * mm, meta)

    c.restoreState()


def _revision_table(data: Dict[str, Any], styles) -> Optional[Table]:
    revs = data.get("revisioni") or []
    if not revs:
        # default minimo
        rev = data.get("rev", "00")
        dt = data.get("data", "")
        revs = [{"Rev": str(rev), "Data": str(dt), "Descrizione": "Emissione documento"}]

    th = ParagraphStyle("rth", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9, leading=10)
    tc = ParagraphStyle("rtc", parent=styles["Normal"], fontName="Helvetica", fontSize=9, leading=10)

    tdata = [[_p("Rev.", th), _p("Data", th), _p("Descrizione", th)]]
    for r in revs:
        tdata.append([
            _p(str(r.get("Rev", "")), tc),
            _p(str(r.get("Data", "")), tc),
            _p(str(r.get("Descrizione", "")), tc),
        ])

    tbl = Table(tdata, colWidths=[18 * mm, 30 * mm, 126 * mm], hAlign="LEFT", repeatRows=1)
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    return tbl


def genera_pdf_relazione_bytes(data: Dict[str, Any]) -> bytes:
    buf = BytesIO()
    styles = getSampleStyleSheet()

    th = ParagraphStyle("th", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8, leading=9)
    tc = ParagraphStyle("tc", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=9)

    h1 = ParagraphStyle("H1", parent=styles["Heading1"], spaceBefore=6, spaceAfter=8)
    h2 = ParagraphStyle("H2", parent=styles["Heading2"], spaceBefore=8, spaceAfter=6)
    h3 = ParagraphStyle("H3", parent=styles["Heading3"], spaceBefore=6, spaceAfter=4)

    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=20 * mm,
        bottomMargin=18 * mm,
        title="Relazione Tecnico-Specialistica",
    )

    story: List[Any] = []

    # 1) COVER
    cover_style = (data.get('cover_style') or 'engineering').lower()
    story.append(EngineeringCoverPage(data) if cover_style.startswith('eng') else LegacyCoverPage(data))
    story.append(PageBreak())

    # 2) REVISIONI
    story.append(_p("ELENCO DELLE REVISIONI", h1))
    rt = _revision_table(data, styles)
    if rt:
        story.append(rt)
        story.append(Spacer(1, 10))

    # 2.1) Dati identificativi documento
    ident = [
        ["DATI IDENTIFICATIVI DOCUMENTO", ""],
        ["Committente", data.get("committente_nome", "")],
        ["Luogo di installazione", data.get("impianto_indirizzo", "")],
        ["Oggetto intervento", data.get("oggetto_intervento", "")],
        ["Tipologia impianto", data.get("tipologia", "")],
        ["Sistema di distribuzione", data.get("sistema", "")],
        ["Tensione/Frequenza", data.get("tensione", "")],
        ["Potenza impegnata / disponibile", data.get("potenza_disp", "")],
        ["Cod. progetto", data.get("cod_progetto", "")],
        ["N. documento", data.get("n_doc", "")],
        ["Revisione", data.get("rev", "")],
        ["Data", data.get("data", "")],
    ]
    story.append(_kv_table(ident, [55 * mm, 119 * mm]))
    story.append(Spacer(1, 10))

    # 2.2) Dati progettista (se forniti)
    progettista_blocco = data.get("progettista_blocco", "")
    if _meaningful(progettista_blocco):
        story.append(_p("TECNICO PROGETTISTA / REDATTORE", h2))
        story.append(_p(progettista_blocco, styles["BodyText"]))
        story.append(Spacer(1, 10))

    # 2.3) Impresa installatrice (se fornita)
    impresa = data.get("impresa", "")
    if _meaningful(impresa):
        story.append(_p("IMPRESA INSTALLATRICE (dati dichiarati)", h2))
        imp_rows = [
            ["Ragione sociale", impresa],
            ["Sede legale", data.get("impresa_sede", "")],
            ["P.IVA / C.F.", data.get("impresa_piva", "")],
            ["CCIAA / REA", data.get("impresa_rea", "")],
            ["Responsabile tecnico", data.get("impresa_resp", "")],
            ["Recapiti", data.get("impresa_cont", "")],
        ]
        story.append(_kv_table([["Dato", "Valore"]] + imp_rows, [55 * mm, 119 * mm]))
        story.append(Spacer(1, 10))

    # Nota calcoli
    disclaimer = data.get(
        "disclaimer_calcoli",
        "I calcoli e le verifiche riportate sono di sintesi e in linea con le normative applicabili"
        "Le raccomandazioni non sostituiscono le verifiche prescrittive previste dalle norme applicabili.",
    )
    if _meaningful(disclaimer):
        story.append(_p("NOTA", h2))
        story.append(_p(disclaimer, styles["BodyText"]))

    story.append(PageBreak())

    # === STRUTTURA DOCUMENTO (template integrato) ===

    # CAPITOLO 1
    story.append(_p("PREMESSA", h2))
    story.append(_p(data.get("premessa", ""), styles["BodyText"]))
    story.append(Spacer(1, 10))

    # CAPITOLO 2
    story.append(_p("LOCALIZZAZIONE DELL'IMPIANTO", h2))
    loc_text = data.get("localizzazione", "") or data.get("impianto_indirizzo", "")
    story.append(_p(loc_text, styles["BodyText"]))
    story.append(Spacer(1, 10))

    # CAPITOLO 3
    story.append(_p("DESCRIZIONE DELL'OPERA E DELLE SCELTE PROGETTUALI", h2))
    descr_text = data.get("descrizione_opera", "") or data.get("descrizione_impianto", "")
    story.append(_p(descr_text, styles["BodyText"]))
    story.append(Spacer(1, 10))

    # CAPITOLO 4
    story.append(_p("LAYOUT D'IMPIANTO", h2))
    lay_text = data.get("layout", "")
    if _meaningful(lay_text):
        story.append(_p(lay_text, styles["BodyText"]))
    story.append(Spacer(1, 10))

    # CAPITOLO 6 (fotografie) - come nel modello Word, qui e non in fondo
    photo_flows = _photo_flowables(data)
    if photo_flows:
        story.append(_p("DOCUMENTAZIONE FOTOGRAFICA", h2))
        story.extend(photo_flows)
        story.append(Spacer(1, 10))

    # CAPITOLO 7
    caratt_short = data.get("caratteristiche_tecniche", "") or data.get("caratteristiche_tecniche_note", "")
    if _meaningful(caratt_short):
        story.append(_p("CARATTERISTICHE TECNICHE", h2))
        story.append(_p(caratt_short, styles["BodyText"]))
        story.append(Spacer(1, 10))

    # CAPITOLO 8
    asp_short = data.get("aspetti_normativi", "") or data.get("norme", "")
    if _meaningful(asp_short):
        story.append(_p("ASPETTI NORMATIVI", h2))
        story.append(_p(asp_short, styles["BodyText"]))
        story.append(Spacer(1, 10))

    # PROGETTO ELETTRICO – RELAZIONE TECNICA (paragrafi 9..58 integrati)
    story.append(PageBreak())
    story.append(_p("PROGETTO ELETTRICO – RELAZIONE TECNICA", h2))
    story.append(Spacer(1, 6))

    # Tabelle sintetiche (manteniamo calcoli/versione precedente, ma nel punto giusto)
    def _quadri_table_if_any():
        quadri = data.get("quadri", []) or []
        if not quadri:
            return None
        tdata = [[
            _p("Quadro", th),
            _p("Ubicazione", th),
            _p("IP", th),
            _p("Interruttore generale", th),
            _p("Differenziale generale", th),
        ]]
        for q in quadri:
            tdata.append([
                _p(str(q.get("Quadro", "")), tc),
                _p(str(q.get("Ubicazione", "")), tc),
                _p(str(q.get("IP", "")), tc),
                _p(str(q.get("Generale", "")), tc),
                _p(str(q.get("Diff", "")), tc),
            ])
        tbl = Table(tdata, colWidths=[18*mm, 48*mm, 12*mm, 50*mm, 46*mm], repeatRows=1, hAlign="LEFT")
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        return tbl

    def _linee_table_if_any():
        linee = data.get("linee", []) or []
        if not linee:
            return None
        tdata = [[
            _p("Circuito/Linea", th),
            _p("Destinazione", th),
            _p("Posa / L (m)", th),
            _p("Cavo (tipo/sezione)", th),
            _p("Protezione", th),
            _p("Differenziale", th),
            _p("ΔV %", th),
            _p("Esito", th),
        ]]
        for ln in linee:
            posa = (ln.get("Posa", "") or "").strip()
            ll = ln.get("L_m", "")
            posa_len = f"{posa}\n{ll}".strip()
            tdata.append([
                _p(str(ln.get("Linea", "")), tc),
                _p(str(ln.get("Uso", "")), tc),
                _p(posa_len, tc),
                _p(str(ln.get("Cavo", "")), tc),
                _p(str(ln.get("Protezione", "")), tc),
                _p(str(ln.get("Diff", "")), tc),
                _p(str(ln.get("DV_perc", "")), tc),
                _p(str(ln.get("Esito", "")), tc),
            ])
        tbl = Table(tdata, colWidths=[22*mm, 40*mm, 20*mm, 35*mm, 30*mm, 30*mm, 12*mm, 15*mm], repeatRows=1, hAlign="LEFT")
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        return tbl

    tpl = data.get("template_sections") or TEMPLATE_SECTIONS_DEFAULT
    printed_evse = False
    for s in tpl:
        try:
            num = int(s.get("num"))
        except Exception:
            num = None
        titolo = str(s.get("title", "")).strip()
        testo = str(s.get("text", "") if "text" in s else s.get("default", "")).strip()
        include = s.get("include", True)
        if not include or not _meaningful(titolo) or not _meaningful(testo):
            continue

        label = f"{num}  {titolo}" if num else titolo
        story.append(_p(label, h3))
        story.append(_p(testo, styles["BodyText"]))

        # Innesti "intelligenti" (manteniamo i calcoli e le tabelle già presenti nella versione precedente)
        if titolo.startswith("QUADRI ELETTRICI"):
            qt = _quadri_table_if_any()
            if qt is not None:
                story.append(Spacer(1, 6))
                story.append(qt)

        if titolo.startswith("CADUTA DI TENSIONE") or titolo.startswith("CORRENTI DI IMPIEGO"):
            lt = _linee_table_if_any()
            if lt is not None:
                story.append(Spacer(1, 6))
                story.append(lt)

        if titolo.startswith("COLONNINA") and not printed_evse:
            ev_tbl = _evse_table(data, styles)
            if ev_tbl is not None:
                story.append(Spacer(1, 6))
                story.append(ev_tbl)
                printed_evse = True

        # Sezione "VERIFICHE" (collaudo) -> stampa tabella verifiche/prove compilata in app
        if titolo == "VERIFICHE":
            # In app: "verifiche" è un testo narrativo; la tabella compilabile è in "verifiche_tabella".
            # Manteniamo compatibilità: accettiamo anche una lista in data["verifiche"] (vecchie versioni).
            ver = data.get("verifiche_tabella")
            if not isinstance(ver, list):
                ver = data.get("verifiche") if isinstance(data.get("verifiche"), list) else []
            filtered = []
            for r in ver:
                if not isinstance(r, dict):
                    continue
                esito = str(r.get("Esito","")).strip()
                note = str(r.get("Note","")).strip()
                if not _meaningful(esito) and not _meaningful(note):
                    continue
                filtered.append(r)
            if filtered:
                story.append(Spacer(1, 6))
                tdata = [[_p("Verifica/Prova", th), _p("Esito", th), _p("Note", th)]]
                for r in filtered:
                    tdata.append([_p(str(r.get("Verifica","")), tc), _p(str(r.get("Esito","")), tc), _p(str(r.get("Note","")), tc)])
                tbl = Table(tdata, colWidths=[70*mm, 25*mm, 79*mm], repeatRows=1, hAlign="LEFT")
                tbl.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ]))
                story.append(tbl)

        story.append(Spacer(1, 8))

    # ALLEGATI e documentazione (in fondo, come richiesto)
    allg = data.get("allegati", "")
    checklist = data.get("checklist") or []
    if _meaningful(allg) or checklist:
        story.append(PageBreak())
        story.append(_p("ALLEGATI E DOCUMENTAZIONE", h2))

        if checklist:
            story.append(_p("Checklist documentale (D.M. 37/08, D.P.R. 462/2001 se applicabile)", h3))
            tdata = [[_p("Documento / Elaborato", th), _p("Stato", th), _p("Note", th)]]
            for r in checklist:
                docu = str(r.get("Documento / Elaborato","")).strip()
                stato = str(r.get("Stato","")).strip()
                note = str(r.get("Note","")).strip()
                if not _meaningful(docu) and not _meaningful(stato) and not _meaningful(note):
                    continue
                if not _meaningful(stato):
                    continue
                tdata.append([_p(docu, tc), _p(stato, tc), _p(note, tc)])
            if len(tdata) > 1:
                tbl = Table(tdata, colWidths=[95*mm, 25*mm, 50*mm], repeatRows=1, hAlign="LEFT")
                tbl.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ]))
                story.append(tbl)
                story.append(Spacer(1, 8))

        if _meaningful(allg):
            story.append(_p(allg, styles["BodyText"]))


    doc.build(
        story,
        onFirstPage=lambda c, d: None,
        onLaterPages=lambda c, d: _draw_header_footer(c, d, data),
        canvasmaker=_NumberedCanvas,
    )
    return buf.getvalue()
