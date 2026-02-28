import streamlit as st
import os, json
_TEMPLATE_JSON_PATH = os.path.join(os.path.dirname(__file__), "template_sections.json")


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

# Sezioni minime consigliate (modello lean): mantiene il senso tecnico senza appesantire
# Carica template completo da template_sections.json (se presente)
try:
    if os.path.exists(_TEMPLATE_JSON_PATH):
        with open(_TEMPLATE_JSON_PATH, 'r', encoding='utf-8') as _f:
            _tpl = json.load(_f)
        if isinstance(_tpl, list) and len(_tpl) >= 40:
            TEMPLATE_SECTIONS_DEFAULT = _tpl
except Exception:
    pass


LEAN_SECTION_NUMS = {9,10,11,12,13,14,15,16,17,18,21,22,23,28,29,32,33,34,35,36,38,39,40,41,48,50,54,55,56,57,58}

import pandas as pd
from datetime import date

from calcoli import corrente_da_potenza, caduta_tensione, verifica_tt_ra_idn, zs_massima_tn
from pdf_generator import genera_pdf_relazione_bytes


def _meaningful(value) -> bool:
    if value is None:
        return False
    s = str(value).strip()
    if not s:
        return False
    low = s.lower()
    bad = {"non pertinente","non applicabile","n/a","na","—","-","nessuna","nessuna / non applicabile"}
    if low in bad:
        return False
    if "xxxx" in low or "inserire" in low:
        return False
    return True

def _clean_records(records, require_keys=None):
    """Rimuove righe vuote e valori non significativi.
    - require_keys: se specificato, la riga viene mantenuta solo se almeno una di queste chiavi è significativa.
    """
    cleaned=[]
    for r in (records or []):
        rr={}
        for k,v in r.items():
            if _meaningful(v):
                rr[k]=v
        if not rr:
            continue
        if require_keys:
            if not any(_meaningful(r.get(k)) for k in require_keys):
                continue
        cleaned.append(rr)
    return cleaned

st.set_page_config(page_title="Relazione Tecnica DiCo – Impianti Elettrici", layout="wide")

st.title("Relazione Tecnica - Impianto Elettrico (Allegato alla DiCo)")
st.caption("Compilazione guidata (stile v7) + calcoli essenziali + generazione PDF.")

PROGETTISTA_BLOCCO = """Ing. Pasquale Senese
Via Francesco Soave 30 - 20135 Milano
Cell: 340 5731381
Email: pasquale.senese@ingpec.eu
P.IVA: 14572980960
"""

CAVI_TIPO = ["FS17", "FG17", "FG16OR16", "FG16OM16"]

with st.sidebar:
    st.header("Parametri calcoli (sintesi)")
    dv_lim = st.number_input("Caduta di tensione max (%)", min_value=1.0, max_value=10.0, value=4.0, step=0.5)
    ul_tt = st.number_input("UL sistema TT (V) – criterio Ra·Idn ≤ UL", min_value=25.0, max_value=100.0, value=50.0, step=5.0)
    st.divider()
    st.markdown("**Nota**: - ")

# =========================
# DATI IDENTIFICATIVI
# =========================
st.subheader("Dati identificativi documento")

c1, c2, c3 = st.columns(3)
with c1:
    committente = st.text_input("Committente", "")
    luogo = st.text_input("Luogo di installazione (indirizzo completo)", "")
    oggetto = st.text_input("Oggetto intervento (descrizione sintetica)", "")
with c2:
    tipologia = st.selectbox("Tipologia impianto", ["Nuova realizzazione", "Ampliamento", "Trasformazione", "Manutenzione straordinaria"], index=3)
    sistema = st.selectbox("Sistema di distribuzione", ["TT", "TN-S", "TN-C-S", "IT"], index=0)
    alimentazione = st.selectbox("Alimentazione", ["Monofase 230 V", "Trifase 400 V"], index=1)
with c3:
    tensione = st.text_input("Tensione/Frequenza", "230/400 V - 50 Hz")
    potenza_disp_kw = st.text_input("Potenza impegnata / disponibile", "")
    cod_progetto = st.text_input("Cod. progetto", "")
    nome_progetto = st.text_input("Nome progetto", "")
    cover_style = st.selectbox("Stile cover", ["Engineering (title-block)", "A riquadri (legacy)"], index=0)

    n_doc = st.text_input("N. documento", "")
    revisione = st.text_input("Revisione", "00")
    data_doc = st.date_input("Data", value=date.today())

st.subheader("Revisioni documento")
rev_df = pd.DataFrame([
    {"Rev": str(revisione), "Data": data_doc.strftime('%d/%m/%Y'), "Descrizione": "Emissione documento"},
])
rev_df = st.data_editor(rev_df, num_rows="dynamic", use_container_width=True, key="revisioni")

# Carica opzionale immagine timbro/firma per la cover
timbro_file = st.file_uploader("Timbro/Firma (PNG) - opzionale", type=["png"], accept_multiple_files=False)
timbro_bytes = timbro_file.getvalue() if timbro_file else None

st.divider()

# =========================
# SEZIONI SPECIFICHE EV / DOCUMENTAZIONE
# =========================
st.subheader("Infrastruttura di ricarica EV (se applicabile)")
st.caption("Compila almeno marca/modello/potenza. Se non pertinente, elimina le righe o lascia campi vuoti.")

default_evse = pd.DataFrame([
    {
        "Tipo": "Wallbox",
        "Marca": "",
        "Modello": "",
        "N. punti": "",
        "Potenza (kW)": "",
        "Alimentazione": "Monofase",
        "Connettore": "Tipo 2",
        "Modo ricarica": "Mode 3",
        "IP/IK": "IPXX / IKXX",
        "RCD richiesto": "Tipo A 30mA + RDC-DD 6mA / Tipo B (se previsto)",
        "Note": "",
    }
])

evse_df = st.data_editor(default_evse, num_rows="dynamic", use_container_width=True, key="evse")

st.subheader("Localizzazione, layout e documentazione fotografica")
c1, c2 = st.columns(2)
with c1:
    localizzazione = st.text_area(
        "Localizzazione dell'impianto (descrizione sintetica)",
        "",
        height=110,
    )
with c2:
    layout = st.text_area(
        "Layout d'impianto (elementi principali)",
        ": quadro/i, linee principali, protezioni, posa, posizione wallbox/colonnina).",
        height=110,
    )

foto_files = st.file_uploader(
    "Foto (JPG/PNG) - opzionale (consigliate max 6)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

st.divider()

# =========================
# SOGGETTI COINVOLTI
# =========================
st.subheader("Soggetti coinvolti")

c1, c2 = st.columns(2)
with c1:
    st.markdown("**Impresa installatrice**")
    impresa = st.text_input("Ragione sociale", "")
    impresa_sede = st.text_input("Sede legale", "")
    impresa_piva = st.text_input("P.IVA / C.F.", "")
    impresa_rea = st.text_input("N. iscrizione CCIAA / REA", "")
    impresa_resp = st.text_input("Responsabile tecnico", "")
    impresa_cont = st.text_input("Recapiti", "")
with c2:
    st.markdown("**Progettista / Tecnico redattore**")
    progettista_blocco = st.text_area(
        "Dati progettista (blocco)",
        value=(
            "Ing. Pasquale Senese\n"
            "Via Francesco Soave 30 - 20135 Milano (MI) - Cell: 340 5731381\n"
            "Email: pasquale.senese@ingpec.eu  P.IVA: 14572980960"
        ),
        height=100,
    )

# Deriva il nominativo (prima riga) per cover/title-block
progettista_nome = (progettista_blocco.strip().splitlines()[0].strip() if progettista_blocco.strip() else "")

st.divider()

# =========================
# DATI TECNICI MINIMI
# =========================
st.subheader("Dati tecnici minimi (da compilare)")

c1, c2, c3, c4 = st.columns(4)
with c1:
    pod = st.text_input("POD / punto di consegna", "")
with c2:
    contatore_ubi = st.text_input("Contatore ubicato in", "")
with c3:
    potenza_prev_kw = st.number_input("Potenza prevista/servita (kW) – per stima Ib", min_value=0.5, max_value=500.0, value=6.0, step=0.5)
with c4:
    cosphi = st.number_input("cosφ (se noto)", min_value=0.3, max_value=1.0, value=0.95, step=0.01)

Ib = corrente_da_potenza(potenza_prev_kw, alimentazione, cosphi=cosphi)
st.info(f"Corrente di impiego indicativa Ib ≈ **{Ib:.1f} A** (stima da potenza {potenza_prev_kw:.1f} kW, cosφ={cosphi:.2f}).")

ambienti = st.multiselect(
    "Destinazione d’uso / ambienti (checklist)",
    ["Ordinario", "Bagno", "Esterno", "Locale tecnico", "Autorimessa", "Maggior rischio incendio", "Cantiere", "Altro"],
    default=["Ordinario"],
)
amb_altro = ""
if "Altro" in ambienti:
    amb_altro = st.text_input("Specificare 'Altro'", "")

# Campi per eliminare XXXX in premessa/norme
st.subheader("Fonti dati e prescrizioni (per evitare 'XXXX' nel PDF)")
c1, c2 = st.columns(2)
with c1:
    fonte_dati = st.text_input("Fonte dati fornitura/condizioni (Committente/Impresa/Gestore)", "Committente")
with c2:
    prescrizioni_enti = st.text_input("Prescrizioni Enti/Autorità locali (se presenti)", "Nessuna / Non applicabile")



st.divider()

# =========================
# DATI DA RACCOGLIERE IN FASE DI PROGETTAZIONE (RILIEVO)
# =========================
st.subheader("Dati raccolti in fase di progettazione (rilievo e ipotesi di progetto)")
st.caption("Questi dati servono a rendere la relazione più completa e coerente con quanto normalmente richiesto in sede di progetto/DiCo (rilievi, vincoli, presupposti).")

c1, c2, c3 = st.columns(3)
with c1:
    data_sopralluogo = st.date_input("Data sopralluogo (se effettuato)", value=date.today())
    referente = st.text_input("Referente in sito / contatto", "")
with c2:
    documenti_disponibili = st.multiselect(
        "Documenti disponibili (rilievo)",
        ["Planimetrie/lay-out", "Schema unifilare esistente", "DiCo/DiRi preesistente", "Contratto/condizioni di fornitura", "Schede tecniche EVSE", "Altro"],
        default=["Planimetrie/lay-out"],
    )
    doc_altro = ""
    if "Altro" in documenti_disponibili:
        doc_altro = st.text_input("Specificare 'Altro' (documenti)", "")
with c3:
    vincoli = st.multiselect(
        "Vincoli/condizioni al contorno",
        ["Condominio (regolamento/autorimessa)", "Vincoli edilizi/architettonici", "Ambiente esterno (IP/IK)", "Ambiente con rischio incendio", "VV.F. / DPR 151/2011", "Nessuno", "Altro"],
        default=["Nessuno"],
    )
    vincoli_note = st.text_input("Note vincoli (se utili)", "—")

c1, c2 = st.columns(2)
with c1:
    icc_presunta = st.text_input("Corrente di cortocircuito presunta a monte (se nota)", "N.D. / da acquisire")
    protezione_monte = st.text_input("Protezione a monte / dispositivo generale (se noto)", "N.D. / da acquisire")
    impianto_esistente = st.selectbox("Impianto esistente", ["Sì (verificato)", "Sì (non verificato)", "No / Nuovo"], index=0)
with c2:
    ra_misurata = st.text_input("Resistenza di terra Ra (Ω) misurata/attesa (se nota)", "N.D. / da misurare")
    percorsi_posa = st.text_area("Percorsi cavi e modalità posa (rilievo sintetico)", ": canalizzazioni/tubazioni/passerelle, attraversamenti, ecc.)", height=90)

note_progettazione = st.text_area("Note di progettazione (presupposti, ipotesi, criticità)", "—", height=100)

# Testo pronto per PDF
doc_list = [d for d in documenti_disponibili if d != "Altro"]
if doc_altro.strip():
    doc_list.append(doc_altro.strip())

vincoli_list = [v for v in vincoli if v != "Altro"]
if "Altro" in vincoli and vincoli_note.strip() and vincoli_note.strip() != "—":
    vincoli_list.append(vincoli_note.strip())

dati_progettazione_txt = f"""Rilievo / progettazione:
- Sopralluogo: {data_sopralluogo.strftime('%d/%m/%Y')} – Referente: {referente}
- Documenti disponibili: {", ".join(doc_list) if doc_list else "N.D."}
- Vincoli/condizioni al contorno: {", ".join(vincoli_list) if vincoli_list else "N.D."}
- Corrente di cortocircuito presunta a monte: {icc_presunta}
- Protezione a monte / dispositivo generale: {protezione_monte}
- Impianto esistente: {impianto_esistente}
- Ra (Ω) misurata/attesa: {ra_misurata}
- Percorsi/posa (sintesi): {percorsi_posa}
- Note: {note_progettazione}
"""

# =========================
# CRITERIO DI PROGETTO (ESTESO - da relazione tecnico-specialistica)
# =========================
st.subheader("Criterio di progetto degli impianti (testo esteso)")
st.caption(
    "Questo capitolo riprende la relazione tecnico-specialistica (con formule). "
    "Se lo disattivi, non verrà stampato nel PDF."
)
includi_criterio = st.checkbox("Includi capitolo 3 esteso (criterio di progetto)", value=True)
cosphi_ricarica = st.number_input(
    "Fattore di potenza (cosφ) per linee prese di ricarica (se presenti)",
    min_value=0.50, max_value=1.00, value=0.99, step=0.01
)
criterio_note = st.text_area("Note/integrazioni al capitolo 3 (opzionale)", "", height=90)


st.divider()

# =========================
# CONFINE INTERVENTO
# =========================
st.subheader("Confini dell’intervento e interfacce")

c1, c2 = st.columns(2)
with c1:
    compresi = st.text_area("L’intervento comprende", " elenco sintetico delle opere incluse).", height=120)
with c2:
    esclusi = st.text_area("Sono esclusi", ", es. parti preesistenti non modificate, linee a monte, apparecchiature non comprese).", height=120)

integrazione = st.selectbox("Integrazione con impianto esistente", ["Sì", "No"], index=0)
integrazione_note = ""
if integrazione == "Sì":
    integrazione_note = st.text_area("Descrizione e condizioni riscontrate/limiti di intervento", ").", height=90)

st.divider()

# =========================
# QUADRI
# =========================
st.subheader("Quadri elettrici e distribuzione (tabella sintetica)")

default_quadri = pd.DataFrame([
    {"Quadro":"QG", "Ubicazione":")", "IP":"XX", "Interruttore generale (tipo/In)":")", "Differenziale generale (tipo/Idn, se presente)":")"},
])
quadri_df = st.data_editor(default_quadri, num_rows="dynamic", use_container_width=True, key="quadri")

st.divider()

# =========================
# LINEE / CIRCUITI
# =========================
st.subheader("Circuiti, cavi e protezioni (con calcolo ΔV)")

st.caption("Per ciascun circuito: scegli **Tipo cavo (FS17/FG17/FG16OR16/FG16OM16)**, sezione, protezione e differenziale. "
           "Il calcolo automatico mostra ΔV% e un esito sintetico.")

default_linee = pd.DataFrame([
    {"Circuito/Linea":"L1", "Destinazione/Utilizzo":"Prese", "Potenza_kW":2.0, "Posa": "", "Lunghezza_m":25,
     "Tipo_cavo":"FG16OM16", "Formazione":"3G", "Sezione_mm2":2.5,
     "Protezione (MT/MTD)":"MT 16A curva C", "Curva":"C", "In_A":16,
     "Differenziale (tipo/Idn)":"Tipo A 30mA", "Tipo_diff":"A", "Idn_mA":30,
     "Ra_Ohm (solo TT)":30.0},
])

linee_df = st.data_editor(
    default_linee,
    num_rows="dynamic",
    use_container_width=True,
    key="linee",
    column_config={
        "Tipo_cavo": st.column_config.SelectboxColumn("Tipo cavo", options=CAVI_TIPO, required=True),
        "Formazione": st.column_config.TextColumn("Formazione (es. 3G / 5G)", help="Esempio: 3G per monofase+PE, 5G per trifase+N+PE."),
        "Tipo_diff": st.column_config.SelectboxColumn("Tipo diff", options=["AC","A","F","B"], required=False),
        "Idn_mA": st.column_config.NumberColumn("Idn (mA)", min_value=0, max_value=3000, step=1),
    }
)

def valuta_linea(row):
    p = float(row.get("Potenza_kW") or 0.0)
    l = float(row.get("Lunghezza_m") or 0.0)
    s = float(row.get("Sezione_mm2") or 0.0)
    curva = str(row.get("Curva") or "C")
    InA = float(row.get("In_A") or 0.0)

    ib_linea = corrente_da_potenza(p, alimentazione, cosphi=cosphi) if p > 0 else Ib

    dv = caduta_tensione(ib_linea, l, s, alimentazione, cosphi=cosphi)
    esito = "OK" if dv.delta_v_percent <= dv_lim else "ΔV"

    note = []
    if sistema == "TT":
        ra = float(row.get("Ra_Ohm (solo TT)") or 0.0)
        idn_a = float(row.get("Idn_mA") or 0.0) / 1000.0
        if ra > 0 and idn_a > 0:
            ok_tt = verifica_tt_ra_idn(ra, idn_a, ul=ul_tt)
            note.append("TT OK" if ok_tt else "TT NO")
            if not ok_tt:
                esito = "TT"
    else:
        if InA > 0:
            zs_max = zs_massima_tn(230.0, curva, InA)
            note.append(f"Zs_max≈{zs_max:.2f}Ω")

    return dv.delta_v_percent, esito, "; ".join(note)

out = [valuta_linea(r) for _, r in linee_df.iterrows()]

linee_df_calc = linee_df.copy()
linee_df_calc["ΔV_%"] = [round(x[0], 2) for x in out]
linee_df_calc["Esito"] = [x[1] for x in out]
linee_df_calc["Note"] = [x[2] for x in out]

st.dataframe(linee_df_calc, use_container_width=True)

st.divider()

# =========================
# SICUREZZA / TERRA / SPD + CPI
# =========================
st.subheader("Sicurezza elettrica, terra, SPD (sintesi)")

c1, c2 = st.columns(2)
with c1:
    terra_cfg = st.selectbox("Configurazione impianto di terra", ["Nuovo", "Esistente verificato", "Esistente non oggetto di intervento (da motivare)"], index=1)
    dispersore = st.text_input("Dispersore (descrizione)", "")
    equipot = st.selectbox("Collegamenti equipotenziali principali", ["Presenti", "Parziali", "Assenti (da adeguare/indicare)"], index=0)
with c2:
    spd_esito = st.selectbox("Protezione contro sovratensioni (SPD) – esito", ["Non previsto", "Previsto", "Presente preesistente"], index=0)
    spd_tipo = st.multiselect("Se installato: tipologia SPD", ["Tipo 1", "Tipo 2", "Tipo 3"], default=[])
    spd_quadro = st.text_input("Quadro di installazione SPD (se pertinente)", "")
    spd_caratt = st.text_input("Caratteristiche principali SPD (se pertinente)", "")

st.subheader("Prevenzione incendi / VV.F. (se pertinente)")
c1, c2, c3 = st.columns(3)
with c1:
    attivita_vvf = st.selectbox("Attività soggetta VV.F. (DPR 151/2011)", ["Non pertinente", "Sì", "No (da verificare)"], index=0)
with c2:
    cpi = st.selectbox("CPI / SCIA antincendio", ["Non pertinente", "Presente", "Non presente", "In corso"], index=0)
with c3:
    vvf_note = st.text_input("Note VV.F. (se pertinente)", "")

st.divider()

# =========================
# VERIFICHE
# =========================
st.subheader("Verifiche, prove e collaudi (registro sintetico)")

ver_df = pd.DataFrame([
    {"Prova / Verifica":"Esame a vista", "Esito":"", "Strumento":"", "Note":""},
    {"Prova / Verifica":"Continuità PE ed equipotenziale", "Esito":"", "Strumento":"", "Note":""},
    {"Prova / Verifica":"Resistenza di isolamento", "Esito":"", "Strumento":"", "Note":""},
    {"Prova / Verifica":"Prova differenziali (Idn/tempo)", "Esito":"", "Strumento":"", "Note":""},
    {"Prova / Verifica":"Polarità / sequenza fasi (se pertinente)", "Esito":"", "Strumento":"", "Note":""},
    {"Prova / Verifica":"TT: misura Ra e coordinamento con Idn (se TT)", "Esito":"", "Strumento":"", "Note":""},
    {"Prova / Verifica":"TN: misura Zs e verifica intervento (se TN)", "Esito":"", "Strumento":"", "Note":""},
    {"Prova / Verifica":"Altre prove (SPD, emergenza, comandi, ecc.)", "Esito":"", "Strumento":"", "Note":""},
])
ver_df = st.data_editor(ver_df, num_rows="dynamic", use_container_width=True, key="verifiche")

st.divider()


# =========================
# CHECKLIST DOCUMENTALE (DM 37/08 - DPR 462/01)
# =========================
st.subheader("Checklist documentale (DM 37/08 e adempimenti correlati)")
st.caption("Compila lo stato degli elaborati/atti da allegare o consegnare. La checklist verrà riportata nel Capitolo 6 del PDF.")

default_check = pd.DataFrame([{"Documento / Elaborato": "", "Stato": "", "Note": ""}])

checklist_df = st.data_editor(
    default_check,
    num_rows="dynamic",
    use_container_width=True,
    key="checklist",
    column_config={
        "Stato": st.column_config.SelectboxColumn(
            "Stato", options=["", "Presente", "Da produrre", "Non applicabile", "Consigliato"], required=False
        )
    }
)

st.divider()

# =========================
# FIRMA
# =========================
st.subheader("Firma (stampa nel PDF)")
c1, c2, c3 = st.columns(3)
with c1:
    luogo_firma = st.text_input("Luogo firma", "")
with c2:
    data_firma = st.date_input("Data firma", value=data_doc)
with c3:
    firmatario = st.text_input("Firmatario", "Ing. Pasquale Senese")

st.divider()

# =========================
# GENERAZIONE PDF
# =========================
st.subheader("Genera PDF")

amb_txt = ", ".join([a for a in ambienti if a != "Altro"])
if "Altro" in ambienti:
    amb_txt += f", Altro: {amb_altro}"

premessa = f"""La presente Relazione Tecnico‑Specialistica è redatta nell’ambito dell’incarico conferito dalla Committenza "{committente}" e riguarda l’intervento "{oggetto}" presso "{luogo}".

FINALITÀ E PERIMETRO
Il documento ha lo scopo di:
• descrivere l’impianto e le opere eseguite/da eseguire, con indicazione dei confini dell’intervento;
• richiamare i riferimenti legislativi e normativi applicabili;
• esplicitare i criteri di progettazione e le verifiche di coordinamento essenziali (correnti, cadute di tensione, protezioni), in coerenza con la regola dell’arte.

VALENZA DOCUMENTALE
La presente Relazione costituisce documento tecnico di progetto e di supporto alla documentazione di conformità ai sensi del D.M. 37/2008; non sostituisce la Dichiarazione di Conformità (DiCo) né i relativi allegati obbligatori, che restano di competenza dell’Impresa installatrice.

RESPONSABILITÀ E DATI DI INGRESSO
Le informazioni relative alla fornitura elettrica (POD, potenza disponibile/contrattuale, caratteristiche del punto di consegna), destinazione d’uso e condizioni di esercizio sono state fornite da "{fonte_dati}" e/o rilevate in sito e/o confermate in data {data_doc.strftime('%d/%m/%Y')}. Eventuali porzioni preesistenti non oggetto di intervento e le interfacce con impianti/parti terze sono indicate nel paragrafo "Confini dell’intervento".

REQUISITI MATERIALI E CONSEGNA
Materiali e componenti devono essere conformi alle norme applicabili, provvisti di marcatura CE e, ove disponibile, marchio di conformità volontario (es. IMQ) o equivalente. Alla consegna l’impianto deve risultare conforme alla regola dell’arte e alle prescrizioni eventualmente impartite da Enti/Autorità competenti.
"""

norme = f"""Si riportano i principali riferimenti legislativi e normativi applicabili (elenco non esaustivo):

• D.M. 22/01/2008 n. 37.
• Legge 01/03/1968 n. 186.
• D.Lgs. 09/04/2008 n. 81 e s.m.i.
• D.P.R. 22/10/2001 n. 462 (ove applicabile).
• Norme CEI applicabili (in particolare CEI 64-8, CEI 64-14, CEI EN 61439, CEI EN 60529; e, se pertinenti, CEI 81-10, CEI 0-10, CEI 0-21/0-16).
• Regolamento Prodotti da Costruzione (UE) 305/2011 (CPR) e norme CEI-UNEL per i cavi (ove applicabile).

Eventuali ulteriori prescrizioni di Enti/Autorità locali: {prescrizioni_enti}.
"""


# =========================
# SEZIONI ESTESE INTEGRATE (Cap. 7-8) – OPZIONALE
# =========================
st.subheader("Sezioni estese integrate (opzionali)")
st.caption(
    "Le sezioni compilate qui vengono integrate nel documento (Capitoli 7 e 8). "
    "Le righe non compilate non verranno stampate nel PDF."
)

precompila_titoli = st.checkbox(
    "Precompila i titoli standard (testi vuoti)",
    value=True,
    help="Inserisce l'elenco dei paragrafi tipici (come nel modello Word). I testi restano vuoti finché non li compili.",
)

def _default_items_cap7():
    titoli = [
        "9 - Area di intervento e tipo di attività",
        "10 - Tipo di impianto",
        "11 - Punto di origine",
        "12 - Sistema di fornitura",
        "13 - Tensione nominale",
        "14 - Sistema di distribuzione",
        "15 - Corrente di corto circuito",
        "16 - Potenza impegnata",
        "17 - Caduta di tensione",
        "18 - Correnti di impiego e portate dei cavi",
        "19 - Sezione minima dei conduttori di fase",
        "20 - Sezione minima dei conduttori di neutro",
        "21 - Sezione minima dei conduttori di protezione (PE)",
        "22 - Sezione minima del conduttore di terra",
        "23 - Colori di identificazione",
        "24 - Sezionamento e comando",
        "28 - Protezione contro i contatti diretti",
        "29 - Protezione contro i contatti indiretti – sistema TT",
        "32 - Protezione delle condutture contro le sovracorrenti",
        "35 - Selettività",
        "36 - Schemi e documentazione",
        "37 - Descrizione degli impianti",
        "38 - Quadri elettrici",
        "39 - Distribuzione principale",
        "40 - Stazione di ricarica / colonnina (se applicabile)",
    ]
    return [{"titolo": t, "testo": ""} for t in titoli]

def _default_items_cap8():
    titoli = [
        "8.1 - Quadro normativo di riferimento (CEI/UNI/Leggi)",
        "8.2 - D.M. 37/08: conformità e documentazione",
        "8.3 - D.P.R. 462/01: verifiche (se applicabile)",
        "8.4 - Infrastrutture di ricarica VE (CEI 64-8/7-722, CEI EN 61851-1)",
        "8.5 - Eventuali prescrizioni VV.F. / prevenzione incendi (se applicabile)",
    ]
    return [{"titolo": t, "testo": ""} for t in titoli]

def _rows_to_items(df: pd.DataFrame):
    items = []
    for _, r in df.iterrows():
        titolo = str(r.get("titolo", "")).strip()
        testo = str(r.get("testo", "")).strip()
        # Salva righe parziali: il PDF stamperà solo titolo+testo compilati
        if titolo or testo:
            items.append({"titolo": titolo, "testo": testo})
    return items

with st.expander("Capitolo 7 – Caratteristiche tecniche (facoltativo)", expanded=False):
    cap7_df = pd.DataFrame(_default_items_cap7() if precompila_titoli else [{"titolo": "", "testo": ""}])
    cap7_df = st.data_editor(
        cap7_df,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "titolo": st.column_config.TextColumn("Titolo paragrafo"),
            "testo": st.column_config.TextColumn("Testo", width="large"),
        },
        key="cap7_editor",
    )
    caratteristiche_tecniche_note = st.text_area(
        "Note tecniche aggiuntive (opzionale)",
        value="",
        height=100,
        placeholder="(opzionale) Se vuoto non verrà stampato.",
        key="cap7_note",
    )

with st.expander("Capitolo 8 – Aspetti normativi (facoltativo)", expanded=False):
    cap8_df = pd.DataFrame(_default_items_cap8() if precompila_titoli else [{"titolo": "", "testo": ""}])
    cap8_df = st.data_editor(
        cap8_df,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "titolo": st.column_config.TextColumn("Titolo paragrafo"),
            "testo": st.column_config.TextColumn("Testo", width="large"),
        },
        key="cap8_editor",
    )
    aspetti_normativi_note = st.text_area(
        "Note normative aggiuntive (opzionale)",
        value="",
        height=100,
        placeholder="(opzionale) Se vuoto non verrà stampato.",
        key="cap8_note",
    )

caratteristiche_tecniche_items = _rows_to_items(cap7_df)
aspetti_normativi_items = _rows_to_items(cap8_df)


dati_tecnici = f"""Tipo sistema di distribuzione: {sistema}. Tensione nominale: {tensione}. Potenza disponibile/contrattuale: {potenza_disp_kw}.
POD: {pod} – contatore ubicato in: {contatore_ubi}.
Alimentazione: {alimentazione}. Potenza prevista/servita (stima): {potenza_prev_kw:.1f} kW (Ib indicativa ≈ {Ib:.1f} A a cosφ={cosphi:.2f}).
Ambientazioni particolari (se presenti): {amb_txt}.
"""

descrizione_impianto = f"""Il sito di intervento è ubicato in {luogo}. L’impianto è alimentato in bassa tensione dal punto di consegna del Distributore (POD: {pod}), tramite contatore/quadretto di misura ubicato in {contatore_ubi}.
Tipo sistema di distribuzione: {sistema}. Tensione nominale: {tensione}. Potenza disponibile/contrattuale: {potenza_disp_kw}.

La ripartizione e distribuzione interna avviene mediante linee in cavo conforme CEI/UNEL e componenti marcati CE (e, ove disponibile, IMQ o equivalente). Le condutture sono posate in tubazioni/canalizzazioni idonee e con protezione meccanica adeguata; i circuiti risultano identificati e separati per destinazione d’uso (illuminazione, prese, ausiliari, ecc.), privilegiando la manutenibilità.

Scopo dell’intervento (descrizione sintetica): {oggetto}

Le opere impiantistiche previste comprendono, in funzione dell’intervento, la realizzazione e/o modifica di linee di alimentazione dedicate, installazione di punti di utilizzo, posa di tubazioni/canalizzazioni, installazione o adeguamento di quadri elettrici (generale e/o di zona), apparecchi di protezione e comando, morsetterie e accessori, nonché collegamenti al sistema di protezione (PE) e ai collegamenti equipotenziali.

I conduttori sono identificati secondo codifica colori (PE giallo-verde, N blu, fasi marrone/nero/grigio) e marcatura/etichettatura dove previsto. I dispositivi di protezione sono coordinati con le linee e con il sistema di distribuzione (TT/TN) in modo coerente con le norme tecniche applicabili.
"""

confini_txt = f"""L’intervento comprende: {compresi}

Sono esclusi: {esclusi}

Integrazione con impianto esistente: {integrazione}. {("Descrizione e limiti: " + integrazione_note) if integrazione == "Sì" else ""}
"""

# Costruisci frase "Idn/tipo" a partire dai dati delle linee (se presenti)
diff_tipici = []
for _, r in linee_df_calc.iterrows():
    td = str(r.get("Tipo_diff") or "").strip()
    idn = int(r.get("Idn_mA") or 0)
    if td and idn:
        diff_tipici.append(f"Tipo {td} {idn} mA")
diff_frase = ", ".join(sorted(set(diff_tipici))) if diff_tipici else "N.D."

vvf_blocco = ""
if attivita_vvf != "Non pertinente" or cpi != "Non pertinente":
    vvf_blocco = f"Prevenzione incendi / VV.F.: attività soggetta: {attivita_vvf}; CPI/SCIA: {cpi}. Note: {vvf_note}."

sicurezza = f"""La protezione contro i contatti diretti è assicurata tramite isolamento delle parti attive, involucri/barriere con grado di protezione adeguato e corretta posa delle condutture.

La protezione contro i contatti indiretti è assicurata mediante interruzione automatica dell’alimentazione, in accordo con CEI 64-8, tramite dispositivi differenziali e/o magnetotermici coordinati con l’impianto di terra (nei sistemi TT) o con il conduttore di protezione (nei sistemi TN).

Protezione differenziale adottata (sintesi): {diff_frase}.

Configurazione impianto di terra: {terra_cfg}. Dispersore: {dispersore}. Collegamenti equipotenziali principali: {equipot}.

Protezione contro le sovratensioni (SPD) – esito: {spd_esito}. {("Tipologia: " + ", ".join(spd_tipo) + " – ") if spd_tipo else ""}quadro: {spd_quadro}. Caratteristiche: {spd_caratt}.

Caduta di tensione: verificata entro il limite adottato in progetto: {dv_lim:.1f}%.
{vvf_blocco}
"""

verifiche = """Ad ultimazione dei lavori, l’impianto è sottoposto alle verifiche previste dalla CEI 64-8 (Parte 6) e dalla CEI 64-14, con esecuzione e registrazione delle prove strumentali pertinenti al sistema di distribuzione (TT/TN) e alla tipologia di impianto. In particolare:\n\n"""
for _, r in ver_df.iterrows():
    verifiche += f"• {r.get('Prova / Verifica','')}: {r.get('Esito','')} – Strumento: {r.get('Strumento','')} – Note: {r.get('Note','')}\n"

manutenzione = """Le attività di esercizio e manutenzione devono essere svolte da personale qualificato e autorizzato, in sicurezza e nel rispetto delle istruzioni dei costruttori e delle norme tecniche applicabili (es. CEI 0-10 / CEI 11-27, ove pertinenti).

PIANO DI MANUTENZIONE (minimo consigliato)
• Quadri elettrici: ispezione visiva, pulizia, verifica serraggi morsetti, integrità targhe/etichette e dispositivi di protezione;
• Dispositivi differenziali: prova periodica con tasto "T" e verifiche strumentali (Idn/tempo) secondo periodicità e criticità del sito;
• Conduttori e condutture: verifica integrità isolamento, fissaggi, protezioni meccaniche e segregazioni;
• Collegamenti equipotenziali e PE: controllo continuità e integrità;
• Comandi/emergenze (se presenti): prova funzionale e ripristino, verifica segnalazioni e cartellonistica;
• Apparecchiature specifiche (es. wallbox/utenze dedicate): ispezione cavi e connettori, prova funzionale e aggiornamenti firmware se previsti dal costruttore.

È raccomandata la tenuta di un registro manutenzione con data, attività eseguite, esito e nominativo dell’operatore."""

allegati = """Completano la presente relazione e/o la DiCo i seguenti allegati. 

- Schema unifilare / multifilare dei quadri interessati: Obbligatorio.
- Elenco linee/circuiti con cavo e protezione (tabella circuiti): Obbligatorio.
- Verbali e report delle misure e prove strumentali
- Schede tecniche principali componenti (quadri, interruttori, SPD, ecc.): Se disponibile.
- Dichiarazioni/Marcature CE (ed eventuale IMQ) dei materiali: Se disponibile.
- Report fotografico essenziale (quadri, targhette, collegamenti di terra, punti significativi): Consigliato.
"""



st.markdown("### Struttura relazione (template integrato)")
stile_relazione = st.radio(
    "Seleziona lo stile del documento",
    ["Lean (essenziale)", "Completo (come template)"],
    index=1,
    help="Lean: include solo i paragrafi che servono quasi sempre. Completo: replica la struttura del template con tutti i paragrafi."
)
edit_avanzata = st.checkbox(
    "Modifica testi dei paragrafi (avanzato)",
    value=False,
    help="Se disattivato, il documento usa il testo standard del template (coerente e pronto)."
)

include_nums = set(LEAN_SECTION_NUMS) if stile_relazione.startswith("Lean") else {s["num"] for s in TEMPLATE_SECTIONS_DEFAULT}

template_sections = []
for sec in TEMPLATE_SECTIONS_DEFAULT:
    num = sec["num"]
    title = sec["title"]
    default_text = sec.get("default", "")
    include = (num in include_nums)

    text_val = default_text
    if edit_avanzata:
        with st.expander(f"{num}  {title}", expanded=False):
            include = st.checkbox("Includi questo paragrafo", value=include, key=f"tpl_inc_{num}")
            text_val = st.text_area("Testo", value=default_text, height=180, key=f"tpl_txt_{num}")

    template_sections.append({
        "num": num,
        "title": title,
        "text": text_val,
        "include": include,
    })
if st.button("Genera PDF"):
    quadri_list = []
    for _, q in quadri_df.iterrows():
        quadri_list.append({
            "Quadro": q.get("Quadro",""),
            "Ubicazione": q.get("Ubicazione",""),
            "IP": q.get("IP",""),
            "Generale": q.get("Interruttore generale (tipo/In)",""),
            "Diff": q.get("Differenziale generale (tipo/Idn, se presente)",""),
        })

    linee_list = []
    for _, r in linee_df_calc.iterrows():
        tipo = r.get("Tipo_cavo","")
        form = r.get("Formazione","")
        sez = r.get("Sezione_mm2","")
        cavo_str = f"{tipo} {form}x{sez} mm²" if tipo and form and sez else ""
        linee_list.append({
            "Linea": r.get("Circuito/Linea",""),
            "Uso": r.get("Destinazione/Utilizzo",""),
            "Posa": r.get("Posa",""),
            "L_m": r.get("Lunghezza_m",""),
            "Cavo": cavo_str,
            "Protezione": r.get("Protezione (MT/MTD)",""),
            "Diff": r.get("Differenziale (tipo/Idn)",""),
            "DV_perc": f"{r.get('ΔV_%','')}",
            "Esito": r.get("Esito",""),
        })


    # === CAPITOLO 3 - CRITERIO DI PROGETTO (ESTESO) ===
    criterio_testo = ""
    if includi_criterio:
        # Tipi cavo usati nelle linee (se presenti)
        try:
            tipi_cavo_usati = ", ".join(sorted(set([str(x) for x in linee_df_calc.get("Tipo_cavo", []).dropna().tolist()])))
        except Exception:
            tipi_cavo_usati = "N.D."

        criterio_testo = (
f"""Tutti i materiali e le apparecchiature utilizzati devono essere di alta qualità, prodotti da aziende affidabili, ben lavorati e adatti all'uso previsto, resistendo a sollecitazioni meccaniche, corrosione, calore, umidità e acque meteoriche (per installazione all’esterno). Devono garantire lunga durata, facilità di ispezione e manutenzione.
È obbligatorio l'uso di componenti con marcatura CE e, se disponibile, marchio IMQ o equivalente europeo. I componenti senza marcatura CE devono avere una dichiarazione di conformità del costruttore ai requisiti di sicurezza delle normative CEI, UNI o IEC.

3.1 Dimensionamento delle linee
Le linee elettriche sono calcolate mediante l’utilizzo dei seguenti criteri progettuali:
• La corrente di impiego (Ib) è calcolata considerando la potenza nominale delle apparecchiature elettriche. La tensione di alimentazione è pari a 230 V per le utenze monofase, 400 V per le utenze trifase. Fattore di potenza pari a {cosphi_ricarica:.2f} per le linee di alimentazione delle prese di ricarica (se presenti).
• La corrente nominale della protezione (In), definita dal costruttore, è considerata come la corrente che l’interruttore può sopportare per un tempo indefinito senza che quest’ultimo subisca alcun danno.
• La portata del cavo (Iz) è calcolata utilizzando le tabelle CEI UNEL 35024 e 35026, tenendo conto delle condizioni di posa, del tipo di isolante del cavo e della temperatura ambiente.
I cavi di alimentazione sono dimensionati in modo da non subire danneggiamento causato da sovraccarichi e cortocircuiti mediante il coordinamento con la corrente nominale (In) del dispositivo di protezione a monte (vedi paragrafi 3.5.1 e 3.5.2).

3.2 Calcolo della sezione del cavo in funzione della corrente di impiego (Ib)
Nota la potenza assorbita dall’utenza, la corrente d’impiego (Ib) può essere calcolata come:
Ib = (Ku · P) / (k · Vn · cosφ)

dove:
• k = 1 per i circuiti monofase; k = √3 per i circuiti trifase;
• Ku è il coefficiente di utilizzazione della potenza nominale del carico;
• P è la potenza totale dell’utenza [W];
• Vn è la tensione nominale del sistema [V].
Determinata la corrente di impiego per ogni utenza, è possibile dimensionare il cavo con portata Iz > Ib.

3.3 Caduta di tensione
Dopo aver determinato la sezione del cavo in funzione della corrente d’impiego, si verifica la caduta di tensione con la formula:
ΔV = K · (R·cosφ + X·sinφ) · L · I

dove:
• K = 2 per le linee monofase (230 V); K = √3 per le linee trifase (400 V);
• R e X sono resistenza e reattanza per unità di lunghezza [Ω/km];
• I è la corrente di impiego;
• L è la lunghezza della linea [m].
La caduta di tensione percentuale è:
ΔV% = (ΔV / Vn) · 100
La caduta di tensione percentuale complessiva non deve superare {dv_lim:.1f}% (rif. CEI 64-8 art. 525).

3.4 Sezione e tipologia dei cavi utilizzati
I cavi utilizzati sono conformi al Regolamento UE 305/2011 (CPR), all’unificazione UNEL e alle norme costruttive CEI.
Per il dimensionamento dei conduttori di neutro e del conduttore di protezione (PE) si fa riferimento alla CEI 64-8/5 par. 543.1.2 tabella 54F:
• per sezione fase Sf ≤ 16 mm²: SPE = Sf
• per 16 < Sf ≤ 35 mm²: SPE = 16 mm²
• per Sf > 35 mm²: SPE = Sf/2
Qualora il PE non faccia parte della conduttura di alimentazione (CEI 64-8/5 par. 543.1.3), valgono i criteri sopra con minimi: 2,5 mm² Cu (con protezione meccanica) o 4 mm² Cu (senza protezione meccanica).

3.4.1 Tipologia dei cavi
I cavi impiegati nel progetto (in funzione delle tratte e delle modalità di posa) appartengono alle tipologie selezionate nei circuiti: {tipi_cavo_usati}.
Esempi (se pertinenti):
• FG16(M)16 / FG16(O)M16 (o similari) per dorsali/esterni Uo/U 0,6/1 kV (HEPR G16 + guaina R16) – CEI UNEL 35318/35322.
• FS17 450/750 V per cablaggi interni quadro e PE (unipolare senza guaina, PVC S17, CPR).

3.4.2 Posa dei cavi
Le tipologie di posa sono indicate nella tabella circuiti (campo “Posa”) e possono comprendere: tubazioni incassate/esterne, canalizzazioni, passerelle, tubazioni interrate, ecc. Gli attraversamenti di pareti/solai saranno ripristinati, ove necessario, con sigillature idonee a mantenere la compartimentazione. 
Nei punti in cui le condutture e/o le tubazioni impiantistiche attraversano elementi di separazione resistenti al fuoco (pareti e solai di compartimentazione), dovrà essere garantito il mantenimento della prestazione di compartimentazione prevista dal progetto antincendio. In conformità ai principi del Codice di Prevenzione Incendi (D.M. 03/08/2015 e s.m.i.) e alle norme di prova e classificazione della resistenza al fuoco dei sistemi di attraversamento, tutti i fori e i passaggi dovranno essere ripristinati mediante sistemi di sigillatura certificati (firestop) con classificazione almeno pari a quella dell’elemento attraversato (es. EI/REI richiesto), installati secondo le istruzioni del produttore.
A titolo esemplificativo, per tubazioni combustibili (PVC, PE, PP – tipicamente scarichi e pluviali) si impiegheranno collari tagliafuoco/REI con materiale termoespandente (intumescente) che, in caso d’incendio, occlude il foro sigillando il passaggio; per cavidotti/cavi e canalizzazioni si utilizzeranno idonei sistemi (malte o sigillanti intumescenti, schiume certificate, bende/manicotti, pannelli o cuscini) compatibili con il tipo di impianto e con le condizioni di posa.
I prodotti utilizzati dovranno essere marcati CE ove applicabile ai sensi del Regolamento (UE) 305/2011 (CPR) oppure corredati da Valutazione Tecnica Europea (ETA) e Dichiarazione di Prestazione (DoP), con rapporti di prova secondo UNI EN 1366-3 (sigillature di attraversamenti) e classificazione secondo UNI EN 13501-2. L’impresa incaricata dell’esecuzione degli attraversamenti e del ripristino dovrà impiegare materiali idonei e certificati, assicurare la continuità della tenuta ai fumi e ai gas caldi e rilasciare idonea documentazione di posa (schede prodotto, istruzioni e, ove richiesto, dichiarazione di corretta installazione) a garanzia del mantenimento della compartimentazione di progetto

Nei punti in cui le condutture e/o le tubazioni impiantistiche attraversano elementi di separazione resistenti al fuoco (pareti e solai di compartimentazione), dovrà essere garantito il mantenimento della prestazione di compartimentazione prevista dal progetto antincendio. In conformità ai principi del Codice di Prevenzione Incendi (D.M. 03/08/2015 e s.m.i.) e alle norme di prova e classificazione della resistenza al fuoco dei sistemi di attraversamento, tutti i fori e i passaggi dovranno essere ripristinati mediante sistemi di sigillatura certificati (firestop) con classificazione almeno pari a quella dell’elemento attraversato (es. EI/REI richiesto), installati secondo le istruzioni del produttore.
I prodotti utilizzati dovranno essere marcati CE ove applicabile ai sensi del Regolamento (UE) 305/2011 (CPR) oppure corredati da Valutazione Tecnica Europea (ETA) e Dichiarazione di Prestazione (DoP), con rapporti di prova secondo UNI EN 1366-3 e classificazione secondo UNI EN 13501-2.
Documentazione minima obbligatoria (a tutela della compartimentazione): l’Impresa incaricata dell’esecuzione degli attraversamenti e del ripristino dovrà consegnare, per ciascun attraversamento, registro attraversamenti (identificativo, ubicazione, elemento attraversato, prestazione richiesta, sistema adottato), schede prodotto/ETA/DoP, istruzioni di posa, e dichiarazione di corretta installazione, corredando il tutto con documentazione fotografica prima/dopo.
In mancanza della suddetta documentazione, la verifica della conformità delle sigillature firestop non si intende effettuata dal Progettista/Tecnico redattore e resta in capo all’Impresa e alla Direzione Lavori/Committente secondo le rispettive competenze.
Nota: Il presente documento non costituisce progetto antincendio né asseverazione ai fini della prevenzione incendi; i requisiti EI/REI e le soluzioni di compartimentazione sono quelli definiti dal progetto antincendio e dalle relative certificazioni.

3.4.3 Colorazione dei conduttori
I conduttori sono identificati secondo CEI-UNEL 00722 e 00712:
• PE: giallo/verde; • Neutro: blu; • Fasi: marrone/grigio/nero.

3.5 Protezioni dalle sovracorrenti
La protezione dalle sovracorrenti è assicurata da interruttori automatici magnetotermici dimensionati affinché le curve I–t si mantengano al di sotto delle curve dei cavi protetti. Gli interruttori devono:
• interrompere sovraccarichi e cortocircuiti prima di danni all’isolamento;
• essere installati all’origine di ogni circuito/derivazione con portate differenti;
• avere PdI/Icu > Icc presunta nel punto di installazione.

3.5.1 Sovraccarichi (CEI 64-8 art. 433.2)
Ib ≤ In ≤ Iz
If ≤ 1,45 · Iz

3.5.2 Cortocircuiti (CEI 64-8 art. 434.3)
I² · t ≤ K² · S²

3.6 Protezione dai contatti indiretti
La protezione contro i contatti indiretti è realizzata mediante interruzione automatica dell’alimentazione (TT/TN) e/o componenti a doppio isolamento.

3.6.1 Sistema TT (CEI 64-8 art. 413.1.4.2)
Idn ≤ UL / Rt
con UL = {ul_tt:.0f} V (ambiente ordinario) e Rt resistenza complessiva terra+conduttori di protezione.

3.7 Protezione dai contatti diretti (CEI 64-8 art. 412)
Isolamento delle parti attive e/o involucri/barriere (minimo IPXXB; superfici orizzontali a portata di mano: IPXXD). Vernici/smalti da soli non sono idonei.

3.8 Potere di interruzione delle apparecchiature
Icc-max < PdI (Icu) del dispositivo di protezione (rif. CEI EN 60947-2).

3.9 Quadri elettrici
Quadri conformi a CEI EN 61439-1/2 (e/o CEI 23-51 per domestici/similari). Cablaggio interno con conduttori idonei (es. FS17 CPR) e dimensionato per corrente nominale e cortocircuito nel punto di installazione.

{("Integrazioni: " + criterio_note) if criterio_note.strip() else ""}"""
        )

        st.markdown("### Struttura relazione (template integrato)")
        stile_relazione = st.radio(
            "Seleziona lo stile del documento",
            ["Lean (essenziale)", "Completo (come template)"],
            index=1,
            help="Lean: include solo i paragrafi che servono quasi sempre. Completo: replica la struttura del template con tutti i paragrafi."
        )
        # (template UI moved outside the button)

    payload = {
        "template_sections": template_sections,
        "revisioni": rev_df.to_dict(orient="records"),
        "committente_nome": committente,
        "impianto_indirizzo": luogo,
        "data": data_doc.strftime("%d/%m/%Y"),
        "data_documento": data_doc.strftime("%d/%m/%Y"),
        "header_titolo": "Relazione Tecnica - Impianto Elettrico (DiCo)",
        "progettista_blocco": progettista_blocco,
        "progettista_nome": progettista_nome,
        "premessa": premessa,
        "norme": norme,
        # Sezioni estese integrate (Cap. 7-8) – stampate solo se compilate
        "caratteristiche_tecniche_items": caratteristiche_tecniche_items,
        "caratteristiche_tecniche_note": caratteristiche_tecniche_note,
        "aspetti_normativi_items": aspetti_normativi_items,
        "aspetti_normativi_note": aspetti_normativi_note,
        "criterio_progetto": criterio_testo,
        "dati_progettazione": dati_progettazione_txt,
        "dati_tecnici": dati_tecnici,
        "descrizione_impianto": descrizione_impianto,
        "confini": confini_txt,
        "quadri": quadri_list,
        "linee": linee_list,
        "evse": evse_df.to_dict(orient="records"),
        "localizzazione": localizzazione,
        "layout": layout,
        "foto": [
            {"name": f.name, "bytes": f.getvalue()} for f in (foto_files or [])
            if hasattr(f, "getvalue")
        ],
        "checklist_documentale": checklist_df.to_dict(orient="records"),
        "verifiche_tabella": ver_df.to_dict(orient="records"),
        "sicurezza": sicurezza,
        "verifiche": verifiche,
        "manutenzione": manutenzione,
        "allegati": allegati,
        "disclaimer_calcoli": "Calcoli e verifiche riportati sono di sintesi e a supporto documentale. Non sostituiscono un progetto esecutivo completo né le verifiche previste dalle norme applicabili.",
        "titolo_cover": "RELAZIONE TECNICA - IMPIANTO ELETTRICO (DiCo)",
        "sottotitolo_cover": oggetto,
        "nome_progetto": nome_progetto,
        "cover_style": ("engineering" if cover_style.startswith("Engineering") else "legacy"),
        "firma": firmatario,
        "timbro_bytes": timbro_bytes,
        "oggetto_intervento": oggetto,
        "tipologia": tipologia,
        "sistema": sistema,
        "tensione": tensione,
        "potenza_disp": potenza_disp_kw,
        "cod_progetto": cod_progetto,
        "n_doc": n_doc,
        "n_documento": n_doc,
        "rev": revisione,
        "revisione": revisione,
        "impresa": impresa,
        "impresa_sede": impresa_sede,
        "impresa_piva": impresa_piva,
        "impresa_rea": impresa_rea,
        "impresa_resp": impresa_resp,
        "impresa_cont": impresa_cont,
        "luogo_firma": luogo_firma,
        "data_firma": data_firma.strftime("%d/%m/%Y"),
    }

    # --- Pulizia: se un campo non è stato compilato non deve comparire nel PDF (niente "XXXX"/placeholder)
    for k, v in list(payload.items()):
        if isinstance(v, str) and not _meaningful(v):
            payload[k] = ""

    # Tabelle/elenchi: rimuove righe non compilate
    payload["evse"] = _clean_records(payload.get("evse"), require_keys=["Potenza (kW)", "Marca", "Modello"])
    payload["checklist_documentale"] = _clean_records(payload.get("checklist_documentale"), require_keys=["Stato"])
    payload["verifiche_tabella"] = _clean_records(payload.get("verifiche_tabella"), require_keys=["Esito", "Note"])


    pdf_bytes = genera_pdf_relazione_bytes(payload)
    st.success("PDF generato.")
    st.download_button(
        "Scarica PDF",
        data=pdf_bytes,
        file_name="Relazione_Tecnica_DiCo_Impianto_Elettrico.pdf",
        mime="application/pdf",
    )
