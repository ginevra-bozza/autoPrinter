# Automatic Printer 🖨️
Ciao 🙂
Questo programma serve a stampare automaticamente i documenti e le foto inviandoli via email

### Scaricare il programma
Vai su `<> Code` (in alto a destra in verde) > _Local_ > _Download ZIP_

Dopo averlo estratto, nella stessa cartella devi aggiungere il file `.env` che contiene le credenziali della mail

### Avviare il programma
Basta che lo avvii quando accendi il computer e poi lo lasci aperto - fai doppio clic sul file chiamato `startPrinter.bat`
Si aprirà una finestra nera con delle scritte. _IMPORTANTE_: Non chiudere questa finestra - finché la finestra è aperta, la stampante è pronta a ricevere

### Per mandare un documento da stampare

- Scrivi una email all'indirizzo: `stampante.bozza1@gmail.com`
- Allega i file che vuoi stampare (puoi lasciare l'oggetto e il testo dell'email vuoti)

### Cosa fa il programma 

- Se mandi dei PDF o Documenti -> li manda direttamente alla stampante
- Se mandi delle Foto (JPG, PNG) -> le prende, le ingrandisce per farle stare su un foglio A4 senza tagliarle, le unisce e le stampa
- (In teoria, chi ha mandato l'email riceverà una risposta automatica di conferma ma forse non funziona)
- Il programma stampa solo documenti "sicuri" (PDF, Word, TXT) e immagini. Se per riceve file .zip o programmi li ignora automaticamente e li cancella senza aprirli

### Risoluzione dei problemi 
Se non viene stampato niente
- controlla di avere il file .env nella stessa cartella in cui ci sono gli altri script
- la stampante che vuoi usare deve essere quella impostata come predefinita da Windows
- se il tuo lettore di PDF predefinito è Microsoft Edge, può essere che il programma non funzioni - in quel caso conviene scaricare Adobe e impostare quello come predefinito (tutti i file, anche immagini, vengono convertiti in PDF per la stampa quindi passano da lì)

Se ci sono errori quando fai partire il file `.bat`: 

potrebbero essere le dependencies di Python - in teoria vengono installate in automatico, ma nel caso non funzionasse puoi installarle a mano.

Dal terminal di Windows (Powershell / CMD), incolla in ordine:

`winget install -e --id Python.Python.3.12`

`pip install imap-tools pillow img2pdf pywin32 python-dotenv`
