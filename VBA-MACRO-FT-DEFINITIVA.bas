' ============================================================
' MACRO VBA DEFINITIVA - FANTA TOSTI 2026
' Aggiornamento completo DB: Listone + Asta Riparazione + Allineamento Date + Assicurazioni + Fix Formule DB + Contratti Invernali
' Mercato Invernale 2026
' ============================================================
' ISTRUZIONI:
' 1. BACKUP del file DB prima di procedere!
' 2. Aprire "Fanta Tosti 2026 - DB completo (06.02.2026).xlsx"
'    Password: 89y3R8HF'(()h7t87gH)(/0?9U38Qyp99
' 3. Alt+F11 > Inserisci > Modulo > Incolla tutto questo codice
' 4. F5 > Seleziona "ESEGUI_TUTTO_FT" > Esegui
' 5. Quando richiesto, seleziona il file listone (Quotazioni_Fantacalcio_...)
' ============================================================

' Colonne Calciatore (1-based VBA) per ogni squadra FT:
' FCK Deportivo=3, Hellas=15, muttley=27, PARTIZAN=39,
' Legenda=51, KFP=63, Millwall=75, CKC=87, Papaie=99, Tronzano=111

Private Const DATA_ASS As String = "14/02/2026"
Private logText As String

' ============================================================
' MAIN: Esegue tutte le operazioni in sequenza
' ============================================================
Sub ESEGUI_TUTTO_FT()
    logText = "=== LOG OPERAZIONI FT - MACRO DEFINITIVA ===" & vbCrLf & vbCrLf

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' FASE 0: Aggiorna LISTA dal listone
    Log "FASE 0: Aggiornamento LISTA dal listone"
    Log "-------------------------------------------"
    Dim wsLista As Worksheet
    Set wsLista = ThisWorkbook.Sheets("LISTA")
    AggiornaListone wsLista, False  ' False = Classic (FT)

    ' FASE 1: Aggiungi giocatori asta riparazione (post 06/02)
    Dim wsSq As Worksheet
    Set wsSq = ThisWorkbook.Sheets("SQUADRE")
    Log ""
    Log "FASE 1: Inserimento giocatori asta riparazione"
    Log "-------------------------------------------"
    InserisciAstaRiparazione wsSq

    ' FASE 2: Allineamento retroattivo date assicurazione (regola triennio rigido)
    Log ""
    Log "FASE 2: Allineamento date assicurazione al ciclo triennale"
    Log "-------------------------------------------"
    AllineaDatePreventive wsSq

    ' FASE 3: Registra assicurazioni
    Log ""
    Log "FASE 3: Registrazione assicurazioni"
    Log "-------------------------------------------"
    RegistraAssicurazioni wsSq

    ' FASE 4: Correggi formule foglio DB (audit 22/02/2026)
    Log ""
    Log "FASE 4: Correzione formule foglio DB"
    Log "-------------------------------------------"
    CorreggiFormuleDB

    ' FASE 5: Calcola e annota quote contratti mercato di riparazione
    Log ""
    Log "FASE 5: Calcolo quote contratti invernali (QUOTE+MONTEPREMI 2026)"
    Log "-------------------------------------------"
    CalcolaContrattiInvernali wsLista

    ' Ricalcola
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Log completamento
    Log ""
    Log "=== COMPLETATO ==="

    ' Crea foglio log
    Dim wsLog As Worksheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("LOG_MACRO")
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = "LOG_MACRO"
    End If
    On Error GoTo 0
    wsLog.Cells.Clear
    wsLog.Cells(1, 1).Value = logText

    MsgBox "Operazioni FT completate!" & vbCrLf & _
           "Controlla il foglio LOG_MACRO per i dettagli.", _
           vbInformation, "Fanta Tosti 2026"
End Sub

' ============================================================
' FASE 0: Aggiorna foglio LISTA dal listone ufficiale
' Legge il file listone (Tutti + Ceduti), match per ID,
' aggiorna/aggiunge giocatori, riordina alfabeticamente.
'
' isMantra = False per FT (Classic), True per FM (Mantra)
' ============================================================
Private Sub AggiornaListone(wsLista As Worksheet, isMantra As Boolean)
    ' 1. Chiedi all'utente di selezionare il file listone
    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
        "File Excel (*.xlsx),*.xlsx", , _
        "Seleziona il file Listone (Quotazioni Fantacalcio)")
    If filePath = False Then
        Log "  ANNULLATO: Nessun file listone selezionato. LISTA non aggiornata."
        Exit Sub
    End If
    Log "  File selezionato: " & CStr(filePath)

    ' 2. Apri il workbook del listone in sola lettura
    Dim wbListone As Workbook
    On Error Resume Next
    Set wbListone = Workbooks.Open(CStr(filePath), ReadOnly:=True)
    On Error GoTo 0
    If wbListone Is Nothing Then
        Log "  ERRORE: Impossibile aprire il file listone."
        Exit Sub
    End If

    ' 3. Leggi dati dal listone (fogli "Tutti" e "Ceduti")
    Dim aggiornati As Long, aggiunti As Long, skippati As Long
    aggiornati = 0: aggiunti = 0: skippati = 0

    ' Trova l'ultima riga con dati nella LISTA
    Dim lastListaRow As Long
    lastListaRow = 1
    Dim scanR As Long
    For scanR = 2 To 2000
        If Trim(CStr(wsLista.Cells(scanR, 1).Value)) <> "" Then
            lastListaRow = scanR
        End If
    Next scanR

    ' Processa foglio "Tutti" (giocatori attivi)
    Dim wsTutti As Worksheet
    On Error Resume Next
    Set wsTutti = wbListone.Sheets("Tutti")
    On Error GoTo 0
    If Not wsTutti Is Nothing Then
        Log "  Elaborazione foglio 'Tutti'..."
        ProcessaFoglioListone wsTutti, wsLista, isMantra, lastListaRow, aggiornati, aggiunti, skippati
    Else
        Log "  ATTENZIONE: Foglio 'Tutti' non trovato nel listone!"
    End If

    ' Processa foglio "Ceduti" (calciatori ceduti)
    Dim wsCeduti As Worksheet
    On Error Resume Next
    Set wsCeduti = wbListone.Sheets("Ceduti")
    On Error GoTo 0
    If Not wsCeduti Is Nothing Then
        Log "  Elaborazione foglio 'Ceduti'..."
        ProcessaFoglioListone wsCeduti, wsLista, isMantra, lastListaRow, aggiornati, aggiunti, skippati
    Else
        Log "  ATTENZIONE: Foglio 'Ceduti' non trovato nel listone."
    End If

    ' 4. Chiudi il listone senza salvare
    wbListone.Close SaveChanges:=False

    ' 5. Imposta formula Eta' per le nuove righe
    Log "  Aggiornamento formula Eta'..."
    Dim etaR As Long
    For etaR = 2 To lastListaRow
        If Trim(CStr(wsLista.Cells(etaR, 2).Value)) <> "" Then
            ' Solo se la cella Eta' e' vuota o e' un numero (non formula)
            If IsEmpty(wsLista.Cells(etaR, 9).Value) Or wsLista.Cells(etaR, 9).Value = "" Then
                wsLista.Cells(etaR, 9).Formula = _
                    "=IFERROR(VLOOKUP(B" & etaR & ",$L:$N,3,FALSE),"""")"
            End If
        End If
    Next etaR

    ' 6. Ordina LISTA per Calciatore (col B) A->Z
    Log "  Ordinamento LISTA per nome calciatore..."
    If lastListaRow > 2 Then
        wsLista.Sort.SortFields.Clear
        wsLista.Sort.SortFields.Add2 _
            Key:=wsLista.Range("B2:B" & lastListaRow), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        With wsLista.Sort
            .SetRange wsLista.Range("A1:I" & lastListaRow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If

    Log "  LISTA aggiornata: " & aggiornati & " aggiornati, " & aggiunti & " aggiunti, " & skippati & " skippati"
End Sub

' ============================================================
' Helper: Processa un singolo foglio del listone (Tutti o Ceduti)
' Legge ogni riga, match per ID, aggiorna o aggiunge nella LISTA
' ============================================================
Private Sub ProcessaFoglioListone(wsSource As Worksheet, wsLista As Worksheet, _
    isMantra As Boolean, ByRef lastListaRow As Long, _
    ByRef aggiornati As Long, ByRef aggiunti As Long, ByRef skippati As Long)

    ' Il listone ha:
    ' Riga 1: titolo ("Quotazioni Fantacalcio Stagione...")
    ' Riga 2: headers (Id, R, RM, Nome, Squadra, Qt.A, Qt.I, Diff., Qt.A M, Qt.I M, Diff.M, FVM, FVM M)
    ' Riga 3+: dati
    '
    ' Colonne listone (1-based VBA):
    '   A(1)=Id, B(2)=R, C(3)=RM, D(4)=Nome, E(5)=Squadra
    '   F(6)=Qt.A, G(7)=Qt.I, H(8)=Diff.
    '   I(9)=Qt.A M, J(10)=Qt.I M, K(11)=Diff.M
    '   L(12)=FVM, M(13)=FVM M

    Dim r As Long
    r = 3 ' Prima riga dati nel listone

    Do While Trim(CStr(wsSource.Cells(r, 1).Value)) <> ""
        Dim listoneId As Variant
        listoneId = wsSource.Cells(r, 1).Value

        ' Leggi nome e rimuovi i punti
        Dim nomeRaw As String
        nomeRaw = Trim(CStr(wsSource.Cells(r, 4).Value))
        Dim nomeClean As String
        nomeClean = Replace(nomeRaw, ".", "")
        ' Rimuovi spazi finali dopo il punto (es. "Martinez L. " -> "Martinez L")
        nomeClean = Trim(nomeClean)

        ' Leggi altri dati dal listone
        Dim ruoloC As String, ruoloM As String, squadra As String
        ruoloC = Trim(CStr(wsSource.Cells(r, 2).Value))
        ruoloM = Trim(CStr(wsSource.Cells(r, 3).Value))
        squadra = Trim(CStr(wsSource.Cells(r, 5).Value))

        ' Quotazioni e FVM in base al tipo di lega
        Dim qtAttuale As Variant, qtIniziale As Variant, fvmVal As Variant
        If isMantra Then
            qtAttuale = wsSource.Cells(r, 9).Value   ' Qt.A M
            qtIniziale = wsSource.Cells(r, 10).Value  ' Qt.I M
            fvmVal = wsSource.Cells(r, 13).Value      ' FVM M
        Else
            qtAttuale = wsSource.Cells(r, 6).Value    ' Qt.A
            qtIniziale = wsSource.Cells(r, 7).Value   ' Qt.I
            fvmVal = wsSource.Cells(r, 12).Value      ' FVM
        End If

        ' Skip se nome vuoto
        If Len(nomeClean) = 0 Then
            skippati = skippati + 1
            GoTo NextRow
        End If

        ' Cerca per ID nella LISTA (col A)
        Dim matchRow As Variant
        matchRow = Application.Match(CLng(listoneId), wsLista.Range("A:A"), 0)

        If Not IsError(matchRow) Then
            ' TROVATO: aggiorna la riga esistente
            Dim mRow As Long
            mRow = CLng(matchRow)
            wsLista.Cells(mRow, 2).Value = nomeClean    ' Calciatore
            wsLista.Cells(mRow, 3).Value = ruoloC        ' Ruolo
            wsLista.Cells(mRow, 4).Value = ruoloM        ' R. Mantra
            wsLista.Cells(mRow, 5).Value = squadra        ' Squadra
            wsLista.Cells(mRow, 6).Value = qtAttuale      ' Q. attuale
            wsLista.Cells(mRow, 7).Value = qtIniziale     ' Q. iniziale
            wsLista.Cells(mRow, 8).Value = fvmVal         ' FVM
            aggiornati = aggiornati + 1
        Else
            ' NON TROVATO: aggiungi come nuova riga
            lastListaRow = lastListaRow + 1
            wsLista.Cells(lastListaRow, 1).Value = CLng(listoneId)  ' ID
            wsLista.Cells(lastListaRow, 2).Value = nomeClean         ' Calciatore
            wsLista.Cells(lastListaRow, 3).Value = ruoloC             ' Ruolo
            wsLista.Cells(lastListaRow, 4).Value = ruoloM             ' R. Mantra
            wsLista.Cells(lastListaRow, 5).Value = squadra             ' Squadra
            wsLista.Cells(lastListaRow, 6).Value = qtAttuale           ' Q. attuale
            wsLista.Cells(lastListaRow, 7).Value = qtIniziale          ' Q. iniziale
            wsLista.Cells(lastListaRow, 8).Value = fvmVal              ' FVM
            aggiunti = aggiunti + 1
        End If

NextRow:
        r = r + 1
    Loop
End Sub

' ============================================================
' FASE 1: Inserisci giocatori acquisiti nell'asta riparazione
' ============================================================
Private Sub InserisciAstaRiparazione(ws As Worksheet)
    ' --- HELLAS MADONNA (col 15) ---
    InserisciGiocatore ws, 15, "Circati", "Par", 1
    InserisciGiocatore ws, 15, "Berisha M", "Lec", 1
    InserisciGiocatore ws, 15, "Durosinmi", "Ven", 3

    ' --- PARTIZAN (col 39) ---
    InserisciGiocatore ws, 39, "Belghali", "Ver", 7
    InserisciGiocatore ws, 39, "Strefezza", "Par", 29
    InserisciGiocatore ws, 39, "Przyborek", "Com", 1

    ' --- KUNG FU PANDEV (col 63) ---
    InserisciGiocatore ws, 63, "Malen", "Ata", 173
    InserisciGiocatore ws, 63, "Vergara", "Laz", 31
    ' NOTA: Kouame' NON inserito - non piu' listato, non assicurabile

    ' --- FC CKC 26 (col 87) ---
    InserisciGiocatore ws, 87, "Tiago Gabriel", "Pis", 1
    InserisciGiocatore ws, 87, "Vaz", "Par", 3
    InserisciGiocatore ws, 87, "Muharemovic", "Ven", 3
    InserisciGiocatore ws, 87, "Baldanzi", "Rom", 3
    InserisciGiocatore ws, 87, "Santos A", "Nap", 1
    InserisciGiocatore ws, 87, "Bijlow", "Fio", 2
    InserisciGiocatore ws, 87, "Bernasconi", "Ata", 1

    ' --- MUTTLEY SUPERSTAR (col 27) ---
    InserisciGiocatore ws, 27, "Ostigard", "Gen", 2
    InserisciGiocatore ws, 27, "Luis Henrique", "Int", 17
    InserisciGiocatore ws, 27, "Solomon", "Tor", 21

    ' --- MILLWALL (col 75) ---
    InserisciGiocatore ws, 75, "Celik", "Rom", 2
    InserisciGiocatore ws, 75, "Ratkov", "Laz", 2
    InserisciGiocatore ws, 75, "Zaragoza", "Mon", 46
    InserisciGiocatore ws, 75, "Perrone", "Mon", 9
    InserisciGiocatore ws, 75, "Paleari", "Cag", 1
    InserisciGiocatore ws, 75, "Boga", "Niz", 1

    ' --- LEGENDA AUREA (col 51) ---
    InserisciGiocatore ws, 51, "Bartesaghi", "Mil", 3
    InserisciGiocatore ws, 51, "Taylor K", "Laz", 42
    InserisciGiocatore ws, 51, "Fagioli", "Fio", 1
    InserisciGiocatore ws, 51, "Ekkelenkamp", "Gen", 4
    InserisciGiocatore ws, 51, "Miretti", "Gen", 1
    InserisciGiocatore ws, 51, "Bonazzoli", "Cre", 2
    InserisciGiocatore ws, 51, "Raspadori", "Ata", 50
End Sub

' ============================================================
' FASE 2: Registra tutte le assicurazioni
' ============================================================
Private Sub RegistraAssicurazioni(ws As Worksheet)
    ' --- HELLAS MADONNA (col 15) ---
    AssicuraG ws, 15, "Sportiello"
    AssicuraG ws, 15, "Circati"
    AssicuraG ws, 15, "Berisha"
    AssicuraG ws, 15, "Moreo"
    AssicuraG ws, 15, "Durosinmi"

    ' --- PARTIZAN (col 39) ---
    AssicuraG ws, 39, "Belghali"
    AssicuraG ws, 39, "Strefezza"
    AssicuraG ws, 39, "Przyborek"

    ' --- KFP (col 63) ---
    AssicuraG ws, 63, "Malen"
    AssicuraG ws, 63, "Vergara"
    AssicuraG ws, 63, "Beukema"
    Log "  SKIP: Kouame' - non piu' listato su Leghe FC, NON assicurabile"

    ' --- FC CKC 26 (col 87) ---
    AssicuraG ws, 87, "Tiago Gabriel"
    AssicuraG ws, 87, "Vaz"
    AssicuraG ws, 87, "Muharemovic"
    AssicuraG ws, 87, "Baldanzi"
    AssicuraG ws, 87, "Santos"       ' = "Allison S." nella comunicazione
    AssicuraG ws, 87, "Bijlow"
    AssicuraG ws, 87, "Bernasconi"
    AssicuraG ws, 87, "Kon"          ' Kone' I.

    ' --- MUTTLEY SUPERSTAR (col 27) ---
    AssicuraG ws, 27, "Ostigard"
    AssicuraG ws, 27, "Luis"         ' Luis Henrique
    AssicuraG ws, 27, "Solomon"

    ' --- MILLWALL (col 75) ---
    AssicuraG ws, 75, "Muric"
    AssicuraG ws, 75, "Celik"
    AssicuraG ws, 75, "Ratkov"
    AssicuraG ws, 75, "Zaragoza"
    AssicuraG ws, 75, "Perrone"
    AssicuraG ws, 75, "Paleari"
    AssicuraG ws, 75, "Boga"
    AssicuraG ws, 75, "Holm"

    ' --- PAPAIE TOP TEAM (col 99) ---
    AssicuraG ws, 99, "Hien"

    ' --- LEGENDA AUREA (col 51) ---
    AssicuraG ws, 51, "Di Gregorio"
    AssicuraG ws, 51, "Sommer"
    AssicuraG ws, 51, "Martinez"
    AssicuraG ws, 51, "Kalulu"
    AssicuraG ws, 51, "Bartesaghi"
    AssicuraG ws, 51, "Lovric"
    AssicuraG ws, 51, "Taylor"
    AssicuraG ws, 51, "Fagioli"
    AssicuraG ws, 51, "Ekkelenkamp"
    AssicuraG ws, 51, "Miretti"
    AssicuraG ws, 51, "Bonazzoli"
    AssicuraG ws, 51, "Raspadori"
    AssicuraG ws, 51, "Vitinha"
End Sub

' ============================================================
' FASE 4: Correzione formule foglio DB (audit 22/02/2026)
' Corregge 3 bug nelle formule del foglio DB:
' 1. BV/BW/BX righe 3-12: usa MAX($J,$AO) invece di solo $AO
' 2. BY/BZ/CA: uniforma tutte le righe a $BF/$BG/$BH/$BI
' 3. BS Portieri: parentesizzazione + coefficiente 0.65 (non 0.75)
' + Rinomina header BP2
' ============================================================
Private Sub CorreggiFormuleDB()
    Dim wsDB As Worksheet
    Set wsDB = ThisWorkbook.Sheets("DB")

    ' Trova l'ultima riga con dati nel DB (colonna A = Ruolo)
    Dim lastRow As Long
    lastRow = 3
    Dim scanR As Long
    For scanR = 3 To 1500
        If Trim(CStr(wsDB.Cells(scanR, 1).Value)) <> "" Then
            lastRow = scanR
        End If
    Next scanR
    Log "  Ultima riga DB con dati: " & lastRow

    ' ----- BUG 1: BV/BW/BX righe 3-12 -----
    ' Le righe 3-12 usano $AO invece di MAX($J,$AO) nella formula INVARIATA
    ' La formula corretta e' presente a partire dalla riga 13
    Log "  BUG 1: BV/BW/BX righe 3-12 -> copia formula corretta da riga 13..."
    wsDB.Range("BV13:BX13").Copy
    wsDB.Range("BV3:BX12").PasteSpecial xlPasteFormulas
    Application.CutCopyMode = False
    Log "    Corretto: BV3:BX12 allineate a MAX($J,$AO)"

    ' ----- BUG 2: BY/BZ/CA — riferimenti colonne POSITIVA -----
    ' La maggior parte delle righe usa $BB/$BC/$BD/$BE (vecchie colonne)
    ' La formula corretta (riga 3) usa $BF/$BG/$BH/$BI (colonne attuali)
    Log "  BUG 2: BY/BZ/CA -> uniforma riferimenti a $BF/$BG/$BH/$BI..."
    wsDB.Range("BY3:CA3").Copy
    wsDB.Range("BY4:CA" & lastRow).PasteSpecial xlPasteFormulas
    Application.CutCopyMode = False
    Log "    Corretto: BY4:CA" & lastRow & " allineate a $BF/$BG/$BH/$BI"

    ' ----- BUG 3: BS — ramo Portieri -----
    ' Il ramo Portiere (A="P") in BS ha due errori:
    '   a) Parentesi sbagliata: MAX(J,AO) - (AL*VLOOKUP(...))*0.75
    '      Dovrebbe essere:    (MAX(J,AO) - (AL*VLOOKUP(...)))*0.65
    '   b) Coefficiente di reparto: 0.75 deve essere 0.65 per i Portieri
    Log "  BUG 3: BS Portieri -> fix parentesi + coefficiente 0.65..."

    Dim bsF As String
    bsF = wsDB.Range("BS3").Formula

    ' Verifica che il pattern errato sia presente nella formula
    If InStr(bsF, "FALSE))*0.75") > 0 Then
        ' Sostituzione 1: Aggiungi ( prima di MAX nel ramo Portiere
        ' Il pattern ="P",$X3<1),MAX( e' unico (solo il ramo P ha "P")
        bsF = Replace(bsF, _
            "=""P"",$X3<1),MAX(", _
            "=""P"",$X3<1),(MAX(")

        ' Sostituzione 2: Cambia FALSE))*0.75 in FALSE)))*0.65
        ' Aggiunge ) per chiudere il nuovo (, e corregge il coefficiente
        bsF = Replace(bsF, _
            "FALSE))*0.75", _
            "FALSE)))*0.65")

        ' Scrivi la formula corretta in BS3 e copia a tutte le righe
        wsDB.Range("BS3").Formula = bsF
        wsDB.Range("BS3").Copy
        wsDB.Range("BS4:BS" & lastRow).PasteSpecial xlPasteFormulas
        Application.CutCopyMode = False
        Log "    Corretto: BS3:BS" & lastRow & " - parentesi + coeff. 0.65 per Portieri"
    Else
        ' Controlla se la correzione e' gia' stata applicata
        If InStr(bsF, "FALSE)))*0.65") > 0 Then
            Log "    GIA' CORRETTO: BS contiene gia' la formula corretta"
        Else
            Log "    ATTENZIONE: Pattern BS non riconosciuto - struttura diversa dal previsto"
            Log "    Formula BS3: " & Left(bsF, 150) & IIf(Len(bsF) > 150, "...", "")
        End If
    End If

    ' ----- Fix header: BP2 -----
    ' Rinomina da "Valore se A e % positiva" a "Valore tabellare per ruolo"
    Log "  Fix header BP2..."
    Dim oldBP As String
    oldBP = CStr(wsDB.Range("BP2").Value)
    If oldBP <> "Valore tabellare per ruolo" Then
        wsDB.Range("BP2").Value = "Valore tabellare per ruolo"
        Log "    Rinominato: '" & oldBP & "' -> 'Valore tabellare per ruolo'"
    Else
        Log "    GIA' CORRETTO: BP2 gia' impostato a 'Valore tabellare per ruolo'"
    End If

    ' ----- Nota: BN ridondante -----
    ' BN (=IFNA(VLOOKUP($C,SQUADRE!...,6,FALSE),"")) e' un duplicato di AO.
    ' Nessuna formula nella cartella la referenzia. Non eliminata per non
    ' spostare le colonne successive, ma segnalata nel log.
    Log "  Nota: BN (col 66) ridondante (duplicato di AO) - nessun riferimento trovato"

    Log "  Correzione formule DB completata."
End Sub

' ============================================================
' Macro standalone per correzione formule DB (FASE 4)
' (eseguibile indipendentemente da ESEGUI_TUTTO_FT)
' ============================================================
Sub CorreggiFormuleDB_FT()
    logText = "=== LOG CORREZIONE FORMULE DB - FT ===" & vbCrLf & vbCrLf

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    CorreggiFormuleDB

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Output log
    Dim wsLog As Worksheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("LOG_MACRO")
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = "LOG_MACRO"
    End If
    On Error GoTo 0
    wsLog.Cells.Clear
    wsLog.Cells(1, 1).Value = logText

    MsgBox "Correzione formule DB completata!" & vbCrLf & _
           "Controlla il foglio LOG_MACRO per i dettagli.", _
           vbInformation, "Fix Formule DB - FT"
End Sub

' ============================================================
' Helper: Inserisci un giocatore nuovo in SQUADRE
' ============================================================
Private Sub InserisciGiocatore(ws As Worksheet, colCalc As Long, nome As String, squadraSerieA As String, spesa As Long)
    Dim r As Long
    For r = 6 To 50
        Dim existName As String
        existName = Trim(CStr(ws.Cells(r, colCalc).Value))
        If UCase(Replace(existName, "'", "")) Like "*" & UCase(Replace(Left(nome, 5), "'", "")) & "*" Then
            Log "  GIA' PRESENTE: " & nome & " (riga " & r & ", come '" & existName & "')"
            If ws.Cells(r, colCalc + 10).Value <> spesa Then
                ws.Cells(r, colCalc + 10).Value = spesa
                Log "    -> Spesa aggiornata a " & spesa
            End If
            Exit Sub
        End If
    Next r

    For r = 6 To 50
        If Trim(CStr(ws.Cells(r, colCalc).Value)) = "" Then
            ws.Cells(r, colCalc).Value = nome
            ws.Cells(r, colCalc + 1).Value = squadraSerieA
            ws.Cells(r, colCalc + 10).Value = spesa
            Log "  INSERITO: " & nome & " (" & squadraSerieA & ", Sp=" & spesa & ") -> riga " & r
            Exit Sub
        End If
    Next r

    Log "  ERRORE: Nessuna riga vuota per " & nome & " nella colonna " & colCalc
End Sub

' ============================================================
' Helper: Assicura un giocatore (cerca per nome, imposta A + data)
' Gestione date rinnovi preventivi:
'   - Se il giocatore e' gia' assicurato (rinnovo) e il triennio
'     non e' ancora scaduto, la nuova data = scadenza triennio
'   - Se triennio scaduto o prima assicurazione, data = DATA_ASS
' ============================================================
Private Function AssicuraG(ws As Worksheet, colCalc As Long, nomeCerca As String) As Boolean
    Dim r As Long
    Dim nc As String
    nc = UCase(Replace(Replace(nomeCerca, "'", ""), ".", ""))

    For r = 6 To 52
        Dim nomeCell As String
        nomeCell = UCase(Trim(CStr(ws.Cells(r, colCalc).Value)))
        nomeCell = Replace(nomeCell, "'", "")
        nomeCell = Replace(nomeCell, ".", "")
        nomeCell = Replace(nomeCell, Chr(232), "e")
        nomeCell = Replace(nomeCell, Chr(233), "e")
        nomeCell = Replace(nomeCell, Chr(242), "o")
        nomeCell = Replace(nomeCell, Chr(224), "a")
        nomeCell = Replace(nomeCell, Chr(249), "u")
        nomeCell = Replace(nomeCell, Chr(200), "E")
        nomeCell = Replace(nomeCell, Chr(201), "E")

        If Len(nomeCell) > 0 And InStr(1, nomeCell, nc, vbTextCompare) > 0 Then
            Dim vecchioFlag As String
            vecchioFlag = Trim(CStr(ws.Cells(r, colCalc + 3).Value))

            ' Determina la data di decorrenza corretta
            Dim dataAss As Date
            dataAss = CDate(DATA_ASS)  ' Default: 14/02/2026

            If vecchioFlag = "A" Then
                ' RINNOVO: verifica se triennio ancora valido
                Dim vecchiaData As Variant
                vecchiaData = ws.Cells(r, colCalc + 7).Value
                If IsDate(vecchiaData) Then
                    Dim scadenzaTriennio As Date
                    scadenzaTriennio = DateAdd("yyyy", 3, CDate(vecchiaData))
                    If scadenzaTriennio > dataAss Then
                        ' Triennio non scaduto: rinnovo preventivo
                        ' La nuova copertura decorre dalla scadenza del triennio corrente
                        dataAss = scadenzaTriennio
                        Log "  RINNOVO PREVENTIVO: " & ws.Cells(r, colCalc).Value & _
                            " (riga " & r & ") - nuova decorrenza " & Format(dataAss, "dd/mm/yyyy") & _
                            " (triennio dal " & Format(CDate(vecchiaData), "dd/mm/yyyy") & ")"
                    Else
                        ' Triennio scaduto: rinnovo con data standard
                        Log "  RINNOVO: " & ws.Cells(r, colCalc).Value & _
                            " (riga " & r & ") - triennio scaduto il " & Format(scadenzaTriennio, "dd/mm/yyyy") & _
                            ", nuova decorrenza " & Format(dataAss, "dd/mm/yyyy")
                    End If
                Else
                    Log "  RINNOVO: " & ws.Cells(r, colCalc).Value & _
                        " (riga " & r & ") - data precedente non valida, usa " & DATA_ASS
                End If
            Else
                Log "  ASSICURATO: " & ws.Cells(r, colCalc).Value & _
                    " (riga " & r & ") - prima assicurazione"
            End If

            ws.Cells(r, colCalc + 3).Value = "A"
            ws.Cells(r, colCalc + 7).Value = dataAss
            ws.Cells(r, colCalc + 7).NumberFormat = "dd/mm/yyyy"

            AssicuraG = True
            Exit Function
        End If
    Next r

    Log "  NON TROVATO: " & nomeCerca & " nella colonna " & colCalc
    AssicuraG = False
End Function

' ============================================================
' Helper: Log
' ============================================================
Private Sub Log(msg As String)
    logText = logText & msg & vbCrLf
    Debug.Print msg
End Sub

' ============================================================
' FASE 2: Allineamento retroattivo date assicurazione
' Regola triennio rigido: la decorrenza e' SEMPRE uno scatto
' triennale dalla Data acquisto, mai la data di comunicazione.
' Algoritmo: per ogni assicurato, cammina la catena
'   Dacq, Dacq+3, Dacq+6, ... e allinea al boundary piu' vicino.
' ============================================================
Private Sub AllineaDatePreventive(ws As Worksheet)
    Dim teamCols As Variant
    teamCols = Array( _
        Array("FCK Deportivo", 3), Array("Hellas Madonna", 15), _
        Array("muttley superstar", 27), Array("PARTIZAN", 39), _
        Array("Legenda Aurea", 51), Array("Kung Fu Pandev", 63), _
        Array("Millwall", 75), Array("FC CKC 26", 87), _
        Array("Papaie Top Team", 99), Array("Tronzano", 111))

    Dim allineati As Long, invariati As Long, noAnchor As Long
    allineati = 0: invariati = 0: noAnchor = 0

    Dim t As Long
    For t = LBound(teamCols) To UBound(teamCols)
        Dim tName As String, col As Long
        tName = teamCols(t)(0): col = teamCols(t)(1)

        Dim r As Long
        For r = 6 To 52
            Dim nome As String
            nome = Trim(CStr(ws.Cells(r, col).Value))
            If Len(nome) = 0 Or nome = "Calciatore" Then GoTo NextPlayerFT

            Dim flag As String
            flag = Trim(CStr(ws.Cells(r, col + 3).Value))
            If flag <> "A" Then GoTo NextPlayerFT

            Dim insDateVal As Variant
            insDateVal = ws.Cells(r, col + 7).Value
            If Not IsDate(insDateVal) Then GoTo NextPlayerFT
            Dim insDate As Date
            insDate = CDate(insDateVal)

            ' Data acquisto come ancora del ciclo triennale
            Dim acqDateVal As Variant
            acqDateVal = ws.Cells(r, col + 4).Value
            If Not IsDate(acqDateVal) Then
                noAnchor = noAnchor + 1
                GoTo NextPlayerFT
            End If
            Dim acqDate As Date
            acqDate = CDate(acqDateVal)

            ' Cammina la catena triennale: trova i due boundary che incorniciano insDate
            Dim prevBound As Date
            prevBound = acqDate
            Do While DateAdd("yyyy", 3, prevBound) <= insDate
                prevBound = DateAdd("yyyy", 3, prevBound)
            Loop
            Dim nextBound As Date
            nextBound = DateAdd("yyyy", 3, prevBound)

            ' Scegli il boundary piu' vicino
            Dim correctDate As Date
            If (insDate - prevBound) <= (nextBound - insDate) Then
                correctDate = prevBound
            Else
                correctDate = nextBound
            End If

            If correctDate <> insDate Then
                ws.Cells(r, col + 7).Value = correctDate
                ws.Cells(r, col + 7).NumberFormat = "dd/mm/yyyy"
                Log "    ALLINEATO: " & tName & " / " & nome & _
                    " - da " & Format(insDate, "dd/mm/yyyy") & _
                    " a " & Format(correctDate, "dd/mm/yyyy") & _
                    " (acquisto: " & Format(acqDate, "dd/mm/yyyy") & ")"
                allineati = allineati + 1
            Else
                invariati = invariati + 1
            End If
NextPlayerFT:
        Next r
    Next t

    Log "  Allineamento: " & allineati & " corretti, " & invariati & _
        " gia' allineati, " & noAnchor & " senza data acquisto (invariati)"
End Sub

' ============================================================
' FASE 5: Calcola quote contratti mercato di riparazione
' Per ogni squadra, somma Qt.Attuale * 0.05 dei giocatori
' acquisiti nell'asta riparazione invernale e scrive il totale
' nel foglio QUOTE+MONTEPREMI 2026 colonna I.
' ============================================================
Private Sub CalcolaContrattiInvernali(wsLista As Worksheet)
    Dim wsQM As Worksheet
    On Error Resume Next
    Set wsQM = ThisWorkbook.Sheets("QUOTE+MONTEPREMI 2026")
    On Error GoTo 0
    If wsQM Is Nothing Then
        Log "  ERRORE: Foglio 'QUOTE+MONTEPREMI 2026' non trovato!"
        Exit Sub
    End If

    ' Definizione acquisti invernali per squadra FT
    ' Array: (nomeSquadraQM, Array di nomi giocatori)
    Dim teamData(0 To 6) As Variant
    teamData(0) = Array("Hellas", Array("Circati", "Berisha", "Durosinmi"))
    teamData(1) = Array("PARTIZAN", Array("Belghali", "Strefezza", "Przyborek"))
    teamData(2) = Array("Kung Fu", Array("Malen", "Vergara"))
    teamData(3) = Array("CKC", Array("Tiago Gabriel", "Vaz", "Muharemovic", "Baldanzi", "Santos", "Bijlow", "Bernasconi"))
    teamData(4) = Array("muttley", Array("Ostigard", "Luis Henrique", "Solomon"))
    teamData(5) = Array("Millwall", Array("Celik", "Ratkov", "Zaragoza", "Perrone", "Paleari", "Boga"))
    teamData(6) = Array("Legenda", Array("Bartesaghi", "Taylor", "Fagioli", "Ekkelenkamp", "Miretti", "Bonazzoli", "Raspadori"))

    Dim totaleGenerale As Double
    totaleGenerale = 0

    Dim i As Long
    For i = 0 To UBound(teamData)
        Dim searchName As String
        searchName = teamData(i)(0)
        Dim players As Variant
        players = teamData(i)(1)

        Dim totaleSquadra As Double
        totaleSquadra = 0
        Dim dettaglio As String
        dettaglio = ""

        Dim p As Long
        For p = LBound(players) To UBound(players)
            Dim playerName As String
            playerName = CStr(players(p))
            Dim qtA As Double
            qtA = CercaQtAttualeInLista(wsLista, playerName)
            Dim costo As Double
            costo = qtA * 0.05
            totaleSquadra = totaleSquadra + costo
            If Len(dettaglio) > 0 Then dettaglio = dettaglio & ", "
            dettaglio = dettaglio & playerName & "(" & Format(qtA, "0") & "*0.05=" & Format(costo, "0.00") & ")"
        Next p

        ' Trova la riga della squadra nel foglio QUOTE+MONTEPREMI 2026
        Dim qmRow As Long
        qmRow = 0
        Dim sr As Long
        For sr = 2 To 15
            Dim cellA As String
            cellA = UCase(Trim(CStr(wsQM.Cells(sr, 1).Value)))
            If InStr(1, cellA, UCase(searchName), vbTextCompare) > 0 Then
                qmRow = sr
                Exit For
            End If
        Next sr

        If qmRow > 0 Then
            wsQM.Cells(qmRow, 9).Value = totaleSquadra
            wsQM.Cells(qmRow, 9).NumberFormat = "#,##0.00"
            Log "  " & searchName & " (riga " & qmRow & "): " & _
                Format(totaleSquadra, "0.00") & " EUR -> " & dettaglio
        Else
            Log "  ATTENZIONE: Squadra '" & searchName & "' non trovata in QUOTE+MONTEPREMI 2026"
            Log "    Quota calcolata: " & Format(totaleSquadra, "0.00") & " EUR"
        End If

        totaleGenerale = totaleGenerale + totaleSquadra
    Next i

    ' Aggiorna riga TOTALI (riga 12)
    Dim totRow As Long
    totRow = 0
    For sr = 2 To 15
        If UCase(Trim(CStr(wsQM.Cells(sr, 1).Value))) Like "*TOTAL*" Then
            totRow = sr
            Exit For
        End If
    Next sr
    If totRow > 0 Then
        wsQM.Cells(totRow, 9).Value = totaleGenerale
        wsQM.Cells(totRow, 9).NumberFormat = "#,##0.00"
        Log "  TOTALE: " & Format(totaleGenerale, "0.00") & " EUR (riga " & totRow & ")"
    End If

    Log "  Quote contratti invernali completate."
End Sub

' ============================================================
' Helper: Cerca Qt.Attuale di un giocatore nel foglio LISTA
' Ritorna il valore di col F (Q.attuale) o 0 se non trovato
' ============================================================
Private Function CercaQtAttualeInLista(wsLista As Worksheet, nomeCerca As String) As Double
    Dim nc As String
    nc = UCase(Replace(Replace(Trim(nomeCerca), "'", ""), ".", ""))

    Dim r As Long
    For r = 2 To 1500
        Dim nomeCell As String
        nomeCell = Trim(CStr(wsLista.Cells(r, 2).Value))
        If Len(nomeCell) = 0 Then GoTo NextListaRow

        Dim nomeUp As String
        nomeUp = UCase(Replace(Replace(nomeCell, "'", ""), ".", ""))

        If InStr(1, nomeUp, nc, vbTextCompare) > 0 Then
            Dim val As Variant
            val = wsLista.Cells(r, 6).Value
            If IsNumeric(val) Then
                CercaQtAttualeInLista = CDbl(val)
            Else
                CercaQtAttualeInLista = 0
            End If
            Exit Function
        End If
NextListaRow:
    Next r

    Log "    ATTENZIONE: '" & nomeCerca & "' non trovato in LISTA per calcolo contratto"
    CercaQtAttualeInLista = 0
End Function

' ============================================================
' Macro di verifica post-esecuzione
' ============================================================
Sub VerificaAssicuratiFT()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SQUADRE")

    Dim teamCols As Variant
    teamCols = Array( _
        Array("FCK Deportivo", 3), Array("Hellas Madonna", 15), _
        Array("muttley superstar", 27), Array("PARTIZAN", 39), _
        Array("Legenda Aurea", 51), Array("Kung Fu Pandev", 63), _
        Array("Millwall", 75), Array("FC CKC 26", 87), _
        Array("Papaie Top Team", 99), Array("Tronzano", 111))

    Dim output As String
    output = "GIOCATORI ASSICURATI - FT (post-macro):" & vbCrLf & vbCrLf

    Dim t As Long
    For t = LBound(teamCols) To UBound(teamCols)
        Dim teamName As String, col As Long, cnt As Long
        teamName = teamCols(t)(0): col = teamCols(t)(1): cnt = 0
        output = output & teamName & ":" & vbCrLf

        Dim r As Long
        For r = 6 To 52
            If UCase(Trim(CStr(ws.Cells(r, col + 3).Value))) = "A" Then
                Dim nome As String
                nome = Trim(CStr(ws.Cells(r, col).Value))
                If Len(nome) > 0 And nome <> "Calciatore" Then
                    Dim sp As String
                    sp = CStr(ws.Cells(r, col + 10).Value)
                    output = output & "  " & nome & " (Sp=" & sp & ")" & vbCrLf
                    cnt = cnt + 1
                End If
            End If
        Next r
        If cnt = 0 Then output = output & "  (nessuno)" & vbCrLf
        output = output & vbCrLf
    Next t

    Debug.Print output
    MsgBox output, vbInformation, "Verifica Assicurati FT"
End Sub
