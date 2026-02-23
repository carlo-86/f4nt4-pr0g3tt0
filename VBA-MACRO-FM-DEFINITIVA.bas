' ============================================================
' MACRO VBA DEFINITIVA - FANTAMANTRA MANAGERIALE 2026
' Aggiornamento completo DB: Listone + Scambi + Asta Riparazione + Allineamento Date + Assicurazioni + Fix Formule DB + Contratti Invernali
' Mercato Invernale 2026
' ============================================================
' ISTRUZIONI:
' 1. BACKUP del file DB prima di procedere!
' 2. Aprire "FantaMantra Manageriale - DB completo (06.02.2026).xlsx"
'    Password: 89y3R8HF'(()h7t87gH)(/0?9U38Qyp99
' 3. Alt+F11 > Inserisci > Modulo > Incolla tutto questo codice
' 4. F5 > Seleziona "ESEGUI_TUTTO_FM" > Esegui
' 5. Quando richiesto, seleziona il file listone (Quotazioni_Fantacalcio_...)
' ============================================================

' Colonne Calciatore (1-based VBA) per ogni squadra FM:
' Papaie Top Team   = 4    Legenda Aurea      = 17
' Lino Banfield FC  = 30   Kung Fu Pandev     = 43
' FICA              = 56   Hellas Madonna      = 69
' MINNESOTA AL MAX  = 82   FC CKC 26          = 95
' H-Q-A Barcelona   = 108  Mastri Birrai      = 121
'
' Offset da colonna Calciatore:
' +0  = Calciatore (nome)
' +1  = Squadra Serie A
' +2  = Reparto (FM ha colonna extra)
' +3  = Flag assicurazione ("A")
' +7  = Data assicurazione/rinnovo
' +10 = Spesa

Private Const DATA_ASS As String = "14/02/2026"
Private Const SHEET_PWD As String = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99"
Private logText As String

' ============================================================
' MAIN: Esegue tutte le operazioni in sequenza
' ============================================================
Sub ESEGUI_TUTTO_FM()
    logText = "=== LOG OPERAZIONI FM - MACRO DEFINITIVA ===" & vbCrLf & vbCrLf

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Salva stato protezione e rimuovi protezione dai fogli
    Dim protLista As Boolean, protSquadre As Boolean, protDB As Boolean, protQM As Boolean
    protLista = ThisWorkbook.Sheets("LISTA").ProtectContents
    protSquadre = ThisWorkbook.Sheets("SQUADRE").ProtectContents
    protDB = ThisWorkbook.Sheets("DB").ProtectContents
    On Error Resume Next
    protQM = ThisWorkbook.Sheets("QUOTE+MONTEPREMI 2026").ProtectContents
    On Error GoTo 0

    On Error Resume Next
    If protLista Then ThisWorkbook.Sheets("LISTA").Unprotect SHEET_PWD
    If protSquadre Then ThisWorkbook.Sheets("SQUADRE").Unprotect SHEET_PWD
    If protDB Then ThisWorkbook.Sheets("DB").Unprotect SHEET_PWD
    If protQM Then ThisWorkbook.Sheets("QUOTE+MONTEPREMI 2026").Unprotect SHEET_PWD
    On Error GoTo 0
    Log "Protezione fogli rimossa dove necessario"
    Log ""

    ' FASE 0: Aggiorna LISTA dal listone
    Log "FASE 0: Aggiornamento LISTA dal listone"
    Log "-------------------------------------------"
    Dim wsLista As Worksheet
    Set wsLista = ThisWorkbook.Sheets("LISTA")
    AggiornaListone wsLista, True  ' True = Mantra (FM)

    ' FASE 1: Gestisci scambi post-06/02
    Dim wsSq As Worksheet
    Set wsSq = ThisWorkbook.Sheets("SQUADRE")
    Log ""
    Log "FASE 1: Scambi post-06/02 (spostamento giocatori)"
    Log "-------------------------------------------"
    GestisciScambi wsSq

    ' FASE 2: Aggiungi giocatori asta riparazione (post 06/02)
    Log ""
    Log "FASE 2: Inserimento giocatori asta riparazione"
    Log "-------------------------------------------"
    InserisciAstaRiparazione wsSq

    ' FASE 3: Allineamento retroattivo date assicurazione (regola triennio rigido)
    Log ""
    Log "FASE 3: Allineamento date assicurazione al ciclo triennale"
    Log "-------------------------------------------"
    AllineaDatePreventive wsSq

    ' FASE 4: Registra assicurazioni
    Log ""
    Log "FASE 4: Registrazione assicurazioni"
    Log "-------------------------------------------"
    RegistraAssicurazioni wsSq

    ' FASE 5: Correggi formule foglio DB (audit 22/02/2026)
    Log ""
    Log "FASE 5: Correzione formule foglio DB"
    Log "-------------------------------------------"
    CorreggiFormuleDB

    ' FASE 6: Calcola e annota quote contratti mercato di riparazione
    Log ""
    Log "FASE 6: Calcolo quote contratti invernali (QUOTE+MONTEPREMI 2026)"
    Log "-------------------------------------------"
    CalcolaContrattiInvernali wsLista

    ' Ripristina protezione fogli (solo quelli che erano protetti)
    On Error Resume Next
    If protLista Then ThisWorkbook.Sheets("LISTA").Protect SHEET_PWD
    If protSquadre Then ThisWorkbook.Sheets("SQUADRE").Protect SHEET_PWD
    If protDB Then ThisWorkbook.Sheets("DB").Protect SHEET_PWD
    If protQM Then ThisWorkbook.Sheets("QUOTE+MONTEPREMI 2026").Protect SHEET_PWD
    On Error GoTo 0
    Log ""
    Log "Protezione fogli ripristinata."

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

    MsgBox "Operazioni FM completate!" & vbCrLf & _
           "Controlla il foglio LOG_MACRO per i dettagli.", _
           vbInformation, "FantaMantra Manageriale 2026"
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
            wsLista.Cells(mRow, 8).Value = fvmVal         ' FVM M
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
            wsLista.Cells(lastListaRow, 8).Value = fvmVal              ' FVM M
            aggiunti = aggiunti + 1
        End If

NextRow:
        r = r + 1
    Loop
End Sub

' ============================================================
' FASE 1: Gestisci scambi post-06/02
' Sposta giocatori tra le colonne squadra nel foglio SQUADRE
' ============================================================
Private Sub GestisciScambi(ws As Worksheet)
    ' --- Scambi 06/02: Legenda <-> Minnesota ---
    ' Questi potrebbero essere gia' nel DB del 06/02
    Log "  Scambi 06/02 (Legenda <-> Minnesota):"
    SpostaGiocatore ws, 17, 82, "Bernab"       ' Legenda -> Minnesota
    SpostaGiocatore ws, 82, 17, "Cancellieri"   ' Minnesota -> Legenda

    ' --- 11/02: Minnesota -> Papaie ---
    Log ""
    Log "  Scambio 11/02 (Minnesota -> Papaie):"
    SpostaGiocatore ws, 82, 4, "Hien"

    ' --- 13/02: Minnesota -> Papaie ---
    Log ""
    Log "  Scambio 13/02 (Minnesota -> Papaie):"
    SpostaGiocatore ws, 82, 4, "Diego Carlos"

    ' --- David: Minnesota -> Hellas ---
    ' David acquisito da Hellas (asta riparazione Sp=307)
    ' Nel DB 06/02 dovrebbe essere in Minnesota (da scambio Mastri->Minnesota 04/02)
    Log ""
    Log "  David: Minnesota -> Hellas (asta riparazione):"
    SpostaGiocatore ws, 82, 69, "David"

    ' --- 13/02: Minnesota -> Lino (4 giocatori) ---
    Log ""
    Log "  Scambio 13/02 (Minnesota -> Lino):"
    SpostaGiocatore ws, 82, 30, "Mazzitelli"
    SpostaGiocatore ws, 82, 30, "Tavares"
    SpostaGiocatore ws, 82, 30, "Koopmeiners"
    ' Bernabe' ora in Minnesota (da FASE 1 step 06/02, o gia' presente)
    SpostaGiocatore ws, 82, 30, "Bernab"

    ' --- 13/02: Lino -> Minnesota (4 giocatori) ---
    Log ""
    Log "  Scambio 13/02 (Lino -> Minnesota):"
    SpostaGiocatore ws, 30, 82, "Fagioli"
    SpostaGiocatore ws, 30, 82, "Bellanova"
    SpostaGiocatore ws, 30, 82, "Gimenez"
    SpostaGiocatore ws, 30, 82, "Miller"
End Sub

' ============================================================
' FASE 2: Inserisci giocatori acquisiti nell'asta riparazione
' Ogni giocatore va nella prima riga vuota della colonna squadra
' ============================================================
Private Sub InserisciAstaRiparazione(ws As Worksheet)
    ' --- KUNG FU PANDEV (col 43) ---
    InserisciGiocatore ws, 43, "Raspadori", "Ata", 119
    ' NOTA: Kouame' NON inserito via asta rip. - e' gia' in rosa (non svincolato)
    ' ma non piu' listato su Leghe Fantacalcio, non assicurabile

    ' --- FC CKC 26 (col 95) ---
    InserisciGiocatore ws, 95, "Durosinmi", "Pis", 19
    InserisciGiocatore ws, 95, "Vergara", "Nap", 71

    ' --- H-Q-A BARCELONA (col 108) ---
    InserisciGiocatore ws, 108, "Britschgi", "Par", 9
    InserisciGiocatore ws, 108, "Taylor K", "Laz", 52
    InserisciGiocatore ws, 108, "Malen", "Rom", 292

    ' --- HELLAS MADONNA (col 69) ---
    ' David: se spostato da Minnesota in FASE 1, aggiorna spesa a 307
    ' Se non trovato in Minnesota, inserisci come nuovo
    InserisciGiocatore ws, 69, "David", "Juv", 307
    InserisciGiocatore ws, 69, "Cheddira", "Lec", 1
    InserisciGiocatore ws, 69, "Zaragoza", "Rom", 10
    InserisciGiocatore ws, 69, "Ekkelenkamp", "Udi", 6
    InserisciGiocatore ws, 69, "Brescianini", "Fio", 4
    InserisciGiocatore ws, 69, "Belghali", "Ver", 1

    ' --- FICA (col 56) ---
    InserisciGiocatore ws, 56, "Luis Henrique", "Int", 1
    InserisciGiocatore ws, 56, "Fullkrug", "Mil", 1

    ' --- LINO BANFIELD FC (col 30) ---
    InserisciGiocatore ws, 30, "Celik", "Rom", 18
    InserisciGiocatore ws, 30, "Obert", "Cag", 1
    InserisciGiocatore ws, 30, "Marcandalli", "Gen", 1
    InserisciGiocatore ws, 30, "Bernasconi", "Ata", 3
    InserisciGiocatore ws, 30, "Bowie", "Ver", 1
    InserisciGiocatore ws, 30, "Vaz", "Rom", 5

    ' --- MINNESOTA AL MAX (col 82) ---
    InserisciGiocatore ws, 82, "Marianucci", "Tor", 4
    InserisciGiocatore ws, 82, "Bakola", "Sas", 1
    InserisciGiocatore ws, 82, "Adzic", "Juv", 2
    InserisciGiocatore ws, 82, "Ratkov", "Laz", 1

    ' --- PAPAIE TOP TEAM (col 4) ---
    InserisciGiocatore ws, 4, "Solomon", "Fio", 28

    ' --- LEGENDA AUREA (col 17) ---
    InserisciGiocatore ws, 17, "Nelsson", "Ver", 1
    InserisciGiocatore ws, 17, "Dossena", "Cag", 1
    InserisciGiocatore ws, 17, "Bartesaghi", "Mil", 40
    InserisciGiocatore ws, 17, "Gandelman", "Lec", 2
    InserisciGiocatore ws, 17, "Barbieri", "Cre", 1
End Sub

' ============================================================
' FASE 3: Registra tutte le assicurazioni
' ============================================================
Private Sub RegistraAssicurazioni(ws As Worksheet)
    ' --- KUNG FU PANDEV (col 43) ---
    AssicuraG ws, 43, "Kon"          ' Kone' I. - match parziale
    AssicuraG ws, 43, "Raspadori"
    ' POSCH: RESPINTO (svincolato da KFP, NON assicurabile)
    Log "  SKIP: Posch - svincolato, NON assicurabile (RESPINTO)"
    AssicuraG ws, 43, "Ferguson"
    ' KOUAME': NON assicurabile (non piu' listato su Leghe Fantacalcio)
    Log "  SKIP: Kouame' - non piu' listato, NON assicurabile"

    ' --- FC CKC 26 (col 95) ---
    AssicuraG ws, 95, "Durosinmi"
    AssicuraG ws, 95, "Vergara"
    AssicuraG ws, 95, "Zaniolo"

    ' --- H-Q-A BARCELONA (col 108) ---
    AssicuraG ws, 108, "Holm"
    AssicuraG ws, 108, "Ndicka"
    AssicuraG ws, 108, "Gallo"
    AssicuraG ws, 108, "Vasquez"
    AssicuraG ws, 108, "Gudmundsson"
    AssicuraG ws, 108, "Frendrup"
    AssicuraG ws, 108, "Britschgi"
    AssicuraG ws, 108, "Sulemana"
    AssicuraG ws, 108, "Taylor"
    AssicuraG ws, 108, "Malen"
    AssicuraG ws, 108, "Sommer"

    ' --- HELLAS MADONNA (col 69) ---
    AssicuraG ws, 69, "David"
    AssicuraG ws, 69, "Cheddira"
    AssicuraG ws, 69, "Zaragoza"
    AssicuraG ws, 69, "Ekkelenkamp"
    AssicuraG ws, 69, "Brescianini"
    AssicuraG ws, 69, "Belghali"
    AssicuraG ws, 69, "Scamacca"

    ' --- FICA (col 56) ---
    AssicuraG ws, 56, "Luis Henrique"
    AssicuraG ws, 56, "Fullkrug"

    ' --- LINO BANFIELD FC (col 30) ---
    AssicuraG ws, 30, "Celik"
    AssicuraG ws, 30, "Obert"
    AssicuraG ws, 30, "Marcandalli"
    AssicuraG ws, 30, "Bernasconi"
    AssicuraG ws, 30, "Bowie"
    AssicuraG ws, 30, "Caprile"
    AssicuraG ws, 30, "Cambiaghi"
    AssicuraG ws, 30, "Vaz"
    AssicuraG ws, 30, "Baldanzi"
    AssicuraG ws, 30, "Koopmeiners"
    ' TAVARES: RESPINTO — vincolo triennale non ancora decorso (scade 14/08/2027)
    Log "  SKIP: Tavares N. - triennale non decorso (scade ago 2027), NON assicurabile"
    AssicuraG ws, 30, "Mazzitelli"

    ' --- MINNESOTA AL MAX (col 82) ---
    AssicuraG ws, 82, "Montip"       ' Montipo' - match parziale
    AssicuraG ws, 82, "Marianucci"
    AssicuraG ws, 82, "Cataldi"
    AssicuraG ws, 82, "Fagioli"
    AssicuraG ws, 82, "Miller"
    AssicuraG ws, 82, "Bakola"
    AssicuraG ws, 82, "Adzic"
    AssicuraG ws, 82, "Ratkov"
    AssicuraG ws, 82, "Bellanova"

    ' --- PAPAIE TOP TEAM (col 4) ---
    AssicuraG ws, 4, "Kolasinac"
    AssicuraG ws, 4, "Hien"
    AssicuraG ws, 4, "Pasalic"
    AssicuraG ws, 4, "Nicolussi"     ' Nicolussi Caviglia - match parziale
    AssicuraG ws, 4, "Solomon"
    AssicuraG ws, 4, "Vlahovic"

    ' --- LEGENDA AUREA (col 17) ---
    AssicuraG ws, 17, "Nelsson"
    AssicuraG ws, 17, "Dossena"
    AssicuraG ws, 17, "Bartesaghi"
    AssicuraG ws, 17, "Gandelman"
    AssicuraG ws, 17, "Barbieri"
    AssicuraG ws, 17, "Leao"
    AssicuraG ws, 17, "Zappa"
End Sub

' ============================================================
' FASE 5: Correzione formule foglio DB (audit 22/02/2026)
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
' Macro standalone per correzione formule DB (FASE 5)
' (eseguibile indipendentemente da ESEGUI_TUTTO_FM)
' ============================================================
Sub CorreggiFormuleDB_FM()
    logText = "=== LOG CORREZIONE FORMULE DB - FM ===" & vbCrLf & vbCrLf

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Rimuovi protezione foglio DB se presente
    Dim protDB As Boolean
    protDB = ThisWorkbook.Sheets("DB").ProtectContents
    On Error Resume Next
    If protDB Then ThisWorkbook.Sheets("DB").Unprotect SHEET_PWD
    On Error GoTo 0

    CorreggiFormuleDB

    ' Ripristina protezione se era presente
    On Error Resume Next
    If protDB Then ThisWorkbook.Sheets("DB").Protect SHEET_PWD
    On Error GoTo 0

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
           vbInformation, "Fix Formule DB - FM"
End Sub

' ============================================================
' Helper: Sposta un giocatore da una colonna squadra a un'altra
' Cerca il giocatore nella colonna sorgente, lo copia nella prima
' riga vuota della colonna destinazione, poi cancella la sorgente.
' Ritorna True se lo spostamento e' avvenuto, False se non trovato.
' ============================================================
Private Function SpostaGiocatore(ws As Worksheet, colSrc As Long, colDst As Long, nomeCerca As String) As Boolean
    Dim r As Long
    Dim nc As String
    nc = UCase(Replace(Replace(nomeCerca, "'", ""), ".", ""))
    nc = Replace(nc, Chr(232), "e")
    nc = Replace(nc, Chr(233), "e")
    nc = Replace(nc, Chr(242), "o")
    nc = Replace(nc, Chr(224), "a")
    nc = Replace(nc, Chr(249), "u")

    ' Cerca il giocatore nella colonna sorgente (righe 6-52)
    For r = 6 To 52
        Dim nomeCell As String
        nomeCell = UCase(Trim(CStr(ws.Cells(r, colSrc).Value)))
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
            ' Trovato nella sorgente. Verifica che non sia gia' nella destinazione.
            Dim nomeGioc As String
            nomeGioc = ws.Cells(r, colSrc).Value

            ' Controlla se gia' presente nella destinazione
            Dim dstCheck As Long
            For dstCheck = 6 To 52
                Dim nomeDst As String
                nomeDst = UCase(Trim(CStr(ws.Cells(dstCheck, colDst).Value)))
                nomeDst = Replace(nomeDst, "'", "")
                nomeDst = Replace(nomeDst, ".", "")
                If Len(nomeDst) > 0 And InStr(1, nomeDst, nc, vbTextCompare) > 0 Then
                    Log "    GIA' IN DESTINAZIONE: " & nomeGioc & " (col " & colDst & " riga " & dstCheck & ") - skip"
                    SpostaGiocatore = False
                    Exit Function
                End If
            Next dstCheck

            ' Cerca prima riga vuota nella destinazione
            Dim dstRow As Long
            For dstRow = 6 To 50
                If Trim(CStr(ws.Cells(dstRow, colDst).Value)) = "" Then
                    ' Copia i dati (colonne +0 a +10)
                    Dim offset As Long
                    For offset = 0 To 10
                        ws.Cells(dstRow, colDst + offset).Value = ws.Cells(r, colSrc + offset).Value
                    Next offset
                    ' Cancella la riga sorgente
                    For offset = 0 To 10
                        ws.Cells(r, colSrc + offset).ClearContents
                    Next offset

                    Log "    SPOSTATO: " & nomeGioc & " (col " & colSrc & " riga " & r & " -> col " & colDst & " riga " & dstRow & ")"
                    SpostaGiocatore = True
                    Exit Function
                End If
            Next dstRow

            Log "    ERRORE: Nessuna riga vuota per spostare " & nomeGioc & " nella colonna " & colDst
            SpostaGiocatore = False
            Exit Function
        End If
    Next r

    Log "    NON TROVATO: " & nomeCerca & " nella colonna " & colSrc & " (potrebbe essere gia' spostato)"
    SpostaGiocatore = False
End Function

' ============================================================
' Helper: Inserisci un giocatore nuovo in SQUADRE
' Cerca la prima riga vuota nella colonna della squadra (righe 6-50)
' Se il giocatore esiste gia', aggiorna solo la spesa se diversa
' ============================================================
Private Sub InserisciGiocatore(ws As Worksheet, colCalc As Long, nome As String, squadraSerieA As String, spesa As Long)
    ' Prima verifica se il giocatore esiste gia'
    Dim r As Long
    For r = 6 To 50
        Dim existName As String
        existName = Trim(CStr(ws.Cells(r, colCalc).Value))
        If UCase(Replace(existName, "'", "")) Like "*" & UCase(Replace(Left(nome, 5), "'", "")) & "*" Then
            Log "  GIA' PRESENTE: " & nome & " (riga " & r & ", come '" & existName & "')"
            ' Aggiorna Spesa se diversa
            If ws.Cells(r, colCalc + 10).Value <> spesa Then
                ws.Cells(r, colCalc + 10).Value = spesa
                Log "    -> Spesa aggiornata a " & spesa
            End If
            Exit Sub
        End If
    Next r

    ' Cerca prima riga vuota
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
' FASE 3: Allineamento retroattivo date assicurazione
' Regola triennio rigido: la decorrenza e' SEMPRE uno scatto
' triennale dalla Data acquisto, mai la data di comunicazione.
' Algoritmo: per ogni assicurato, cammina la catena
'   Dacq, Dacq+3, Dacq+6, ... e allinea al boundary piu' vicino.
' ============================================================
Private Sub AllineaDatePreventive(ws As Worksheet)
    Dim teamCols As Variant
    teamCols = Array( _
        Array("Papaie Top Team", 4), Array("Legenda Aurea", 17), _
        Array("Lino Banfield FC", 30), Array("Kung Fu Pandev", 43), _
        Array("FICA", 56), Array("Hellas Madonna", 69), _
        Array("MINNESOTA AL MAX", 82), Array("FC CKC 26", 95), _
        Array("H-Q-A Barcelona", 108), Array("Mastri Birrai", 121))

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
            If Len(nome) = 0 Or nome = "Calciatore" Then GoTo NextPlayerFM

            Dim flag As String
            flag = Trim(CStr(ws.Cells(r, col + 3).Value))
            If flag <> "A" Then GoTo NextPlayerFM

            Dim insDateVal As Variant
            insDateVal = ws.Cells(r, col + 7).Value
            If Not IsDate(insDateVal) Then GoTo NextPlayerFM
            Dim insDate As Date
            insDate = CDate(insDateVal)

            ' Data acquisto come ancora del ciclo triennale
            Dim acqDateVal As Variant
            acqDateVal = ws.Cells(r, col + 4).Value
            If Not IsDate(acqDateVal) Then
                noAnchor = noAnchor + 1
                GoTo NextPlayerFM
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
NextPlayerFM:
        Next r
    Next t

    Log "  Allineamento: " & allineati & " corretti, " & invariati & _
        " gia' allineati, " & noAnchor & " senza data acquisto (invariati)"
End Sub

' ============================================================
' FASE 6: Calcola quote contratti mercato di riparazione
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

    ' Definizione acquisti invernali per squadra FM
    ' Array: (nomeSquadraQM, Array di nomi giocatori)
    Dim teamData(0 To 8) As Variant
    teamData(0) = Array("Kung Fu", Array("Raspadori"))
    teamData(1) = Array("CKC", Array("Durosinmi", "Vergara"))
    teamData(2) = Array("H-Q-A", Array("Britschgi", "Taylor", "Malen"))
    teamData(3) = Array("Hellas", Array("David", "Cheddira", "Zaragoza", "Ekkelenkamp", "Brescianini", "Belghali"))
    teamData(4) = Array("Federazione", Array("Luis Henrique", "Fullkrug"))
    teamData(5) = Array("Lino", Array("Celik", "Obert", "Marcandalli", "Bernasconi", "Bowie", "Vaz"))
    teamData(6) = Array("Minnesota", Array("Marianucci", "Bakola", "Adzic", "Ratkov"))
    teamData(7) = Array("Papaie", Array("Solomon"))
    teamData(8) = Array("Legenda", Array("Nelsson", "Dossena", "Bartesaghi", "Gandelman", "Barbieri"))

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
        If Len(nomeCell) = 0 Then GoTo NextListaRowFM

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
NextListaRowFM:
    Next r

    Log "    ATTENZIONE: '" & nomeCerca & "' non trovato in LISTA per calcolo contratto"
    CercaQtAttualeInLista = 0
End Function

' ============================================================
' Macro di verifica post-esecuzione
' ============================================================
Sub VerificaAssicuratiFM()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SQUADRE")

    Dim teamCols As Variant
    teamCols = Array( _
        Array("Papaie Top Team", 4), Array("Legenda Aurea", 17), _
        Array("Lino Banfield FC", 30), Array("Kung Fu Pandev", 43), _
        Array("FICA", 56), Array("Hellas Madonna", 69), _
        Array("MINNESOTA AL MAX", 82), Array("FC CKC 26", 95), _
        Array("H-Q-A Barcelona", 108), Array("Mastri Birrai", 121))

    Dim output As String
    output = "GIOCATORI ASSICURATI - FM (post-macro):" & vbCrLf & vbCrLf

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
    MsgBox output, vbInformation, "Verifica Assicurati FM"
End Sub
