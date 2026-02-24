' ============================================================
' MACRO VBA DEFINITIVA - FANTAMANTRA MANAGERIALE 2026
' Aggiornamento completo DB: Listone + Svincoli + Scambi + Asta Riparazione + Riordinamento SQUADRE + Allineamento Date + Assicurazioni + Fix Formule DB + Contratti Invernali
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
' -1  = Reparto (P/D/C/A, pre-compilato, FM ha colonna extra)
' +0  = Calciatore (nome)
' +1  = Ruolo (formula CERCA.VERT - NON scrivere!)
' +2  = Squadra Serie A (formula CERCA.VERT - NON scrivere!)
' +3  = Flag assicurazione ("A")
' +4  = Data acquisto (serial)
' +5  = Q. all'acquisto
' +6  = FVM Prop. all'acquisto
' +7  = Data assicurazione/rinnovo
' +8  = Q. rinn. ass. ("/" se prima assicurazione)
' +9  = FVM Prop. rinn.
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

    ' SALVA QUOTAZIONI 06/02 da LISTA prima dell'aggiornamento
    ' (servono in FASE 3 per Qt all'acquisto e FVM dei nuovi giocatori)
    Dim wsLista As Worksheet
    Set wsLista = ThisWorkbook.Sheets("LISTA")
    Dim dictQuot As Object
    Set dictQuot = CreateObject("Scripting.Dictionary")
    Dim dqR As Long
    For dqR = 2 To 600
        Dim dqName As String
        dqName = NormName(CStr(wsLista.Cells(dqR, 2).Value))
        If dqName <> "" Then
            ' Array: (0)=Ruolo Classic, (1)=Qt.Attuale, (2)=FVM
            dictQuot(dqName) = Array( _
                UCase(Trim(CStr(wsLista.Cells(dqR, 3).Value))), _
                wsLista.Cells(dqR, 6).Value, _
                wsLista.Cells(dqR, 8).Value)
        End If
    Next dqR
    Log "Salvate " & dictQuot.Count & " quotazioni pre-aggiornamento da LISTA"
    Log ""

    ' FASE 0: Aggiorna LISTA dal listone (enhanced: delistati a 601+, date nascita)
    Log "FASE 0: Aggiornamento LISTA dal listone"
    Log "-------------------------------------------"
    AggiornaListone wsLista, True  ' True = Mantra (FM)

    ' FASE 1: Svincoli (rosa attiva -> elenco storico)
    Dim wsSq As Worksheet
    Set wsSq = ThisWorkbook.Sheets("SQUADRE")
    Log ""
    Log "FASE 1: Svincoli (spostamento da rosa attiva a elenco storico)"
    Log "-------------------------------------------"
    EseguiSvincoliFM wsSq

    ' FASE 2: Gestisci scambi post-06/02
    Log ""
    Log "FASE 2: Scambi post-06/02 (spostamento giocatori)"
    Log "-------------------------------------------"
    GestisciScambi wsSq

    ' FASE 3: Aggiungi giocatori asta riparazione (post 06/02)
    Log ""
    Log "FASE 3: Inserimento giocatori asta riparazione"
    Log "-------------------------------------------"
    InserisciAstaRiparazione wsSq, dictQuot

    ' FASE 4: Riordinamento SQUADRE per reparto/spesa
    Log ""
    Log "FASE 4: Riordinamento SQUADRE per reparto e spesa"
    Log "-------------------------------------------"
    RiordinaSquadreFM wsSq

    ' FASE 5: Allineamento retroattivo date assicurazione (regola triennio rigido)
    Log ""
    Log "FASE 5: Allineamento date assicurazione al ciclo triennale"
    Log "-------------------------------------------"
    AllineaDatePreventive wsSq

    ' FASE 6: Registra assicurazioni
    Log ""
    Log "FASE 6: Registrazione assicurazioni"
    Log "-------------------------------------------"
    RegistraAssicurazioni wsSq

    ' FASE 7: Correggi formule foglio DB (audit 22/02/2026)
    Log ""
    Log "FASE 7: Correzione formule foglio DB"
    Log "-------------------------------------------"
    CorreggiFormuleDB

    ' FASE 8: Calcola e annota quote contratti mercato di riparazione
    Log ""
    Log "FASE 8: Calcolo quote contratti invernali (QUOTE+MONTEPREMI 2026)"
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
    Dim logLines() As String, logR As Long
    logLines = Split(logText, vbCrLf)
    For logR = 0 To UBound(logLines)
        wsLog.Cells(logR + 1, 1).Value = logLines(logR)
    Next logR

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

    ' 3. Dizionario per tracciare tutti gli ID nel listone
    Dim dictListoneIDs As Object
    Set dictListoneIDs = CreateObject("Scripting.Dictionary")
    Dim dictNewPlayers As Object
    Set dictNewPlayers = CreateObject("Scripting.Dictionary")

    Dim aggiornati As Long, aggiunti As Long, skippati As Long
    aggiornati = 0: aggiunti = 0: skippati = 0

    ' Trova l'ultima riga con dati nella tabella principale (solo righe 2-600)
    Dim lastListaRow As Long
    lastListaRow = 1
    Dim scanR As Long
    For scanR = 2 To 600
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
        ProcessaFoglioListone wsTutti, wsLista, isMantra, lastListaRow, _
            aggiornati, aggiunti, skippati, dictListoneIDs, dictNewPlayers
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
        ProcessaFoglioListone wsCeduti, wsLista, isMantra, lastListaRow, _
            aggiornati, aggiunti, skippati, dictListoneIDs, dictNewPlayers
    Else
        Log "  ATTENZIONE: Foglio 'Ceduti' non trovato nel listone."
    End If

    ' 4. Chiudi il listone senza salvare
    wbListone.Close SaveChanges:=False

    Log "  LISTA aggiornata: " & aggiornati & " aggiornati, " & aggiunti & " aggiunti, " & skippati & " skippati"

    ' ========================================
    ' 5. GESTIONE DELISTATI: sposta dalla tabella principale (2-600) alla sezione storica (601+)
    ' ========================================
    Log "  Gestione delistati..."
    Dim delistCount As Long: delistCount = 0

    Dim lastHistRow As Long: lastHistRow = 600
    For scanR = 601 To 2000
        If Trim(CStr(wsLista.Cells(scanR, 1).Value)) <> "" Then
            lastHistRow = scanR
        End If
    Next scanR

    For scanR = 2 To lastListaRow
        Dim listaId As Variant
        listaId = wsLista.Cells(scanR, 1).Value
        If Trim(CStr(listaId)) <> "" Then
            On Error Resume Next
            Dim idNum As Long
            idNum = CLng(listaId)
            On Error GoTo 0
            If idNum > 0 And Not dictListoneIDs.Exists(idNum) Then
                lastHistRow = lastHistRow + 1
                Dim colC As Long
                For colC = 1 To 9
                    wsLista.Cells(lastHistRow, colC).Value = wsLista.Cells(scanR, colC).Value
                Next colC
                For colC = 1 To 9
                    wsLista.Cells(scanR, colC).ClearContents
                Next colC
                delistCount = delistCount + 1
            End If
            idNum = 0
        End If
    Next scanR
    Log "  Delistati spostati a 601+: " & delistCount

    ' ========================================
    ' 6. COMPATTAMENTO tabella principale
    ' ========================================
    Log "  Compattamento tabella principale..."
    Dim mainCount As Long: mainCount = 0
    For scanR = 2 To 600
        If Trim(CStr(wsLista.Cells(scanR, 2).Value)) <> "" Then
            mainCount = mainCount + 1
        End If
    Next scanR

    If mainCount > 0 Then
        ReDim mainData(1 To mainCount, 1 To 9) As Variant
        Dim mIdx As Long: mIdx = 0
        For scanR = 2 To 600
            If Trim(CStr(wsLista.Cells(scanR, 2).Value)) <> "" Then
                mIdx = mIdx + 1
                For colC = 1 To 9
                    mainData(mIdx, colC) = wsLista.Cells(scanR, colC).Value
                Next colC
            End If
        Next scanR

        For scanR = 2 To 600
            mIdx = scanR - 1
            If mIdx <= mainCount Then
                For colC = 1 To 9
                    wsLista.Cells(scanR, colC).Value = mainData(mIdx, colC)
                Next colC
            Else
                For colC = 1 To 9
                    wsLista.Cells(scanR, colC).ClearContents
                Next colC
            End If
        Next scanR
    End If
    lastListaRow = mainCount + 1
    Log "  Tabella principale compattata: " & mainCount & " calciatori"

    ' ========================================
    ' 7. DATE DI NASCITA per nuovi calciatori (L:N)
    ' ========================================
    If dictNewPlayers.Count > 0 Then
        Log "  Ricerca date di nascita per " & dictNewPlayers.Count & " nuovi calciatori..."
        Dim newKeys As Variant
        newKeys = dictNewPlayers.Keys
        Dim nk As Long
        For nk = LBound(newKeys) To UBound(newKeys)
            Dim newId As Long: newId = newKeys(nk)
            Dim newInfo As String: newInfo = CStr(dictNewPlayers(newId))
            Dim pipePos As Long: pipePos = InStr(newInfo, "|")
            Dim newNome As String: newNome = Left(newInfo, pipePos - 1)
            Dim newSquadra As String: newSquadra = Mid(newInfo, pipePos + 1)

            Dim alreadyInLN As Boolean: alreadyInLN = False
            For scanR = 1 To 2000
                If UCase(Trim(CStr(wsLista.Cells(scanR, 12).Value))) = UCase(newNome) Then
                    alreadyInLN = True
                    Exit For
                End If
            Next scanR

            If Not alreadyInLN Then
                Dim lnRow As Long: lnRow = 0
                For scanR = 2 To 2000
                    If Trim(CStr(wsLista.Cells(scanR, 12).Value)) = "" Then
                        lnRow = scanR
                        Exit For
                    End If
                Next scanR

                If lnRow > 0 Then
                    wsLista.Cells(lnRow, 12).Value = newNome
                    Dim birthDate As Variant
                    birthDate = CercaDataNascitaOnline(newId, newNome, newSquadra)
                    If Not IsEmpty(birthDate) Then
                        wsLista.Cells(lnRow, 13).Value = birthDate
                        wsLista.Cells(lnRow, 13).NumberFormat = "dd/mm/yyyy"
                        Log "    " & newNome & " (ID " & newId & "): data nascita trovata"
                    Else
                        Log "    " & newNome & " (ID " & newId & "): data nascita NON trovata - inserire manualmente"
                    End If
                    wsLista.Cells(lnRow, 14).Formula = _
                        "=IF(M" & lnRow & "="""","""",INT(($N$1-M" & lnRow & ")/365.25))"
                End If
            End If
        Next nk
    End If

    ' ========================================
    ' 8. Formula Eta'
    ' ========================================
    Log "  Aggiornamento formula Eta'..."
    Dim etaR As Long
    For etaR = 2 To lastListaRow
        If Trim(CStr(wsLista.Cells(etaR, 2).Value)) <> "" Then
            Dim etaValFM As Variant
            etaValFM = wsLista.Cells(etaR, 9).Value
            Dim needsEtaFM As Boolean
            needsEtaFM = False
            If IsEmpty(etaValFM) Then
                needsEtaFM = True
            ElseIf IsError(etaValFM) Then
                needsEtaFM = True
            ElseIf etaValFM = "" Then
                needsEtaFM = True
            End If
            If needsEtaFM Then
                wsLista.Cells(etaR, 9).Formula = _
                    "=IFERROR(VLOOKUP(B" & etaR & ",$L:$N,3,FALSE),"""")"
            End If
        End If
    Next etaR

    ' ========================================
    ' 9. Ordina tabella principale A->Z
    ' ========================================
    Log "  Ordinamento tabella principale..."
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

    ' ========================================
    ' 10. Ordina sezione storica (601+) A->Z
    ' ========================================
    If lastHistRow > 601 Then
        Log "  Ordinamento sezione storica (601-" & lastHistRow & ")..."
        wsLista.Sort.SortFields.Clear
        wsLista.Sort.SortFields.Add2 _
            Key:=wsLista.Range("B601:B" & lastHistRow), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        With wsLista.Sort
            .SetRange wsLista.Range("A601:I" & lastHistRow)
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If

    Log "  FASE 0 completata."
End Sub

' ============================================================
' Helper: Processa un singolo foglio del listone (Tutti o Ceduti)
' Legge ogni riga, match per ID, aggiorna o aggiunge nella LISTA
' ============================================================
Private Sub ProcessaFoglioListone(wsSource As Worksheet, wsLista As Worksheet, _
    isMantra As Boolean, ByRef lastListaRow As Long, _
    ByRef aggiornati As Long, ByRef aggiunti As Long, ByRef skippati As Long, _
    dictIDs As Object, dictNew As Object)

    Dim r As Long
    r = 3

    Do While Trim(CStr(wsSource.Cells(r, 1).Value)) <> ""
        Dim listoneId As Variant
        listoneId = wsSource.Cells(r, 1).Value

        Dim nomeRaw As String
        nomeRaw = Trim(CStr(wsSource.Cells(r, 4).Value))
        Dim nomeClean As String
        nomeClean = Replace(nomeRaw, ".", "")
        nomeClean = Trim(nomeClean)

        Dim ruoloC As String, ruoloM As String, squadra As String
        ruoloC = Trim(CStr(wsSource.Cells(r, 2).Value))
        ruoloM = Trim(CStr(wsSource.Cells(r, 3).Value))
        squadra = Trim(CStr(wsSource.Cells(r, 5).Value))

        Dim qtAttuale As Variant, qtIniziale As Variant, fvmVal As Variant
        If isMantra Then
            qtAttuale = wsSource.Cells(r, 9).Value
            qtIniziale = wsSource.Cells(r, 10).Value
            fvmVal = wsSource.Cells(r, 13).Value
        Else
            qtAttuale = wsSource.Cells(r, 6).Value
            qtIniziale = wsSource.Cells(r, 7).Value
            fvmVal = wsSource.Cells(r, 12).Value
        End If

        If Len(nomeClean) = 0 Then
            skippati = skippati + 1
            GoTo NextRowFM
        End If

        On Error Resume Next
        dictIDs.Item(CLng(listoneId)) = True
        On Error GoTo 0

        Dim matchRow As Variant
        matchRow = Application.Match(CLng(listoneId), wsLista.Range("A1:A600"), 0)

        If Not IsError(matchRow) Then
            Dim mRow As Long
            mRow = CLng(matchRow)
            wsLista.Cells(mRow, 2).Value = nomeClean
            wsLista.Cells(mRow, 3).Value = ruoloC
            wsLista.Cells(mRow, 4).Value = ruoloM
            wsLista.Cells(mRow, 5).Value = squadra
            wsLista.Cells(mRow, 6).Value = qtAttuale
            wsLista.Cells(mRow, 7).Value = qtIniziale
            wsLista.Cells(mRow, 8).Value = fvmVal
            aggiornati = aggiornati + 1
        Else
            Dim histMatch As Variant
            histMatch = Empty
            Dim sR As Long
            For sR = 601 To 2000
                If Trim(CStr(wsLista.Cells(sR, 1).Value)) = "" Then GoTo SkipHistRowFM
                If CLng(wsLista.Cells(sR, 1).Value) = CLng(listoneId) Then
                    histMatch = sR
                    Exit For
                End If
SkipHistRowFM:
            Next sR

            If Not IsEmpty(histMatch) Then
                lastListaRow = lastListaRow + 1
                Dim colC As Long
                For colC = 1 To 9
                    wsLista.Cells(lastListaRow, colC).Value = wsLista.Cells(CLng(histMatch), colC).Value
                Next colC
                For colC = 1 To 9
                    wsLista.Cells(CLng(histMatch), colC).ClearContents
                Next colC
                wsLista.Cells(lastListaRow, 2).Value = nomeClean
                wsLista.Cells(lastListaRow, 3).Value = ruoloC
                wsLista.Cells(lastListaRow, 4).Value = ruoloM
                wsLista.Cells(lastListaRow, 5).Value = squadra
                wsLista.Cells(lastListaRow, 6).Value = qtAttuale
                wsLista.Cells(lastListaRow, 7).Value = qtIniziale
                wsLista.Cells(lastListaRow, 8).Value = fvmVal
                aggiornati = aggiornati + 1
            Else
                lastListaRow = lastListaRow + 1
                If lastListaRow > 600 Then
                    Log "    ATTENZIONE: Tabella principale piena, " & nomeClean & " non aggiunto"
                    lastListaRow = lastListaRow - 1
                    skippati = skippati + 1
                    GoTo NextRowFM
                End If
                wsLista.Cells(lastListaRow, 1).Value = CLng(listoneId)
                wsLista.Cells(lastListaRow, 2).Value = nomeClean
                wsLista.Cells(lastListaRow, 3).Value = ruoloC
                wsLista.Cells(lastListaRow, 4).Value = ruoloM
                wsLista.Cells(lastListaRow, 5).Value = squadra
                wsLista.Cells(lastListaRow, 6).Value = qtAttuale
                wsLista.Cells(lastListaRow, 7).Value = qtIniziale
                wsLista.Cells(lastListaRow, 8).Value = fvmVal
                aggiunti = aggiunti + 1
                dictNew.Item(CLng(listoneId)) = nomeClean & "|" & squadra
            End If
        End If

NextRowFM:
        r = r + 1
    Loop
End Sub

' ============================================================
' Helper: Cerca data di nascita online da fantacalcio.it
' ============================================================
Private Function CercaDataNascitaOnline(playerID As Long, playerName As String, squadraAbbr As String) As Variant
    CercaDataNascitaOnline = Empty
    On Error GoTo BirthDateErrFM

    Dim teamUrl As String: teamUrl = ""
    Dim sq As String: sq = LCase(Trim(squadraAbbr))
    Select Case sq
        Case "ata": teamUrl = "atalanta"
        Case "bol": teamUrl = "bologna"
        Case "cag": teamUrl = "cagliari"
        Case "com": teamUrl = "como"
        Case "cre": teamUrl = "cremonese"
        Case "emp": teamUrl = "empoli"
        Case "fio": teamUrl = "fiorentina"
        Case "gen": teamUrl = "genoa"
        Case "int": teamUrl = "inter"
        Case "juv": teamUrl = "juventus"
        Case "laz": teamUrl = "lazio"
        Case "lec": teamUrl = "lecce"
        Case "mil": teamUrl = "milan"
        Case "mon": teamUrl = "monza"
        Case "nap": teamUrl = "napoli"
        Case "par": teamUrl = "parma"
        Case "pis": teamUrl = "pisa"
        Case "rom": teamUrl = "roma"
        Case "sas": teamUrl = "sassuolo"
        Case "tor": teamUrl = "torino"
        Case "udi": teamUrl = "udinese"
        Case "ven": teamUrl = "venezia"
        Case "ver": teamUrl = "verona"
        Case "niz": teamUrl = "nizza"
        Case Else: Exit Function
    End Select

    Dim nomeUrl As String
    nomeUrl = LCase(Trim(playerName))
    nomeUrl = Replace(nomeUrl, " ", "-")
    nomeUrl = Replace(nomeUrl, "'", "")
    nomeUrl = Replace(nomeUrl, ".", "")
    nomeUrl = Replace(nomeUrl, Chr(232), "e")
    nomeUrl = Replace(nomeUrl, Chr(233), "e")
    nomeUrl = Replace(nomeUrl, Chr(236), "i")
    nomeUrl = Replace(nomeUrl, Chr(242), "o")
    nomeUrl = Replace(nomeUrl, Chr(224), "a")
    nomeUrl = Replace(nomeUrl, Chr(249), "u")

    Dim url As String
    url = "https://www.fantacalcio.it/serie-a/squadre/" & teamUrl & "/" & nomeUrl & "/" & playerID

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send

    If http.Status <> 200 Then GoTo BirthDateErrFM

    Dim html As String
    html = http.responseText

    Dim pos As Long
    pos = InStr(1, html, "nascita", vbTextCompare)
    If pos = 0 Then pos = InStr(1, html, "nato il", vbTextCompare)
    If pos = 0 Then GoTo BirthDateErrFM

    Dim chunk As String
    chunk = Mid(html, pos, 200)
    Dim ci As Long
    For ci = 1 To Len(chunk) - 9
        Dim ch As String
        ch = Mid(chunk, ci, 10)
        If Mid(ch, 3, 1) = "/" And Mid(ch, 6, 1) = "/" Then
            If IsNumeric(Left(ch, 2)) And IsNumeric(Mid(ch, 4, 2)) And IsNumeric(Right(ch, 4)) Then
                CercaDataNascitaOnline = CDate(ch)
                Exit For
            End If
        End If
    Next ci

    Set http = Nothing
    Exit Function

BirthDateErrFM:
    CercaDataNascitaOnline = Empty
    On Error Resume Next
    Set http = Nothing
    On Error GoTo 0
End Function

' ============================================================
' FASE 1: Svincoli (sposta giocatori dalla rosa attiva
'         all'elenco storico dei calciatori ceduti)
' ============================================================
Private Sub EseguiSvincoliFM(ws As Worksheet)
    ' --- PAPAIE TOP TEAM (col 4) - usa asterischi nell'elenco storico ---
    SvincolaGiocatore ws, 4, "Masina", True
    SvincolaGiocatore ws, 4, "Mari", True             ' Mari'
    SvincolaGiocatore ws, 4, "Lookman", True

    ' --- LEGENDA AUREA (col 17) ---
    SvincolaGiocatore ws, 17, "Bianchetti", False
    SvincolaGiocatore ws, 17, "Angelino", False
    SvincolaGiocatore ws, 17, "Zerbin", False
    SvincolaGiocatore ws, 17, "Pobega", False
    SvincolaGiocatore ws, 17, "Sohm", False

    ' --- LINO BANFIELD FC (col 30) ---
    SvincolaGiocatore ws, 30, "Mandas", False
    SvincolaGiocatore ws, 30, "Viti", False
    SvincolaGiocatore ws, 30, "Asllani", False
    SvincolaGiocatore ws, 30, "Gronbaek", False
    SvincolaGiocatore ws, 30, "Bravo", False
    SvincolaGiocatore ws, 30, "Castellanos", False
    SvincolaGiocatore ws, 30, "Masini", False

    ' --- KUNG FU PANDEV (col 43) ---
    SvincolaGiocatore ws, 43, "Posch", False
    SvincolaGiocatore ws, 43, "Stengs", False
    SvincolaGiocatore ws, 43, "Israel", False
    SvincolaGiocatore ws, 43, "Ehizibue", False

    ' --- FICA (col 56) ---
    SvincolaGiocatore ws, 56, "Guendouzi", False
    SvincolaGiocatore ws, 56, "Vasquez D", False

    ' --- HELLAS MADONNA (col 69) ---
    SvincolaGiocatore ws, 69, "De Vrij", False
    SvincolaGiocatore ws, 69, "Stanciu", False
    SvincolaGiocatore ws, 69, "Bailey", False
    SvincolaGiocatore ws, 69, "Anjorin", False
    SvincolaGiocatore ws, 69, "Sorensen", False        ' Sorensen O.
    SvincolaGiocatore ws, 69, "Lucca", False

    ' --- MINNESOTA AL MAX (col 82) ---
    SvincolaGiocatore ws, 82, "Troilo", False
    SvincolaGiocatore ws, 82, "Musah", False
    SvincolaGiocatore ws, 82, "Almqvist", False

    ' --- FC CKC 26 (col 95) ---
    SvincolaGiocatore ws, 95, "Zanoli", False
    SvincolaGiocatore ws, 95, "Carboni V", False
    SvincolaGiocatore ws, 95, "Lang", False

    ' --- H-Q-A BARCELONA (col 108) ---
    SvincolaGiocatore ws, 108, "Lykogiannis", False
    SvincolaGiocatore ws, 108, "Pierotti", False
    SvincolaGiocatore ws, 108, "Fazzini", False
    SvincolaGiocatore ws, 108, "Luvumbo", False

    ' --- MASTRI BIRRAI (col 121) ---
    SvincolaGiocatore ws, 121, "Dele-Bashiru", False
End Sub

' ============================================================
' Helper: Svincola un giocatore (rosa attiva -> elenco storico)
' Cerca il giocatore nella rosa attiva (righe 6-52), copia
' tutti i dati nell'elenco storico sottostante, cancella la
' riga dalla rosa attiva. useAsterisk=True per Papaie (FM)
' e FCK (FT) che usano * nei nomi storici.
' ============================================================
Private Function SvincolaGiocatore(ws As Worksheet, colCalc As Long, nomeCerca As String, useAsterisk As Boolean) As Boolean
    Dim nc As String
    nc = UCase(Replace(Replace(Trim(nomeCerca), "'", ""), ".", ""))
    nc = Replace(nc, Chr(232), "e")
    nc = Replace(nc, Chr(233), "e")
    nc = Replace(nc, Chr(236), "i")
    nc = Replace(nc, Chr(242), "o")
    nc = Replace(nc, Chr(224), "a")
    nc = Replace(nc, Chr(249), "u")

    ' Cerca il giocatore nella rosa attiva (righe 6-52)
    Dim r As Long
    For r = 6 To 52
        Dim nomeCell As String
        nomeCell = UCase(Trim(CStr(ws.Cells(r, colCalc).Value)))
        If Len(nomeCell) = 0 Or nomeCell = "CALCIATORE" Then GoTo NextSvRowFM
        nomeCell = Replace(nomeCell, "'", "")
        nomeCell = Replace(nomeCell, ".", "")
        nomeCell = Replace(nomeCell, Chr(232), "e")
        nomeCell = Replace(nomeCell, Chr(233), "e")
        nomeCell = Replace(nomeCell, Chr(236), "i")
        nomeCell = Replace(nomeCell, Chr(242), "o")
        nomeCell = Replace(nomeCell, Chr(224), "a")
        nomeCell = Replace(nomeCell, Chr(249), "u")
        nomeCell = Replace(nomeCell, Chr(200), "E")
        nomeCell = Replace(nomeCell, Chr(201), "E")

        If InStr(1, nomeCell, nc, vbTextCompare) > 0 Then
            ' Trovato! Salva tutti i dati (offsets +0 a +10)
            Dim savedData(0 To 10) As Variant
            Dim offset As Long
            For offset = 0 To 10
                savedData(offset) = ws.Cells(r, colCalc + offset).Value
            Next offset

            ' Trova l'elenco storico: cerca "Elenco storico" nella colonna del team
            Dim histStart As Long
            histStart = 0
            Dim sr As Long
            For sr = 48 To 200
                If InStr(1, CStr(ws.Cells(sr, colCalc).Value), "Elenco storico", vbTextCompare) > 0 Then
                    histStart = sr
                    Exit For
                End If
            Next sr

            If histStart = 0 Then
                Log "  ERRORE SVINCOLO: Elenco storico non trovato per colonna " & colCalc
                SvincolaGiocatore = False
                Exit Function
            End If

            ' L'header e' 3 righe sotto "Elenco storico", i dati partono da 4 righe sotto
            Dim histDataStart As Long
            histDataStart = histStart + 4

            ' Trova la prima riga vuota nell'elenco storico
            Dim histRow As Long
            histRow = histDataStart
            Do While Trim(CStr(ws.Cells(histRow, colCalc).Value)) <> ""
                histRow = histRow + 1
                If histRow > 570 Then
                    Log "  ERRORE SVINCOLO: Elenco storico pieno per colonna " & colCalc
                    SvincolaGiocatore = False
                    Exit Function
                End If
            Loop

            ' Scrivi i dati nell'elenco storico
            For offset = 0 To 10
                ws.Cells(histRow, colCalc + offset).Value = savedData(offset)
            Next offset

            ' Formatta le colonne data
            If IsDate(ws.Cells(histRow, colCalc + 4).Value) Or IsNumeric(ws.Cells(histRow, colCalc + 4).Value) Then
                ws.Cells(histRow, colCalc + 4).NumberFormat = "dd/mm/yyyy"
            End If
            If IsDate(ws.Cells(histRow, colCalc + 7).Value) Or IsNumeric(ws.Cells(histRow, colCalc + 7).Value) Then
                ws.Cells(histRow, colCalc + 7).NumberFormat = "dd/mm/yyyy"
            End If

            ' Gestisci asterischi se necessario (solo Papaie per FM)
            If useAsterisk Then
                ws.Cells(histRow, colCalc).Value = CStr(savedData(0)) & "*"
                Dim insFlag As String
                insFlag = Trim(CStr(ws.Cells(histRow, colCalc + 3).Value))
                If insFlag = "A" Then
                    ws.Cells(histRow, colCalc + 3).Value = "A*"
                End If
            End If

            ' Cancella la riga nella rosa attiva
            For offset = 0 To 10
                ws.Cells(r, colCalc + offset).ClearContents
            Next offset

            Log "  SVINCOLATO: " & CStr(savedData(0)) & " (riga " & r & " -> storico riga " & histRow & ")" & _
                IIf(useAsterisk, " [con *]", "")
            SvincolaGiocatore = True
            Exit Function
        End If
NextSvRowFM:
    Next r

    Log "  NON TROVATO per svincolo: " & nomeCerca & " nella colonna " & colCalc
    SvincolaGiocatore = False
End Function

' ============================================================
' FASE 2: Gestisci scambi post-06/02
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
Private Sub InserisciAstaRiparazione(ws As Worksheet, dictQuot As Object)
    ' 36 giocatori totali - date da "Mercato ASTA CLASSICA"
    ' Formato: InserisciGiocatore ws, colCalc, nome, spesa, dataAcquisto, dictQuot
    ' NON scrive Ruolo (+1) e Squadra (+2): sono formule CERCA.VERT
    ' Scrive: Nome (+0), DataAcq (+4), Qt.Acq (+5), FVM (+6), "/" (+8), Spesa (+10)
    ' NOTA: David (Juv, 307) NON e' asta riparazione - e' scambio (gestito in FASE 2)

    ' --- KUNG FU PANDEV (col 43) ---
    InserisciGiocatore ws, 43, "Raspadori", 119, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 43, "Kone' I.", 41, DateSerial(2026, 2, 11), dictQuot

    ' --- FC CKC 26 (col 95) ---
    InserisciGiocatore ws, 95, "Vergara", 71, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 95, "Zaniolo", 38, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 95, "Durosinmi", 19, DateSerial(2026, 2, 11), dictQuot

    ' --- H-Q-A BARCELONA (col 108) ---
    InserisciGiocatore ws, 108, "Malen", 292, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 108, "Taylor K.", 52, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 108, "Britschgi", 9, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 108, "Sulemana I.", 1, DateSerial(2026, 2, 12), dictQuot

    ' --- HELLAS MADONNA (col 69) ---
    InserisciGiocatore ws, 69, "Zaragoza", 10, DateSerial(2026, 2, 12), dictQuot
    InserisciGiocatore ws, 69, "Davis K.", 9, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 69, "Ekkelenkamp", 6, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 69, "Brescianini", 4, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 69, "Cheddira", 1, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 69, "Belghali", 1, DateSerial(2026, 2, 12), dictQuot

    ' --- LINO BANFIELD FC (col 30) ---
    InserisciGiocatore ws, 30, "Celik", 18, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 30, "Vaz", 5, DateSerial(2026, 2, 13), dictQuot
    InserisciGiocatore ws, 30, "Bernasconi", 3, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 30, "Miller L.", 1, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 30, "Obert", 1, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 30, "Marcandalli", 1, DateSerial(2026, 2, 12), dictQuot
    InserisciGiocatore ws, 30, "Bowie", 1, DateSerial(2026, 2, 12), dictQuot

    ' --- LEGENDA AUREA (col 17) ---
    InserisciGiocatore ws, 17, "Bartesaghi", 40, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 17, "Gandelman", 2, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 17, "Dossena", 1, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 17, "Nelsson", 1, DateSerial(2026, 2, 12), dictQuot
    InserisciGiocatore ws, 17, "Barbieri", 1, DateSerial(2026, 2, 12), dictQuot

    ' --- MINNESOTA AL MAX (col 82) ---
    InserisciGiocatore ws, 82, "Marianucci", 4, DateSerial(2026, 2, 12), dictQuot
    InserisciGiocatore ws, 82, "Adzic", 2, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 82, "Mazzitelli", 1, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 82, "Bakola", 1, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 82, "Ratkov", 1, DateSerial(2026, 2, 13), dictQuot

    ' --- PAPAIE TOP TEAM (col 4) ---
    InserisciGiocatore ws, 4, "Solomon", 28, DateSerial(2026, 2, 11), dictQuot

    ' --- FICA (col 56) ---
    InserisciGiocatore ws, 56, "Fullkrug", 1, DateSerial(2026, 2, 13), dictQuot
    InserisciGiocatore ws, 56, "Luis Henrique", 1, DateSerial(2026, 2, 13), dictQuot

    ' --- MASTRI BIRRAI (col 121) ---
    InserisciGiocatore ws, 121, "Ndour", 1, DateSerial(2026, 2, 11), dictQuot
End Sub

' ============================================================
' FASE 4: Riordinamento SQUADRE per reparto/spesa
' Per ogni squadra: ordina ogni sezione di reparto per Spesa
' DESC, Q.acq DESC come tiebreaker. Poi ordina l'elenco
' storico con la stessa logica. Gestisce la colonna Reparto
' (colCalc-1) specifica di FM.
' ============================================================
Private Sub RiordinaSquadreFM(ws As Worksheet)
    Dim teamCols As Variant
    teamCols = Array( _
        Array("Papaie Top Team", 4), Array("Legenda Aurea", 17), _
        Array("Lino Banfield FC", 30), Array("Kung Fu Pandev", 43), _
        Array("FICA", 56), Array("Hellas Madonna", 69), _
        Array("MINNESOTA AL MAX", 82), Array("FC CKC 26", 95), _
        Array("H-Q-A Barcelona", 108), Array("Mastri Birrai", 121))

    Dim t As Long
    For t = LBound(teamCols) To UBound(teamCols)
        Dim tName As String, col As Long
        tName = teamCols(t)(0): col = teamCols(t)(1)

        ' --- Rosa attiva: trova i confini delle sezioni ---
        Dim hdrRows(1 To 10) As Long
        Dim hCount As Long: hCount = 0
        Dim r As Long
        For r = 5 To 52
            If Trim(CStr(ws.Cells(r, col).Value)) = "Calciatore" Then
                hCount = hCount + 1
                hdrRows(hCount) = r
            End If
        Next r

        ' Ordina ogni sezione tra header consecutivi
        Dim h As Long
        For h = 1 To hCount
            Dim secStart As Long, secEnd As Long
            secStart = hdrRows(h) + 1
            If h < hCount Then
                secEnd = hdrRows(h + 1) - 1
            Else
                secEnd = 52
            End If

            ' Salva il valore Reparto (colCalc-1) dalla prima riga non vuota
            Dim repartoVal As String: repartoVal = ""
            For r = secStart To secEnd
                If Trim(CStr(ws.Cells(r, col).Value)) <> "" Then
                    repartoVal = Trim(CStr(ws.Cells(r, col - 1).Value))
                    If Len(repartoVal) > 0 Then Exit For
                End If
            Next r

            ' Ordina la sezione
            OrdinaSezione ws, col, secStart, secEnd

            ' Ripristina colonna Reparto per tutte le righe non vuote
            If Len(repartoVal) > 0 Then
                For r = secStart To secEnd
                    If Trim(CStr(ws.Cells(r, col).Value)) <> "" Then
                        ws.Cells(r, col - 1).Value = repartoVal
                    Else
                        ws.Cells(r, col - 1).ClearContents
                    End If
                Next r
            End If
        Next h

        ' --- Elenco storico: ordina per Spesa DESC ---
        Dim histStart As Long: histStart = 0
        For r = 48 To 200
            If InStr(1, CStr(ws.Cells(r, col).Value), "Elenco storico", vbTextCompare) > 0 Then
                histStart = r
                Exit For
            End If
        Next r

        If histStart > 0 Then
            Dim histDataStart As Long
            histDataStart = histStart + 4
            Dim histDataEnd As Long: histDataEnd = histDataStart - 1
            For r = histDataStart To 570
                If Trim(CStr(ws.Cells(r, col).Value)) <> "" Then
                    histDataEnd = r
                End If
            Next r
            If histDataEnd >= histDataStart Then
                OrdinaSezione ws, col, histDataStart, histDataEnd
            End If
        End If

        Log "    " & tName & ": rosa + storico riordinati"
    Next t
End Sub

' ============================================================
' Helper: Ordina una sezione di giocatori per Spesa DESC,
'         Q.acquisto DESC come tiebreaker (bubble sort)
' firstRow..lastRow = range di righe dati (no header)
' ============================================================
Private Sub OrdinaSezione(ws As Worksheet, colCalc As Long, firstRow As Long, lastRow As Long)
    Dim pCount As Long: pCount = 0
    Dim r As Long
    For r = firstRow To lastRow
        If Trim(CStr(ws.Cells(r, colCalc).Value)) <> "" Then
            pCount = pCount + 1
        End If
    Next r

    If pCount <= 1 Then Exit Sub

    ReDim pData(1 To pCount, 0 To 10) As Variant
    Dim idx As Long: idx = 0
    For r = firstRow To lastRow
        If Trim(CStr(ws.Cells(r, colCalc).Value)) <> "" Then
            idx = idx + 1
            Dim c As Long
            For c = 0 To 10
                pData(idx, c) = ws.Cells(r, colCalc + c).Value
            Next c
        End If
    Next r

    Dim i As Long, j As Long
    For i = 1 To pCount - 1
        For j = 1 To pCount - i
            Dim spJ As Double, spJ1 As Double
            spJ = 0: spJ1 = 0
            If IsNumeric(pData(j, 10)) Then spJ = CDbl(pData(j, 10))
            If IsNumeric(pData(j + 1, 10)) Then spJ1 = CDbl(pData(j + 1, 10))

            Dim doSwap As Boolean: doSwap = False
            If spJ1 > spJ Then
                doSwap = True
            ElseIf spJ1 = spJ Then
                Dim qJ As Double, qJ1 As Double
                qJ = 0: qJ1 = 0
                If IsNumeric(pData(j, 5)) Then qJ = CDbl(pData(j, 5))
                If IsNumeric(pData(j + 1, 5)) Then qJ1 = CDbl(pData(j + 1, 5))
                If qJ1 > qJ Then doSwap = True
            End If

            If doSwap Then
                Dim tmp As Variant
                For c = 0 To 10
                    tmp = pData(j, c)
                    pData(j, c) = pData(j + 1, c)
                    pData(j + 1, c) = tmp
                Next c
            End If
        Next j
    Next i

    ' Separa eventuali celle unite e riscrivi dati ordinati
    ' On Error Resume Next per gestire celle unite residue o protette
    On Error Resume Next
    ws.Range(ws.Cells(firstRow, colCalc), ws.Cells(lastRow, colCalc + 10)).UnMerge
    idx = 0
    For r = firstRow To lastRow
        idx = idx + 1
        If idx <= pCount Then
            For c = 0 To 10
                ws.Cells(r, colCalc + c).Value = pData(idx, c)
            Next c
        Else
            For c = 0 To 10
                ws.Cells(r, colCalc + c).ClearContents
            Next c
        End If
    Next r
    On Error GoTo 0
End Sub

' ============================================================
' FASE 5: Registra tutte le assicurazioni
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
    Dim logLines() As String, logR As Long
    logLines = Split(logText, vbCrLf)
    For logR = 0 To UBound(logLines)
        wsLog.Cells(logR + 1, 1).Value = logLines(logR)
    Next logR

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
' Helper: Normalizza un nome (uppercase, senza accenti/apostrofi/punti)
' ============================================================
Private Function NormName(ByVal s As String) As String
    Dim n As String
    n = UCase(Trim(s))
    n = Replace(n, "'", "")
    n = Replace(n, ".", "")
    n = Replace(n, Chr(232), "E")  ' e-grave
    n = Replace(n, Chr(233), "E")  ' e-acuto
    n = Replace(n, Chr(236), "I")  ' i-grave
    n = Replace(n, Chr(237), "I")  ' i-acuto
    n = Replace(n, Chr(242), "O")  ' o-grave
    n = Replace(n, Chr(243), "O")  ' o-acuto
    n = Replace(n, Chr(224), "A")  ' a-grave
    n = Replace(n, Chr(225), "A")  ' a-acuto
    n = Replace(n, Chr(249), "U")  ' u-grave
    n = Replace(n, Chr(250), "U")  ' u-acuto
    NormName = n
End Function

' ============================================================
' Helper: Inserisci un giocatore nella sezione corretta di SQUADRE
' - Trova la sezione per reparto (P/D/C/A) tramite lookup in dictQuot
' - NON scrive +1 (Ruolo) e +2 (Squadra): contengono formule CERCA.VERT
' - Scrive: +0 Nome, +4 Data acquisto, +5 Qt.acq, +6 FVM, +8 "/", +10 Spesa
' ============================================================
Private Sub InserisciGiocatore(ws As Worksheet, colCalc As Long, nome As String, spesa As Long, dataAcq As Date, dictQuot As Object)
    Dim nomeUp As String
    nomeUp = UCase(Trim(nome))
    Dim lookupKey As String
    lookupKey = NormName(nome)

    ' 1. Verifica se il giocatore esiste gia' (righe 6-52, tutte le sezioni)
    Dim r As Long
    For r = 6 To 52
        Dim existName As String
        existName = Trim(CStr(ws.Cells(r, colCalc).Value))
        If existName <> "" Then
            If NormName(existName) = lookupKey Then
                Log "  GIA' PRESENTE: " & nome & " (riga " & r & ", come '" & existName & "')"
                ' Aggiorna spesa se diversa
                If ws.Cells(r, colCalc + 10).Value <> spesa Then
                    ws.Cells(r, colCalc + 10).Value = spesa
                    Log "    -> Spesa aggiornata a " & spesa
                End If
                ' Compila dati mancanti se vuoti
                Dim cv4 As Variant: cv4 = ws.Cells(r, colCalc + 4).Value
                If IsEmpty(cv4) Or cv4 = "" Or (IsNumeric(cv4) And CLng(cv4) = 0) Then
                    ws.Cells(r, colCalc + 4).Value = CLng(dataAcq)
                    Log "    -> Data acquisto compilata"
                End If
                If IsEmpty(ws.Cells(r, colCalc + 8).Value) Or ws.Cells(r, colCalc + 8).Value = "" Then
                    ws.Cells(r, colCalc + 8).Value = "/"
                End If
                ' Qt e FVM da dictQuot se mancanti
                If dictQuot.Exists(lookupKey) Then
                    Dim cv5 As Variant: cv5 = ws.Cells(r, colCalc + 5).Value
                    If IsEmpty(cv5) Or cv5 = "" Or (IsNumeric(cv5) And CLng(cv5) = 0) Then
                        ws.Cells(r, colCalc + 5).Value = dictQuot(lookupKey)(1)
                        Log "    -> Qt.acq compilata"
                    End If
                    Dim cv6 As Variant: cv6 = ws.Cells(r, colCalc + 6).Value
                    If IsEmpty(cv6) Or cv6 = "" Or (IsNumeric(cv6) And CLng(cv6) = 0) Then
                        ws.Cells(r, colCalc + 6).Value = dictQuot(lookupKey)(2)
                        Log "    -> FVM compilato"
                    End If
                End If
                Exit Sub
            End If
        End If
    Next r

    ' 2. Lookup ruolo (P/D/C/A), Qt, FVM da dictQuot
    Dim lookupKey2 As String
    lookupKey2 = NormName(nome)
    Dim ruolo As String, qtAcq As Variant, fvm As Variant
    ruolo = ""
    qtAcq = Empty
    fvm = Empty
    If dictQuot.Exists(lookupKey2) Then
        ruolo = CStr(dictQuot(lookupKey2)(0))  ' Ruolo Classic
        qtAcq = dictQuot(lookupKey2)(1)         ' Qt.Attuale
        fvm = dictQuot(lookupKey2)(2)           ' FVM
    Else
        Log "  ATTENZIONE: " & nome & " non trovato in LISTA - ruolo/Qt/FVM non disponibili"
    End If

    ' 3. Trova le sezioni per reparto contando gli header "Calciatore"
    ' FM: 1a occorrenza=P, 2a=D, 3a=C, 4a=A
    Dim hdrRows(1 To 4) As Long
    Dim hCount As Long: hCount = 0
    For r = 5 To 52
        If Trim(CStr(ws.Cells(r, colCalc).Value)) = "Calciatore" Then
            hCount = hCount + 1
            If hCount <= 4 Then hdrRows(hCount) = r
        End If
    Next r

    Dim targetSec As Long: targetSec = 0
    Select Case UCase(Left(ruolo, 1))
        Case "P": targetSec = 1
        Case "D": targetSec = 2
        Case "C": targetSec = 3
        Case "A": targetSec = 4
    End Select

    If targetSec = 0 Or targetSec > hCount Then
        Log "  ERRORE: Ruolo '" & ruolo & "' non valido per " & nome & " (col " & colCalc & ")"
        Exit Sub
    End If

    ' 4. Trova una riga vuota nella sezione corretta
    Dim secStart As Long, secEnd As Long
    secStart = hdrRows(targetSec) + 1
    If targetSec < hCount Then
        secEnd = hdrRows(targetSec + 1) - 1
    Else
        secEnd = 52
    End If

    For r = secStart To secEnd
        If Trim(CStr(ws.Cells(r, colCalc).Value)) = "" Then
            ' Scrivi dati giocatore
            ws.Cells(r, colCalc).Value = nomeUp              ' +0: Nome (UPPERCASE)
            ' +1 (Ruolo) e +2 (Squadra): NON toccare - formule CERCA.VERT
            ws.Cells(r, colCalc + 4).Value = CLng(dataAcq)   ' +4: Data acquisto (serial)
            If Not IsEmpty(qtAcq) Then
                ws.Cells(r, colCalc + 5).Value = qtAcq        ' +5: Qt all'acquisto
            End If
            If Not IsEmpty(fvm) Then
                ws.Cells(r, colCalc + 6).Value = fvm          ' +6: FVM Prop. all'acquisto
            End If
            ws.Cells(r, colCalc + 8).Value = "/"              ' +8: Qt rinn ass
            ws.Cells(r, colCalc + 10).Value = spesa           ' +10: Spesa
            Log "  INSERITO: " & nomeUp & " (Sp=" & spesa & ", Sez=" & Choose(targetSec, "P", "D", "C", "A") & ") -> riga " & r
            Exit Sub
        End If
    Next r

    Log "  ERRORE: Nessuna riga vuota nella sezione " & Choose(targetSec, "P", "D", "C", "A") & " per " & nome & " (col " & colCalc & ")"
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
