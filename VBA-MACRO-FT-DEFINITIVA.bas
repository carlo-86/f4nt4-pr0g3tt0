' ============================================================
' MACRO VBA DEFINITIVA - FANTA TOSTI 2026
' Aggiornamento completo DB: Listone + Svincoli + Asta Riparazione + Riordinamento SQUADRE + Allineamento Date + Assicurazioni + Fix Formule DB + Contratti Invernali
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
Private Const SHEET_PWD As String = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99"
Private logText As String

' ============================================================
' MAIN: Esegue tutte le operazioni in sequenza
' ============================================================
Sub ESEGUI_TUTTO_FT()
    logText = "=== LOG OPERAZIONI FT - MACRO DEFINITIVA ===" & vbCrLf & vbCrLf

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
    ' (servono in FASE 2 per Qt all'acquisto e FVM dei nuovi giocatori)
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

    ' FASE 0: Aggiorna LISTA dal listone (+ delistati a 601+ e date nascita)
    Log "FASE 0: Aggiornamento LISTA dal listone"
    Log "-------------------------------------------"
    AggiornaListone wsLista, False  ' False = Classic (FT)

    Dim wsSq As Worksheet
    Set wsSq = ThisWorkbook.Sheets("SQUADRE")

    ' FASE 1: Svincoli (rosa attiva -> elenco storico)
    Log ""
    Log "FASE 1: Svincoli (spostamento nell'elenco storico)"
    Log "-------------------------------------------"
    EseguiSvincoliFT wsSq

    ' FASE 2: Asta riparazione (inserimento nuovi giocatori)
    Log ""
    Log "FASE 2: Inserimento giocatori asta riparazione"
    Log "-------------------------------------------"
    InserisciAstaRiparazione wsSq, dictQuot

    ' FASE 3: Riordinamento SQUADRE per reparto/spesa
    Log ""
    Log "FASE 3: Riordinamento SQUADRE per reparto e spesa"
    Log "-------------------------------------------"
    RiordinaSquadreFT wsSq

    ' FASE 4: Allineamento retroattivo date assicurazione (regola triennio rigido)
    Log ""
    Log "FASE 4: Allineamento date assicurazione al ciclo triennale"
    Log "-------------------------------------------"
    AllineaDatePreventive wsSq

    ' FASE 5: Registra assicurazioni
    Log ""
    Log "FASE 5: Registrazione assicurazioni"
    Log "-------------------------------------------"
    RegistraAssicurazioni wsSq

    ' FASE 6: Correggi formule foglio DB (audit 22/02/2026)
    Log ""
    Log "FASE 6: Correzione formule foglio DB"
    Log "-------------------------------------------"
    CorreggiFormuleDB

    ' FASE 7: Calcola e annota quote contratti mercato di riparazione
    Log ""
    Log "FASE 7: Calcolo quote contratti invernali (QUOTE+MONTEPREMI 2026)"
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

    ' Trova l'ultima riga occupata nella sezione storica (601+)
    Dim lastHistRow As Long: lastHistRow = 600
    For scanR = 601 To 2000
        If Trim(CStr(wsLista.Cells(scanR, 1).Value)) <> "" Then
            lastHistRow = scanR
        End If
    Next scanR

    ' Cerca nella tabella principale gli ID non presenti nel listone
    For scanR = 2 To lastListaRow
        Dim listaId As Variant
        listaId = wsLista.Cells(scanR, 1).Value
        If Trim(CStr(listaId)) <> "" Then
            On Error Resume Next
            Dim idNum As Long
            idNum = CLng(listaId)
            On Error GoTo 0
            If idNum > 0 And Not dictListoneIDs.Exists(idNum) Then
                ' Delistato! Sposta nella sezione storica
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
    ' 6. COMPATTAMENTO tabella principale (rimuovi righe vuote)
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
    lastListaRow = mainCount + 1  ' Aggiorna lastListaRow dopo compattamento
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
            ' newInfo = "nome|squadra"
            Dim pipePos As Long: pipePos = InStr(newInfo, "|")
            Dim newNome As String: newNome = Left(newInfo, pipePos - 1)
            Dim newSquadra As String: newSquadra = Mid(newInfo, pipePos + 1)

            ' Verifica se il calciatore ha gia' una voce in L:N
            Dim alreadyInLN As Boolean: alreadyInLN = False
            For scanR = 1 To 2000
                If UCase(Trim(CStr(wsLista.Cells(scanR, 12).Value))) = UCase(newNome) Then
                    alreadyInLN = True
                    Exit For
                End If
            Next scanR

            If Not alreadyInLN Then
                ' Trova la prima riga vuota in colonna L
                Dim lnRow As Long: lnRow = 0
                For scanR = 2 To 2000
                    If Trim(CStr(wsLista.Cells(scanR, 12).Value)) = "" Then
                        lnRow = scanR
                        Exit For
                    End If
                Next scanR

                If lnRow > 0 Then
                    wsLista.Cells(lnRow, 12).Value = newNome
                    ' Tenta ricerca data nascita online
                    Dim birthDate As Variant
                    birthDate = CercaDataNascitaOnline(newId, newNome, newSquadra)
                    If Not IsEmpty(birthDate) Then
                        wsLista.Cells(lnRow, 13).Value = birthDate
                        wsLista.Cells(lnRow, 13).NumberFormat = "dd/mm/yyyy"
                        Log "    " & newNome & " (ID " & newId & "): data nascita trovata"
                    Else
                        Log "    " & newNome & " (ID " & newId & "): data nascita NON trovata - inserire manualmente"
                    End If
                    ' Formula eta' in colonna N
                    wsLista.Cells(lnRow, 14).Formula = _
                        "=IF(M" & lnRow & "="""","""",INT(($N$1-M" & lnRow & ")/365.25))"
                End If
            End If
        Next nk
    End If

    ' ========================================
    ' 8. Formula Eta' per le righe della tabella principale
    ' ========================================
    Log "  Aggiornamento formula Eta'..."
    Dim etaR As Long
    For etaR = 2 To lastListaRow
        If Trim(CStr(wsLista.Cells(etaR, 2).Value)) <> "" Then
            Dim etaVal As Variant
            etaVal = wsLista.Cells(etaR, 9).Value
            Dim needsEta As Boolean
            needsEta = False
            If IsEmpty(etaVal) Then
                needsEta = True
            ElseIf IsError(etaVal) Then
                needsEta = True
            ElseIf etaVal = "" Then
                needsEta = True
            End If
            If needsEta Then
                wsLista.Cells(etaR, 9).Formula = _
                    "=IFERROR(VLOOKUP(B" & etaR & ",$L:$N,3,FALSE),"""")"
            End If
        End If
    Next etaR

    ' ========================================
    ' 9. Ordina tabella principale (A1:I) per Calciatore A->Z
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
    ' 10. Ordina sezione storica (601+) per Calciatore A->Z
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

        ' Traccia ID nel dizionario
        On Error Resume Next
        dictIDs.Item(CLng(listoneId)) = True
        On Error GoTo 0

        ' Cerca per ID nella LISTA (col A, solo tabella principale 1-600)
        Dim matchRow As Variant
        matchRow = Application.Match(CLng(listoneId), wsLista.Range("A1:A600"), 0)

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
            ' Controlla anche nella sezione storica (601+) - potrebbe essere stato delistato prima
            Dim histMatch As Variant
            histMatch = Empty
            Dim sR As Long
            For sR = 601 To 2000
                If Trim(CStr(wsLista.Cells(sR, 1).Value)) = "" Then GoTo SkipHistRow
                If CLng(wsLista.Cells(sR, 1).Value) = CLng(listoneId) Then
                    histMatch = sR
                    Exit For
                End If
SkipHistRow:
            Next sR

            If Not IsEmpty(histMatch) Then
                ' Era nella sezione storica: ri-sposta nella tabella principale
                lastListaRow = lastListaRow + 1
                Dim colC As Long
                For colC = 1 To 9
                    wsLista.Cells(lastListaRow, colC).Value = wsLista.Cells(CLng(histMatch), colC).Value
                Next colC
                For colC = 1 To 9
                    wsLista.Cells(CLng(histMatch), colC).ClearContents
                Next colC
                ' Aggiorna con i nuovi dati
                wsLista.Cells(lastListaRow, 2).Value = nomeClean
                wsLista.Cells(lastListaRow, 3).Value = ruoloC
                wsLista.Cells(lastListaRow, 4).Value = ruoloM
                wsLista.Cells(lastListaRow, 5).Value = squadra
                wsLista.Cells(lastListaRow, 6).Value = qtAttuale
                wsLista.Cells(lastListaRow, 7).Value = qtIniziale
                wsLista.Cells(lastListaRow, 8).Value = fvmVal
                aggiornati = aggiornati + 1
            Else
                ' Veramente nuovo: aggiungi nella tabella principale
                lastListaRow = lastListaRow + 1
                If lastListaRow > 600 Then
                    Log "    ATTENZIONE: Tabella principale piena (>600 righe), " & nomeClean & " non aggiunto"
                    lastListaRow = lastListaRow - 1
                    skippati = skippati + 1
                    GoTo NextRow
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
                ' Traccia come nuovo per ricerca data nascita
                dictNew.Item(CLng(listoneId)) = nomeClean & "|" & squadra
            End If
        End If

NextRow:
        r = r + 1
    Loop
End Sub

' ============================================================
' Helper: Cerca data di nascita online da fantacalcio.it
' Ritorna la data come Variant (Date) o Empty se non trovata.
' URL: https://www.fantacalcio.it/serie-a/squadre/{team}/{player}/{id}
' ============================================================
Private Function CercaDataNascitaOnline(playerID As Long, playerName As String, squadraAbbr As String) As Variant
    CercaDataNascitaOnline = Empty
    On Error GoTo BirthDateErr

    ' Mappa abbreviazioni squadre -> nomi URL
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

    ' Formatta nome per URL (lowercase, spazi->trattini, rimuovi accenti)
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

    If http.Status <> 200 Then GoTo BirthDateErr

    Dim html As String
    html = http.responseText

    ' Cerca pattern data nascita: "dd/mm/yyyy" vicino a "nascita" o "nato"
    Dim pos As Long
    pos = InStr(1, html, "nascita", vbTextCompare)
    If pos = 0 Then pos = InStr(1, html, "nato il", vbTextCompare)
    If pos = 0 Then GoTo BirthDateErr

    ' Cerca un pattern data (dd/mm/yyyy) nei 200 caratteri successivi
    Dim chunk As String
    chunk = Mid(html, pos, 200)
    Dim datePattern As String
    Dim ci As Long
    For ci = 1 To Len(chunk) - 9
        Dim ch As String
        ch = Mid(chunk, ci, 10)
        ' Pattern: dd/mm/yyyy
        If Mid(ch, 3, 1) = "/" And Mid(ch, 6, 1) = "/" Then
            If IsNumeric(Left(ch, 2)) And IsNumeric(Mid(ch, 4, 2)) And IsNumeric(Right(ch, 4)) Then
                datePattern = ch
                Exit For
            End If
        End If
    Next ci

    If Len(datePattern) > 0 Then
        CercaDataNascitaOnline = CDate(datePattern)
    End If

    Set http = Nothing
    Exit Function

BirthDateErr:
    CercaDataNascitaOnline = Empty
    On Error Resume Next
    Set http = Nothing
    On Error GoTo 0
End Function

' ============================================================
' FASE 1: Svincoli (sposta giocatori dalla rosa attiva
'         all'elenco storico dei calciatori ceduti)
' ============================================================
Private Sub EseguiSvincoliFT(ws As Worksheet)
    ' --- FCK DEPORTIVO (col 3) - usa asterischi nell'elenco storico ---
    SvincolaGiocatore ws, 3, "Viti", True
    SvincolaGiocatore ws, 3, "Guendouzi", True

    ' --- HELLAS MADONNA (col 15) ---
    SvincolaGiocatore ws, 15, "Sava", False
    SvincolaGiocatore ws, 15, "Floriani", False       ' Floriani Mussolini
    SvincolaGiocatore ws, 15, "Dzeko", False
    SvincolaGiocatore ws, 15, "Sanabria", False
    SvincolaGiocatore ws, 15, "Sohm", False

    ' --- MUTTLEY SUPERSTAR (col 27) ---
    SvincolaGiocatore ws, 27, "Bravo", False
    SvincolaGiocatore ws, 27, "Lang", False

    ' --- PARTIZAN (col 39) ---
    SvincolaGiocatore ws, 39, "Vasquez D", False
    SvincolaGiocatore ws, 39, "Lovik", False
    SvincolaGiocatore ws, 39, "Stanciu", False

    ' --- LEGENDA AUREA (col 51) ---
    SvincolaGiocatore ws, 51, "Biraghi", False
    SvincolaGiocatore ws, 51, "Fazzini", False
    SvincolaGiocatore ws, 51, "Venturino", False
    SvincolaGiocatore ws, 51, "Collocolo", False
    SvincolaGiocatore ws, 51, "Gronbaek", False
    SvincolaGiocatore ws, 51, "Cutrone", False
    SvincolaGiocatore ws, 51, "Lucca", False
    SvincolaGiocatore ws, 51, "Belotti", False

    ' --- KUNG FU PANDEV (col 63) ---
    SvincolaGiocatore ws, 63, "Lazzari", False
    SvincolaGiocatore ws, 63, "Vazquez", False
    SvincolaGiocatore ws, 63, "Luvumbo", False
    SvincolaGiocatore ws, 63, "Tameze", False

    ' --- MILLWALL (col 75) ---
    SvincolaGiocatore ws, 75, "Israel", False
    SvincolaGiocatore ws, 75, "Leali", False
    SvincolaGiocatore ws, 75, "Posch", False
    SvincolaGiocatore ws, 75, "Bailey", False
    SvincolaGiocatore ws, 75, "Castellanos", False
    SvincolaGiocatore ws, 75, "Lookman", False

    ' --- FC CKC 26 (col 87) ---
    SvincolaGiocatore ws, 87, "Mandas", False
    SvincolaGiocatore ws, 87, "Ghilardi", False
    SvincolaGiocatore ws, 87, "Zemura", False
    SvincolaGiocatore ws, 87, "Carboni V", False
    SvincolaGiocatore ws, 87, "Gilmour", False
    SvincolaGiocatore ws, 87, "Anjorin", False
    SvincolaGiocatore ws, 87, "Immobile", False
    SvincolaGiocatore ws, 87, "Ngonge", False
    SvincolaGiocatore ws, 87, "Almqvist", False

    ' --- PAPAIE TOP TEAM (col 99) ---
    SvincolaGiocatore ws, 99, "Sportiello", False

    ' --- TRONZANO (col 111) ---
    SvincolaGiocatore ws, 111, "Mari", False          ' Mari'
    SvincolaGiocatore ws, 111, "Tete", False           ' Tete Morente
End Sub

' ============================================================
' Helper: Svincola un giocatore (rosa attiva -> elenco storico)
' Cerca il giocatore nella rosa attiva (righe 6-52), copia
' tutti i dati nell'elenco storico sottostante, cancella la
' riga dalla rosa attiva. useAsterisk=True per FCK (FT) e
' Papaie (FM) che usano * nei nomi storici.
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
        If Len(nomeCell) = 0 Or nomeCell = "CALCIATORE" Then GoTo NextSvRow
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

            ' Gestisci asterischi se necessario (solo FCK per FT, Papaie per FM)
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
NextSvRow:
    Next r

    Log "  NON TROVATO per svincolo: " & nomeCerca & " nella colonna " & colCalc
    SvincolaGiocatore = False
End Function

' ============================================================
' FASE 2: Inserisci giocatori acquisiti nell'asta riparazione
' ============================================================
Private Sub InserisciAstaRiparazione(ws As Worksheet, dictQuot As Object)
    ' 38 giocatori totali - date da "Mercato ASTA CLASSICA"
    ' Formato: InserisciGiocatore ws, colCalc, nome, spesa, dataAcquisto, dictQuot
    ' NON scrive Ruolo (+1) e Squadra (+2): sono formule CERCA.VERT
    ' Scrive: Nome (+0), DataAcq (+4), Qt.Acq (+5), FVM (+6), "/" (+8), Spesa (+10)

    ' --- KUNG FU PANDEV (col 63) ---
    InserisciGiocatore ws, 63, "Malen", 173, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 63, "Vergara", 31, DateSerial(2026, 2, 11), dictQuot

    ' --- FC CKC 26 (col 87) ---
    InserisciGiocatore ws, 87, "Kone' I.", 16, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 87, "Bijlow", 2, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 87, "Bernasconi", 1, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 87, "Santos A.", 1, DateSerial(2026, 2, 12), dictQuot
    InserisciGiocatore ws, 87, "Baldanzi", 3, DateSerial(2026, 2, 12), dictQuot
    InserisciGiocatore ws, 87, "Muharemovic", 3, DateSerial(2026, 2, 12), dictQuot
    InserisciGiocatore ws, 87, "Vaz", 3, DateSerial(2026, 2, 12), dictQuot
    InserisciGiocatore ws, 87, "Tiago Gabriel", 1, DateSerial(2026, 2, 12), dictQuot

    ' --- HELLAS MADONNA (col 15) ---
    InserisciGiocatore ws, 15, "Moreo", 1, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 15, "Durosinmi", 3, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 15, "Berisha M.", 1, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 15, "Sportiello", 1, DateSerial(2026, 2, 12), dictQuot
    InserisciGiocatore ws, 15, "Circati", 1, DateSerial(2026, 2, 12), dictQuot

    ' --- MUTTLEY SUPERSTAR (col 27) ---
    InserisciGiocatore ws, 27, "Solomon", 21, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 27, "Ostigard", 2, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 27, "Luis Henrique", 17, DateSerial(2026, 2, 12), dictQuot

    ' --- MILLWALL (col 75) ---
    InserisciGiocatore ws, 75, "Zaragoza", 46, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 75, "Muric", 21, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 75, "Perrone", 9, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 75, "Celik", 2, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 75, "Ratkov", 2, DateSerial(2026, 2, 12), dictQuot
    InserisciGiocatore ws, 75, "Boga", 1, DateSerial(2026, 2, 13), dictQuot
    InserisciGiocatore ws, 75, "Paleari", 1, DateSerial(2026, 2, 13), dictQuot

    ' --- LEGENDA AUREA (col 51) ---
    InserisciGiocatore ws, 51, "Raspadori", 50, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 51, "Taylor K.", 42, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 51, "Vitinha O.", 5, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 51, "Bartesaghi", 3, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 51, "Bonazzoli", 2, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 51, "Ekkelenkamp", 4, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 51, "Fagioli", 1, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 51, "Miretti", 1, DateSerial(2026, 2, 12), dictQuot

    ' --- PARTIZAN (col 39) ---
    InserisciGiocatore ws, 39, "Strefezza", 29, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 39, "Belghali", 7, DateSerial(2026, 2, 11), dictQuot
    InserisciGiocatore ws, 39, "Przyborek", 1, DateSerial(2026, 2, 11), dictQuot

    ' --- A.S. TRONZANO (col 111) ---
    InserisciGiocatore ws, 111, "Fullkrug", 2, DateSerial(2026, 2, 11), dictQuot
End Sub

' ============================================================
' FASE 3: Riordinamento SQUADRE per reparto/spesa
' Per ogni squadra: ordina ogni sezione di reparto per Spesa
' DESC, Q.acq DESC come tiebreaker. Poi ordina l'elenco
' storico con la stessa logica (ma senza divisione per reparto).
' ============================================================
Private Sub RiordinaSquadreFT(ws As Worksheet)
    Dim teamCols As Variant
    teamCols = Array( _
        Array("FCK Deportivo", 3), Array("Hellas Madonna", 15), _
        Array("muttley superstar", 27), Array("PARTIZAN", 39), _
        Array("Legenda Aurea", 51), Array("Kung Fu Pandev", 63), _
        Array("Millwall", 75), Array("FC CKC 26", 87), _
        Array("Papaie Top Team", 99), Array("Tronzano", 111))

    Dim t As Long
    For t = LBound(teamCols) To UBound(teamCols)
        Dim tName As String, col As Long
        tName = teamCols(t)(0): col = teamCols(t)(1)

        ' --- Rosa attiva: trova i confini delle sezioni ---
        ' Le sezioni sono delimitate dalle righe dove col+0 = "Calciatore"
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
            OrdinaSezione ws, col, secStart, secEnd
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
    ' Conta giocatori non vuoti
    Dim pCount As Long: pCount = 0
    Dim r As Long
    For r = firstRow To lastRow
        If Trim(CStr(ws.Cells(r, colCalc).Value)) <> "" Then
            pCount = pCount + 1
        End If
    Next r

    If pCount <= 1 Then Exit Sub

    ' Leggi dati giocatori in array 2D
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

    ' Bubble sort DESC: Spesa (+10), tiebreak Q.acquisto (+5)
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
           vbInformation, "Fix Formule DB - FT"
End Sub

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
    Dim ruolo As String, qtAcq As Variant, fvm As Variant
    ruolo = ""
    qtAcq = Empty
    fvm = Empty
    If dictQuot.Exists(lookupKey) Then
        ruolo = CStr(dictQuot(lookupKey)(0))  ' Ruolo Classic
        qtAcq = dictQuot(lookupKey)(1)         ' Qt.Attuale
        fvm = dictQuot(lookupKey)(2)           ' FVM
    Else
        Log "  ATTENZIONE: " & nome & " non trovato in LISTA - ruolo/Qt/FVM non disponibili"
    End If

    ' 3. Trova le sezioni per reparto contando gli header "Calciatore"
    ' FT: 1a occorrenza=P, 2a=D, 3a=C, 4a=A
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
