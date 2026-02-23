' ============================================================
' MACRO VBA COMPLETA - FANTAMANTRA MANAGERIALE 2026
' Aggiornamento DB: Scambi + Asta Riparazione + Assicurazioni
' Mercato Invernale 2026
' ============================================================
' ISTRUZIONI:
' 1. BACKUP del file DB prima di procedere!
' 2. Aprire "FantaMantra Manageriale - DB completo (06.02.2026).xlsx"
' 3. Alt+F11 > Inserisci > Modulo > Incolla tutto questo codice
' 4. F5 > Seleziona "ESEGUI_TUTTO_FM" > Esegui
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
Private logText As String

' ============================================================
' MAIN: Esegue tutte le operazioni in sequenza
' ============================================================
Sub ESEGUI_TUTTO_FM()
    logText = "=== LOG OPERAZIONI FM ===" & vbCrLf & vbCrLf

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SQUADRE")

    ' FASE 1: Gestisci scambi post-06/02
    Log "FASE 1: Scambi post-06/02 (spostamento giocatori)"
    Log "-------------------------------------------"
    GestisciScambi ws

    ' FASE 2: Aggiungi giocatori asta riparazione (post 06/02)
    Log ""
    Log "FASE 2: Inserimento giocatori asta riparazione"
    Log "-------------------------------------------"
    InserisciAstaRiparazione ws

    ' FASE 3: Registra assicurazioni
    Log ""
    Log "FASE 3: Registrazione assicurazioni"
    Log "-------------------------------------------"
    RegistraAssicurazioni ws

    ' Mostra log
    Log ""
    Log "=== COMPLETATO ==="

    ' Crea foglio log
    Dim wsLog As Worksheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("LOG_MACRO")
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
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
    ' NOTA: Kouame' NON inserito via asta rip. — e' gia' in rosa (non svincolato)
    ' ma non piu' listato su Leghe Fantacalcio, non assicurabile

    ' --- FC CKC 26 (col 95) ---
    InserisciGiocatore ws, 95, "Durosinmi", "Pis", 19
    InserisciGiocatore ws, 95, "Vergara", "Nap", 71

    ' --- H-Q-A BARCELONA (col 108) ---
    InserisciGiocatore ws, 108, "Britschgi", "Par", 9
    InserisciGiocatore ws, 108, "Taylor K.", "Laz", 52
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
    AssicuraG ws, 30, "Tavares"
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

            ws.Cells(r, colCalc + 3).Value = "A"
            ws.Cells(r, colCalc + 7).Value = CDate(DATA_ASS)
            ws.Cells(r, colCalc + 7).NumberFormat = "dd/mm/yyyy"

            If vecchioFlag = "A" Then
                Log "  RINNOVO: " & ws.Cells(r, colCalc).Value & " (riga " & r & ") - era gia' assicurato"
            Else
                Log "  ASSICURATO: " & ws.Cells(r, colCalc).Value & " (riga " & r & ")"
            End If
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
