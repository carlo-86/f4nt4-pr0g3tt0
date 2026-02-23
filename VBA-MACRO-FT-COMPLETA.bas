' ============================================================
' MACRO VBA COMPLETA - FANTA TOSTI 2026
' Aggiornamento DB: Asta Riparazione + Assicurazioni
' Mercato Invernale 2026
' ============================================================
' ISTRUZIONI:
' 1. BACKUP del file DB prima di procedere!
' 2. Aprire "Fanta Tosti 2026 - DB completo (06.02.2026).xlsx"
' 3. Alt+F11 > Inserisci > Modulo > Incolla tutto questo codice
' 4. F5 > Seleziona "ESEGUI_TUTTO_FT" > Esegui
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
    logText = "=== LOG OPERAZIONI FT ===" & vbCrLf & vbCrLf

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SQUADRE")

    ' FASE 1: Aggiungi giocatori asta riparazione (post 06/02)
    Log "FASE 1: Inserimento giocatori asta riparazione"
    Log "-------------------------------------------"
    InserisciAstaRiparazione ws

    ' FASE 2: Registra assicurazioni
    Log ""
    Log "FASE 2: Registrazione assicurazioni"
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

    MsgBox "Operazioni FT completate!" & vbCrLf & _
           "Controlla il foglio LOG_MACRO per i dettagli.", _
           vbInformation, "Fanta Tosti 2026"
End Sub

' ============================================================
' FASE 1: Inserisci giocatori acquisiti nell'asta riparazione
' Ogni giocatore va nella prima riga vuota della colonna squadra
' ============================================================
Private Sub InserisciAstaRiparazione(ws As Worksheet)
    ' --- HELLAS MADONNA (col 15) ---
    ' Sportiello e Moreo gia' presenti dal 06/02
    InserisciGiocatore ws, 15, "Circati", "Par", 1
    InserisciGiocatore ws, 15, "Berisha M.", "Lec", 1
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
    InserisciGiocatore ws, 87, "Santos A.", "Nap", 1
    InserisciGiocatore ws, 87, "Bijlow", "Fio", 2
    InserisciGiocatore ws, 87, "Bernasconi", "Ata", 1

    ' --- MUTTLEY SUPERSTAR (col 27) ---
    InserisciGiocatore ws, 27, "Ostigard", "Gen", 2
    InserisciGiocatore ws, 27, "Luis Henrique", "Int", 17
    InserisciGiocatore ws, 27, "Solomon", "Tor", 21

    ' --- MILLWALL (col 75) ---
    ' Holm gia' presente (Sp=5)
    InserisciGiocatore ws, 75, "Celik", "Rom", 2
    InserisciGiocatore ws, 75, "Ratkov", "Laz", 2
    InserisciGiocatore ws, 75, "Zaragoza", "Mon", 46
    InserisciGiocatore ws, 75, "Perrone", "Mon", 9
    InserisciGiocatore ws, 75, "Paleari", "Cag", 1
    InserisciGiocatore ws, 75, "Boga", "Niz", 1

    ' --- LEGENDA AUREA (col 51) ---
    ' Di Gregorio, Sommer, Martinez Jo., Kalulu, Lovric, Vitinha gia' presenti
    InserisciGiocatore ws, 51, "Bartesaghi", "Mil", 3
    InserisciGiocatore ws, 51, "Taylor K.", "Laz", 42
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
    ' KOUAME: NON assicurabile (non piu' listato)
    Log "  SKIP: Kouame' - non piu' listato, non assicurabile"

    ' --- FC CKC 26 (col 87) ---
    AssicuraG ws, 87, "Tiago Gabriel"
    AssicuraG ws, 87, "Vaz"
    AssicuraG ws, 87, "Muharemovic"
    AssicuraG ws, 87, "Baldanzi"
    AssicuraG ws, 87, "Santos"       ' = "Allison S." nella comunicazione
    AssicuraG ws, 87, "Bijlow"
    AssicuraG ws, 87, "Bernasconi"
    AssicuraG ws, 87, "Kon"          ' match parziale per Kone' I.

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
' Helper: Inserisci un giocatore nuovo in SQUADRE
' Cerca la prima riga vuota nella colonna della squadra (righe 6-50)
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
