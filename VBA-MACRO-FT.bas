' ============================================================
' MACRO VBA - FANTA TOSTI 2026
' Aggiornamento Assicurazioni - Mercato Invernale 2026
' ============================================================
' ISTRUZIONI:
' 1. Aprire il file "Fanta Tosti 2026 - DB completo (06.02.2026).xlsx"
' 2. Premere Alt+F11 per aprire l'editor VBA
' 3. Menu Inserisci > Modulo
' 4. Incollare TUTTO questo codice nel modulo
' 5. Premere F5 oppure andare su Esegui > Esegui Sub
' 6. Selezionare "AggiornaAssicurazioniFT" e fare clic su Esegui
' ============================================================

' Colonne relative per ogni blocco squadra in SQUADRE (offset da colonna Calciatore):
' +0  = Calciatore (nome)
' +3  = Flag assicurazione ("A")
' +7  = Data assicurazione/rinnovo
' +10 = Spesa

' Posizioni colonna Calciatore per ogni squadra FT:
' FCK Deportivo = Col C (3)  -> col index 3
' Hellas Madonna = Col O (15) -> col index 15
' muttley superstar = Col AA (27) -> col index 27
' PARTIZAN = Col AM (39) -> col index 39
' Legenda Aurea = Col AY (51) -> col index 51
' Kung Fu Pandev = Col BK (63) -> col index 63
' Millwall = Col BW (75) -> col index 75
' FC CKC 26 = Col CI (87) -> col index 87
' Papaie Top Team = Col CU (99) -> col index 99
' Tronzano = Col DG (111) -> col index 111
' NOTA: In VBA le colonne sono 1-based, non 0-based come in JavaScript!

Private Const DATA_ASS As String = "14/02/2026"

Sub AggiornaAssicurazioniFT()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SQUADRE")

    Dim count As Long
    count = 0

    ' === HELLAS MADONNA (Col Calciatore = 15) ===
    count = count + AssicuraGiocatore(ws, 15, "Sportiello")
    count = count + AssicuraGiocatore(ws, 15, "Circati")
    count = count + AssicuraGiocatore(ws, 15, "Berisha")
    count = count + AssicuraGiocatore(ws, 15, "Moreo")
    count = count + AssicuraGiocatore(ws, 15, "Durosinmi")

    ' === PARTIZAN (Col Calciatore = 39) ===
    count = count + AssicuraGiocatore(ws, 39, "Belghali")
    count = count + AssicuraGiocatore(ws, 39, "Strefezza")
    count = count + AssicuraGiocatore(ws, 39, "Przyborek")

    ' === KUNG FU PANDEV (Col Calciatore = 63) ===
    count = count + AssicuraGiocatore(ws, 63, "Malen")
    count = count + AssicuraGiocatore(ws, 63, "Vergara")
    count = count + AssicuraGiocatore(ws, 63, "Beukema")
    count = count + AssicuraGiocatore(ws, 63, "Kouam")  ' nome troncato per match parziale

    ' === FC CKC 26 (Col Calciatore = 87) ===
    count = count + AssicuraGiocatore(ws, 87, "Tiago Gabriel")
    count = count + AssicuraGiocatore(ws, 87, "Vaz")
    count = count + AssicuraGiocatore(ws, 87, "Muharemovic")
    count = count + AssicuraGiocatore(ws, 87, "Baldanzi")
    count = count + AssicuraGiocatore(ws, 87, "Allison")    ' NON trovato nel DB - potrebbe mancare
    count = count + AssicuraGiocatore(ws, 87, "Bijlow")
    count = count + AssicuraGiocatore(ws, 87, "Bernasconi")
    count = count + AssicuraGiocatore(ws, 87, "Kon")  ' match parziale per Kone/Kone'

    ' === MUTTLEY SUPERSTAR (Col Calciatore = 27) ===
    count = count + AssicuraGiocatore(ws, 27, "Ostigard")
    count = count + AssicuraGiocatore(ws, 27, "Luis")  ' match parziale per Luis Henrique/Luis Enrique
    count = count + AssicuraGiocatore(ws, 27, "Solomon")

    ' === MILLWALL (Col Calciatore = 75) ===
    count = count + AssicuraGiocatore(ws, 75, "Muric")
    count = count + AssicuraGiocatore(ws, 75, "Celik")
    count = count + AssicuraGiocatore(ws, 75, "Ratkov")
    count = count + AssicuraGiocatore(ws, 75, "Zaragoza")
    count = count + AssicuraGiocatore(ws, 75, "Perrone")
    count = count + AssicuraGiocatore(ws, 75, "Paleari")
    count = count + AssicuraGiocatore(ws, 75, "Boga")
    count = count + AssicuraGiocatore(ws, 75, "Holm")

    ' === PAPAIE TOP TEAM (Col Calciatore = 99) ===
    count = count + AssicuraGiocatore(ws, 99, "Hien")

    ' === LEGENDA AUREA (Col Calciatore = 51) ===
    count = count + AssicuraGiocatore(ws, 51, "Di Gregorio")
    count = count + AssicuraGiocatore(ws, 51, "Sommer")
    count = count + AssicuraGiocatore(ws, 51, "Martinez")
    count = count + AssicuraGiocatore(ws, 51, "Kalulu")
    count = count + AssicuraGiocatore(ws, 51, "Bartesaghi")
    count = count + AssicuraGiocatore(ws, 51, "Lovric")
    count = count + AssicuraGiocatore(ws, 51, "Taylor")
    count = count + AssicuraGiocatore(ws, 51, "Fagioli")
    count = count + AssicuraGiocatore(ws, 51, "Ekkelenkamp")
    count = count + AssicuraGiocatore(ws, 51, "Miretti")
    count = count + AssicuraGiocatore(ws, 51, "Bonazzoli")
    count = count + AssicuraGiocatore(ws, 51, "Raspadori")
    count = count + AssicuraGiocatore(ws, 51, "Vitinha")

    MsgBox "Assicurazioni FT aggiornate!" & vbCrLf & _
           "Giocatori trovati e aggiornati: " & count & vbCrLf & _
           "Data assicurazione: " & DATA_ASS, vbInformation, "Fanta Tosti 2026"
End Sub

' ============================================================
' Funzione helper: cerca un giocatore nella colonna specificata
' e imposta il flag assicurazione e la data
' ============================================================
Private Function AssicuraGiocatore(ws As Worksheet, colCalciatore As Long, nomeGiocatore As String) As Long
    Dim r As Long
    Dim nomeCell As String
    Dim nomeCerca As String

    nomeCerca = UCase(Replace(nomeGiocatore, "'", ""))

    ' Cerca nelle righe 6-52 (area giocatori attivi)
    For r = 6 To 52
        nomeCell = UCase(Trim(CStr(ws.Cells(r, colCalciatore).Value)))
        nomeCell = Replace(nomeCell, "'", "")
        nomeCell = Replace(nomeCell, Chr(232), "e")  ' e grave
        nomeCell = Replace(nomeCell, Chr(233), "e")  ' e acute
        nomeCell = Replace(nomeCell, Chr(242), "o")  ' o grave
        nomeCell = Replace(nomeCell, Chr(224), "a")  ' a grave
        nomeCell = Replace(nomeCell, Chr(249), "u")  ' u grave

        ' Match: contiene il termine di ricerca
        If Len(nomeCell) > 0 And InStr(1, nomeCell, nomeCerca, vbTextCompare) > 0 Then
            ' Imposta flag assicurazione
            ws.Cells(r, colCalciatore + 3).Value = "A"

            ' Imposta data assicurazione
            ws.Cells(r, colCalciatore + 7).Value = CDate(DATA_ASS)
            ws.Cells(r, colCalciatore + 7).NumberFormat = "dd/mm/yyyy"

            ' Log nel foglio Immediate (Debug)
            Debug.Print "OK: " & nomeGiocatore & " -> trovato come '" & ws.Cells(r, colCalciatore).Value & "' (riga " & r & ", col " & colCalciatore & ")"

            AssicuraGiocatore = 1
            Exit Function
        End If
    Next r

    ' Non trovato
    Debug.Print "ATTENZIONE: " & nomeGiocatore & " NON TROVATO nella colonna " & colCalciatore
    MsgBox "ATTENZIONE: Giocatore '" & nomeGiocatore & "' non trovato!" & vbCrLf & _
           "Colonna Calciatore: " & colCalciatore & vbCrLf & _
           "Potrebbe essere un acquisto post-06/02 non ancora nel DB.", _
           vbExclamation, "Giocatore non trovato"

    AssicuraGiocatore = 0
End Function

' ============================================================
' Macro di verifica: mostra tutti i giocatori assicurati
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
    output = "GIOCATORI ASSICURATI - FT:" & vbCrLf & vbCrLf

    Dim t As Long
    For t = LBound(teamCols) To UBound(teamCols)
        Dim teamName As String
        Dim col As Long
        teamName = teamCols(t)(0)
        col = teamCols(t)(1)

        output = output & teamName & ":" & vbCrLf
        Dim count As Long
        count = 0

        Dim r As Long
        For r = 6 To 52
            If UCase(Trim(CStr(ws.Cells(r, col + 3).Value))) = "A" Then
                Dim nome As String
                nome = Trim(CStr(ws.Cells(r, col).Value))
                If Len(nome) > 0 And nome <> "Calciatore" Then
                    output = output & "  " & nome & vbCrLf
                    count = count + 1
                End If
            End If
        Next r

        If count = 0 Then output = output & "  (nessuno)" & vbCrLf
        output = output & vbCrLf
    Next t

    Debug.Print output
    MsgBox output, vbInformation, "Verifica Assicurati FT"
End Sub
