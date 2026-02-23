' ============================================================
' MACRO VBA - FANTAMANTRA MANAGERIALE
' Aggiornamento Assicurazioni - Mercato Invernale 2026
' ============================================================
' ISTRUZIONI:
' 1. Aprire il file "FantaMantra Manageriale - DB completo (06.02.2026).xlsx"
' 2. Premere Alt+F11 per aprire l'editor VBA
' 3. Menu Inserisci > Modulo
' 4. Incollare TUTTO questo codice nel modulo
' 5. Premere F5 oppure andare su Esegui > Esegui Sub
' 6. Selezionare "AggiornaAssicurazioniFM" e fare clic su Esegui
' ============================================================

' Colonne Calciatore per FM (1-based, VBA):
' Papaie Top Team  = 4
' Legenda Aurea    = 17
' Lino Banfield FC = 30
' Kung Fu Pandev   = 43
' FICA             = 56
' Hellas Madonna   = 69
' MINNESOTA AL MAX = 82
' FC CKC 26        = 95
' H-Q-A Barcelona  = 108
' Mastri Birrai    = 121
' NOTA: FM ha colonna extra "Reparto", quindi le colonne sono +1 rispetto a FT

Private Const DATA_ASS As String = "14/02/2026"

Sub AggiornaAssicurazioniFM()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SQUADRE")

    Dim count As Long
    count = 0

    ' === KUNG FU PANDEV (Col Calciatore = 43) ===
    count = count + AssicuraGiocatore(ws, 43, "Kon")      ' Kone' I. - match parziale
    count = count + AssicuraGiocatore(ws, 43, "Raspadori")
    ' POSCH: RESPINTO (svincolato, NON assicurabile)
    count = count + AssicuraGiocatore(ws, 43, "Ferguson")
    count = count + AssicuraGiocatore(ws, 43, "Kouam")     ' Kouame' - match parziale

    ' === FC CKC 26 (Col Calciatore = 95) ===
    count = count + AssicuraGiocatore(ws, 95, "Durosinmi")
    count = count + AssicuraGiocatore(ws, 95, "Vergara")
    count = count + AssicuraGiocatore(ws, 95, "Zaniolo")

    ' === H-Q-A BARCELONA (Col Calciatore = 108) ===
    count = count + AssicuraGiocatore(ws, 108, "Holm")
    count = count + AssicuraGiocatore(ws, 108, "Ndicka")
    count = count + AssicuraGiocatore(ws, 108, "Gallo")
    count = count + AssicuraGiocatore(ws, 108, "Vasquez")
    count = count + AssicuraGiocatore(ws, 108, "Gudmundsson")
    count = count + AssicuraGiocatore(ws, 108, "Frendrup")
    count = count + AssicuraGiocatore(ws, 108, "Britschgi")
    count = count + AssicuraGiocatore(ws, 108, "Sulemana")
    count = count + AssicuraGiocatore(ws, 108, "Taylor")
    count = count + AssicuraGiocatore(ws, 108, "Malen")
    count = count + AssicuraGiocatore(ws, 108, "Sommer")

    ' === HELLAS MADONNA (Col Calciatore = 69) ===
    ' NOTA: David potrebbe NON essere nella colonna Hellas (era in Minnesota)
    ' Se non trovato qui, cercarlo in Minnesota (col 82) o aggiungere manualmente
    count = count + AssicuraGiocatore(ws, 69, "David")
    count = count + AssicuraGiocatore(ws, 69, "Cheddira")
    count = count + AssicuraGiocatore(ws, 69, "Zaragoza")
    count = count + AssicuraGiocatore(ws, 69, "Ekkelenkamp")
    count = count + AssicuraGiocatore(ws, 69, "Brescianini")
    count = count + AssicuraGiocatore(ws, 69, "Belghali")
    count = count + AssicuraGiocatore(ws, 69, "Scamacca")

    ' === FICA (Col Calciatore = 56) ===
    count = count + AssicuraGiocatore(ws, 56, "Luis Henrique")
    count = count + AssicuraGiocatore(ws, 56, "Fullkrug")

    ' === LINO BANFIELD FC (Col Calciatore = 30) ===
    count = count + AssicuraGiocatore(ws, 30, "Celik")
    count = count + AssicuraGiocatore(ws, 30, "Obert")
    count = count + AssicuraGiocatore(ws, 30, "Marcandalli")
    count = count + AssicuraGiocatore(ws, 30, "Bernasconi")
    count = count + AssicuraGiocatore(ws, 30, "Bowie")
    count = count + AssicuraGiocatore(ws, 30, "Caprile")
    count = count + AssicuraGiocatore(ws, 30, "Cambiaghi")
    count = count + AssicuraGiocatore(ws, 30, "Vaz")
    count = count + AssicuraGiocatore(ws, 30, "Baldanzi")
    ' I seguenti 3 sono da Minnesota (scambio 13/02) - potrebbero non essere
    ' nella colonna Lino nel DB 06/02. Se non trovati, vanno aggiunti manualmente
    ' o cercati nella colonna Minnesota (82)
    count = count + AssicuraGiocatore(ws, 30, "Koopmeiners")
    count = count + AssicuraGiocatore(ws, 30, "Tavares")
    count = count + AssicuraGiocatore(ws, 30, "Mazzitelli")

    ' === MINNESOTA AL MAX (Col Calciatore = 82) ===
    count = count + AssicuraGiocatore(ws, 82, "Montip")     ' Montipo' - match parziale
    count = count + AssicuraGiocatore(ws, 82, "Marianucci")
    count = count + AssicuraGiocatore(ws, 82, "Cataldi")
    ' I seguenti 3 sono da Lino Banfield (scambio 13/02)
    count = count + AssicuraGiocatore(ws, 82, "Fagioli")
    count = count + AssicuraGiocatore(ws, 82, "Miller")
    count = count + AssicuraGiocatore(ws, 82, "Bakola")
    count = count + AssicuraGiocatore(ws, 82, "Adzic")
    count = count + AssicuraGiocatore(ws, 82, "Ratkov")
    count = count + AssicuraGiocatore(ws, 82, "Bellanova")

    ' === PAPAIE TOP TEAM (Col Calciatore = 4) ===
    count = count + AssicuraGiocatore(ws, 4, "Kolasinac")
    count = count + AssicuraGiocatore(ws, 4, "Hien")   ' Da Minnesota (acquisto 11/02)
    count = count + AssicuraGiocatore(ws, 4, "Pasalic")
    count = count + AssicuraGiocatore(ws, 4, "Nicolussi")  ' match parziale per Nicolussi Caviglia
    count = count + AssicuraGiocatore(ws, 4, "Solomon")
    count = count + AssicuraGiocatore(ws, 4, "Vlahovic")

    ' === LEGENDA AUREA (Col Calciatore = 17) ===
    count = count + AssicuraGiocatore(ws, 17, "Nelsson")
    count = count + AssicuraGiocatore(ws, 17, "Dossena")
    count = count + AssicuraGiocatore(ws, 17, "Bartesaghi")
    count = count + AssicuraGiocatore(ws, 17, "Gandelman")
    count = count + AssicuraGiocatore(ws, 17, "Barbieri")
    count = count + AssicuraGiocatore(ws, 17, "Leao")
    count = count + AssicuraGiocatore(ws, 17, "Zappa")

    MsgBox "Assicurazioni FM aggiornate!" & vbCrLf & _
           "Giocatori trovati e aggiornati: " & count & vbCrLf & _
           "Data assicurazione: " & DATA_ASS & vbCrLf & vbCrLf & _
           "NOTA: Giocatori scambiati dopo il 06/02 (Koopmeiners, Tavares, " & vbCrLf & _
           "Mazzitelli, Fagioli, Miller, Bellanova, Hien, David) " & vbCrLf & _
           "potrebbero non essere nelle colonne aggiornate.", _
           vbInformation, "FantaMantra Manageriale"
End Sub

' ============================================================
' Funzione helper per FM
' ============================================================
Private Function AssicuraGiocatore(ws As Worksheet, colCalciatore As Long, nomeGiocatore As String) As Long
    Dim r As Long
    Dim nomeCell As String
    Dim nomeCerca As String

    nomeCerca = UCase(Replace(nomeGiocatore, "'", ""))

    For r = 6 To 52
        nomeCell = UCase(Trim(CStr(ws.Cells(r, colCalciatore).Value)))
        nomeCell = Replace(nomeCell, "'", "")
        nomeCell = Replace(nomeCell, Chr(232), "e")
        nomeCell = Replace(nomeCell, Chr(233), "e")
        nomeCell = Replace(nomeCell, Chr(242), "o")
        nomeCell = Replace(nomeCell, Chr(224), "a")
        nomeCell = Replace(nomeCell, Chr(249), "u")

        If Len(nomeCell) > 0 And InStr(1, nomeCell, nomeCerca, vbTextCompare) > 0 Then
            ws.Cells(r, colCalciatore + 3).Value = "A"
            ws.Cells(r, colCalciatore + 7).Value = CDate(DATA_ASS)
            ws.Cells(r, colCalciatore + 7).NumberFormat = "dd/mm/yyyy"

            Debug.Print "OK: " & nomeGiocatore & " -> '" & ws.Cells(r, colCalciatore).Value & "' (riga " & r & ", col " & colCalciatore & ")"
            AssicuraGiocatore = 1
            Exit Function
        End If
    Next r

    Debug.Print "ATTENZIONE: " & nomeGiocatore & " NON TROVATO nella colonna " & colCalciatore
    MsgBox "ATTENZIONE: Giocatore '" & nomeGiocatore & "' non trovato!" & vbCrLf & _
           "Colonna Calciatore: " & colCalciatore & vbCrLf & _
           "Potrebbe essere un acquisto/scambio post-06/02.", _
           vbExclamation, "Giocatore non trovato"
    AssicuraGiocatore = 0
End Function

' ============================================================
' Macro di verifica
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
    output = "GIOCATORI ASSICURATI - FM:" & vbCrLf & vbCrLf

    Dim t As Long
    For t = LBound(teamCols) To UBound(teamCols)
        Dim teamName As String
        Dim col As Long
        teamName = teamCols(t)(0)
        col = teamCols(t)(1)

        output = output & teamName & ":" & vbCrLf
        Dim cnt As Long
        cnt = 0

        Dim r As Long
        For r = 6 To 52
            If UCase(Trim(CStr(ws.Cells(r, col + 3).Value))) = "A" Then
                Dim nome As String
                nome = Trim(CStr(ws.Cells(r, col).Value))
                If Len(nome) > 0 And nome <> "Calciatore" Then
                    output = output & "  " & nome & vbCrLf
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
