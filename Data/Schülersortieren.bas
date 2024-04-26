Attribute VB_Name = "Modul1"
Sub ErstelleAnwesenheitslisten()
    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    Dim veranstaltungsDict As Object
    Set veranstaltungsDict = CreateObject("Scripting.Dictionary")
    Dim veranstaltungsNr As String
    Dim veranstaltungsBlatt As Worksheet
    Dim currentRow As Long

    ' Schleife über jede Zeile im Datenblatt
    For currentRow = 2 To lastRow
        veranstaltungsNr = wsSource.Cells(currentRow, 4).Value ' Veranstaltungsnummer aus Spalte D
        If Not veranstaltungsDict.Exists(veranstaltungsNr) Then
            ' Erstelle ein neues Blatt, wenn es noch nicht existiert
            Set veranstaltungsBlatt = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            veranstaltungsBlatt.Name = "Veranstaltung_" & veranstaltungsNr
            veranstaltungsBlatt.Cells(1, 1).Resize(1, 5).Value = Array("Klasse", "Name", "Vorname", "Veranstaltungsnummer", "Slot")
            Set veranstaltungsDict(veranstaltungsNr) = veranstaltungsBlatt
        Else
            Set veranstaltungsBlatt = veranstaltungsDict(veranstaltungsNr)
        End If

        ' Bestimme die erste leere Zeile im Zielblatt
        Dim emptyRow As Long
        emptyRow = veranstaltungsBlatt.Cells(veranstaltungsBlatt.Rows.Count, "A").End(xlUp).Row + 1
        ' Kopiere die Zeile aus dem Quellblatt ins Zielblatt
        veranstaltungsBlatt.Cells(emptyRow, 1).Resize(1, 5).Value = wsSource.Cells(currentRow, 1).Resize(1, 5).Value
    Next currentRow

    ' Sortiere die Schüler in jedem Blatt nach dem Slot
    Dim key As Variant
    For Each key In veranstaltungsDict.Keys
        Set veranstaltungsBlatt = veranstaltungsDict(key)
        With veranstaltungsBlatt.Sort
            .SortFields.Clear
            .SortFields.Add key:=veranstaltungsBlatt.Cells(2, 5), Order:=xlAscending
            .SetRange veranstaltungsBlatt.Range("A1:E" & veranstaltungsBlatt.Cells(veranstaltungsBlatt.Rows.Count, "A").End(xlUp).Row)
            .Header = xlYes
            .Apply
        End With
    Next key

    MsgBox "Anwesenheitslisten für jede Veranstaltung wurden erstellt und sortiert."
End Sub

