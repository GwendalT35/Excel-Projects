Attribute VB_Name = "Module3"
Sub Import()

Dim Numero As String
Dim ID As String
Dim Rng As Range
Dim Nom As String
Dim adresse As String

Numero = InputBox("Which set would you like to import (Please state set ID) ?", "Import")
If Trim(Numero) <> "" Then
    With Sheets("Database").Range("C2:C19516")
        Set Rng = .Find(What:=Numero, _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False)
        If Not Rng Is Nothing Then ' If Numero exist
            Application.Goto Rng, True
            Worksheets("Dashboard").Range("I1") = Numero
            Worksheets("Dashboard").Range("J1").FormulaR1C1 = "=XLOOKUP(RC[-1],Table2[set_Numero_Boite],Table2[Set_ID])"
            Worksheets("Dashboard").Range("K1").FormulaR1C1 = "=XLOOKUP(RC[-2],Table2[set_Numero_Boite],Table2[Set_Nom])"
            ID = Worksheets("Dashboard").Range("J1")
            Nom = Worksheets("Dashboard").Range("K1")
            Sheets("Dashboard").Cells(2, "I").Value = "https://rebrickable.com/inventory/" & ID & "/parts/?format=rbpartscsv&inc_spares"
            adresse = Sheets("Dashboard").Cells(2, "I").Value
            Worksheets("Dashboard").Range("I1:K2").Clear
            ActiveWorkbook.Queries.Add Name:=Nom, Formula:= _
            "let" & Chr(13) & "" & Chr(10) & "    Source = Csv.Document(Web.Contents(""" & adresse & """),[Delimiter="","", Columns=4, Encoding=1252, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted H" & _
            "eaders"",{{""Part"", type text}, {""Color"", Int64.Type}, {""Quantity"", Int64.Type}, {""Is Spare"", type logical}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
        ActiveWorkbook.Worksheets.Add
        With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & Nom & ";Extended Properties="""""), Destination:=Range("$A$1")).QueryTable
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM  " & Nom & "")
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            .ListObject.DisplayName = Nom
            .Refresh BackgroundQuery:=False
        End With
        ActiveSheet.Name = Nom
        Else ' If Numero doesnt exist
            MsgBox "This set doesn't exist in Lego Database" '
        End If
    End With
End If

End Sub
Sub Add()

Dim Numero As String
Dim ID As String
Dim Rng As Range
Dim Nom As String
Dim adresse As String
Dim Color As String
Dim Part As String
Dim Quantity As Double
Dim y As Integer
Dim r As Variant
Dim c As Variant
Dim x As Integer
Dim i As Integer
Dim fc As Worksheet
Dim z As Integer

y = 2
z = 1

Numero = InputBox("Which set would you like to import (Please state set ID) ?", "Import")

If Trim(Numero) <> "" Then 'If set exists
    With Sheets("Database").Range("C2:C19516")
        Set Rng = .Find(What:=Numero, _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False)
        If Not Rng Is Nothing Then 'Set exists
            Application.Goto Rng, True
            Worksheets("Dashboard").Range("I1") = Numero
            Worksheets("Dashboard").Range("J1").FormulaR1C1 = "=XLOOKUP(RC[-1],Table2[set_Numero_Boite],Table2[Set_ID])"
            Worksheets("Dashboard").Range("K1").FormulaR1C1 = "=XLOOKUP(RC[-2],Table2[set_Numero_Boite],Table2[Set_Nom])"
            ID = Worksheets("Dashboard").Range("J1")
            Nom = Worksheets("Dashboard").Range("K1")
            Sheets("Dashboard").Cells(2, "I").Value = "https://rebrickable.com/inventory/" & ID & "/parts/?format=rbpartscsv&inc_spares"
            adresse = Sheets("Dashboard").Cells(2, "I").Value
            Worksheets("Dashboard").Range("I1:K2").Clear
            If Trim(Nom) <> "" Then ' If set in DB
                With Sheets("My_Sets").Range("A:A")
                    Set Rng = .Find(What:=Nom, _
                                    After:=.Cells(.Cells.Count), _
                                    LookIn:=xlValues, _
                                    LookAt:=xlWhole, _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlNext, _
                                    MatchCase:=False)
                    If Not Rng Is Nothing Then 'if yes
                        Application.Goto Rng, True
                        Sheets(Nom).Visible = True
                        Sheets(Nom).Select
                        x = InputBox("How many copies of this set would you like to add ?", "Import")
                        For i = 1 To x
                            For y = 2 To 500
                                If IsEmpty(Sheets(Nom).Cells(y, 1)) = False Then
                                    Sheets(Nom).Select
                                    Part = Cells(y, ColumnIndex:="A").Value
                                    Color = Cells(y, ColumnIndex:="B").Value
                                    Quantity = Cells(y, ColumnIndex:="C").Value
                                    r = Sheets("My_Parts").Columns("A").Find(Part).row
                                    c = Sheets("My_Parts").Rows("1").Find(Color).column
                                    Sheets("My_Parts").Cells(r, c) = Quantity + Sheets("My_Parts").Cells(r, c)
                                End If
                            Next
                        Next
                    Else 'if not
                        ActiveWorkbook.Queries.Add Name:=Nom, Formula:= _
                        "let" & Chr(13) & "" & Chr(10) & "    Source = Csv.Document(Web.Contents(""" & adresse & """),[Delimiter="","", Columns=4, Encoding=1252, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted H" & _
                        "eaders"",{{""Part"", type text}, {""Color"", Int64.Type}, {""Quantity"", Int64.Type}, {""Is Spare"", type logical}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
                        ActiveWorkbook.Worksheets.Add
                        With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
                            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & Nom & ";Extended Properties="""""), Destination:=Range("$A$1")).QueryTable
                            .CommandType = xlCmdSql
                            .CommandText = Array("SELECT * FROM  " & Nom & "")
                            .RowNumbers = False
                            .FillAdjacentFormulas = False
                            .PreserveFormatting = True
                            .RefreshOnFileOpen = False
                            .BackgroundQuery = True
                            .RefreshStyle = xlInsertDeleteCells
                            .SavePassword = False
                            .SaveData = True
                            .AdjustColumnWidth = True
                            .RefreshPeriod = 0
                            .PreserveColumnInfo = True
                            .ListObject.DisplayName = Nom
                            .Refresh BackgroundQuery:=False
                        End With
                        ActiveSheet.Name = Nom
                        For Each fc In Worksheets
                            If fc.Name <> "My_Sets" And fc.Name <> "My_Parts" And fc.Name <> "Database" And fc.Name <> "Dashboard" And fc.Name <> "sets" And fc.Name <> "inventories" And fc.Name <> "themes" And fc.Name <> "colors" And fc.Name <> "elements" And fc.Name <> "inventory_minifigs" And fc.Name <> "inventory_parts" And fc.Name <> "inventory_sets" And fc.Name <> "minifigs" And fc.Name <> "part_categories" And fc.Name <> "part_relationships" And fc.Name <> "parts" And fc.Name <> "Missing_Parts" Then
                                Sheets("My_Sets").Cells(z, 1) = fc.Name
                                z = z + 1
                            End If
                        Next fc
                        With Sheets("My_Sets").Range("A:A")
                            Set Rng = .Find(What:=Nom, _
                                            After:=.Cells(.Cells.Count), _
                                            LookIn:=xlValues, _
                                            LookAt:=xlWhole, _
                                            SearchOrder:=xlByRows, _
                                            SearchDirection:=xlNext, _
                                            MatchCase:=False)
                            If Not Rng Is Nothing Then 'if yes
                                Application.Goto Rng, True
                                Sheets(Nom).Select
                                x = InputBox("Combien d'exemplaires de ce set souhaitez vous ajouter ?", "Import")
                                For i = 1 To x
                                    For y = 2 To 500
                                        If IsEmpty(Sheets(Nom).Cells(y, 1)) = False Then
                                            Sheets(Nom).Select
                                            Part = Cells(y, ColumnIndex:="A").Value
                                            Color = Cells(y, ColumnIndex:="B").Value
                                            Quantity = Cells(y, ColumnIndex:="C").Value
                                            r = Sheets("My_Parts").Columns("A").Find(Part).row
                                            c = Sheets("My_Parts").Rows("1").Find(Color).column
                                            Sheets("My_Parts").Cells(r, c) = Quantity + Sheets("My_Parts").Cells(r, c)
                                        End If
                                    Next
                                Next
                            End If
                        End With
                    End If
                End With
            End If
        Sheets("Dashboard").Select
        Sheets(Nom).Visible = False
        Else ' If Numero doesnt exist
            MsgBox "This set doesn't exist in Lego Database" '
        End If
    End With
    Sheets("Dashboard").Range("I1:K2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Sheets("My_Sets").Range("A:A").Clear
    For Each fc In Worksheets
        If fc.Name <> "My_Sets" And fc.Name <> "My_Parts" And fc.Name <> "Database" And fc.Name <> "Dashboard" And fc.Name <> "sets" And fc.Name <> "inventories" And fc.Name <> "themes" And fc.Name <> "colors" And fc.Name <> "elements" And fc.Name <> "inventory_minifigs" And fc.Name <> "inventory_parts" And fc.Name <> "inventory_sets" And fc.Name <> "minifigs" And fc.Name <> "part_categories" And fc.Name <> "part_relationships" And fc.Name <> "parts" And fc.Name <> "Missing_Parts" Then
         Sheets("My_Sets").Cells(z, 1) = fc.Name
         z = z + 1
        End If
    Next fc
    Sheets("My_Sets").Range("A:A").Select
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
    End With
    Sheets("Dashboard").Select
    Sheets("Dashboard").Range("A1").Select
End If

End Sub
Sub Delete()

Dim Nom As String
Dim Rng As Range
Dim Rep As Integer
Dim Color As String
Dim Part As String
Dim Quantity As Double
Dim i As Integer
Dim x As Integer
Dim y As Integer
Dim r As Variant
Dim c As Variant
Dim No As String
Dim Yes As String
Dim fc As Worksheet
Dim z As Integer
z = 1

Nom = InputBox("Which set would you like to delete (Please state set Name) ?", "Delete")

If Trim(Nom) <> "" Then
    With Sheets("My_Sets").Range("A:A")
        Set Rng = .Find(What:=Nom, _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False)
        If Not Rng Is Nothing Then 'Y
            x = InputBox("How many copies of this set would you like do delete ?")
            For i = 1 To x
                For y = 2 To 500
                    If IsEmpty(Sheets(Nom).Cells(y, 1)) = False Then
                        Sheets(Nom).Visible = True
                        Sheets(Nom).Select
                        Part = Cells(y, ColumnIndex:="A").Value
                        Color = Cells(y, ColumnIndex:="B").Value
                        Quantity = Cells(y, ColumnIndex:="C").Value
                        r = Sheets("My_Parts").Columns("A").Find(Part).row
                        c = Sheets("My_Parts").Rows("1").Find(Color).column
                        Sheets("My_Parts").Cells(r, c) = Sheets("My_Parts").Cells(r, c) - Quantity
                    End If
                    If Sheets("My_Parts").Cells(r, c) < 0 Then
                        Sheets("My_Parts").Cells(r, c) = "0"
                    End If
                Next
            Next
            Yes = MsgBox("Do you also want to delete the sheet corresponding to this set ?", vbQuestion + vbYesNo, "Delete")
            If Yes = vbYes Then
                Sheets(Nom).Delete
                ActiveWorkbook.Queries(Nom).Delete
                MsgBox ("Sheet deleted")
            End If
            MsgBox (x & " copies of the set : " & Nom & " have been deleted")
        Else
            MsgBox ("This set doesn't exist in your Lego Database")
        End If
    End With
End If
Sheets("My_Sets").Range("A:A").Clear

For Each fc In Worksheets
    If fc.Name <> "My_Sets" And fc.Name <> "My_Parts" And fc.Name <> "Database" And fc.Name <> "Dashboard" And fc.Name <> "sets" And fc.Name <> "inventories" And fc.Name <> "themes" And fc.Name <> "colors" And fc.Name <> "elements" And fc.Name <> "inventory_minifigs" And fc.Name <> "inventory_parts" And fc.Name <> "inventory_sets" And fc.Name <> "minifigs" And fc.Name <> "part_categories" And fc.Name <> "part_relationships" And fc.Name <> "parts" And fc.Name <> "Missing_Parts" Then
     Sheets("My_Sets").Cells(z, 1) = fc.Name
     z = z + 1
    End If
Next fc
Sheets("My_Sets").Range("A:A").Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
End With
Sheets("Dashboard").Select
Sheets("Dashboard").Range("A1").Select

End Sub
Sub List()

Dim fc As Worksheet
Dim z As Integer
z = 1

Sheets("My_Sets").Range("A:A").Clear

For Each fc In Worksheets
    If fc.Name <> "My_Sets" And fc.Name <> "My_Parts" And fc.Name <> "Database" And fc.Name <> "Dashboard" And fc.Name <> "sets" And fc.Name <> "inventories" And fc.Name <> "themes" And fc.Name <> "colors" And fc.Name <> "elements" And fc.Name <> "inventory_minifigs" And fc.Name <> "inventory_parts" And fc.Name <> "inventory_sets" And fc.Name <> "minifigs" And fc.Name <> "part_categories" And fc.Name <> "part_relationships" And fc.Name <> "parts" And fc.Name <> "Missing_Parts" Then
     Sheets("My_Sets").Cells(z, 1) = fc.Name
     z = z + 1
    End If
Next fc
Sheets("My_Sets").Range("A:A").Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
End With

End Sub
Sub Clear()

Dim Yes As String
Dim No As String

Yes = MsgBox("Would you like to reset you parts collection ?", vbQuestion + vbYesNo, "Reset")

If Yes = vbYes Then
    Sheets("My_Parts").Range("C2", "HK47701").Value = Null
    MsgBox ("Parts Collection reseted")
Else
    MsgBox ("Reset Aborted")
End If

End Sub
Sub Instructions()

Dim Numero As Variant

Set objShell = CreateObject("Wscript.Shell")
Numero = InputBox("Which set's instructions would you like ? (State set ID)")

'intMessage = MsgBox("Do you want to be redirected towards Brickset to get access to this set's instructions ?", vbYesNo)

'If intMessage = vbYes Then
If StrPtr(Numero) = 0 Then
    Exit Sub
ElseIf Numero = vbNullInteger Then
    Exit Sub
Else
    objShell.Run ("https://brickset.com/sets/" & Numero & "")
End If
'End If

End Sub
Sub Create()

Dim Numero As String
Dim ID As String
Dim Rng As Range
Dim Nom As String
Dim adresse As String
Dim Color As String
Dim Part As String
Dim Quantity As Double
Dim y As Integer
Dim r As Variant
Dim c As Variant
Dim x As Integer
Dim i As Integer
Dim myRange As Range
Dim myCell As Range
Dim pc As Integer
Dim mp As Integer
Dim nb As Integer
Dim w As Integer
Dim fc As Worksheet
Dim z As Integer



y = 2
x = 1
z = 1
mp = 0
nb = 0

Set objShell = CreateObject("Wscript.Shell")
Numero = InputBox("Which set would you like to create ?", "Create")

If Trim(Numero) <> "" Then 'If set exists
    With Sheets("Database").Range("C2:C19516")
        Set Rng = .Find(What:=Numero, _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False)
        If Not Rng Is Nothing Then 'Set exists
            Application.Goto Rng, True
            Worksheets("Dashboard").Range("I1") = Numero
            Worksheets("Dashboard").Range("J1").FormulaR1C1 = "=XLOOKUP(RC[-1],Table2[set_Numero_Boite],Table2[Set_ID])"
            Worksheets("Dashboard").Range("K1").FormulaR1C1 = "=XLOOKUP(RC[-2],Table2[set_Numero_Boite],Table2[Set_Nom])"
            ID = Worksheets("Dashboard").Range("J1")
            Nom = Worksheets("Dashboard").Range("K1")
            Sheets("Dashboard").Cells(2, "I").Value = "https://rebrickable.com/inventory/" & ID & "/parts/?format=rbpartscsv&inc_spares"
            adresse = Sheets("Dashboard").Cells(2, "I").Value
            Worksheets("Dashboard").Range("I1:K2").Clear
            If Trim(Nom) <> "" Then ' If set in DB
                With Sheets("My_Sets").Range("A:A")
                    Set Rng = .Find(What:=Nom, _
                                    After:=.Cells(.Cells.Count), _
                                    LookIn:=xlValues, _
                                    LookAt:=xlWhole, _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlNext, _
                                    MatchCase:=False)
                    If Not Rng Is Nothing Then 'if yes
                        Application.Goto Rng, True
                        Sheets(Nom).Visible = True
                        Sheets(Nom).Select
                        Set myRange = Worksheets(Nom).Range("E2:E500")
                        For Each cell In myRange
                            If IsEmpty(Worksheets(Nom).Range("C" & y)) = False Then
                                w = w + 1
                            End If
                            y = y + 1
                        Next cell
                        Range("E1").FormulaR1C1 = "Available"
                        Range("F1").FormulaR1C1 = "Times"
                        w = w + 1
                        For y = 2 To w
                            Sheets(Nom).Select
                            Part = Cells(y, ColumnIndex:="A").Value
                            Color = Cells(y, ColumnIndex:="B").Value
                            Quantity = Cells(y, ColumnIndex:="C").Value
                            r = Sheets("My_Parts").Columns("A").Find(Part).row
                            c = Sheets("My_Parts").Rows("1").Find(Color).column
                            Worksheets(Nom).Range("E" & y).FormulaR1C1 = Sheets("My_Parts").Cells(r, c)
                            Worksheets(Nom).Range("F" & y).FormulaR1C1 = "=[@Available]/[@Quantity]"
                            Set myRange = Worksheets(Nom).Range("E2:E" & w)
                            If IsEmpty(Worksheets(Nom).Range("E" & y)) = True Then
                                Worksheets(Nom).Range("E" & y).FormulaR1C1 = "0"
                            End If
                        Next
                        y = 2
                        For Each cell In myRange
                            pc = pc + 1
                            If Worksheets(Nom).Range("E" & y) = 0 Then
                                mp = mp + Worksheets(Nom).Range("C" & y).Value
                                nb = nb + 1
                            End If
                            y = y + 1
                        Next cell
                    Else ' if
                        ActiveWorkbook.Queries.Add Name:=Nom, Formula:= _
                        "let" & Chr(13) & "" & Chr(10) & "    Source = Csv.Document(Web.Contents(""" & adresse & """),[Delimiter="","", Columns=4, Encoding=1252, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted H" & _
                        "eaders"",{{""Part"", type text}, {""Color"", Int64.Type}, {""Quantity"", Int64.Type}, {""Is Spare"", type logical}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
                        ActiveWorkbook.Worksheets.Add
                        With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
                            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & Nom & ";Extended Properties="""""), Destination:=Range("$A$1")).QueryTable
                            .CommandType = xlCmdSql
                            .CommandText = Array("SELECT * FROM  " & Nom & "")
                            .RowNumbers = False
                            .FillAdjacentFormulas = False
                            .PreserveFormatting = True
                            .RefreshOnFileOpen = False
                            .BackgroundQuery = True
                            .RefreshStyle = xlInsertDeleteCells
                            .SavePassword = False
                            .SaveData = True
                            .AdjustColumnWidth = True
                            .RefreshPeriod = 0
                            .PreserveColumnInfo = True
                            .ListObject.DisplayName = Nom
                            .Refresh BackgroundQuery:=False
                        End With
                        ActiveSheet.Name = Nom
                        For Each fc In Worksheets
                            If fc.Name <> "My_Sets" And fc.Name <> "My_Parts" And fc.Name <> "Database" And fc.Name <> "Dashboard" And fc.Name <> "sets" And fc.Name <> "inventories" And fc.Name <> "themes" And fc.Name <> "colors" And fc.Name <> "elements" And fc.Name <> "inventory_minifigs" And fc.Name <> "inventory_parts" And fc.Name <> "inventory_sets" And fc.Name <> "minifigs" And fc.Name <> "part_categories" And fc.Name <> "part_relationships" And fc.Name <> "parts" And fc.Name <> "Missing_Parts" Then
                                Sheets("My_Sets").Cells(z, 1) = fc.Name
                                z = z + 1
                            End If
                        Next fc
                        With Sheets("My_Sets").Range("A:A")
                            Set Rng = .Find(What:=Nom, _
                                            After:=.Cells(.Cells.Count), _
                                            LookIn:=xlValues, _
                                            LookAt:=xlWhole, _
                                            SearchOrder:=xlByRows, _
                                            SearchDirection:=xlNext, _
                                            MatchCase:=False)
                            If Not Rng Is Nothing Then 'if yes
                                Application.Goto Rng, True
                                Sheets(Nom).Visible = True
                                Sheets(Nom).Select
                                Set myRange = Worksheets(Nom).Range("E2:E500")
                                For Each cell In myRange
                                    If IsEmpty(Worksheets(Nom).Range("C" & y)) = False Then
                                        w = w + 1
                                    End If
                                    y = y + 1
                                Next cell
                                Range("E1").FormulaR1C1 = "Available"
                                Range("F1").FormulaR1C1 = "Times"
                                For y = 2 To w
                                    Sheets(Nom).Select
                                    Part = Cells(y, ColumnIndex:="A").Value
                                    Color = Cells(y, ColumnIndex:="B").Value
                                    Quantity = Cells(y, ColumnIndex:="C").Value
                                    r = Sheets("My_Parts").Columns("A").Find(Part).row
                                    c = Sheets("My_Parts").Rows("1").Find(Color).column
                                    Worksheets(Nom).Range("E" & y).FormulaR1C1 = Sheets("My_Parts").Cells(r, c)
                                    Worksheets(Nom).Range("F" & y).FormulaR1C1 = "=[@Available]/[@Quantity]"
                                    Set myRange = Worksheets(Nom).Range("E2:E" & w)
                                    If IsEmpty(Worksheets(Nom).Range("E" & y)) = True Then
                                        Worksheets(Nom).Range("E" & y).FormulaR1C1 = "0"
                                    End If
                                Next
                                y = 2
                                For Each cell In myRange
                                    pc = pc + 1
                                    If Worksheets(Nom).Range("E" & y) = 0 Then
                                        mp = mp + Worksheets(Nom).Range("C" & y).Value
                                        nb = nb + 1
                                    End If
                                    y = y + 1
                                Next cell
                            End If
                        End With
                    End If
                End With
            End If
        If nb > 0 Then
            MsgBox ("In order to create the set " & Nom & ", you miss " & nb & " different parts out of " & pc & " parts, for a total of " & mp & " parts.")
        Else
            MsgBox ("You have all the parts necessary to create this set !")
        End If
        intMessage = MsgBox("Do you want to be redirected towards Brickset to get access to this set's instructions ?", vbYesNo)
        If intMessage = vbYes Then
            objShell.Run ("https://brickset.com/sets/" & Numero & "")
        End If
        Else ' If Numero doesnt exist
            MsgBox "This set doesn't exist in Lego Database" '
        End If
    End With
End If
End Sub
Sub Reset()

Dim Yes As String
Dim No As String
Dim fc As Worksheet
Dim Nom As String

Yes = MsgBox("Do you want to reset this whole file", vbQuestion + vbYesNo + vbCritical, "RESET")

Sheets("My_Sets").Range("A:A").Clear

For Each fc In Worksheets
     If fc.Name <> "My_Sets" And fc.Name <> "My_Parts" And fc.Name <> "Database" And fc.Name <> "Dashboard" And fc.Name <> "sets" And fc.Name <> "inventories" And fc.Name <> "themes" And fc.Name <> "colors" And fc.Name <> "elements" And fc.Name <> "inventory_minifigs" And fc.Name <> "inventory_parts" And fc.Name <> "inventory_sets" And fc.Name <> "minifigs" And fc.Name <> "part_categories" And fc.Name <> "part_relationships" And fc.Name <> "parts" And fc.Name <> "Missing_Parts" Then
     Nom = fc.Name
     Sheets(Nom).Delete
     ActiveWorkbook.Queries(Nom).Delete
     End If
Next fc
Sheets("My_Parts").Range("C2", "HK47701").Value = Null

End Sub
Sub Feedback()

Set objShell = CreateObject("Wscript.Shell")

intMessage = MsgBox("Do you want to be taken to this project's GitHub page ?", vbYesNo)

If intMessage = vbYes Then
    objShell.Run ("https://github.com/Acerioz/Excel-Projects/issues/new")
End If

End Sub
