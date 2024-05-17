
'below pulled from http://www.cpearson.com/excel/findall.aspx
Function FindAll(SearchRange As Range, _
                FindWhat As Variant, _
               Optional LookIn As XlFindLookIn = xlValues, _
                Optional LookAt As XlLookAt = xlWhole, _
                Optional SearchOrder As XlSearchOrder = xlByRows, _
                Optional MatchCase As Boolean = False, _
                Optional BeginsWith As String = vbNullString, _
                Optional EndsWith As String = vbNullString, _
                Optional BeginEndCompare As VbCompareMethod = vbTextCompare) As Range
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FindAll
' This searches the range specified by SearchRange and returns a Range object
' that contains all the cells in which FindWhat was found. The search parameters to
' this function have the same meaning and effect as they do with the
' Range.Find method. If the value was not found, the function return Nothing. If
' BeginsWith is not an empty string, only those cells that begin with BeginWith
' are included in the result. If EndsWith is not an empty string, only those cells
' that end with EndsWith are included in the result. Note that if a cell contains
' a single word that matches either BeginsWith or EndsWith, it is included in the
' result.  If BeginsWith or EndsWith is not an empty string, the LookAt parameter
' is automatically changed to xlPart. The tests for BeginsWith and EndsWith may be
' case-sensitive by setting BeginEndCompare to vbBinaryCompare. For case-insensitive
' comparisons, set BeginEndCompare to vbTextCompare. If this parameter is omitted,
' it defaults to vbTextCompare. The comparisons for BeginsWith and EndsWith are
' in an OR relationship. That is, if both BeginsWith and EndsWith are provided,
' a match if found if the text begins with BeginsWith OR the text ends with EndsWith.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim FoundCell As Range
Dim FirstFound As Range
Dim LastCell As Range
Dim ResultRange As Range
Dim XLookAt As XlLookAt
Dim Include As Boolean
Dim CompMode As VbCompareMethod
Dim Area As Range
Dim MaxRow As Long
Dim MaxCol As Long
Dim BeginB As Boolean
Dim EndB As Boolean


CompMode = BeginEndCompare
If BeginsWith <> vbNullString Or EndsWith <> vbNullString Then
    XLookAt = xlPart
Else
    XLookAt = LookAt
End If

' this loop in Areas is to find the last cell
' of all the areas. That is, the cell whose row
' and column are greater than or equal to any cell
' in any Area.

For Each Area In SearchRange.Areas
    With Area
        If .Cells(.Cells.Count).Row > MaxRow Then
            MaxRow = .Cells(.Cells.Count).Row
        End If
        If .Cells(.Cells.Count).Column > MaxCol Then
            MaxCol = .Cells(.Cells.Count).Column
        End If
    End With
Next Area
Set LastCell = SearchRange.Worksheet.Cells(MaxRow, MaxCol)

On Error GoTo 0
Set FoundCell = SearchRange.Find(What:=FindWhat, _
        After:=LastCell, _
        LookIn:=LookIn, _
        LookAt:=XLookAt, _
        SearchOrder:=SearchOrder, _
        MatchCase:=MatchCase)

If Not FoundCell Is Nothing Then
    Set FirstFound = FoundCell
    Do Until False ' Loop forever. We'll "Exit Do" when necessary.
        Include = False
        If BeginsWith = vbNullString And EndsWith = vbNullString Then
            Include = True
        Else
            If BeginsWith <> vbNullString Then
                If StrComp(Left(FoundCell.Text, Len(BeginsWith)), BeginsWith, BeginEndCompare) = 0 Then
                    Include = True
                End If
            End If
            If EndsWith <> vbNullString Then
                If StrComp(Right(FoundCell.Text, Len(EndsWith)), EndsWith, BeginEndCompare) = 0 Then
                    Include = True
                End If
            End If
        End If
        If Include = True Then
            If ResultRange Is Nothing Then
                Set ResultRange = FoundCell
            Else
                Set ResultRange = Application.Union(ResultRange, FoundCell)
            End If
        End If
        Set FoundCell = SearchRange.FindNext(After:=FoundCell)
        If (FoundCell Is Nothing) Then
            Exit Do
        End If
        If (FoundCell.Address = FirstFound.Address) Then
            Exit Do
        End If

    Loop
End If
    
Set FindAll = ResultRange

End Function
Sub remove_nondata_rows()
'
' removes nondata rows
'

'
    Range("A1").Select
    Dim rng1 As Range
    Set rng1 = Range("A:A").Find("https", , xlValues, xlPart)
    rng1.Select
    ActiveSheet.Range(Selection, Selection).EntireRow.Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Delete Shift:=xlUp


'
' bolds header

    Rows("1:1").Select
    Selection.Font.Bold = True
    Range("A1").Select
    
' replaces blanks with 0s
    
    Range("A1").CurrentRegion.Select
    Selection.Replace What:="", Replacement:="0", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("A1").Select
    
End Sub
Sub delete_unwanted_columns()
'
' delete_unwanted_columns Macro
'

'
    
    Dim rng2 As Range
    Set rng2 = FindAll(Range("1:1"), "p-value", xlValues, xlPart)
    Dim rng3 As Range
    Set rng3 = FindAll(Range("1:1"), "type", xlValues, xlPart)
    
    With ActiveSheet
        Set Rng = Union(rng2, rng3)
    End With
    
    Rng.EntireColumn.Select
    Selection.Delete
    Range("A1").Select
    
End Sub


Sub add_adjusted_score_columns()
'
' add_adjusted_score_columns Macro
'

'

    Dim rng4 As Range
    Set rng4 = FindAll(Range("1:1"), "Score", xlValues, xlPart, , True)
    
    rng4.EntireColumn.Select
    Selection.Insert Shift:=xlToRight
    
    
    rng4.Select
    Dim rng5 As Range
    Set rng5 = rng4.Offset(0, -1)
    rng5.Select
    
    
    Dim c As Range, sel As Range, i As Integer
    i = 1
    
    Set sel = Selection
    
    For Each c In sel.Cells
        c.Formula = "Adjusted Score " & i
        i = i + 1
    Next c
        
    
'    Selection.Formula = "Adjusted Score"
    Selection.EntireColumn.Interior.Color = 65535

'
' adjscore_formulas Macro
'

    rng5.Select
    Set rng6 = rng5.Offset(1, 0)
    rng6.Select

    
    Dim ws As Worksheet, lastR As Long
  
    Set ws = ActiveSheet
    lastR = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    Intersect(Selection.EntireColumn, ws.Rows(Selection.Row & ":" & lastR)).Formula = _
                                                          "=IF(RC[2]<0,RC[1]*-1,RC[1])"
    
    
    Range("A1").Select
End Sub

Sub make_analysis_worksheet()
'
' make_analysis_worksheet Macro
'

'
    
    Sheets(1).Copy After:=Worksheets(Worksheets.Count)
    Sheets(Worksheets.Count).Name = "analysis worksheet"
End Sub
Sub make_final_worksheet()
'
' make_final_worksheet Macro
'

'
    Sheets.Add After:=Worksheets(Worksheets.Count)
    Sheets(Worksheets.Count).Name = "final worksheet"
    
End Sub
Sub populate_final_worksheet()
'
' populate_final_worksheet Macro
'

'
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Gene"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "GEO"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Virus"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Disease Type"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Disease Severity"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Comparison"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Weeks P.I."
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Biosource"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Tissue Type"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "GEO Link"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Samples"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Platform"
    Range("A1").Select
End Sub
Sub move_gene_col_to_final()
'
' move_gene_col_to_final Macro
'

'
    Sheets("analysis worksheet").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("final worksheet").Select
    Range("M1").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("A1").Select
End Sub
Sub move_biosets_to_final()
'
' move_biosets_to_final Macro
'

'
    Sheets(1).Select
    Range("A1").Select
    Cells.Find(What:="https", After:=ActiveCell, LookIn:=xlValues, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, _
        SearchFormat:=False).Activate
    Selection.Offset(-1, 0).Select
    Range("A2", Selection).Select
    Selection.Copy
    Sheets("final worksheet").Select
    Range("A2").Select
    ActiveSheet.Paste
End Sub


