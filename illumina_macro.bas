Sub remove_nondata_rows()
'
' remove_nondata_rows Macro
'

'
    Range("A1").Select
    Dim rng1 As Range
    Set rng1 = Range("A:A").Find("https", , xlValues, xlPart)
    rng1.Select
    ActiveSheet.Range(Selection, Selection).EntireRow.Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Delete Shift:=xlUp


End Sub

