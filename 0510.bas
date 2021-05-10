Attribute VB_Name = "Module1"
Option Explicit
Sub 取消合併()
Dim rowcnt, mergerow As Long
Dim myrng As Range
rowcnt = Sheets(1).UsedRange.Rows.Count
For Each myrng In Range(Cells(2, "A"), Cells(rowcnt, "A"))
    myrng.Select
    mergerow = myrng.MergeArea.Count
    MsgBox "目前是" & mergerow & "列合併"
    myrng.UnMerge
    myrng.Resize(mergerow, 1) = myrng
Next
Sheets(1).Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
End Sub
Sub cancel()
Dim shtidx As Integer
For shtidx = 1 To Sheets.Count
Sheets(shtidx).Activate
Dim rowcnt, mergerow As Long
Dim myrng As Range
rowcnt = Sheets(shtidx).UsedRange.Rows.Count
For Each myrng In Range(Cells(2, "A"), Cells(rowcnt, "A"))
    myrng.Select
    mergerow = myrng.MergeArea.Count
    myrng.UnMerge
    myrng.Resize(mergerow, 1) = myrng
Next

Sheets(shtidx).Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
Next
End Sub
