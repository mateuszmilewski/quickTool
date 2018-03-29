Attribute VB_Name = "CorailOpenPN"
Public Sub openPartNumber(ictrl As IRibbonControl)

    innerOpenPartNumber
End Sub


Private Sub innerOpenPartNumber()

    Dim r As Range
    If ActiveSheet.Name = QT.G_SH_NM_IN Then
        
        Set r = Cells(ActiveCell.Row, 1)
        
        GoToPNForm.TextBoxPlt = r.Value
        GoToPNForm.TextBoxPN = r.Offset(0, 1).Value
        GoToPNForm.TextBoxCorailType = r.Offset(0, 2).Value
    
    Else
        
        If ActiveSheet.Cells(4, 2).Value = "PART" And ActiveSheet.Cells(4, 3).Value = "Plant" Then
        
            Set r = Cells(ActiveCell.Row, 3)
        
            GoToPNForm.TextBoxPlt = r.Value
            GoToPNForm.TextBoxPN = r.Offset(0, -1).Value
            GoToPNForm.TextBoxCorailType = getCorailType(r.Value)
        Else
            GoToPNForm.TextBoxPlt = ""
            GoToPNForm.TextBoxPN = ""
            GoToPNForm.TextBoxCorailType = ""
        End If
    End If
    
    GoToPNForm.Show vbModeless
End Sub


Private Function getCorailType(str) As String

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(QT.G_SH_NM_PLT_LIST)
    
    Dim r As Range
    Set r = sh.Range("A2")
    Do
        If r.Value = str Then
            getCorailType = r.Offset(0, 3).Value
            Exit Do
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Function
