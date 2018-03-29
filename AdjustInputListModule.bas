Attribute VB_Name = "AdjustInputListModule"
Public Sub adjustInputList()

    
    innerAdjustInputList
End Sub


Private Sub innerAdjustInputList()
    
    Dim sh As Worksheet
    Dim i As Worksheet
    Set sh = ThisWorkbook.Sheets(QT.G_SH_NM_PLT_LIST)
    Set i = ThisWorkbook.Sheets(QT.G_SH_NM_IN)
    
    
    Dim ir As Range
    Set ir = i.Range("A2")
    Dim pltr As Range
    
    
    Do
        Set pltr = sh.Range("A2")
        
        If Len(ir.Value) = 1 Then
        
        
            Set pltr = sh.Range("A2")
            Do
                If ir.Value = pltr.Value Then
                    ir.Offset(0, 2).Value = pltr.Offset(0, 3).Value
                    Exit Do
                End If
                Set pltr = pltr.Offset(1, 0)
            Loop Until Trim(pltr) = ""
            
            
            If ir.Offset(0, 2).Value = "" Then
                ir.Offset(0, 2).Value = "MANUAL"
            End If
        
        Else
            
            ' najpierw dopasuj nazwe plantu
            Set pltr = sh.Range("A2")
            Do
                If UCase(ir.Value) Like "*" & UCase(Trim(Replace(CStr(pltr.Offset(0, 1).Value), "Corail", ""))) & "*" Then
                    ir.Value = pltr.Value
                    ir.Offset(0, 2).Value = pltr.Offset(0, 3).Value
                    Exit Do
                End If
                Set pltr = pltr.Offset(1, 0)
            Loop Until Trim(pltr) = ""
            
            If ir.Offset(0, 2).Value = "" Then
                ir.Offset(0, 2).Value = "MANUAL"
            End If
            
        End If
        Set ir = ir.Offset(1, 0)
    Loop Until Trim(ir) = ""
End Sub
