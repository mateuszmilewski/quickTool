Attribute VB_Name = "ObsoleteModule"
' Dim toggle As Boolean
' toggle = False
' Dim ir As Range
' For Each ir In tmp
' '
' ' If ir.Value = "PART" Then
' ' ' ir.EntireColumn.ColumnWidth = 14
' ' '
' ' '
' ' ' Set innerTmp = tmp.Parent.Range(ir, ir.End(xlDown))
' ' ' Me.fillSolidGridLines innerTmp, colors.colorMattBlack
' ' '
' ' ' innerTmp.Interior.Color = colors.colorMattBlueLight
' ' '
' ' End If
' '
' ' If ir.Value = "Plant" Or ir.Value = "Stock" Then
' ' ' ir.EntireColumn.ColumnWidth = 8
' ' ' ir.Font.Bold = True
' ' ' ir.Interior.Color = colors.colorMattBlueLight
' ' End If
' '
' '
' ' If ir.Value = "RQM" Then
' ' '
' ' ' Set innerTmp = tmp.Parent.Range(ir, ir.End(xlDown))
' ' ' Me.fillSolidGridLines innerTmp, colors.colorMattBlueDark
' ' '
' ' '
' ' '
' ' End If
' '
' '
' '
' '

' ' If ir.Value = "BALANCE" Then
' ' '
' ' ' Set innerTmp = tmp.Parent.Range(ir, ir.End(xlDown))
' ' ' innerTmp.Font.Bold = True
' ' ' innerTmp.Font.Size = 13
' ' ' innerTmp.Font.Color = RGB(255, 255, 255)
' ' '
' ' ' If toggle Then
' ' ' ' innerTmp.Interior.Color = colors.colorMattBlueDark
' ' ' Else
' ' ' ' innerTmp.Interior.Color = colors.colorMattBlueMain
' ' ' End If
' ' '
' ' '
' ' ' If IsDate(ir.Offset(-1, -3).Value) Then
' ' ' '
' ' ' ' wd = Weekday(ir.Offset(-1, -3).Value, vbMonday)
' ' ' '
' ' ' ' If Int(wd) = 6 Or Int(wd) = 7 Then
' ' ' '
' ' ' '
' ' ' ' ' innerTmp.Font.Italic = True
' ' ' ' ' innerTmp.Font.Size = 11
' ' ' ' ' innerTmp.Font.Color = colors.colorMattBlack
' ' ' ' '
' ' ' ' '
' ' ' ' ' If toggle Then
' ' ' ' ' ' innerTmp.Interior.Color = colors.colorMattBlueLight
' ' ' ' ' Else
' ' ' ' ' ' innerTmp.Interior.Color = colors.colorPurpleLight
' ' ' ' ' End If
' ' ' ' '
' ' ' ' End If
' ' ' End If
' ' '
' ' '
' ' '
' ' '
' ' ' If toggle Then
' ' ' ' toggle = False
' ' ' Else
' ' ' ' toggle = True
' ' ' End If
' ' '
' ' End If
' '
' '
' '
' '
' '
' '
' '
' '
' Next ir
'
' Me.fillSolidGridLines tmp, colors.colorMattBlack
' tmp.Font.Bold = True
' tmp.Font.Size = 11
