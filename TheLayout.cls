VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TheLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'
'Copyright (c) 2017 FORREST
' Mateusz Milewski mateusz.milewski@opel.com aka FORREST
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.


Private colors As ColorsHandler
Private jezyk As E_LANG




Public Sub setLang(eTmp As E_LANG)
    
    jezyk = eTmp
End Sub


Public Sub init(ByRef sh As Worksheet)
    
    ActiveWindow.Zoom = 80
    ActiveWindow.DisplayGridlines = False
End Sub


' TheLayout class
Public Function fillFirstRawLabels(r As Range) As Range

    
    r.Value = "PART"
    r.Offset(0, 1).Value = "Plant"
    r.Offset(0, 2).Value = "Stock"
    
    
    Set fillFirstRawLabels = r.Offset(0, 3)
End Function


Public Function getDatesFromDropperAndFillLabels(r As Range, c As Collection) As Range

    Dim d1 As Date
    Dim d2 As Date
    
    Dim firstTime As Boolean
    firstTime = True

    Dim el As DropperHandler
    For Each el In c
        
        If CBool(firstTime) Then
            
            d1 = el.getFirstDate
            d2 = el.getLastDate
        Else
            
            If d1 > el.getFirstDate Then
                d1 = el.getFirstDate
            End If
            
            If d2 < el.getLastDate Then
                d2 = el.getLastDate
            End If
        End If
        firstTime = False
    Next el
    
    
    Dim tempDate As Date
    tempDate = d1
    Do
        r.Offset(-1, 0).Value = CStr(tempDate)
        r.Offset(-1, 3).Value = CStr(Weekday(tempDate, vbMonday))
        
        r.Offset(0, 0).Value = "RQM"
        r.Offset(0, 1).Value = "ORDER"
        r.Offset(0, 2).Value = "SHIP"
        r.Offset(0, 3).Value = "BALANCE"
        
        Set r = r.Offset(0, 4)
    
        tempDate = tempDate + 1
    Loop Until CDate(tempDate) > CDate(d2)
    
    
    Set getDatesFromDropperAndFillLabels = r
    
End Function




' ponizej znajduja sie suby i funkcje,
' ktore pracuja na juz wykonanym raporcie,
' ktory narazie przechowuje tylko i wylacznie dane
' ------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------



' TOP ROW WITH DATES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub addDayNamesAndMergeTopDateRows(r As Range)
    
    addDayNames r
    mergeProperly r
End Sub

Private Sub addDayNames(ByRef r As Range)

    Dim regSh As Worksheet
    Set regSh = ThisWorkbook.Sheets("register")
    
    ' r = E3
    Dim tmp As Range
    Set tmp = r
    Set tmp = tmp.Offset(0, 3)
    
    Do
        tmp.Value = CStr(regSh.Cells(Int(tmp.Value), Int(jezyk)).Value)
        Set tmp = tmp.Offset(0, 4)
    Loop Until Trim(tmp) = ""

End Sub

Private Sub mergeProperly(ByRef r As Range)

    Dim sh  As Worksheet
    Set sh = r.Parent
    Dim tmp As Range, tmp1 As Range, tmp2 As Range
    
    Dim toggle As Boolean
    toggle = False
    
    Do
        Set tmp = r
        Set tmp1 = tmp
        Set tmp2 = r.Offset(0, 2)
        
        Set r = r.Offset(0, 3)
        
        tmp1.HorizontalAlignment = xlCenter
        sh.Range(tmp1, tmp2).Merge
        
        
        tmp1.Font.Bold = True
        r.Font.Bold = True
        
        
        If toggle Then
        
            With tmp1
                .Interior.Color = colors.colorMattBlueDark
                .Font.Color = RGB(255, 255, 255)
            End With
            'Me.fillThinFrame tmp1, colors.colorMattBlueLight
            
            With r
                .Interior.Color = colors.colorMattBlueDark
                .Font.Color = RGB(255, 255, 255)
    
            End With
            'Me.fillThinFrame r, colors.colorMattBlack
        Else
            
            With tmp1
                .Interior.Color = colors.colorMattBlueMain
                .Font.Color = RGB(255, 255, 255)
            End With
            'Me.fillThinFrame tmp1, colors.colorMattBlueLight
            
            With r
                .Interior.Color = colors.colorMattBlueMain
                .Font.Color = RGB(255, 255, 255)
    
            End With
            'Me.fillThinFrame r, colors.colorMattBlack
            
        End If
        Set r = r.Offset(0, 1)
        
        If toggle Then
            toggle = False
        Else
            toggle = True
        End If
        
    Loop Until Trim(r) = ""

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' PART PLANT STOCK ROW TO FILL
Public Sub fillWithColorsAndGrid(ByRef r As Range)


    Dim tmp As Range, innerTmp As Range, obszar As Range
    Set tmp = r.Parent.Range(r, r.End(xlToRight))
    
    Dim lastRow As Long
    Dim lastCol As Long
    
    lastRow = r.End(xlDown).Row
    lastCol = r.End(xlToRight).Column
    
    Set obszar = r.Parent.Range(r, r.Parent.Cells(lastRow, lastCol))
    Me.fillThinGridLines obszar, colors.colorMattBlack
    
    
    
End Sub



' DYNAMIC COLORS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub colourBalances(ByRef r As Range)

    ' change dynamically formats
    adjustBalanceCellsFormat r
    adjustRqmCellsFormat r

End Sub


Private Sub adjustBalanceCellsFormat(ByRef r As Range)

    Dim tmpr As Range, area As Range
    Set tmpr = r
    Do
        Set area = prepareArea(tmpr)
        area.NumberFormat = "0_ ;[Red]-0 ;[Green]0_ "
        Set tmpr = tmpr.Offset(0, 4)
    Loop Until Trim(tmpr) = ""
End Sub

Private Sub adjustRqmCellsFormat(ByRef r As Range)

    Dim tmpr As Range, area As Range
    Set tmpr = r
    Do
        Set area = prepareArea(tmpr)
        area.Offset(0, -3).NumberFormat = "[Blue]0_ ;[Red]-0 ;0_ "
        Set tmpr = tmpr.Offset(0, 4)
    Loop Until Trim(tmpr) = ""
End Sub


Private Function prepareArea(r As Range) As Range
    
    Dim tmp As Range
    
    Set tmp = r
    Do
        Set tmp = tmp.Offset(1, 0)
    Loop Until Trim(tmp) = ""
    
    Set tmp = r.Parent.Range(r, tmp.Offset(-1, 0))
    
    
    Set prepareArea = tmp
    
    
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Public Sub backColor(rng As Range, kolor As Long)

    With rng.Interior
        .Pattern = xlSolid
        .Color = kolor
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
 
Public Sub changeTxtOrientation(rng As Range, angle As Integer)

    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = Int(angle)
        
    End With
End Sub


Public Sub fillSolidFrame(rng As Range, kolor As Long)


    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
End Sub

Public Sub fillSolidGridLines(rng As Range, kolor As Long)
    
    
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
End Sub

Public Sub fillThinFrame(rng As Range, kolor As Long)


    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Public Sub fillThinGridLines(rng As Range, kolor As Long)

    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub






' ------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------


Private Sub Class_Initialize()
    Set colors = New ColorsHandler
End Sub

Private Sub Class_Terminate()
    Set colors = Nothing
End Sub
