VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DropperHandler"
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


Private dane As ConvertedData
Private plt As String
Private pn As String

Public Sub init(argPlt, argPn, argDane As ConvertedData)

    Set dane = argDane
    plt = argPlt
    pn = argPn
End Sub


Public Sub run(sh As Worksheet, repRow As Long)

    Dim r As Range
    
    'pn
    sh.Cells(repRow, 2).Value = pn
    'plt
    sh.Cells(repRow, 3).Value = plt
    
    
    If Not dane Is Nothing Then
        'stock + recv
        sh.Cells(repRow, 4).Value = dane.iteration.getStock + dane.iteration.getRecv
        
    Else
        sh.Cells(repRow, 4).Value = 0
        
    End If
    
    
    If Not dane Is Nothing Then
    
    
        Dim item As CorailItem
        Set r = sh.Cells(repRow, 5)
        i = 0
        Do
            dstr = CStr(r.Parent.Cells(3, r.Column).Value)
            
            If Trim(dstr) <> "" Then
        
                For Each item In dane.iteration.pItems
                
                
                
                    If CStr(dstr) = CStr(item.getStringDate) Then
                
                        r.Value = (item.besoinsBC + item.besoinsPC)
                        r.Offset(0, 1).Value = item.order
                        r.Offset(0, 2).Value = item.ship
                        
                        Exit For
                    End If
    
                
                Next item
            
                r.Offset(0, 3).FormulaR1C1 = "=RC[-1]-RC[-3]+RC[-4]"
            
                i = i + 1
            
                Set r = r.Offset(0, 4)
            Else
                Exit Do
            End If
        Loop Until CStr(dstr) = ""
    Else
        
        Set r = sh.Cells(repRow, 5)
        Do
        
            dstr = CStr(r.Parent.Cells(3, r.Column).Value)
            
            If Trim(dstr) <> "" Then
            
            
                r.Offset(0, 3).FormulaR1C1 = "=RC[-1]-RC[-3]+RC[-4]"
                Set r = r.Offset(0, 4)
                
                
            Else
                Exit Do
            End If
        Loop Until CStr(dstr) = ""
    End If
End Sub



Private Sub runObsolete(sh As Worksheet, repRow As Long)




    Dim r As Range
    
    Set r = sh.Cells(Int(repRow), 1)
    r.Value = plt
    r.Offset(0, 1).Value = pn
    
    Set r = r.Offset(0, 2)

    If Not dane Is Nothing Then
    
        Dim item As CorailItem
        For Each item In dane.iteration.pItems
            ' Debug.Print (item.besoinsBC + item.besoinsPC)
            
            ' date
            r.Offset(-1, 0).Value = item.getStringDate()
            
            ' rqms sum
            r.Value = (item.besoinsPC + item.besoinsBC)
            
            ' order
            r.Offset(0, 1).Value = item.order
            
            ' ship
            r.Offset(0, 2).Value = item.ship
    
            Set r = r.Offset(0, 3)
        Next item
    Else
        
        MsgBox "fatal error in run obsolete"
        End
    End If
End Sub


Public Function getFirstDate() As Date

    If Not dane Is Nothing Then
    
        Dim item As CorailItem
        Set item = dane.iteration.pItems(1)
        Dim tmpDate As Date
        tmpDate = CDate(item.getDate())
        
    
        For Each item In dane.iteration.pItems
            
            If CDate(item.getDate) < CDate(tmpDate) Then
                tmpDate = item.getDate
            End If
            
        Next item
        
        getFirstDate = tmpDate
        
    Else
        getFirstDate = Date
    End If
End Function


Public Function getLastDate() As Date

    If Not dane Is Nothing Then
    
        Dim item As CorailItem
        Set item = dane.iteration.pItems(1)
        Dim tmpDate As Date
        tmpDate = CDate(item.getDate())
        
    
        For Each item In dane.iteration.pItems
            
            If CDate(item.getDate) > CDate(tmpDate) Then
                tmpDate = item.getDate
            End If
            
        Next item
        
        getLastDate = tmpDate
        
    Else
        getLastDate = Date
    End If
End Function
