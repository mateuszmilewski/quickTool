VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParserHandler"
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

Private dateCatcher As String

Private besoinsPCCatcher As String
Private besoinsBCCatcher As String

Private commandCatcher As String
Private expCatcher As String


Public Property Get pDateCatcher() As String
    pDateCatcher = dateCatcher
End Property

Public Property Get pBesoinsPCCattcher() As String
    pBesoinsPCCattcher = besoinsPCCatcher
End Property

Public Property Get pBesoinsBCCatcher() As String
    pBesoinsBCCatcher = besoinsBCCatcher
End Property

Public Property Get pCmdCatcher() As String
    pCmdCatcher = commandCatcher
End Property

Public Property Get pExpCatcher() As String
    pExpCatcher = expCatcher
End Property




Public Function htmlTableToRawMatrix(tb As HTMLTable, dom As DOMHandler, e As E_TYPE_OF_CORAIL) As ConvertedData

    Dim htmlString As String
    Dim tmp As ConvertedData
    
    
    Dim i As CorailItem
    Dim ii As New CorailIteration
    Dim innerTb As HTMLTable

    Dim r As HTMLTableRow
    Dim c As HTMLTableCell
    
    Dim s As String
    
    Dim tmpStock As String
    Dim tmpRecv As String
    Dim tmpBacklog As String
    
    
    htmlString = tb.innerHTML
    If checkIfHTMLTable(htmlString) Then
    
        If e = ORANGE Then
            
            
            Set tmp = New ConvertedData
            Set innerTb = tb.getElementsByTagName("table")(0)

            avoidHeading = True
            For Each r In innerTb.Rows
            
                Set i = Nothing
                Set i = New CorailItem
            
                For Each c In r.Cells
                    
                    
                    If avoidHeading Then
                        Exit For
                    Else
                        ' Debug.Print c.innerHTML
                        
                        
                        If UCase(c.innerHTML) Like "*" & UCase(Me.pDateCatcher) & "*" Then
                            
                            
                            s = CStr(c.innerHTML)
                            arr = Split(s, ">")
                            
                            s = Left(arr(1), 10)
                            
                            i.parseStringToDate s
                            
                            
                        ElseIf UCase(c.innerHTML) Like "*" & UCase(Me.pBesoinsPCCattcher) & "*" Then
                            
                            s = CStr(c.innerHTML)
                            arr = Split(s, ">")
                            s = arr(1)
                            s = Replace(UCase(s), "</DIV", "")
                            s = Replace(s, ".0", "")
                            s = Replace(s, ",0", "")
                            
                            i.besoinsPC = CLng(s)
                            
                        ElseIf UCase(c.innerHTML) Like "*" & UCase(Me.pBesoinsBCCatcher) & "*" Then
                        
                            s = CStr(c.innerHTML)
                            arr = Split(s, ">")
                            s = arr(1)
                            s = Replace(UCase(s), "</DIV", "")
                            s = Replace(s, ".0", "")
                            s = Replace(s, ",0", "")
                            
                            i.besoinsBC = CLng(s)
                            
                        ElseIf UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher) & "*" Then
                        
                            s = CStr(c.innerHTML)
                            arr = Split(s, ">")
                            s = arr(1)
                            s = Replace(UCase(s), "</DIV", "")
                            s = Replace(s, ".0", "")
                            s = Replace(s, ",0", "")
                            
                            i.order = CLng(s)
                            
                        ElseIf UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher) & "*" Then
                        
                            s = CStr(c.innerHTML)
                            arr = Split(s, ">")
                            s = arr(1)
                            s = Replace(UCase(s), "</DIV", "")
                            s = Replace(s, ".0", "")
                            s = Replace(s, ",0", "")
                            
                            i.ship = CLng(s)
                            
                        End If
                        
                    End If
                    
                Next c
                
                If Not avoidHeading Then
                
                    ii.addItem i
                End If
                
                avoidHeading = False
            Next r
            
            
            
            tmpStock = dom.stockInputElement.Value
            tmpStock = Replace(tmpStock, ".0", "")
            tmpStock = Replace(tmpStock, ",0", "")
            
            If Not IsNumeric(tmpStock) Then tmpStock = "0"
            
            tmpRecv = dom.recvInputElement.Value
            tmpRecv = Replace(tmpRecv, ".0", "")
            tmpRecv = Replace(tmpRecv, ",0", "")
            
            If Not IsNumeric(tmpRecv) Then tmpRecv = "0"
            
            tmpBacklog = dom.backlogInputElement.Value
            tmpBacklog = Replace(tmpBacklog, ".0", "")
            tmpBacklog = Replace(tmpBacklog, ",0", "")
            
            If Not IsNumeric(tmpBacklog) Then tmpBacklog = "0"
            
            Set tmp.iteration = ii
            tmp.iteration.setStock CLng(tmpStock)
            tmp.iteration.setRecv CLng(tmpRecv)
            tmp.iteration.setBacklog CLng(tmpBacklog)
            Set htmlTableToRawMatrix = tmp
            
        
        ElseIf e = BLUE Then
        
        

        
            
            Set tmp = New ConvertedData
            Set innerTb = tb.getElementsByTagName("table")(0)
            avoidHeading = True
            For Each r In innerTb.Rows
            
                Set i = Nothing
                Set i = New CorailItem
            
                For Each c In r.Cells
                    
                    
                    If avoidHeading Then
                        Exit For
                    Else
                        ' Debug.Print c.innerHTML
                        
                        
                        If UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & UCase(Me.pDateCatcher) & "*" Then
                            
                            
                            s = CStr(c.innerHTML)
                            arr = Split(s, ">")
                            
                            s = Left(arr(1), 10)
                            
                            i.parseStringToDate s
                            
                            
                        ElseIf UCase(c.innerHTML) Like "*" & UCase(Me.pBesoinsPCCattcher) & "*" Then
                            
                            s = CStr(c.innerHTML)
                            arr = Split(s, ">")
                            s = arr(1)
                            s = Replace(UCase(s), "</DIV", "")
                            s = Replace(s, ".0", "")
                            s = Replace(s, ",0", "")
                            
                            i.besoinsPC = CLng(s)
                            
                        ElseIf UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & UCase(Me.pBesoinsBCCatcher) & "*" Then
                        
                            s = CStr(c.innerHTML)
                            arr = Split(s, ">")
                            s = arr(1)
                            s = Replace(UCase(s), "</DIV", "")
                            s = Replace(s, ".0", "")
                            s = Replace(s, ",0", "")
                            
                            i.besoinsBC = CLng(s)
                            
                        ElseIf UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher) & "*" Then
                        
                            s = CStr(c.innerHTML)
                            arr = Split(s, ">")
                            s = arr(1)
                            s = Replace(UCase(s), "</DIV", "")
                            s = Replace(s, ".0", "")
                            s = Replace(s, ",0", "")
                            
                            i.order = CLng(s)
                            
                        ElseIf UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher) & "*" Then
                        
                            s = CStr(c.innerHTML)
                            arr = Split(s, ">")
                            s = arr(1)
                            s = Replace(UCase(s), "</DIV", "")
                            s = Replace(s, ".0", "")
                            s = Replace(s, ",0", "")
                            
                            i.ship = CLng(s)
                            
                        End If
                        
                    End If
                    
                Next c
                
                If Not avoidHeading Then
                
                    ii.addItem i
                End If
                
                avoidHeading = False
            Next r
            
            tmpStock = dom.stockInputElement.Value
            tmpStock = Replace(tmpStock, ".0", "")
            tmpStock = Replace(tmpStock, ",0", "")
            
            If Not IsNumeric(tmpStock) Then tmpStock = "0"
            
            tmpRecv = dom.recvInputElement.Value
            tmpRecv = Replace(tmpRecv, ".0", "")
            tmpRecv = Replace(tmpRecv, ",0", "")
            
            If Not IsNumeric(tmpRecv) Then tmpRecv = "0"
            
            tmpBacklog = dom.backlogInputElement.Value
            tmpBacklog = Replace(tmpBacklog, ".0", "")
            tmpBacklog = Replace(tmpBacklog, ",0", "")
            
            If Not IsNumeric(tmpBacklog) Then tmpBacklog = "0"
            
            Set tmp.iteration = ii
            tmp.iteration.setStock CLng(tmpStock)
            tmp.iteration.setRecv CLng(tmpRecv)
            tmp.iteration.setBacklog CLng(tmpBacklog)
            Set htmlTableToRawMatrix = tmp
        
        
        
        
        
        ElseIf e = MAESTRO Then
        
            Set tmp = New ConvertedData
            Set innerTb = tb
        
        
            For Each r In innerTb.Rows
            
                Set i = Nothing
                Set i = New CorailItem
                
                cCounter = 0
            
                For Each c In r.Cells
                    
                    ' Debug.Print c.innerHTML
                    
                    ' 0: date
                    ' 1: rqm
                    ' 2: order
                    ' 3: it
                    
                    ' always
                    s = CStr(c.innerHTML)
                    
                    If cCounter = 0 Then
                        i.parseStringToDate Trim(s)
                    ElseIf cCounter = 1 Then
                    
                        arr = Split(s, "<")
                        s = Trim(arr(LBound(arr)))
                        s = Replace(s, ".0", "")
                        s = Replace(s, ",0", "")
                        
                        i.besoinsBC = 0
                        i.besoinsPC = 0
                        If IsNumeric(s) Then i.besoinsBC = CLng(s)
                        
                    ElseIf cCounter = 2 Then
                    
                        arr = Split(s, "<")
                        s = Trim(arr(LBound(arr)))
                        s = Replace(s, ".0", "")
                        s = Replace(s, ",0", "")
                        i.order = 0
                        If IsNumeric(s) Then i.order = CLng(s)
                        
                    ElseIf cCounter = 3 Then
                        If s Like "<FONT*" Then
                            arr = Split(s, ">")
                            arr = Split(arr(LBound(arr) + 1), "</")
                        Else
                            arr = Split(s, "<")
                            arr = Split(arr(LBound(arr)), "</")
                        End If
                        s = Trim(arr(LBound(arr)))
                        s = Replace(s, ".0", "")
                        s = Replace(s, ",0", "")
                        
                        i.ship = 0
                        If IsNumeric(s) Then i.ship = CLng(s)
                    End If
                    
                    cCounter = cCounter + 1
                Next c
                
                ii.addItem i
            Next r
            
            If Not dom.istockElement Is Nothing Then
                tmpStock = dom.istockElement.innerHTML
                tmpStock = Replace(tmpStock, ".0", "")
                tmpStock = Replace(tmpStock, ",0", "")
                tmpStock = Trim(Replace(tmpStock, G_NBSP, ""))
            End If
            
            If Not IsNumeric(tmpStock) Then tmpStock = "0"
            
            If Not dom.ibacklogElement Is Nothing Then
                tmpBacklog = dom.ibacklogElement.innerHTML
                tmpBacklog = Replace(tmpBacklog, ".0", "")
                tmpBacklog = Replace(tmpBacklog, ",0", "")
                tmpBacklog = Trim(Replace(tmpBacklog, G_NBSP, ""))
            End If
            
            If Not IsNumeric(tmpBacklog) Then tmpBacklog = "0"
            
            
            Set tmp.iteration = ii
            tmp.iteration.setStock CLng(tmpStock)
            tmp.iteration.setRecv CLng(0)
            tmp.iteration.setBacklog CLng(tmpBacklog)
            Set htmlTableToRawMatrix = tmp
            
        Else
            MsgBox "brak zdefiniowanego systemu!"
            End
        End If
        
        ' -----------------------------------------------------------------
    
    End If
End Function



Private Function checkIfHTMLTable(htmlString As String) As Boolean


    Dim tmp As String
    tmp = CStr(htmlString)
    
    If UCase(tmp) Like UCase("*tr*td*td*tr*") Then
    
        checkIfHTMLTable = True
    Else
    
        checkIfHTMLTable = False
    End If
End Function

Private Sub Class_Initialize()
    dateCatcher = "<div class=date>"

    besoinsPCCatcher = "<div class=" & Chr(34) & "number fbpcDetailByDay" & Chr(34) & ""
    besoinsBCCatcher = "<div class=number>"

    commandCatcher = "<div class=" & Chr(34) & "ecwButtonTexteOverOrdered number" & Chr(34) & ""
    expCatcher = "<div class=" & Chr(34) & "ecwButtonTexteOverShipped number" & Chr(34) & ""
End Sub
