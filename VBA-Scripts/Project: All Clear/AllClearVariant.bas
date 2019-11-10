Sub VariantAllClearMirror()
Application.ScreenUpdating = False

Dim answer As Integer
answer = MsgBox("Highlight matching rows?", vbYesNo + vbQuestion, "All Clear")

If answer = vbYes Then
        
    Call VariantAllClearForward
    Call VariantAllClearReverse
    
Else
' Do Nothing

End If

Application.ScreenUpdating = True
End Sub

Sub VariantAllClearForward()

' This code was created to cleanse a database of invoices, removing cleared invoices based on another sheet.
' For this code to work, the sheets must be named "ALL" and "CLEAR" based on their contents.
' In this particular code iteration, the key columns between the sheets are:
' ALL   = Column #13 / "M"
' CLEAR = Column #02 / "B"


        Dim fullCounter As Integer
        Dim nomatchCounter As Long
        Dim i As Integer
        Dim varAll As String
        Dim varClear As String
        
        Dim all As String
        Dim clear As String
        

                ' ALL sheet contains the list of all values with duplicates.
                ' CLEAR sheet contains the list of duplicates which will be deleted from the ALL sheet.
                
                all = "ALL"
                clear = "CLEAR"
        
                        ' 1st Loop - Loop through each value in the CLEAR sheet column.
                        ' 2nd Loop - Loop through each value in the ALL sheet column.
                        ' fullCounter is the maximum loop # which will decrease by "1" until it reaches "1".
                        ' "1 Step -1" prevents the loop from accessing the header column.
                        ' "i" is the interval between the loops.
                        ' "x" is one cell for each cell in the CLEAR sheet colum.
                        
                        varAll = Sheets("home").Range("G22").Value
                        varClear = Sheets("home").Range("H22").Value
                        
                        fullCounter = Sheets(all).Cells(Rows.Count, varAll).End(xlUp).Row
                        matchList = Sheets(clear).Cells(Rows.Count, varClear).End(xlUp).Row
                        matchCounter = 0
                        
                        For Each x In Sheets(clear).Range(varClear & "1:" & varClear & Sheets(clear).Cells(Rows.Count, varClear).End(xlUp).Row)
                        
                              With x
                                      .NumberFormat = "0"
                                      .Value = .Value
                              End With
                                
                           For i = fullCounter To 1 Step -1

                              If x.Value = Sheets(all).Cells(i, varAll).Value Then
                                      Sheets(all).Cells(i, varAll).EntireRow.Interior.ColorIndex = 6
                                    matchCounter = matchCounter + 1
                              End If
                           Next i
                          
                        Next
                    'MsgBox matchCounter


End Sub

Sub VariantAllClearReverse()

' This code was created to cleanse a database of invoices, removing cleared invoices based on another sheet.
' For this code to work, the sheets must be named "ALL" and "CLEAR" based on their contents.
' In this particular code iteration, the key columns between the sheets are:
' ALL   = Column #13 / "M"
' CLEAR = Column #02 / "B"


        Dim fullCounter As Integer
        Dim nonMatchCounter As Integer
        Dim matchCounter As Integer
        Dim i As Integer
        Dim varAll As String
        Dim varClear As String
        
        Dim all As String
        Dim clear As String

        
                all = "ALL"
                clear = "CLEAR"
                        
                        varAll = Sheets("home").Range("G22").Value
                        varClear = Sheets("home").Range("H22").Value
                        
                        fullCounter = Sheets(clear).Cells(Rows.Count, varClear).End(xlUp).Row
                        matchList = Sheets(all).Cells(Rows.Count, varAll).End(xlUp).Row
                        matchCounter = 0
                        
                        For Each x In Sheets(all).Range(varAll & "1:" & varAll & Sheets(all).Cells(Rows.Count, varAll).End(xlUp).Row)
                        
                              With x
                                      .NumberFormat = "0"
                                      .Value = .Value
                              End With
                                
                           For i = fullCounter To 1 Step -1

                              If x.Value = Sheets(clear).Cells(i, varClear).Value Then
                              Sheets(clear).Cells(i, varClear).EntireRow.Interior.ColorIndex = 6
                              matchCounter = matchCounter + 1
                                End If
                            
                           Next i
                        Next
                        
                        nonMatchCounter = fullCounter - matchCounter - 1
                        
                        If nonMatchCounter = (fullCounter - 1) Then
                            MsgBox "CLEAR sheet results:" & vbNewLine & "No matches identified. Please double check if reference key columns were assigned correctly.", , "All Clear...?"
                        Else
                            MsgBox "CLEAR sheet results: " & vbNewLine & matchCounter & " matches identified." & vbNewLine & nonMatchCounter & " records had no matches.", , "Success."
                        End If
    
End Sub

Sub CheckLookUp()

        Dim all As String
        Dim clear As String
        Dim rLastCell As Range
        Dim vLookupCell As Range
        Dim LookupRange As Range
        Dim fullRangeAll As Range
        Dim fullRangeClear As Range
        
        clear = "CLEAR"
        all = "ALL"
        varAll = Sheets("home").Range("G22").Value
        varClear = Sheets("home").Range("H22").Value
        varClearReturn = Sheets("home").Range("I22").Value

        Set rLastCell = Sheets(all).Cells.Find(What:="*", _
                                                            After:=Sheets(all).Cells(1, 1), _
                                                            LookIn:=xlFormulas, _
                                                            LookAt:=xlPart, _
                                                            SearchOrder:=xlByColumns, _
                                                            SearchDirection:=xlPrevious, _
                                                            MatchCase:=False)

        fullCounter = Sheets(all).Cells(Rows.Count, 1).End(xlUp).Row
        fullCounterClear = Sheets(clear).Cells(Rows.Count, 1).End(xlUp).Row
        
        Sheets(all).Cells(1, rLastCell.Column + 1).Interior.ColorIndex = 6
        Sheets(all).Cells(1, rLastCell.Column + 1).Value = "Check Lookup"

        Set vLookupCell = Sheets(all).Cells.Find(What:="Check Lookup", _
                                                            After:=Sheets("ALL").Cells(1, 1), _
                                                            LookIn:=xlFormulas, _
                                                            LookAt:=xlPart, _
                                                            SearchOrder:=xlByColumns, _
                                                            SearchDirection:=xlPrevious, _
                                                            MatchCase:=False)

        For i = 2 To fullCounter
            Sheets(all).Cells(i, vLookupCell.Column).Formula = "=INDEX(CLEAR!$A$1:$P$81,MATCH(" & varAll & i & ",CLEAR!$" & varClear & "$2:$" & varClear & "$" & fullCounterClear & ",0),4)"
                                                                                                            ' Last Work : You were trying to figure out how to convert the #4 to "E"
            If Not IsError(Sheets(all).Cells(i, vLookupCell.Column).Value) Then
                              Sheets(all).Cells(i, vLookupCell.Column).Interior.ColorIndex = 6
                                End If
        Next i

End Sub
