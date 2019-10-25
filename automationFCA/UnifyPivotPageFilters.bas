Sub UnifyPivotPageFilters()

    Dim allSheet As Worksheet
    Dim allPivot As PivotTable

    Dim itemName As String
    Dim filterField As String
    Dim answer As Integer
    
    filterField = "Nameplate"
    answer = MsgBox("Change filters for " & filterField & "?", vbYesNo + vbQuestion)

        If answer = vbYes Then
            itemName = InputBox(Prompt:="Input "& filterField & ":", Title:="Page Filter Name")
            For Each allSheet In Worksheets
                For Each allPivot In allSheet.PivotTables
                    allPivot.PivotFields(filterField).ClearAllFilters
                    allPivot.PivotFields(filterField).CurrentPage = _
                    itemName
                Next allPivot
            Next allSheet
            
            MsgBox "Success. All " & filterField & " page filters set to " & itemName & "."
        
        Else
            ' Do nothing.
        
        End If

End Sub
