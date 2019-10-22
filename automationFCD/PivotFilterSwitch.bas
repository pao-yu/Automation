Sub PivotFilterSwitch()

    Dim allSheet As Worksheet
    Dim allPivot As PivotTable
    
    Dim itemName As String
    Dim filterField As String
    Dim answer As Integer
    
    filterField = "Nameplate"
    answer = MsgBox("Change filters for "& filterField & "?", vbYesNo + vbQuestion)
    
        If answer = vbYes Then
        
        itemName = InputBox(Prompt:="Input name of filter item:", Title:="ENTER FILTER")

            For Each allSheet In Worksheets
                For Each allPivot In allSheet.PivotTables
                   'allPivot.PivotFields(filterField).ClearAllFilters
                    allPivot.PivotFields(filterField).CurrentPage = itemName
                Next allPivot
            Next allSheet

            MsgBox "Success."

        Else
        ' Do nothing.
        End If

End Sub
