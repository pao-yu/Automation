Option Explicit

' --------------------------------------------------------------------------------------------------------
'
'   "UnifyPivotFilters" Page-level version. Changes all pivot page filters into one through user input. 
'
'   Average run-time:   3~5 seconds (45 Pivot Tables)
'   Requirements:       Filter on pivot tables set to page level (Filters above the table).
'   Effect Hierarchy:   All pivot tables with user input page filter applied.
'   Created by:         Pao Yu
'
' --------------------------------------------------------------------------------------------------------

Sub UnifyPivotFilters_PageLevel()

    Dim allSheet As Worksheet                                                               ' Define all objects and object types.
    Dim allPivot As PivotTable

    Dim itemName As String                                                                  ' Define all variables and variable types.
    Dim filterLabel As String
    Dim answer As Integer
    
    filterLabel = "Nameplate"                                                               ' Set the LABEL of the page-level filter (case-sensitive).
    
    answer = MsgBox("Change filters for " & filterLabel & "?", vbYesNo + vbQuestion)

        If answer = vbYes Then
        itemName = InputBox(Prompt:="Enter "& filterLabel & ":", Title:="Filter Select")    ' Input an ITEM within the page-level filter (case-sensitive).
            For Each allSheet In Worksheets
                For Each allPivot In allSheet.PivotTables
                    allPivot.PivotFields(filterLabel).ClearAllFilters
                    allPivot.PivotFields(filterLabel).CurrentPage = _
                    itemName
                Next allPivot
            Next allSheet
            
            MsgBox "Success. " & filterLabel & " filters set to " & itemName & "."
        
        Else
            ' Do nothing.
        
        End If

End Sub
