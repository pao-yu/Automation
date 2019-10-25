Option Explicit

' --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'   "UnifyPivotFilters" Field-level version. Changes specific field line items based on an excel named table list. 
'
'   Average run-time:   30 seconds (45 Pivot Tables)
'   Requirements:       Filter on pivots set to field level (filters on the table headers, items on rows).
'                       Excel named table containing a list of unique items to be applied as filter items.
'   Effect Hierarchy:   Pre-set worksheet, pivot table and excel named table.
'   Created by:         Pao Yu
'
' --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub UnifyPivotFilters_FieldLevel()

    Dim list As Range                                                                        ' Define all objects.
    Dim listItem As Range
    Dim filterTable As PivotTable
    Dim filterLabel As PivotField
    Dim filterItem As PivotItem
    
    Dim listName As String                                                                   ' Define all variables and variable types.
    Dim pivotTableSheetName As String
    Dim pivotTableName As String
    Dim pivotFieldLabel As String
    
    listName            = ""                                                                 ' Set all object names. Must be pre-set on the code level (case-sensitive).
    pivotTableSheet     = ""
    pivotTableName      = ""
    pivotFieldLabel     = ""
    
    Set list = Range(listName)                                                               
        Set filterTable = Sheets(pivotTableSheet).PivotTables(pivotTableName)                
            Set filterLabel = filterTable.PivotFields(pivotFieldLabel)
    
        With filterLabel                                                                     ' Indicate specific object to be manipulated.

            For Each filterItem In .PivotItems                                               ' Loop through each item inside the fieldLabel.
                filterItem.Visible = False                                                   ' Hide all pivot items by setting their visibility to false.
                    For Each listItem In list                                                ' Show a pivot item IF it appears in the list criteria.
                        If filterItem.Caption = listItem.Text Then
                        filterItem.Visible = True
                        Exit For
                        End If         
                    Next listItem
            Next filterItem

        End With

MsgBox "Success. All "& filterLabel & " filters applied using criteria from " & listName & "."

End Sub
