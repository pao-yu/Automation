Sub PPivotFieldSwitch_Multi()

    Call PivotFieldSwitch
    MsgBox "Success."

End Sub

' --------------------------------------------------------------------------------------------------------

Sub PivotFieldSwitch()

    Dim list As Range
    Dim listItem As Range
    Dim filterTable As PivotTable
    Dim filterField As PivotField
    Dim filterItem As PivotItem
    
    Set list = Range("LIST CRITERIA TABLE NAME")
    Set filterTable = Sheets("SHEET WITH FILTERED TABLE").PivotTables("FILTERED TABLE NAME")
    Set filterField = filterTable.PivotFields("FILTERED TABLE FIELD")
    
        With filterField

            For Each filterItem In .PivotItems
                filterItem.Visible = False

                    For Each listItem In list

                        If filterItem.Caption = listItem.Text Then
                        filterItem.Visible = True
                        Exit For
                        End If
                        
                    Next listItem
                    
            Next filterItem
            
        End With

End Sub
