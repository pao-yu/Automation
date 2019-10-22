' ----------------------------------------------------------------------------------------------------------------------------------------------------------
' -------- AutomationMNR.bas -------------------------------------------------------------------------------------------------------------------------------
' ----------------------------------------------------------------------------------------------------------------------------------------------------------
'
'
' This file is a documemntation and automation of a series of actions on a reporting data set.
' The prerequisite actions of this automation module are:
'
'     1. The oldSheetName and newSheetName string variables must be defined by the user.
'     2. The data in oldSheetName must be converted into an Excel Table (list object).
'
' From these inputs, the automation module then proceeds with the individual process steps,
' called through the master subroutine.calling procedure, "AutomationMNR_Cluster()".
'
'
' ----------------------------------------------------------------------------------------------------------------------------------------------------------
' -------- Master Program ----------------------------------------------------------------------------------------------------------------------------------
' ----------------------------------------------------------------------------------------------------------------------------------------------------------


Sub AutomationMNR_Cluster()                                                         '  VBA Calling Procedure

  Call AutomationMNR_C0000_TableGeneration
  Call AutomationMNR_C0001_CloneClearIfA
  Call AutomationMNR_C0002_CloneClearIfB
  
End Sub


Sub AutomationMNR_C0000_CloneRename()                                               ' Duplicate sheet and paste table as values
  
    Dim oldSheetName As String
    Dim newSheetName As String
    
    oldSheetName = "Data"
    newSheetName = "Tier 2"
    
    Worksheets(oldSheetName).Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = newSheetName
  
    Worksheets(newSheetName).ListObjects.Item(1).DataBodyRange.Copy
    Worksheets(newSheetName).ListObjects.Item(1).DataBodyRange.PasteSpecial xlValues
  
End Sub



Sub AutomationMNR_C0001_CloneClearIfA()
    
  Dim fullCounter As Long
  fullCounter = ActiveWorkbook.Worksheets("Tier 2").ListObjects.Item(1).DataBodyRange.Rows.Count

  For i = 2 To fullCounter + 1
    If Not Worksheets("Tier 2").Cells(i, 2).Value = "Tier 2" Then
    Sheets("Tier 2").Cells(i, 2).EntireRow.Interior.ColorIndex = 6 '.Delete
    End If
  Next i
  
End Sub



Sub AutomationMNR_C0002_CloneClearIfB()

  Dim fullCounter As Long
  fullCounter = ActiveWorkbook.Worksheets("Tier 2").ListObjects.Item(1).DataBodyRange.Rows.Count

  For i = 2 To fullCounter + 1
  
    If  Worksheets("Tier 2").Cells(i, 2).Value = "Tier 2" And _
        Worksheets("Tier 2").Cells(i, 3).Value = "OOH" Or _
        Worksheets("Tier 2").Cells(i, 3).Value = "Local Newspapers" Or _
        Worksheets("Tier 2").Cells(i, 3).Value = "Magazines" Then
        
  Sheets("Tier 2").Cells(i, 2).EntireRow.Interior.ColorIndex = 6 '.Delete
        
     End If
Next i

End Sub

