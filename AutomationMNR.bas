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

  Call AutomationMNR_C0000_CloneRename
  Call AutomationMNR_C0001_CloneClearIf
  
End Sub


Sub AutomationMNR_C0000_CloneRename()                                           ' Duplicate sheet and paste table as values
  
    Dim oldSheetName As String
    Dim newSheetName As String
    
    oldSheetName = "Data"
    newSheetName = "Tier 2"
    
    Worksheets(oldSheetName).Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = newSheetName
  
    Worksheets(newSheetName).ListObjects.Item(1).DataBodyRange.Copy
    Worksheets(newSheetName).ListObjects.Item(1).DataBodyRange.PasteSpecial xlValues
  
End Sub



Sub AutomationMNR_C0001_CloneClearIf()
  
Dim fullCounter As Long

  ' Set fullCounter for 1st Criteria
  fullCounter = ActiveWorkbook.Worksheets("Tier 2").ListObjects.Item(1).DataBodyRange.Rows.Count

  For i = 2 To fullCounter + 1
    If Not Worksheets("Tier 2").Cells(i, 2).Value = "Tier 2" Then
      Sheets("Tier 2").Cells(i, 2).EntireRow.Delete
    End If
  Next i

   ' Reset fullCounter for 2nd Criteria
   fullCounter = ActiveWorkbook.Worksheets("Tier 2").ListObjects.Item(1).DataBodyRange.Rows.Count
    For i = 2 To fullCounter + 1
      If Worksheets("Tier 2").Cells(i, 2).Value = "Tier 2" And _
          Worksheets("Tier 2").Cells(i, 3).Value = "OOH" Then
          Sheets("Tier 2").Cells(i, 2).EntireRow.Delete
       End If
       Next i

   ' Reset fullCounter for 3rd Criteria
   fullCounter = ActiveWorkbook.Worksheets("Tier 2").ListObjects.Item(1).DataBodyRange.Rows.Count
    For i = 2 To fullCounter + 1
      If Worksheets("Tier 2").Cells(i, 2).Value = "Tier 2" And _
          Worksheets("Tier 2").Cells(i, 3).Value = "Local Newspapers" Then
          Sheets("Tier 2").Cells(i, 2).EntireRow.Delete
       End If
            Next i

   ' Reset fullCounter for 4th Criteria
   fullCounter = ActiveWorkbook.Worksheets("Tier 2").ListObjects.Item(1).DataBodyRange.Rows.Count
    For i = 2 To fullCounter + 1
      If Worksheets("Tier 2").Cells(i, 2).Value = "Tier 2" And _
          Worksheets("Tier 2").Cells(i, 3).Value = "Magazines" Then
          Sheets("Tier 2").Cells(i, 2).EntireRow.Delete
       End If

  Next i

End Sub
