' This file is a documemntation and automation of a series of actions on a reporting data set.
' It is named as "C_Master" because of it's categorization as the third phase of a specific process.
' The prerequisite actions of this automation module are:
'
'     1. The oldSheetName and newSheetName variables must be defined by the user.
'     2. The data in oldSheetName must be converted into an Excel Table (list object).
'
' From these inputs, the automation module then proceeds with the individual process steps,
' called through the master subroutine, "AutomationMNR_CMaster()".


' ----------------------------------------------------------------------------------------------------------
' -------- Master Program ----------------------------------------------------------------------------------
' ----------------------------------------------------------------------------------------------------------


Sub AutomationMNR_CMaster()                         ' VBA Calling Procedure

  Dim oldSheetName As String
  Dim newSheetName As String

  oldSheetName = "Data"
  newSheetName = "Tier 2"

      Call AutomationMNR_C0001_CopySheet(oldSheetName, newSheetName)
      Call AutomationMNR_C0002_CopyPasteAsValues(oldSheetName, newSheetName)
  
End Sub


' ----------------------------------------------------------------------------------------------------------
' -------- Automation Steps --------------------------------------------------------------------------------
' ----------------------------------------------------------------------------------------------------------


' ----------------------------------------------------------------------------------------------------------


Sub AutomationMNR_C0001_CopySheet(ByVal oldSheetName As String, ByVal newSheetName As String)
  
  Sheets(oldSheetName).Copy(After:=Sheets(Sheets.Count))
  ActiveSheet.Name = newSheetName
  
End Sub


' ----------------------------------------------------------------------------------------------------------


Sub AutomationMNR_C0002_CopyPasteAsValues(ByVal oldSheetName As String, ByVal newSheetName As String)
  
  Worksheets(newSheetName).ListObjects.Item(1).DataBodyRange.Copy
  Worksheets(newSheetName).ListObjects.Item(1).DataBodyRange.PasteSpecial xlValues
  
End Sub




