Option Explicit

' --------------------------------------------------------------------------------------------------------
' "UnifyPivotData"    Changes and unifies all pivot data sources (and caches) into one. 
'
' Average run-time:   15:00 mins (45 Pivot Tables / 500K data rows)
' Requirements:       Excel Named Table as Data Source
' Effect Hierarchy:   Active Workbook Only
' Created by:         Pao Yu
' --------------------------------------------------------------------------------------------------------


Sub UnifyPivotData()

  Dim wb As Workbook
  Dim ws As Worksheet
  Dim pt As PivotTable
  Set wb = ActiveWorkbook
  
  Dim answer As Integer
  Dim sourceTableName As String
    
  answer = MsgBox("Unify Pivot Data?", vbYesNo + vbQuestion)

  If answer = vbYes Then
      
      sourceTableName = InputBox(Prompt:="Enter table name of new source data (case-sensitive).", Title:="Unify Source Data")
        
        If sourceTableName = "" Then
          MsgBox "Name not detected."
          Exit Sub
        Else                                                              ' Loop through and change data source for all Pivot Tables in the Active Workbook.
          For Each ws In wb.Worksheets                                    ' The data source is set from the user's input. Name must be from an Excel Named Table.
            For Each pt In ws.PivotTables
              If pt.PivotCache.OLAP = False Then
                pt.ChangePivotCache _
                  wb.PivotCaches.Create(SourceType:=xlDatabase, _
                                        SourceData:=sourceTableName)
              End If
            Next pt
          Next ws
        End If
  
    For Each ws In ActiveWorkbook.Worksheets                              ' Loop through and unify pivot caches for all Pivot Tables in the Active Workbook.
          For Each pt In ws.PivotTables                                   ' The cache is set from the first pivot table that appears in the first worksheet.
            pt.CacheIndex = wb.Worksheets(1).PivotTables(1).CacheIndex
          Next pt
        Next ws

  MsgBox MsgBox "Success. Pivot data unified."

  Else
  ' Do nothing.
  End If

End Sub
