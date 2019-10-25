Option Explicit

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

        Else
          For Each ws In wb.Worksheets
            For Each pt In ws.PivotTables

              If pt.PivotCache.OLAP = False Then
                pt.ChangePivotCache _
                  wb.PivotCaches.Create(SourceType:=xlDatabase, _
                                        SourceData:=sourceTableName)
              End If
            Next pt
          Next ws

        End If
  
        For Each ws In ActiveWorkbook.Worksheets
          For Each pt In ws.PivotTables
            pt.CacheIndex = wb.Worksheets(1).PivotTables(1).CacheIndex
          Next pt
        Next ws

        MsgBox "Pivot datasource unification success." & _
                "Unified " & ActiveWorkbook.PivotTables.Count & "pivot tables into" & _
                          ActiveWorkbook.PivotCaches.Count _
                          & " pivot cache."
  Else
  ' Do nothing.
  End If

End Sub
