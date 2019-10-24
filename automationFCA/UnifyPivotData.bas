Option Explicit


  Sub UnifyPivotData()
  
    Call UnifyAllPivotSources
    Call UnifyAllPivotCaches
  
  End Sub


      Sub UnifyAllPivotSources()

        Dim wb As Workbook
        Dim ws As Worksheet
        Dim pt As PivotTable
        Set wb = ActiveWorkbook

        Dim sourceTableName As String
        sourceTableName = InputBox(Prompt:="Enter Table Name", Title:="Source Data")

              If sourceTableName = "" Then
                MsgBox "Cancelled."
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

        MsgBox "Pivot datasource unification success."& _
                      "There are " _
                      & ActiveWorkbook.PivotCaches.Count _
                      & " pivot caches in the active workook."

      End Sub


      Sub UnifyAllPivotCaches()

        Dim wb As Workbook
        Dim ws As Worksheet
        Dim pt As PivotTable
        Set wb = ActiveWorkbook

        Dim answer As Integer
        answer = MsgBox("Unify all Pivot Caches?", vbYesNo + vbQuestion)

            If answer = vbYes Then

              For Each ws In ActiveWorkbook.Worksheets
                For Each pt In ws.PivotTables
                  pt.CacheIndex = wb.PivotTables(1).CacheIndex
                Next pt
              Next ws

              MsgBox "Pivot cache unification success. "& _
                      "There are " _
                      & ActiveWorkbook.PivotCaches.Count _
                      & " pivot caches in the active workook."
            Else
            ' Do nothing.

            End If

      End Sub
