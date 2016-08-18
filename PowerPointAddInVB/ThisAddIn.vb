Imports Microsoft.Office.Interop

Public Class ThisAddIn
    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub
    Private Sub app_WindowSelectionChange(Sel As PowerPoint.Selection) Handles Application.WindowSelectionChange
        Dim shpRange As PowerPoint.ShapeRange = Nothing
        Dim eWorkbook As Excel.Workbook = Nothing
        Dim eWorksheet As Excel.Worksheet = Nothing
        Dim pChart As PowerPoint.Chart

        If Not Sel.Type = 0 Then
            shpRange = Sel.ShapeRange
        End If
        Try
            If Not (shpRange Is Nothing) Then
                For Each shp As PowerPoint.Shape In shpRange
                    If shp.HasChart = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        pChart = shp.Chart

                        Dim pChartData As PowerPoint.ChartData = pChart.ChartData
                        If Not (pChartData Is Nothing) Then
                            'pChartData.Activate()
                            eWorkbook = pChartData.Workbook
                            eWorksheet = eWorkbook.Worksheets(1)
                            RemoveHandler eWorksheet.Change, AddressOf Worksheet_Change
                            AddHandler eWorksheet.Change, AddressOf Worksheet_Change
                            Exit For

                        End If

                    End If
                Next
            End If
        Catch
        End Try
    End Sub

    Private Sub Worksheet_Change(ByVal Target As Excel.Range)
        MsgBox("t")
    End Sub


End Class
