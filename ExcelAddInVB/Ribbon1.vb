'Option Strict On
Imports Microsoft.Office.Tools.Ribbon
Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub ExcelTemplate_Click(sender As Object, e As RibbonControlEventArgs) Handles ExcelTemplate.Click
        Dim a As Microsoft.Office.Tools.Excel.Worksheet
        a = Globals.ThisAddIn.Application.Worksheets.Add(Type:="C:\Users\v-tazho\Desktop\Test.xlsx")
    End Sub

    Private Sub CopyPowerPivot_Click(sender As Object, e As RibbonControlEventArgs) Handles CopyPowerPivot.Click
        Dim ow As Microsoft.Office.Interop.Excel.Workbooks
        Dim aw As Microsoft.Office.Tools.Excel.Worksheet
        Dim excelapp As Microsoft.Office.Interop.Excel.Application

        excelapp = New Microsoft.Office.Interop.Excel.Application
        excelapp.Visible = True
        ow = DirectCast(excelapp.Workbooks.Open("C:\Users\v-tazho\Desktop\Test.xlsx"), Microsoft.Office.Interop.Excel.Workbook)
        Globals.ThisAddIn.Application.Worksheets.Copy(After:=ow.Application.Worksheets("Sheet3"))
    End Sub

    Private Sub CTypeChangebtn_Click(sender As Object, e As RibbonControlEventArgs) Handles CTypeChangebtn.Click
        Dim app As Microsoft.Office.Interop.Excel.Application
        Dim range As Microsoft.Office.Interop.Excel.Range
        Dim t As Object
        app = Globals.ThisAddIn.Application
        range = app.Columns
        t = range(("A:A"))
        't = TypeName(app.Columns("A:A"))
        'range.Select()
    End Sub

    Private Sub ShowMsg_Click(sender As Object, e As RibbonControlEventArgs) Handles ShowMsg.Click
        Dim xlMain As New NativeWindow()

        Dim frmCB As New frmCmpBrowse

        'gXLApp = CType(Globals.ThisAddIn.Application, Excel.Application)

        frmCB.ShowDialog(xlMain)

        frmCB.Dispose()

        xlMain.ReleaseHandle()

        'gXLApp = Nothing

    End Sub

    Private chart As Excel.Chart
    Private Sub AddDataLabel_Click(sender As Object, e As RibbonControlEventArgs)
    End Sub

    Private Sub DataLabelPosition_Click(sender As Object, e As RibbonControlEventArgs)

    End Sub



    Private Sub AddDataLabel_Click_1(sender As Object, e As RibbonControlEventArgs) Handles AddDataLabel.Click
        Dim worksheet As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        ' Add chart.
        Dim charts = TryCast(worksheet.ChartObjects(), Microsoft.Office.Interop.Excel.ChartObjects)
        Dim chartObject = TryCast(charts.Add(60, 10, 300, 300), Microsoft.Office.Interop.Excel.ChartObject)
        chart = chartObject.Chart

        ' Set chart range.
        Dim range = worksheet.Range("A1", "B3")
        chart.SetSourceData(range)

        ' Set chart properties.
        chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterSmooth
        chart.ChartWizard(Source:=range, Title:="graphTitle", CategoryTitle:="xAxis", ValueTitle:="yAxis")
        chart.SetElement(MsoChartElementType.msoElementDataLabelTop)

    End Sub

    Private Sub DataLabelPosition_Click_1(sender As Object, e As RibbonControlEventArgs) Handles DataLabelPosition.Click
        Dim series As Excel.Series = chart.FullSeriesCollection(1)
        Dim db As Excel.DataLabels = series.DataLabels()
        'db.Select()
        'db.Position = Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionBelow
        For Each dl As Excel.DataLabel In db
            dl.Top = dl.Top + 100
        Next
    End Sub
End Class
