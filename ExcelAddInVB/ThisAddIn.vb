Imports Microsoft.Office.Interop.Excel
Public Class ThisAddIn

    Dim instance As WorkbookEvents_Event
    Dim handler As WorkbookEvents_BeforeSaveEventHandler


    Sub ThisWorkbook_BeforeSave(ByVal workbook As Excel.Workbook, ByVal SaveAsUI As Boolean, _
        ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        Console.Write(SaveAsUI)
        Dim wsS As Microsoft.Office.Interop.Excel.Worksheet

    End Sub


    Private Sub ThisAddIn_Startup() Handles Me.Startup
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub






End Class
