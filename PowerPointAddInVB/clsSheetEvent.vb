Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop

Public Class clsSheetEvent

    Public WithEvents Worksheet As Worksheet

    Private Sub Worksheet_Change(ByVal Target As Excel.Range)
        MsgBox("t")
    End Sub

    Private Sub Worksheet_SelectionChange(ByVal Target As Range)
        MsgBox("tt")
    End Sub

    'Private Sub Class_Terminate()
    '    Worksheet = Nothing
    'End Sub

End Class
