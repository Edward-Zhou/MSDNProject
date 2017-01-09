Imports Microsoft.Office.Interop
Imports System.Timers

Public Class ThisAddIn
    Private _wn As PowerPoint.SlideShowWindow
    Private WithEvents _tmrScroller As System.Timers.Timer
    Private _counter As Integer
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        _tmrScroller = New System.Timers.Timer(3000)
        _tmrScroller.AutoReset = False
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub
    'Private Sub app_WindowSelectionChange(Sel As PowerPoint.Selection) Handles Application.WindowSelectionChange
    '    Dim shpRange As PowerPoint.ShapeRange = Nothing
    '    Dim eWorkbook As Excel.Workbook = Nothing
    '    Dim eWorksheet As Excel.Worksheet = Nothing
    '    Dim pChart As PowerPoint.Chart

    '    If Not Sel.Type = 0 Then
    '        shpRange = Sel.ShapeRange
    '    End If
    '    Try
    '        If Not (shpRange Is Nothing) Then
    '            For Each shp As PowerPoint.Shape In shpRange
    '                If shp.HasChart = Microsoft.Office.Core.MsoTriState.msoTrue Then
    '                    pChart = shp.Chart

    '                    Dim pChartData As PowerPoint.ChartData = pChart.ChartData
    '                    If Not (pChartData Is Nothing) Then
    '                        'pChartData.Activate()
    '                        eWorkbook = pChartData.Workbook
    '                        eWorksheet = eWorkbook.Worksheets(1)
    '                        RemoveHandler eWorksheet.Change, AddressOf Worksheet_Change
    '                        AddHandler eWorksheet.Change, AddressOf Worksheet_Change
    '                        Exit For

    '                    End If

    '                End If
    '            Next
    '        End If
    '    Catch
    '    End Try
    'End Sub

    'Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    '    MsgBox("t")
    'End Sub
    Private Sub UpdateTable(sh As PowerPoint.Shape)
        _counter = _counter + 1
        Dim tablerowcount As Integer = sh.Table.Rows.Count
        Dim tablecolcount As Integer = sh.Table.Columns.Count
        For x = 1 To tablerowcount
            For y = 1 To tablecolcount
                sh.Table.Cell(x, y).Shape.TextFrame.TextRange.Text = _counter & "/" & Now.Millisecond
            Next
        Next
    End Sub
    Public Sub UpdateTablebak()
        For Each _slide As PowerPoint.Slide In Globals.ThisAddIn.Application.ActivePresentation.Slides
            For Each _shape As PowerPoint.Shape In _slide.Shapes
                If _shape.Type = Microsoft.Office.Core.MsoShapeType.msoTable Then
                    _counter = _counter + 1
                    Dim tablerowcount As Integer = _shape.Table.Rows.Count
                    Dim tablecolcount As Integer = _shape.Table.Columns.Count
                    For x = 1 To tablerowcount
                        For y = 1 To tablecolcount
                            _shape.Table.Cell(x, y).Shape.TextFrame.TextRange.Text = _counter & "/" & Now.Millisecond
                        Next
                    Next

                End If
            Next
        Next

    End Sub

    Private Sub Application_SlideShowBegin(Wn As PowerPoint.SlideShowWindow) Handles Application.SlideShowBegin
        _wn = Wn
        For Each _slide As PowerPoint.Slide In Wn.Presentation.Slides
            For Each _shape As PowerPoint.Shape In _slide.Shapes
                If _shape.Type = Microsoft.Office.Core.MsoShapeType.msoTable Then
                    ' UpdateTable(_shape)
                End If
            Next
        Next
        _tmrScroller.Start()
    End Sub
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
      (ByVal lpClassName As String, _
       ByVal lpWindowName As Long) As Long
    Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function LockWindowUpdate Lib "user32" _
      (ByVal hwndLock As Long) As Long
    Private Sub _tmrScroller_Elapsed(sender As Object, e As ElapsedEventArgs) Handles _tmrScroller.Elapsed
        Static hwnd As Long
        Dim VersionNo As String
        VersionNo = Left(Globals.ThisAddIn.Application.Version, _
                    InStr(1, Globals.ThisAddIn.Application.Version, ".") - 1)
        hwnd = FindWindow("PPTFrameClass", 0&)
        'LockWindowUpdate(0&)

        'hwnd = 0
        UpdateTablebak()
        UpdateWindow(hwnd)
        'Try
        '_tmrScroller.Stop()
        'UpdateTablebak()
        '__tmrScroller.Start()
        'For Each _slide As PowerPoint.Slide In _wn.Presentation.Slides
        '    For Each _shape As PowerPoint.Shape In _slide.Shapes
        '        If _shape.Type = Microsoft.Office.Core.MsoShapeType.msoTable Then
        '            UpdateTablebak(_shape)
        '        End If
        '    Next
        'Next

        'Catch ex As Exception

        'Finally
        '    If _tmrScroller IsNot Nothing Then
        '        _tmrScroller.Start()
        '    End If
        'End Try
    End Sub


End Class


