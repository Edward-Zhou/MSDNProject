Imports System.Text
Imports DocumentFormat.OpenXml.Packaging
Imports Microsoft.Office.Interop
Imports System
Imports System.IO
Imports System.Xml
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Excel

Public Class Form1

    Private Sub CreateDoc_Click(sender As Object, e As EventArgs) Handles CreateDoc.Click
        Dim dc As New CreateDocument()
        dc.CreateDocFromTem()
        MessageBox.Show("ok")
    End Sub

    Private Sub CreateDocbak_Click(sender As Object, e As EventArgs) Handles CreateDocbak.Click
        Dim sFile As String = "D:\OfficeDev\OpenXML\Word\Hello World.dotx"
        If File.Exists(sFile.Replace("dotx", "docx")) Then
            File.Delete(sFile.Replace("dotx", "docx"))
        End If
        File.Copy(sFile, sFile.Replace("dotx", "docx"))
        Dim uniEncoding As New UnicodeEncoding()
        Dim fs As New FileStream(sFile, FileMode.Open, FileAccess.Read)
        Dim data As Byte() = File.ReadAllBytes(sFile.Replace("dotx", "docx"))
        Dim stream As New MemoryStream(data)

        ' Modify the document to ensure it is correctly marked as a Document and not Template
        Using document As WordprocessingDocument = WordprocessingDocument.Open(stream, True)
            document.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document)
        End Using

        ' Add the XML into the document and save to the correct location.               
        File.WriteAllBytes(sFile.Replace("dotx", "docx"), stream.ToArray())

        MessageBox.Show("ok")
    End Sub

    Private Sub ExcelMove_Click(sender As Object, e As EventArgs) Handles ExcelMove.Click
        Dim SourceExcelWorkbook As Excel.Workbook = Nothing
        Dim TargetExcelWorkbook As Excel.Workbook = Nothing
        Dim TargetExcelSheets As Excel.Sheets = Nothing
        Dim SourceExcelSheets As Excel.Sheets = Nothing
        Dim CopyWorkSheet As Excel.Worksheet = Nothing

        Dim XLApp As Excel.Application
        'XLApp = New Excel.Application
        XLApp = CreateObject("Excel.Application", "")
        XLApp.Visible = True
        XLApp.DisplayAlerts = True
        XLApp.ScreenUpdating = False

        Dim pobjExcelWorkbooks As Excel.Workbooks = XLApp.Workbooks

        SourceExcelWorkbook = pobjExcelWorkbooks.Open("D:\Book1.xlsm")
        TargetExcelWorkbook = pobjExcelWorkbooks.Open("D:\Book2.xlsm")

        TargetExcelSheets = TargetExcelWorkbook.Worksheets
        SourceExcelSheets = SourceExcelWorkbook.Worksheets

        Dim OriginalSheetCount As Integer = TargetExcelSheets.Count
        Dim SheetCount As Integer = OriginalSheetCount
        Dim SheetsToBeCopiedCount As Integer = SourceExcelSheets.Count

        While SheetsToBeCopiedCount > 1
            'Dim lobjAfterSheet As Object = TargetExcelSheets.Item(1)
            Dim lobjAfterSheet As Object = TargetExcelSheets.Item(SheetCount)
            CopyWorkSheet = SourceExcelSheets.Item(1)
            CopyWorkSheet.Move(After:=lobjAfterSheet)
            SheetCount = SheetCount + 1
            TargetExcelWorkbook.Save()
            SheetsToBeCopiedCount = SheetsToBeCopiedCount - 1
        End While
        SourceExcelWorkbook.Save()
        TargetExcelWorkbook.Save()
        SourceExcelWorkbook.Close()
        TargetExcelWorkbook.Close()
        pobjExcelWorkbooks.Close()
    End Sub

    Private Sub OutlookBtn_Click(sender As Object, e As EventArgs) Handles OutlookBtn.Click
        SendOutlookMail("Test", "v-tazho@hotmail.com", "Hello Word")
    End Sub

    Public Sub SendOutlookMail(Subject As String, Recipient As _
 String, Message As String)

        On Error GoTo errorHandler
        Dim oLapp As Outlook.Application
        Dim oItem As Object
        oLapp = CreateObject("Outlook.Application")
        oItem = oLapp.CreateItem(0)

        With oItem
            .Subject = Subject
            .To = Recipient
            .body = Message
            .Display()
        End With

        oLapp = Nothing
        oItem = Nothing


        Exit Sub

errorHandler:
        oLapp = Nothing
        oItem = Nothing
        Exit Sub


    End Sub

    Private Sub AccessObjectbtn_Click(sender As Object, e As EventArgs) Handles AccessObjectbtn.Click
        Dim a As Access.Application
        Dim b As Access.Application
        a = GetObject(, "Access.Application")


        a.Visible = True
        MessageBox.Show(a.Visible.ToString())
    End Sub

    Private Sub SetRTFBodybtn_Click(sender As Object, e As EventArgs) Handles SetRTFBodybtn.Click
        Dim oApp As Outlook.Application = New Outlook.Application
        Dim t As Outlook.AppointmentItem = DirectCast(oApp.CreateItem(Outlook.OlItemType.olAppointmentItem), Outlook.AppointmentItem)
        Dim sRTF As String
        t.Start = Convert.ToDateTime("11/05/2015")
        t.End = Convert.ToDateTime("11/05/2015")
        t.Subject = "VB.Net test"
        sRTF = "{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl" & _
       "{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}" & _
       "{\f2\froman\fprq2 Times New Roman;<AngularNoBind>}}</AngularNoBind>" & _
       "{\colortbl\red0\green0\blue0;\red255\green0\blue0;}" & _
       "\deflang1033\horzdoc{\*\fchars }{\*\lchars }" & _
       "\pard\plain\f2\fs24 Line 1 of \plain\f2\fs24\cf1" & _
       "inserted\plain\f2\fs24  file.\par }"
        t.RTFBody = System.Text.Encoding.ASCII.GetBytes(sRTF)
        't.RTFBody = System.Text.Encoding.ASCII.GetBytes("{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1033{\fonttbl{\f0\fswiss\fcharset0 Arial;}}{\*\generator Riched20 15.0.4599}{\*\mmathPr\mwrapIndent1440 }\viewkind4\uc1 \pard\f0\fs22 Test Body: First Line\parSecond Line of Text\par}")
        t.Display()
    End Sub
    Public db_path As String

    Dim oAccess As Access.Application
    Private Sub AccessReport_Click(sender As Object, e As EventArgs) Handles AccessReport.Click
        oAccess = CreateObject("Access.Application")
        oAccess.Visible = True
        db_path = "C:\Users\v-tazho\Documents\Test.accdb"
        oAccess.OpenCurrentDatabase(db_path)
        oAccess.DoCmd.OpenReport(ReportName:="cTime", View:=Access.AcView.acViewPreview)

    End Sub

    Private Sub ExcelApp_Click(sender As Object, e As EventArgs) Handles ExcelApp.Click
        Dim wkb As Workbook
        Dim app As Excel.Application
        app = CType(GetObject(, "Excel.Application"), Excel.Application)
        With Me.ComboBox1.Items
            For Each wkb In app.Workbooks
                .Add(wkb.Name())
            Next wkb
        End With
    End Sub

    Private Sub AddParagraph_Click(sender As Object, e As EventArgs) Handles AddParagraph.Click
        'Dim app As Word.Application
        'Dim doc As Word.Document
        'Dim p As Word.Paragraph
        'Dim rng As Word.Range
        'app = New Word.Application
        'app.Visible = True
        'doc = app.Documents.Open("D:\work\Palladium\Thread Practise\201603\20160325.docx")
        'p = doc.Paragraphs.Add
        'With p
        '    .Range.Text = "test"
        '    .Range.Font.Bold = 1
        '    .Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        'End With
        'rng = doc.Range(Start:=DirectCast(doc.Content.End - 1, Object), End:=DirectCast(doc.Content.End, Object))
        'rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        'rng.InsertFile(FileName:="C:\Users\v-tazho\Desktop\Test (2).docx".ToString)
    End Sub

    Private Sub ExcelGetObject_Click(sender As Object, e As EventArgs) Handles ExcelGetObject.Click
        Dim oExcel As Excel.Workbook = Nothing

        oExcel = GetObject("C:\Users\v-tazho\Desktop\Test.xlsx", "Excel.Application")

        oExcel.Application.Visible = True

        oExcel.Windows(1).Visible = True
        MsgBox("OK")

    End Sub

    Private Sub ExcelCreateApp_Click(sender As Object, e As EventArgs) Handles ExcelCreateApp.Click
        Dim eApp As Excel.Application
        eApp = New Excel.Application
        eApp.Workbooks.Add()
        eApp.Visible = True
    End Sub

    Private Sub EmbedExcel_Click(sender As Object, e As EventArgs) Handles EmbedExcel.Click
        Dim Exl As Object = CreateObject("Excel.Application")
        Dim tasks As New List(Of Task)
        ' Execute the task 10 times.
        tasks.Add(Task.Factory.StartNew(Sub()

                                            Dim Bok As Object
                                            Dim Shet As Object

                                            Exl.visible = True
                                            Exl.WindowState = -4143 'maximze window
                                            Bok = Exl.Workbooks.Open("C:\Users\v-tazho\Desktop\Test.xlsx")
                                            Shet = Exl.Workbooks(1).Sheets(1)
                                        End Sub
         ))

        Task.WaitAll(tasks.ToArray())
        SetParent(Exl.Hwnd, Me.Handle)
    End Sub
    Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
    
End Class
