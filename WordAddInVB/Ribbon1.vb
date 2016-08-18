Imports Microsoft.Office.Tools.Ribbon
Imports Word = Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub ShapeFormat_Click(sender As Object, e As RibbonControlEventArgs) Handles ShapeFormat.Click
        With Globals.ThisAddIn.Application.ActiveDocument.Shapes(1)
            .RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
            .RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin
        End With
    End Sub

    Private Sub WordNew_Click(sender As Object, e As RibbonControlEventArgs) Handles WordNew.Click
        Dim doc As Word.Application

    End Sub

    Private Sub AddCommandBar_Click(sender As Object, e As RibbonControlEventArgs) Handles AddCommandBar.Click
        Dim oCommandBars As CommandBar
        Dim oStandardBar As CommandBar
        Dim MyButton As CommandBarControl
        On Error Resume Next
        ' Set up a custom button on the "Standard" command bar.
        oCommandBars = Globals.ThisAddIn.Application.CommandBars
        If oCommandBars Is Nothing Then
            ' Outlook has the CommandBars collection on the Explorer object.
            oCommandBars = Globals.ThisAddIn.Application.CommandBars.ActiveExplorer.CommandBars
        End If

        oStandardBar = oCommandBars.Item("Standard")

        ' In case the button was not deleted, use the exiting one.
        MyButton = oStandardBar.Controls.Item("SaveDocument")
        If MyButton Is Nothing Then

            MyButton = oStandardBar.Controls.Add(1)
            With MyButton
                .Caption = "SaveDocument"
                .Style = MsoButtonStyle.msoButtonCaption
                .Tag = "SaveDocument"
                .OnAction = "!<MyCOMAddin.Connect>"

                .Visible = True
            End With
        End If

        ' Display a simple message to show which application you started in.
        'MsgBox("Started in " & applicationObject.Name & ".")

        oStandardBar = Nothing
        oCommandBars = Nothing

    End Sub

    Private Sub SaveAsTemplate_Click(sender As Object, e As RibbonControlEventArgs) Handles SaveAsTemplate.Click
        Dim FileName As Object = "C:\Users\v-tazho\Desktop\myfile.dotx"
        Dim FileFormat As Object = Word.WdSaveFormat.wdFormatXMLTemplate
        Dim LockComments As Object = False
        Dim AddToRecentFiles As Object = True
        Dim ReadOnlyRecommended As Object = False
        Dim EmbedTrueTypeFonts As Object = False
        Dim SaveNativePictureFormat As Object = True
        Dim SaveFormsData As Object = True
        Dim SaveAsAOCELetter As Object = False
        Dim Encoding As Object = MsoEncoding.msoEncodingUSASCII
        Dim InsertLineBreaks As Object = False
        Dim AllowSubstitutions As Object = False
        Dim LineEnding As Object = Word.WdLineEndingType.wdCRLF
        Dim AddBiDiMarks As Object = False
        Dim wdCompatibilityMode As Object = 15
        Dim missing As Object = Type.Missing
        'var a = Globals.ThisDocument.Application.ActiveDocument.Content.Text;
        'Globals.ThisDocument.Application.ActiveDocument.SaveAs2(ref FileName, ref FileFormat, ref LockComments,
        'ref missing, ref AddToRecentFiles, ref missing,
        'ref ReadOnlyRecommended, ref EmbedTrueTypeFonts,
        'ref SaveNativePictureFormat, ref SaveFormsData,
        'ref SaveAsAOCELetter, ref missing, ref missing,
        'ref missing, ref missing, ref missing, ref wdCompatibilityMode);
        Globals.ThisAddIn.Application.ActiveDocument.SaveAs(FileName, FileFormat, LockComments, missing, AddToRecentFiles, missing, _
            ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, missing, _
            missing, missing, missing, missing)
        MsgBox("ok")

    End Sub
End Class
