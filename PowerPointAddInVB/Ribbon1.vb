Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub InsertShape_Click(sender As Object, e As RibbonControlEventArgs) Handles InsertShape.Click

    End Sub
    Private Sub InsertAutoShape47()

        'Dim oActiveWindow As Object = Me.HostApplication.ActiveWindow()
        'Dim oPresentation As PowerPoint.Presentation = oActiveWindow.Presentation

        'Dim oSlides As PowerPoint.Slides = oPresentation.Slides
        'Dim oCustomLayout As PowerPoint.CustomLayout = oSlides(1).CustomLayout
        'Dim oSlidesNew As PowerPoint.Slides = oPresentation.Slides
        'Dim oSlide As PowerPoint.Slide = oSlidesNew.AddSlide(oSlides.Count + 1, oCustomLayout)
        'Dim oSlideShapes As PowerPoint.Shapes = oSlide.Shapes
        'Dim oShapenew As PowerPoint.Shape = oSlideShapes.AddShape(47, 250, 250, 200, 20)


        'GC.Collect()
        'GC.WaitForPendingFinalizers()
        'GC.Collect()
        'GC.WaitForPendingFinalizers()

    End Sub
    Dim sli As PowerPoint.Slide
    Dim s As PowerPoint.Shape
    Private Sub addShape_Click(sender As Object, e As RibbonControlEventArgs) Handles addShape.Click
        sli = Globals.ThisAddIn.Application.ActivePresentation.Slides(1)
        s = sli.Shapes.AddShape(47, 10, 10, 240, 60)
    End Sub

    Private Sub changeShape_Click(sender As Object, e As RibbonControlEventArgs) Handles changeShape.Click
        'With s.Adjustments
        '    .Item(1) += 0.1
        '    .Item(2) += 0.2
        '    .Item(3) += 0.3
        'End With
        s.Height = 480
        s.Width = 240

    End Sub

    Private Sub getShape_Click(sender As Object, e As RibbonControlEventArgs) Handles getShape.Click
        With s.Adjustments
            MsgBox("width: " & s.Width & " ;height: " & s.Height & " ;item(1): " & .Item(1) & "; item(2): " & .Item(2) & "; item(3): " & .Item(3) & "; item(4): " & .Item(4))
        End With
    End Sub

    Private Sub UpdateTable_Click(sender As Object, e As RibbonControlEventArgs) Handles UpdateTable.Click
        Globals.ThisAddIn.UpdateTablebak()
    End Sub
End Class
