Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.ShapeFormat = Me.Factory.CreateRibbonButton
        Me.WordNew = Me.Factory.CreateRibbonButton
        Me.AddCommandBar = Me.Factory.CreateRibbonButton
        Me.SaveAsTemplate = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.ShapeFormat)
        Me.Group1.Items.Add(Me.WordNew)
        Me.Group1.Items.Add(Me.AddCommandBar)
        Me.Group1.Items.Add(Me.SaveAsTemplate)
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'ShapeFormat
        '
        Me.ShapeFormat.Label = "ShapeFormat"
        Me.ShapeFormat.Name = "ShapeFormat"
        '
        'WordNew
        '
        Me.WordNew.Label = "WordNew"
        Me.WordNew.Name = "WordNew"
        '
        'AddCommandBar
        '
        Me.AddCommandBar.Label = "AddCommandBar"
        Me.AddCommandBar.Name = "AddCommandBar"
        '
        'SaveAsTemplate
        '
        Me.SaveAsTemplate.Label = "SaveAsTemplate"
        Me.SaveAsTemplate.Name = "SaveAsTemplate"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ShapeFormat As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents WordNew As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AddCommandBar As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SaveAsTemplate As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
