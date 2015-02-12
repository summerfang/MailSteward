<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DisplayFoldersBar
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DisplayFoldersBar))
        Me.imgNavigator = New System.Windows.Forms.ImageList(Me.components)
        Me.SuspendLayout()
        '
        'imgNavigator
        '
        Me.imgNavigator.ImageStream = CType(resources.GetObject("imgNavigator.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgNavigator.TransparentColor = System.Drawing.Color.Transparent
        Me.imgNavigator.Images.SetKeyName(0, "PreviousFolder.bmp")
        Me.imgNavigator.Images.SetKeyName(1, "NextFolder.bmp")
        '
        'DisplayFoldersBar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Name = "DisplayFoldersBar"
        Me.Size = New System.Drawing.Size(596, 23)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents imgNavigator As System.Windows.Forms.ImageList

End Class
