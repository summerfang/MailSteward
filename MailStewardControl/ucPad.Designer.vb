<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucPad
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
        Me.txtCommandLine = New System.Windows.Forms.TextBox
        Me.dfbFolder = New MailStewardControl.DisplayFoldersBar
        Me.SuspendLayout()
        '
        'txtCommandLine
        '
        Me.txtCommandLine.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCommandLine.BackColor = System.Drawing.SystemColors.Control
        Me.txtCommandLine.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtCommandLine.Location = New System.Drawing.Point(4, 3)
        Me.txtCommandLine.Name = "txtCommandLine"
        Me.txtCommandLine.Size = New System.Drawing.Size(643, 13)
        Me.txtCommandLine.TabIndex = 0
        '
        'dfbFolder
        '
        Me.dfbFolder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dfbFolder.Location = New System.Drawing.Point(4, 22)
        Me.dfbFolder.Name = "dfbFolder"
        Me.dfbFolder.Size = New System.Drawing.Size(643, 23)
        Me.dfbFolder.TabIndex = 1
        '
        'ucPad
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.dfbFolder)
        Me.Controls.Add(Me.txtCommandLine)
        Me.Name = "ucPad"
        Me.Size = New System.Drawing.Size(651, 50)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtCommandLine As System.Windows.Forms.TextBox
    Friend WithEvents dfbFolder As MailStewardControl.DisplayFoldersBar

End Class
