<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGenerateCodeOptions
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGenerateCodeOptions))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtAbbrev = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.optFile = New System.Windows.Forms.RadioButton()
        Me.optClipboard = New System.Windows.Forms.RadioButton()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOkay = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(111, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Program Name:"
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(141, 25)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(201, 20)
        Me.txtName.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 69)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(111, 23)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Program Abbreviation:"
        '
        'txtAbbrev
        '
        Me.txtAbbrev.Location = New System.Drawing.Point(141, 69)
        Me.txtAbbrev.MaxLength = 2
        Me.txtAbbrev.Name = "txtAbbrev"
        Me.txtAbbrev.Size = New System.Drawing.Size(47, 20)
        Me.txtAbbrev.TabIndex = 3
        Me.txtAbbrev.Text = "XX"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 113)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(111, 23)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Output to:"
        '
        'optFile
        '
        Me.optFile.AutoSize = True
        Me.optFile.Checked = True
        Me.optFile.Location = New System.Drawing.Point(141, 113)
        Me.optFile.Name = "optFile"
        Me.optFile.Size = New System.Drawing.Size(41, 17)
        Me.optFile.TabIndex = 5
        Me.optFile.TabStop = True
        Me.optFile.Text = "&File"
        Me.optFile.UseVisualStyleBackColor = True
        '
        'optClipboard
        '
        Me.optClipboard.AutoSize = True
        Me.optClipboard.Location = New System.Drawing.Point(141, 136)
        Me.optClipboard.Name = "optClipboard"
        Me.optClipboard.Size = New System.Drawing.Size(69, 17)
        Me.optClipboard.TabIndex = 6
        Me.optClipboard.Text = "&Clipboard"
        Me.optClipboard.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.Location = New System.Drawing.Point(364, 113)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(63, 63)
        Me.cmdCancel.TabIndex = 8
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdOkay
        '
        Me.cmdOkay.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOkay.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOkay.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOkay.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOkay.Image = CType(resources.GetObject("cmdOkay.Image"), System.Drawing.Image)
        Me.cmdOkay.Location = New System.Drawing.Point(364, 25)
        Me.cmdOkay.Name = "cmdOkay"
        Me.cmdOkay.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOkay.Size = New System.Drawing.Size(63, 63)
        Me.cmdOkay.TabIndex = 7
        Me.cmdOkay.Text = "&Okay"
        Me.cmdOkay.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdOkay.UseVisualStyleBackColor = False
        '
        'frmGenerateCodeOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(439, 205)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOkay)
        Me.Controls.Add(Me.optClipboard)
        Me.Controls.Add(Me.optFile)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtAbbrev)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmGenerateCodeOptions"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Generate VB Code Options"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtAbbrev As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents optFile As System.Windows.Forms.RadioButton
    Friend WithEvents optClipboard As System.Windows.Forms.RadioButton
    Public WithEvents cmdCancel As System.Windows.Forms.Button
    Public WithEvents cmdOkay As System.Windows.Forms.Button
End Class
