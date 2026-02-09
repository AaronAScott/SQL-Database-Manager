<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Logo
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents lblTitle As System.Windows.Forms.Label
    Public WithEvents lblWarning As System.Windows.Forms.Label
	Public WithEvents lblVersion As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
		Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
		Me.lblTitle = New System.Windows.Forms.Label()
		Me.lblWarning = New System.Windows.Forms.Label()
		Me.lblVersion = New System.Windows.Forms.Label()
		Me.lblLicense = New System.Windows.Forms.Label()
		Me.SuspendLayout()
		'
		'Timer1
		'
		Me.Timer1.Enabled = True
		Me.Timer1.Interval = 4000
		'
		'lblTitle
		'
		Me.lblTitle.BackColor = System.Drawing.Color.Transparent
		Me.lblTitle.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTitle.Font = New System.Drawing.Font("Arial", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTitle.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
		Me.lblTitle.Location = New System.Drawing.Point(277, 152)
		Me.lblTitle.Name = "lblTitle"
		Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTitle.Size = New System.Drawing.Size(241, 105)
		Me.lblTitle.TabIndex = 3
		Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.TopCenter
		'
		'lblWarning
		'
		Me.lblWarning.BackColor = System.Drawing.Color.Transparent
		Me.lblWarning.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblWarning.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblWarning.ForeColor = System.Drawing.Color.White
		Me.lblWarning.Location = New System.Drawing.Point(12, 298)
		Me.lblWarning.Name = "lblWarning"
		Me.lblWarning.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblWarning.Size = New System.Drawing.Size(521, 13)
		Me.lblWarning.TabIndex = 1
		Me.lblWarning.Text = " Warning: This product is protected by U.S. and International Copyrights."
		'
		'lblVersion
		'
		Me.lblVersion.BackColor = System.Drawing.Color.Transparent
		Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblVersion.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblVersion.ForeColor = System.Drawing.Color.White
		Me.lblVersion.Location = New System.Drawing.Point(30, 152)
		Me.lblVersion.Name = "lblVersion"
		Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblVersion.Size = New System.Drawing.Size(241, 33)
		Me.lblVersion.TabIndex = 0
		Me.lblVersion.TextAlign = System.Drawing.ContentAlignment.TopCenter
		'
		'lblLicense
		'
		Me.lblLicense.BackColor = System.Drawing.Color.Transparent
		Me.lblLicense.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLicense.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLicense.ForeColor = System.Drawing.Color.White
		Me.lblLicense.Location = New System.Drawing.Point(30, 224)
		Me.lblLicense.Name = "lblLicense"
		Me.lblLicense.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLicense.Size = New System.Drawing.Size(241, 33)
		Me.lblLicense.TabIndex = 4
		Me.lblLicense.TextAlign = System.Drawing.ContentAlignment.TopCenter
		'
		'Logo
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.BackgroundImage = My.Resources.Resources.SiriusLogo
		Me.ClientSize = New System.Drawing.Size(550, 320)
		Me.Controls.Add(Me.lblLicense)
		Me.Controls.Add(Me.lblTitle)
		Me.Controls.Add(Me.lblWarning)
		Me.Controls.Add(Me.lblVersion)
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ForeColor = System.Drawing.SystemColors.WindowText
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Location = New System.Drawing.Point(120, 98)
		Me.Name = "Logo"
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.ResumeLayout(False)

	End Sub
	Public WithEvents lblLicense As System.Windows.Forms.Label
#End Region 
End Class