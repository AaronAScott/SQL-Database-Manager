<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class About
	Inherits System.Windows.Forms.Form

	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()>
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
	<System.Diagnostics.DebuggerStepThrough()>
	Private Sub InitializeComponent()
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(About))
		Me.Panel1 = New System.Windows.Forms.Panel()
		Me.lblTitle = New System.Windows.Forms.Label()
		Me.lblVersion = New System.Windows.Forms.Label()
		Me.lblCopyright = New System.Windows.Forms.Label()
		Me.lblWSID = New System.Windows.Forms.Label()
		Me.lblUser = New System.Windows.Forms.Label()
		Me.lblExpiration = New System.Windows.Forms.Label()
		Me.lblLicensee = New System.Windows.Forms.Label()
		Me.btnNotes = New System.Windows.Forms.Button()
		Me.lblDBVersion = New System.Windows.Forms.Label()
		Me.btnClose = New System.Windows.Forms.Button()
		Me.SuspendLayout()
		'
		'Panel1
		'
		Me.Panel1.BackgroundImage = CType(resources.GetObject("Panel1.BackgroundImage"), System.Drawing.Image)
		Me.Panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
		Me.Panel1.Location = New System.Drawing.Point(6, 9)
		Me.Panel1.Name = "Panel1"
		Me.Panel1.Size = New System.Drawing.Size(272, 275)
		Me.Panel1.TabIndex = 0
		'
		'lblTitle
		'
		Me.lblTitle.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTitle.Location = New System.Drawing.Point(297, 9)
		Me.lblTitle.Name = "lblTitle"
		Me.lblTitle.Size = New System.Drawing.Size(257, 20)
		Me.lblTitle.TabIndex = 1
		Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'lblVersion
		'
		Me.lblVersion.Location = New System.Drawing.Point(297, 39)
		Me.lblVersion.Name = "lblVersion"
		Me.lblVersion.Size = New System.Drawing.Size(257, 20)
		Me.lblVersion.TabIndex = 2
		Me.lblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'lblCopyright
		'
		Me.lblCopyright.Location = New System.Drawing.Point(297, 69)
		Me.lblCopyright.Name = "lblCopyright"
		Me.lblCopyright.Size = New System.Drawing.Size(257, 20)
		Me.lblCopyright.TabIndex = 3
		Me.lblCopyright.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'lblWSID
		'
		Me.lblWSID.Location = New System.Drawing.Point(297, 130)
		Me.lblWSID.Name = "lblWSID"
		Me.lblWSID.Size = New System.Drawing.Size(257, 20)
		Me.lblWSID.TabIndex = 4
		Me.lblWSID.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'lblUser
		'
		Me.lblUser.Location = New System.Drawing.Point(297, 160)
		Me.lblUser.Name = "lblUser"
		Me.lblUser.Size = New System.Drawing.Size(257, 20)
		Me.lblUser.TabIndex = 5
		Me.lblUser.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'lblExpiration
		'
		Me.lblExpiration.Location = New System.Drawing.Point(297, 191)
		Me.lblExpiration.Name = "lblExpiration"
		Me.lblExpiration.Size = New System.Drawing.Size(257, 20)
		Me.lblExpiration.TabIndex = 6
		Me.lblExpiration.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'lblLicensee
		'
		Me.lblLicensee.Location = New System.Drawing.Point(297, 221)
		Me.lblLicensee.Name = "lblLicensee"
		Me.lblLicensee.Size = New System.Drawing.Size(257, 20)
		Me.lblLicensee.TabIndex = 7
		Me.lblLicensee.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'btnNotes
		'
		Me.btnNotes.Location = New System.Drawing.Point(297, 265)
		Me.btnNotes.Name = "btnNotes"
		Me.btnNotes.Size = New System.Drawing.Size(135, 20)
		Me.btnNotes.TabIndex = 8
		Me.btnNotes.Text = "View &Release Notes"
		Me.btnNotes.UseVisualStyleBackColor = True
		'
		'lblDBVersion
		'
		Me.lblDBVersion.Location = New System.Drawing.Point(296, 100)
		Me.lblDBVersion.Name = "lblDBVersion"
		Me.lblDBVersion.Size = New System.Drawing.Size(257, 20)
		Me.lblDBVersion.TabIndex = 9
		Me.lblDBVersion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'btnClose
		'
		Me.btnClose.Location = New System.Drawing.Point(453, 263)
		Me.btnClose.Name = "btnClose"
		Me.btnClose.Size = New System.Drawing.Size(107, 20)
		Me.btnClose.TabIndex = 10
		Me.btnClose.Text = "&Close"
		Me.btnClose.UseVisualStyleBackColor = True
		'
		'About
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(570, 300)
		Me.Controls.Add(Me.btnClose)
		Me.Controls.Add(Me.lblDBVersion)
		Me.Controls.Add(Me.btnNotes)
		Me.Controls.Add(Me.lblLicensee)
		Me.Controls.Add(Me.lblExpiration)
		Me.Controls.Add(Me.lblUser)
		Me.Controls.Add(Me.lblWSID)
		Me.Controls.Add(Me.lblCopyright)
		Me.Controls.Add(Me.lblVersion)
		Me.Controls.Add(Me.lblTitle)
		Me.Controls.Add(Me.Panel1)
		Me.Name = "About"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Text = "About"
		Me.ResumeLayout(False)

	End Sub

	Friend WithEvents Panel1 As Panel
	Friend WithEvents lblTitle As Label
	Friend WithEvents lblVersion As Label
	Friend WithEvents lblCopyright As Label
	Friend WithEvents lblWSID As Label
	Friend WithEvents lblUser As Label
	Friend WithEvents lblExpiration As Label
	Friend WithEvents lblLicensee As Label
	Friend WithEvents btnNotes As Button
	Friend WithEvents lblDBVersion As Label
	Friend WithEvents btnClose As Button
End Class
