Option Strict Off
Option Explicit On
Friend Class Logo
	Inherits System.Windows.Forms.Form
	'**********************************************************
	' Logo for Visual Basic Programs
    ' LOGO.VB
	' Written: October 1993
	' Programmer: Aaron Scott
	' Copyright (C) 1993 Sirius Software All Rights Reserved
	'**********************************************************
	
    '**********************************************************
	'
	' The form is loaded.
	'
	'**********************************************************
	Private Sub Logo_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        lblTitle.Text = ProgramName & " for Windows"
		lblVersion.Text = "Version " & Version
		lblLicense.Text = LicenseInfo
		Application.DoEvents()
	End Sub
	

	
	'**********************************************************
	'
	' When the timer "clicks", unload the logo.
	'
	'**********************************************************
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        Me.Hide()
	End Sub
End Class