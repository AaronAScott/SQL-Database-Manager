Option Strict Off
Option Explicit On
Friend Class About
	Inherits System.Windows.Forms.Form
	'**********************************************************
	' About Box for Visual Basic Programs
    ' ABOUT.VB
	' Written: October 1997
    '  Updated December 2001
    '  Updated November 2013
    '  Updated January 2017
	' Programmer: Aaron Scott
    ' Copyright (C) 1996-2017 Sirius Software All Rights Reserved
	' All Rights Reserved
	
	' Required modules: none
	'**********************************************************
	

	'**************************************************
	'
	' The okay button is clicked.  Unload the form.
	'
	'**************************************************
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Me.Close()
	End Sub
	
	'**********************************************************
	'
	' The form is loaded. Get the system resources to display
	'
	'**********************************************************
    Private Sub About_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        ' Declare variables

        Dim zx0 As String

        Me.Text = "About " & ProgramName
        Label2_0.Text = ProgramName
        Label2_1.Text = ProgramName
        lblTitle.Text = ProgramName & " for Windows"
        If My.Application.Info.Copyright <> "" Then lblCopyright.Text = My.Application.Info.Copyright
        On Error Resume Next
        zx0 = CStr(FileDateTime(ProgramPath & My.Application.Info.AssemblyName & ".exe"))
        On Error GoTo 0
		lblVersion.Text = "Version " & CStr(My.Application.Info.Version.Major) & "." & CStr(My.Application.Info.Version.Minor) & CStr(My.Application.Info.Version.Build) & CStr(My.Application.Info.Version.MinorRevision) & " (" & zx0 & ")"
        If WSID() <> "" Then lblWSID.Text = "Workstation ID: " & WSID()
        lblUser.Text = "User: " & My.User.Name
        lblRegistration.Text = ""
        lblLicensee.Text = ""

        System.Windows.Forms.Application.DoEvents() ' give things a chance to display
    End Sub
	
End Class