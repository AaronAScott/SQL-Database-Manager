Option Strict Off
Option Explicit On
Imports System.Reflection.Emit
Imports System.Net.Http
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel


Friend Class About
	Inherits System.Windows.Forms.Form
	'**********************************************************
	' About Box for Visual Basic Programs 
	' ABOUT.VB
	' Written: February 2026
	' Copyright (C) 2026 Sirius Software All Rights Reserved

	' Required modules: none
	'**********************************************************

	Private Shared VersionNotes
	Private Shared ReadOnly Client As New HttpClient()


	'**********************************************************
	'
	' The form is loaded. Get the system resources to display
	'
	'**********************************************************
	Private Sub About_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

		' Declare variables

		Dim zx0 As String

		Me.Text = "About " & ProgramName
		lblTitle.Text = ProgramName & " for Windows"
		If My.Application.Info.Copyright <> "" Then lblCopyright.Text = My.Application.Info.Copyright
		zx0 = CStr(FileDateTime(Application.ExecutablePath))
		lblVersion.Text = "Version " & CStr(My.Application.Info.Version.Major) & "." & CStr(My.Application.Info.Version.Minor) & CStr(My.Application.Info.Version.Build) & CStr(My.Application.Info.Version.MinorRevision) & " (" & zx0 & ")"
		If WSID() <> "" Then lblWSID.Text = "Workstation ID: " & WSID()
		If UserName <> "" Then lblUser.Text = "User: " & UserName
		If DBVersion = "" Then lblDBVersion.Text = "Database Version (none)" Else lblDBVersion.Text = "Database Version " & DBVersion
		lblLicensee.Text = LicenseInfo
		Me.Icon = frmMain.Icon

		VersionNotes = GetCurrentVersionInfo()

		' Disable the version notes button if there are none found.

		If VersionNotes = "" Then btnNotes.Enabled = False Else btnNotes.Enabled = True

		System.Windows.Forms.Application.DoEvents() ' give things a chance to display
	End Sub
	'**************************************************
	'
	' The okay button is clicked.  Unload the form.
	'
	'**************************************************
	Private Sub btnClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnClose.Click
		Me.Close()
	End Sub

	'**************************************************

	' The View Release Notes button is clicked.

	'**************************************************
	Private Sub btnNotes_Click(sender As Object, e As EventArgs) Handles btnNotes.Click

		' Declare variables

		Dim fNotes As Form

		' Retrieve CurrentVersion.txt from the download website.

		' Create a form for viewing the release notes.

		fNotes = New Form
		With fNotes
			.AutoScaleDimensions = New SizeF(6.0!, 13.0!)
			.AutoScaleMode = AutoScaleMode.Font
			.ClientSize = New Size(600, 400)
			.FormBorderStyle = FormBorderStyle.FixedDialog
			.ControlBox = True
			.MaximizeBox = False
			.MinimizeBox = False
			.Name = "frmReleaseNotes"
			.ShowIcon = False
			.ShowInTaskbar = False
			.StartPosition = FormStartPosition.CenterScreen
			.Text = ProgramName & " Release Notes"
		End With

		' Add the text box.

		Dim t As New TextBox
		t.Multiline = True
		t.Dock = DockStyle.Fill
		t.Location = New Point(0, 0)
		t.ScrollBars = ScrollBars.Vertical
		t.ReadOnly = True
		t.Font = New Font("Arial", 10)
		t.Text = VersionNotes
		t.SelectionStart = t.Text.Length
		fNotes.Controls.Add(t)

		' Show the notes.

		fNotes.Show()

	End Sub
	'**********************************************************

	' Function to retrieve the contents of the CurrentVersion.txt
	' file from the update website.

	'**********************************************************

	Private Function GetCurrentVersionInfo() As String

		' Declare variables.

		Dim zx As String = ""

		' Retrieve the CurrentVersion.txt file contents.

		Try
			zx = Client.GetStringAsync(ProgramUpdateURL & "CurrentVersion.txt").GetAwaiter().GetResult()
		Catch ex As Exception
		End Try

		Return zx
	End Function
End Class