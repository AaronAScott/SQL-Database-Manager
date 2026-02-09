Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Data.SqlClient
Friend Class frmMain
	Inherits System.Windows.Forms.Form
	'*******************************************************************
	' SQL Data Manager Main Module (frmMain)
	' DM_MAIN.VB
	' Written: October 2017
	' Programmer: Aaron Scott
	' Copyright (C) 2017-2021 Sirius Software All Rights Reserved
	'*******************************************************************

	'***********************************************************
	'
	' The form is loaded
	'
	'***********************************************************
	Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

		' Declare variables

		Dim zx As String

		' Set the Public variables for the name of the system,
		' the name of the database and the version.  This variant
		' of SQL Data Manager began with A/P database version 1.6.

		ProgramName = "SQL Database Manager"
		Version = CStr(My.Application.Info.Version.Major) & "." & CStr(My.Application.Info.Version.Minor) & CStr(My.Application.Info.Version.Build) & CStr(My.Application.Info.Version.MinorRevision)

		' Check for updates.

		CheckForUpdates()

		' Get the path to the program.

		ProgramPath = AddDirSeparator(My.Application.Info.DirectoryPath)

		' Set the background colors based on the selected theme.

		frmTheme.SetBackgroundColor(Me)

		' Show the logo while preparing system

		Logo.Show()
		System.Windows.Forms.Application.DoEvents()

		' Make the main form visible immediately

		Me.Show()

		' Set the position of the main form and its size

		zx = GetSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MainWindow", "Size")
		If zx <> "" Then
			Me.Top = Val(ParseString(zx))
			Me.Left = Val(ParseString(zx))
			Me.Height = Val(ParseString(zx))
			Me.Width = Val(zx)
		End If

		'Get the server name.  The default is the old 2012 version of SQL Server Express.

		ServerName = "(LocalDB)\" & GetSetting("SiriusSoftwareGlobal", "SQLServer", "InstanceName", "v11.0")

		' Allow the MDI child forms to resize, now.  This would have been
		' turned off during the resizing of this form on startup.

		DontAllowResize = False

		' Fill the MRU list

		FillMRU()

		' If a database name is passed on the command line, make sure it
		' hasn't been supplied with quotation marks.

		If Command() <> "" Then Databasename = SRep(Trim(Command), 1, """", "")

		' If the database name was supplied, open it now

		If Databasename <> "" Then DbOpen = OpenADatabase(Databasename, False)


	End Sub

	'***********************************************************
	'
	' The form is is about to be unloaded
	'
	'***********************************************************
	Private Sub frmMain_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

		' Declare variables

		Dim xx As Integer

		' Make sure the user wants to terminate the program

		System.Windows.Forms.Application.DoEvents()
		xx = MsgBox("This will terminate your SQL Database Manager session.  Are you sure you want to do this?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.MsgBoxSetForeground, "Exit SQL Database Manager")
		If xx = MsgBoxResult.No Then eventArgs.Cancel = True

	End Sub

	'***********************************************************
	'
	' The form is unloaded
	'
	'***********************************************************
	Private Sub frmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

		' Declare variables

		Dim jj As Integer
		Dim zx As String

		' Save the window size and position

		If Me.Left > 0 And Me.Top > 0 Then
			zx = Me.Top & "," & Me.Left & "," & Me.Height & "," & Me.Width
			SaveSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MainWindow", "Size", zx)
		End If

		' Close the database if it is open

		If DbOpen Then CloseDatabase()

		' Unload all forms

		For jj = My.Application.OpenForms.Count - 1 To 0 Step -1
			If Not My.Application.OpenForms.Item(jj) Is Me Then My.Application.OpenForms(jj).Close()
		Next jj

	End Sub
	'***********************************************************

	' Something is dropped on the form.

	'***********************************************************
	Private Sub frmMain_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop

		' Declare variables

		Dim xx() As String

		' Geth the list of files dropped.

		xx = e.Data.GetData(System.Windows.Forms.DataFormats.FileDrop)

		' If the first/only item is a file name, open it.

		Databasename = xx(0)
		DbOpen = OpenADatabase(Databasename, False)


	End Sub
	'***********************************************************
	'
	' An object is dragged over this form.  See if it's a valid object:
	' it must be a file name.
	'
	'***********************************************************
	Private Sub Main_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter

		If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
			e.Effect = DragDropEffects.Copy
		Else
			e.Effect = DragDropEffects.None
		End If

	End Sub

	'***********************************************************

	' The form is moved.

	'***********************************************************
	Private Sub frmMain_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged

		' Set the flag that will keep the MDI child forms from being automatically
		' moved or resized by this form.

		DontAllowResize = True
	End Sub

	'***********************************************************

	' The form is beginning to be resized.

	'***********************************************************
	Private Sub frmMain_ResizeBegin(sender As Object, e As EventArgs) Handles Me.ResizeBegin

		' Set the flag that will keep the MDI child forms from being automatically
		' moved or resized by this form.

		DontAllowResize = True
	End Sub
	'*******************************************************************

	' The main window has finished resizing.

	'*******************************************************************
	Private Sub frmMain_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd

		' Clear the flag that will keep the MDI child forms from being automatically
		' moved or resized by this form.

		DontAllowResize = False
	End Sub
	'***********************************************************
	'
	' One of the Most Recently Used database menu
	' items is clicked.
	'
	'***********************************************************
	Public Sub MRUToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MRU1.Click, MRU2.Click, MRU3.Click, MRU4.Click, MRU5.Click, MRU6.Click, MRU7.Click, MRU8.Click

		' Declare variables

		Dim zx As String = sender.ToString

		' Get the name of the selected database and open it.

		Databasename = Mid(zx, 4)
		DbOpen = OpenADatabase(Databasename, False)

	End Sub

	'*******************************************************************
	'
	' The Open database menu option is clicked.
	'
	'*******************************************************************
	Public Sub OpenDatabaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpenDatabase.Click
		LaunchProcess("OpenDatabase")
	End Sub
	'*******************************************************************

	' The Organize Database History menu option is selected.

	'*******************************************************************
	Public Sub OrganizeDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOrganizeDB.Click
		frmOrganizeMRU.ShowDialog()
	End Sub
	'*******************************************************************

	' The Change Color Theme menu option is selected.

	'*******************************************************************
	Public Sub ColorTheme_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTheme.Click
		frmTheme.ShowDialog()
	End Sub
	'*******************************************************************
	'
	' Sub to open a database
	'
	'*******************************************************************
	Private Sub OpenDB()

		' Declare variables

		Dim zx0 As String

		' Get the current directory (the common dialog box may
		' change it).

		zx0 = FileSystem.CurDir

		' Use the common dialog open box to get the file name.

		Me.OpenFileDialog1.FileName = ""
		Me.OpenFileDialog1.Filter = "SQL Databases (*.MDF)|*.MDF|All Files (*.*)|*.*"
		Me.OpenFileDialog1.FilterIndex = 1
		Me.OpenFileDialog1.DefaultExt = "MDF"
		Me.OpenFileDialog1.Title = "Open SQL Database"
		If Me.OpenFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
			Databasename = Me.OpenFileDialog1.FileName

			' Restore the current directory

			FileSystem.ChDir(zx0)

			' Open the specified database

			DbOpen = OpenADatabase(Databasename, False)

		End If
	End Sub

	'*******************************************************************
	'
	' The close database menu option is clicked.
	'
	'*******************************************************************
	Public Sub CloseDatabaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCloseDatabase.Click
		If DbOpen Then CloseDatabase()
	End Sub
	'*******************************************************************
	'
	' The "new database" menu option is selected.
	'
	'*******************************************************************
	Public Sub mnuCompactDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCompactDB.Click
		LaunchProcess("compactdb")
	End Sub
	'*******************************************************************
	'
	' The "compact database" menu option is selected.
	'
	'*******************************************************************
	Public Sub NewDatabaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewDatabase.Click
		LaunchProcess("NewDatabase")
	End Sub
	'***********************************************************
	'
	' The Exit menu option is selected.
	'
	'***********************************************************
	Public Sub ExitToolStripMenuItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
		LaunchProcess("Shutdown")
	End Sub
	'*******************************************************************
	'
	' The About SQL Data Manager menu item is clicked.
	'
	'*******************************************************************
	Private Sub About_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAbout.Click
		About.Picture1.Image = Me.Icon.ToBitmap
		About.ShowDialog()
	End Sub

	'*******************************************************************

	' When one or more Most Recently Used items are added,
	' this event will cause the divider to appear.

	'*******************************************************************
	Private Sub MRU1_VisibleCh5anged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MRU1.VisibleChanged
		MRUSeparator.Visible = MRU1.Visible
	End Sub
	'*******************************************************************

	' The user has selected the table designer.

	'*******************************************************************
	Private Sub mnuTableDesigner_Click(sender As Object, e As EventArgs) Handles mnuTableDesigner.Click

		' Display the script builder

		frmTableDesigner.MdiParent = Me
		frmTableDesigner.Show()
		LaunchProcess("ArrangeWindows")

	End Sub

	'*******************************************************************

	' The user has selected the script editor.

	'*******************************************************************
	Private Sub mnuScriptEditor_Click(sender As Object, e As EventArgs) Handles mnuScriptEditor.Click

		' Display the script editor

		Dim se As New frmScriptEditor
		se.MdiParent = Me
		se.Show()
		LaunchProcess("ArrangeWindows")
	End Sub

	'*******************************************************************

	' The user has selected the relationship editor.

	'*******************************************************************
	Private Sub mnuRelationships_Click(sender As Object, e As EventArgs) Handles mnuRelationships.Click

		' Display the script builder

		frmRelationships.MdiParent = Me
		frmRelationships.Show()
		LaunchProcess("ArrangeWindows")

	End Sub
	'*******************************************************************

	' The "Generate Commands" context menu is selected.

	'*******************************************************************
	Private Sub mnuGenerateCommands_Click(sender As Object, e As EventArgs) Handles mnuGenerateCommands.Click

		' Declare variables

		Dim xx As DialogResult

		' Get the options

		frmGenerateCodeOptions.ShowDialog()
		If frmGenerateCodeOptions.DialogResult = Windows.Forms.DialogResult.OK Then

			If frmGenerateCodeOptions.optFile.Checked Then

				' Get the name of the new database.

				SaveFileDialog1.AddExtension = True
				SaveFileDialog1.Filter = "Visual Basic Module (*.vb)|*.vb|Text File (*.txt)|*.txt"
				SaveFileDialog1.DefaultExt = "vb"
				SaveFileDialog1.InitialDirectory = System.IO.Directory.GetCurrentDirectory
				SaveFileDialog1.OverwritePrompt = True
				SaveFileDialog1.SupportMultiDottedExtensions = True
				SaveFileDialog1.FileName = frmGenerateCodeOptions.txtAbbrev.Text & "_DataAdapterQueries.vb"
				xx = SaveFileDialog1.ShowDialog()

				' If the dialog returns a valid file name, proceed.

				GenerateDataAccessCommandModule(SaveFileDialog1.FileName)
			Else
				GenerateDataAccessCommandModule()
			End If
		End If
	End Sub

	'*******************************************************************
	'
	' When the mouse enters the main form, set the focus to
	' the main form.
	'
	'*******************************************************************
	Private Sub frmMain_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseEnter
		Me.Focus()
	End Sub
	'*******************************************************************

	' Following are routines that process various button and
	' menu click events: they merely call the process launch
	' routine with the name of the appropriate action to be
	' performed.  Why document each of them?

	'*******************************************************************
	'*******************************************************************
	'
	' Sub to launch a process.  This  may be called from
	' anyplace in the project to launch something.
	'
	'*******************************************************************
	Public Sub LaunchProcess(ByVal Process As String)

		' Declare variables

		Dim zx As String

		' Get a form of the process name that has no spaces and
		' is all lower case.

		zx = LCase(SRep(Process, 1, " ", ""))

		' Look for a matching process in our list to begin.

		Select Case zx

			Case "newdatabase"
				CreateNewDatabase()
			Case "opendatabase"
				OpenDB()
			Case "arrangewindows"
				If mnuTile.Checked = True Then
					Me.LayoutMdi(MdiLayout.TileHorizontal)
					mnuCascade.Checked = False
				ElseIf mnuCascade.Checked = True Then
					Me.LayoutMdi(MdiLayout.Cascade)
					mnuTile.Checked = False
				End If
			Case "compactdb"
				CompactDatabase()
			Case "shutdown"
				Me.Close()
			Case Else
				MsgBox("No process """ & Process & """ is available.", vbInformation, "Launch Process")
		End Select
	End Sub

	'*******************************************************************

	' The Tile Windows option is selected

	'*******************************************************************
	Private Sub mnuTile_Click(sender As Object, e As EventArgs) Handles mnuTile.Click
		mnuNoArrange.Checked = Not mnuNoArrange.Checked
		mnuCascade.Checked = Not mnuCascade.Checked
		LaunchProcess("ArrangeWindows")
	End Sub

	'*******************************************************************

	' The Cascade Windows option is selected

	'*******************************************************************
	Private Sub mnuCascade_Click(sender As Object, e As EventArgs) Handles mnuCascade.Click
		mnuNoArrange.Checked = Not mnuNoArrange.Checked
		mnuTile.Checked = Not mnuTile.Checked
		LaunchProcess("ArrangeWindows")
	End Sub
	'*******************************************************************

	' The No Arrange Windows option is selected

	'*******************************************************************
	Private Sub mnuNoArrange_Click(sender As Object, e As EventArgs) Handles mnuNoArrange.Click
		mnuTile.Checked = False
		mnuCascade.Checked = False
	End Sub

	'*******************************************************************

	' The Reset User Settings menu option is selected.

	'*******************************************************************
	Private Sub mnuResetSettings_Click(sender As Object, e As EventArgs) Handles mnuResetSettings.Click
		My.Settings.LastDatabaseUsed = ""
		My.Settings.SaveSQLTip = True
		frmRegistryEditor.ShowDialog()
	End Sub

End Class