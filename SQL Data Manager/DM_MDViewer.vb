Imports System.IO
Public Class frmMDViewer
	'***********************************************************************
	' Form to display .md files in Visual Basic programs.
	' EP_MDVIEWER.VB
	' Written: March 2026
	' Programmer: Aaron Scott
	' Copyright 2026 Sirius Software All Rights Reserved.
	'***********************************************************************

	' Loca variables for the properties.

	Private mFileName As String
	Private mEditing As Boolean

	'***********************************************************************

	' The form is loaded.

	'***********************************************************************
	Private Sub frmMDViewer_Load(sender As Object, e As EventArgs) Handles Me.Load

		If mEditing Then TextBox1.Text = MarkdownViewer1.RawText
		TextBox1.Font = New Font("Arial", 12)

	End Sub
	'***********************************************************************

	' Sub to load a file to be displayed.

	'***********************************************************************
	Public Sub LoadFile(path As String)
		MarkdownViewer1.LoadFile(path)
		mFileName = path
	End Sub
	'***********************************************************************

	' The EnableEditing property can be set to allow the user to both view
	' and edit an .md file.

	'***********************************************************************
	Public Property EnableEditing As Boolean
		Get
			Return mEditing
		End Get
		Set(value As Boolean)
			mEditing = value

			If value Then
				MenuStrip1.Visible = True
			Else
				MenuStrip1.Visible = False
			End If
		End Set
	End Property
	'***********************************************************************

	' The View Render menu option is clicked.  Display the rendered .md text.

	'***********************************************************************
	Private Sub mnuViewRender_Click(sender As Object, e As EventArgs) Handles mnuViewRender.Click
		MarkdownViewer1.Visible = True
		TextBox1.Visible = False
		MarkdownViewer1.RawText = TextBox1.Text
	End Sub
	'***********************************************************************

	' The View Editor menu option is clicked.  Show the editing box.

	'***********************************************************************
	Private Sub mnuViewEditor_Click(sender As Object, e As EventArgs) Handles mnuViewEditor.Click
		MarkdownViewer1.Visible = False
		TextBox1.Visible = True

	End Sub
	'***********************************************************************

	' The Save menu option is clicked.

	'***********************************************************************
	Private Sub mnuSave_Click(sender As Object, e As EventArgs) Handles mnuSave.Click

		If mFileName = "" Then mnuSaveAs_Click(sender, e)

		Try
			My.Computer.FileSystem.WriteAllText(mFileName, TextBox1.Text, False)
		Catch ex As Exception
			MsgBox("Cannot write to file " & mFileName & "." & vbCrLf & ex.Message, MsgBoxStyle.Information, "Cannot Save File")
		End Try

	End Sub
	'***********************************************************************

	' The SaveAs menu option is clicked.

	'***********************************************************************
	Private Sub mnuSaveAs_Click(sender As Object, e As EventArgs) Handles mnuSaveAs.Click

		Dim sf As New SaveFileDialog
		sf.Title = "Save File As"
		sf.FileName = "Untitled.md"
		sf.Filter = "Markdown Files (*.md)|*.md|All Files (*.*)|*.*"
		sf.ShowDialog()
		If sf.FileName <> "" Then
			mFileName = sf.FileName
			My.Computer.FileSystem.WriteAllText(TextBox1.Text, mFileName, False)
		End If
	End Sub
End Class