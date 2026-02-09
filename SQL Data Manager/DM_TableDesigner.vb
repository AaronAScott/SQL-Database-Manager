Option Strict Off
Option Explicit On
Imports System.Data.SqlClient
Imports VB = Microsoft.VisualBasic

Public Class frmTableDesigner
	Inherits System.Windows.Forms.Form
	'**********************************************************************
	' SQL Data Manager Table Designer form
	' DM_TABLEDESIGNER.VB
	' Written: October 2017
	' Programmer: Aaron Scott
	' Copyright (C) 2017 Sirius Software All Rights Reserved
	'**********************************************************************

	' Declare variables local to this module

	Private PanelsInitialized As Boolean = False
	Private FieldRowCount As Integer
	Private IndexRowCount As Integer
	Private CurrentTableName As String = ""
	Private CurrentIndexName As String = ""
	Private CurrentTableRow As Integer
	Private CurrentIndexRow As Integer
	Private ScriptFileName As String
	Private TokenFlag As String = Chr(0)
	Private LinkLabel1 As LinkLabel
	Private LinkLabel2 As LinkLabel

	' Create some collections to keep track of changes made to a table, if we are modifying
	' an existing table.

	Private TableDefinition As New Dictionary(Of Integer, TableRowValues)
	Private IndexesDefinition As New Dictionary(Of String, String)
	Private FieldsAdded As New Dictionary(Of Integer, TableRowValues)
	Private FieldsDropped As New Dictionary(Of Integer, String)
	Private FieldsModified As New Dictionary(Of Integer, TableRowValues)
	Private FieldNamesChanged As New Dictionary(Of Integer, String)
	Private IndexesAdded As New Dictionary(Of String, String)
	Private IndexesDropped As New Dictionary(Of String, String)

	Private Enum TableEditModeEnum
		TableCreate
		TableAlter
	End Enum
	Private TableEditMode As TableEditModeEnum = TableEditModeEnum.TableCreate

	' Define a structure for holding information about one field of a table.

	Private Structure TableRowValues
		Dim Field_Name As String
		Dim Data_Type As String
		Dim Field_Length As String
		Dim Is_Primary_Key As Boolean
		Dim Is_Nullable As Boolean
		Dim Is_Identity As Boolean
		Dim Is_Unique As Boolean
		Dim Is_Unicode As Boolean
		Dim Has_Default As Boolean
		Dim Default_Value As String
	End Structure

	' Define a structure for holding information about one index field to a table.

	Private Structure IndexRowValues
		Dim Is_Primary_Field As Boolean
		Dim Is_Clustered As Boolean
		Dim Is_Unique As Boolean
		Dim Sort As Byte
		Dim Index_Name As String
		Dim Field_Name As String
	End Structure

	' Define the text of the tip window

	Const TipTitle As String = "Saving an SQL Script"
	' Use the following RTF attributes in the tip text, as needed.
	' To emphasize a part of the text, use \i\cf1 emphasized text \cf0\plain\f3.  The cfx/cfo attributes change the font color. The fx attributes set and reset
	' a font.  Use \par for paragraphs.  Use \bullet to create a bulleted list.
	Const TipText As String = "When you save an SQL script from the Table Designer, the saved script will always be a CREATE TABLE script, which defines the new or altered design of the table in the designer.  If you are modifying an existing table, you may look at the SQL script that effects your changes by selecting the \b\cf1 View in Script Editor \cf0\plain\f3 menu item of the Script menu.  This allows you to save a script from which the table may be created in another database at any time. \par\par If you want a saved copy of the ALTER TABLE script that is generated, select \b\cf1 Save As \cf0\plain\f3 from the SQL Script Editor Script menu."
	Const ControlItem = "SaveSQLTip"

	'**********************************************************************

	' The form is loaded.

	'**********************************************************************
	Private Sub frmTableDesigner_Load(sender As Object, e As EventArgs) Handles MyBase.Load

		' Declare variables

		Dim zx As String

		' Put up a message

		frmMain.StatusLabel.Text = "Loading Form.  Please Wait..."
		System.Windows.Forms.Application.DoEvents()

		' Set the position of the main form and its size

		zx = GetSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "DesignerWindow", "Size")
		If zx <> "" Then
			Me.Top = Val(ParseString(zx))
			Me.Left = Val(ParseString(zx))
			Me.Height = Val(ParseString(zx))
			Me.Width = Val(zx)
		End If

		' Initialize both the table and index builder panels

		If Not PanelsInitialized Then InitializePanels()

		' Clear the message

		frmMain.StatusLabel.Text = ""
		System.Windows.Forms.Application.DoEvents()

	End Sub
	'**********************************************************************

	' The form is closed.

	'**********************************************************************
	Private Sub frmTableDesigner_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

		' Declare variables

		Dim zx As String
		Dim c As Control

		' Save the window size and position

		If Me.Left > 0 And Me.Top > 0 Then
			zx = Me.Top & "," & Me.Left & "," & Me.Height & "," & Me.Width
			SaveSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "DesignerWindow", "Size", zx)
		End If

		' Remove event handlers for the controls in each window.

		For Each c In Panel2.Controls
			Select Case VB.Left(c.Name, 1)
				Case "F", "L", "C", "D"
					RemoveHandler c.KeyDown, AddressOf TextBox_KeyDown
					RemoveHandler c.MouseDown, AddressOf TextBox_MouseDown
			End Select
		Next c

		For Each c In Panel4.Controls
			Select Case VB.Left(c.Name, 1)
				Case "I", "F"
					RemoveHandler c.KeyDown, AddressOf TextBox_KeyDown
					RemoveHandler c.MouseDown, AddressOf TextBox_MouseDown
			End Select
		Next c


		' Remove all controls

		Panel2.Controls.Clear()
		Panel4.Controls.Clear()

	End Sub
	'**********************************************************************

	' The Close button is clicked

	'**********************************************************************
	Private Sub mnuClose_Click(sender As Object, e As EventArgs) Handles mnuClose.Click


		Me.Close()

	End Sub

	'**********************************************************

	' The Open Script menu option is selected.

	'**********************************************************
	Private Sub mnuOpenSQL_Click(sender As Object, e As EventArgs) Handles mnuOpenSQL.Click

		' Declare variables

		Dim zx0 As String

		' Get the current directory (the common dialog box may
		' change it).

		zx0 = FileSystem.CurDir

		' Use the common dialog open box to get the file nafrmMain.

		frmMain.OpenFileDialog1.FileName = ""
		frmMain.OpenFileDialog1.Filter = "SQL Script Files (*.SQL)|*.SQL|All Files (*.*)|*.*"
		frmMain.OpenFileDialog1.FilterIndex = 1
		frmMain.OpenFileDialog1.DefaultExt = "SQL"
		frmMain.OpenFileDialog1.Title = "Open SQL Script File"
		If frmMain.OpenFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then

			' Save the file name

			ScriptFileName = frmMain.OpenFileDialog1.FileName

			' Restore the current directory

			FileSystem.ChDir(zx0)

			' Open the specified script file

			zx0 = My.Computer.FileSystem.ReadAllText(ScriptFileName)

			' Assign the script file to the SQLScript property.

			SQLScript = zx0

		End If
	End Sub
	'**********************************************************************

	' Save Script was selected

	'**********************************************************************
	Private Sub mnuSave_Click(sender As Object, e As EventArgs) Handles mnuSave.Click

		' Declare variables

		Dim zx As String

		' If there is no script file name, do a "save as" instead.

		If ScriptFileName = "" Then
			mnuSaveAs_Click(sender, New EventArgs)

			' Otherwise, save the modified script to the file.

		Else
			' Open the specified script file and write the script.

			zx = SQLScript
			My.Computer.FileSystem.WriteAllText(ScriptFileName, zx, False)

		End If
	End Sub
	'**********************************************************************

	' The Save As menu option is selected.

	'**********************************************************************
	Private Sub mnuSaveAs_Click(sender As Object, e As EventArgs) Handles mnuSaveAs.Click

		' Declare variables

		Dim zx As String
		Dim zx0 As String

		' If not turned off, show the tips window

		If My.Settings.Item(ControlItem) Then
			frmTips.TipTitle = TipTitle
			frmTips.TipText = TipText
			frmTips.ControlItem = ControlItem
			frmTips.ShowDialog()
		End If

		' Get the current directory (the common dialog box may
		' change it).

		zx0 = FileSystem.CurDir

		' Use the common dialog open box to get the file nafrmMain.

		frmMain.SaveFileDialog1.FileName = ""
		frmMain.SaveFileDialog1.Filter = "SQL Script Files (*.SQL)|*.SQL|All Files (*.*)|*.*"
		frmMain.SaveFileDialog1.FilterIndex = 1
		frmMain.SaveFileDialog1.DefaultExt = "sql"
		frmMain.SaveFileDialog1.AddExtension = True
		frmMain.SaveFileDialog1.Title = "Save SQL Script File"
		If frmMain.SaveFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then

			' Save the file name

			ScriptFileName = frmMain.SaveFileDialog1.FileName

			' Restore the current directory

			FileSystem.ChDir(zx0)

			' Open the specified script file and write the script.

			zx = SQLScript
			My.Computer.FileSystem.WriteAllText(ScriptFileName, zx, False)

		End If
	End Sub
	'**********************************************************************

	' The Scroll bar is clicked.

	'**********************************************************************
	Private Sub VScrollBar1_Scroll(sender As Object, e As ScrollEventArgs) Handles VScrollBar1.Scroll
		If e.Type = ScrollEventType.EndScroll Then
			Panel2.Top = -VScrollBar1.Value
			Panel4.Top = -VScrollBar1.Value
		End If
	End Sub
	'**********************************************************************

	' An field attributes checked list box is checked.

	'**********************************************************************
	Private Sub FieldAttributesChanged(sender As Object, e As EventArgs)

		' Declare variables

		Dim CkCount As Integer = 0
		Dim RowNum As Integer
		Dim FieldName As String
		Dim c As Control
		Dim Lb As CheckedListBox
		Dim T As TextBox

		' Get the row number

		RowNum = Val(Mid(sender.Name, 3))

		' Set the current field name.

		FieldName = Panel2.Controls("F" & RowNum).Text

		' Cast the sender as a checked list box.

		Lb = CType(Panel2.Controls("Lb" & RowNum), CheckedListBox)

		' Check the state of all the attributes checked list boxes in the table definition.

		For Each c In Panel2.Controls
			If TypeOf (c) Is CheckedListBox Then

				' Get the row number

				RowNum = Val(Mid(c.Name, 3))

				' Cast the control as a checked list box.

				Lb = CType(c, CheckedListBox)

				If VB.Left(Lb.Name, 2) = "Lb" Then

					' If more than one "primary key"  box was checked, alert the user
					' and uncheck the box.

					If Lb.GetItemChecked(0) Then ' 0 is index of "primary key" attribute
						CkCount += 1
						If CkCount > 1 Then
							MsgBox("Only one primary key may be defined in a table.", MsgBoxStyle.Information, "Add Field")
							Lb.SetItemChecked(0, False)
						End If
					End If

					' If the "Identity" box was checked, place the "(1,1)" seed value/increment values in the Default field.
					' Get the default value text box for this row.

					T = CType(Panel2.Controls("D" & RowNum), TextBox)

					' 2 is index of "Identity" attribute.  If it's checked, see if the default
					' column already has the seed value/increment info and add it if not. If the
					' indentity attribute is not checked, see if we need to remove a seed value/increment.

					If Lb.GetItemChecked(2) Then
						If T.Text <> "(1,1)" Then T.Text = "(1,1)"
					Else
						If T.Text = "(1,1)" Then T.Text = ""
					End If

					' If the "Default" attribute has been removed, clear the default value field.

					If Not Lb.GetItemChecked(2) And Not Lb.GetItemChecked(3) Then T.Text = ""

				End If

			End If
		Next c

	End Sub
	'**********************************************************************

	' A Primary/Clustered/Nonclustered index attribute checked list box is checked.

	'**********************************************************************
	Private Sub IndexAttributesChanged(sender As Object, e As EventArgs)

		' Declare variables

		Dim PrimaryKeyDefined As Boolean
		Dim CkCount As Integer = 0
		Dim c As Control
		Dim Lb As CheckedListBox

		' Determine if there is a primary key defined for this table.

		For Each c In Panel2.Controls
			If TypeOf (c) Is CheckedListBox Then
				Lb = CType(c, CheckedListBox)

				' If a "primary key"  box was checked, set the PrimaryKeyDefined flag to true.

				If VB.Left(Lb.Name, 2) = "Lb" And Lb.GetItemChecked(0) Then ' 0 is index of "primary key" attribute
					PrimaryKeyDefined = True
					Exit For
				End If
			End If
		Next c

		' Check for more than one index with the "clustered" attribute checked, or an attempt to
		' specify a clustered index when the table contains a primary key.

		For Each c In Panel4.Controls
			If TypeOf (c) Is CheckedListBox Then
				Lb = CType(c, CheckedListBox)

				' If more than one "clustered index" box was checked, alert the user
				' and uncheck the box.

				If VB.Left(Lb.Name, 2) = "Lb" And Lb.GetItemChecked(0) Then ' 0 is index of "clustered" attribute
					CkCount += 1
					If CkCount > 1 Then
						MsgBox("Only one clustered index may be defined in a table.", MsgBoxStyle.Information, "Create Index")
						Lb.SetItemChecked(0, False)
						Exit For

						' If a primary key exists, alert the user and uncheck the box.

					ElseIf PrimaryKeyDefined Then
						MsgBox("A clustered index may not be defined in a table with a primary key.", MsgBoxStyle.Information, "Create Index")
						Lb.SetItemChecked(0, False)
						Exit For
					End If
				End If
			End If
		Next c

	End Sub
	'**********************************************************************

	' The Add New Field Item (Table) link label is clicked. This is bound
	' at run time.

	'**********************************************************************
	Private Sub LinkLabel1_Click(sender As Object, e As EventArgs)

		' Add a new, empty row

		AddTableRow()

	End Sub
	'**********************************************************************

	' The Add New Field Item (Index) link label is clicked. This is bound
	' at run time.

	'**********************************************************************
	Private Sub LinkLabel2_Click(sender As Object, e As EventArgs)

		' Add a new, empty row

		AddIndexRow()

	End Sub
	'**********************************************************

	' Sub to handle KeyDown events for all the field/index/relation
	' text boxes.  This is bound at run time.

	'**********************************************************

	Private Sub TextBox_KeyDown(sender As Object, e As KeyEventArgs)

		' Declare variables and get the row number of the text box
		' in which the KeyDown event was triggered.

		Dim RowNum As Integer = Val(Mid(sender.name, 2))
		Dim zx As String = VB.Left(sender.name, 1)
		Dim T As TextBox

		' We only are interested in up and down arrows.

		If e.KeyCode = Keys.Up Or e.KeyCode = Keys.Down Then

			' Determine in which panel the textbox in which the key was pressed resides.

			Select Case sender.parent.name
				Case "Panel2"
					If e.KeyCode = Keys.Down And RowNum + 1 <= FieldRowCount Then RowNum += 1
					If e.KeyCode = Keys.Up And RowNum - 1 >= 1 Then RowNum -= 1
					T = Panel2.Controls(zx & RowNum)
					T.Focus()
					T.SelectAll()
				Case "Panel4"
					If e.KeyCode = Keys.Down And RowNum + 1 <= IndexRowCount Then RowNum += 1
					If e.KeyCode = Keys.Up And RowNum - 1 >= 1 Then RowNum -= 1
					T = Panel4.Controls(zx & RowNum)
					T.Focus()
					T.SelectAll()
			End Select

			' If the key was an up or down arrow, then indicate we've handled it here.

			e.Handled = True
		End If

	End Sub
	'**********************************************************

	' Sub to handle MouseDown events for the late-bound text boxes.  This
	' is bound at run time.

	'**********************************************************
	Private Sub TextBox_MouseDown(sender As Object, e As MouseEventArgs)

		Dim RowNum As Integer

		' Extract the row number from the object name.

		RowNum = Val(Mid(sender.name, 2))

		' Determine to which panel the text box which raised this event belongs.

		Select Case sender.parent.name

			Case "Panel2"
				CurrentTableRow = RowNum


			Case "Panel4"
				CurrentIndexRow = RowNum
		End Select

	End Sub

	'**********************************************************************

	' One of the option buttons, "Table" or "Index" is checked.

	'**********************************************************************

	Private Sub RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles optTable.CheckedChanged, optIndex.CheckedChanged

		' The "Table" radio button is checked.

		If sender Is optTable Then
			If CType(sender, RadioButton).Checked Then
				Panel1.Visible = True
				Panel3.Visible = False
				VScrollBar1.Left = Panel1.Left + Panel1.Width
				VScrollBar1.Maximum = Panel2.Height
			End If

			' The "Index" radio button is checked.  Because overlaying panel 1 with panel 3 at design time
			' makes panel 3 a child of panel 1, it had to be located off to the side.  So, when it's
			' made visible, move it to the same location as panel1 at run time.

		ElseIf sender Is optIndex Then
			If CType(sender, RadioButton).Checked Then
				Panel1.Visible = False
				Panel3.Visible = True
				Panel3.Location = Panel1.Location
				VScrollBar1.Left = Panel3.Left + Panel3.Width
				VScrollBar1.Maximum = Panel4.Height
			End If
		End If

		' Reset the scroll bar to the beginning whenever a view changes.

		VScrollBar1.Value = VScrollBar1.Minimum
	End Sub
	'**********************************************************************

	' The Upload to Database menu option is selected.

	'**********************************************************************

	Sub mnuUpload_Click(sender As Object, e As EventArgs) Handles mnuUpload.Click

		' Declare variables.

		Dim jj As Integer
		Dim zx As String = AssembleSQLScript(TableEditMode)
		Dim IndexName As String
		Dim Cmd = New SqlCommand
		Dim f As TableRowValues
		Dim Transaction As SqlTransaction

		' Remove the "GO" keywords, which are included for use with T-SQL.exe, but which
		' are not supported by ExecuteNonQuery.

		zx = SRep(zx, 1, "GO" & vbCrLf, "", CompareMethod.Binary)

		' Begin a transaction

		Transaction = DB.BeginTransaction

		' Set the command parameters.

		Cmd.Connection = DB
		Cmd.CommandText = zx
		Cmd.CommandType = CommandType.Text
		Cmd.Transaction = Transaction

		' Execute the query and watch for errors.

		Try
			Cmd.ExecuteNonQuery()
			StatusLabel1.Text = "Script executed successfully"
			Transaction.Commit()

			' Rebuild the object list.

			frmObjectList.RefreshObjectList()

			' After creating or modifying a table, clear all the collections that recorded the changes.

			FieldsAdded = New Dictionary(Of Integer, TableRowValues)
			FieldsDropped = New Dictionary(Of Integer, String)
			FieldsModified = New Dictionary(Of Integer, TableRowValues)
			IndexesAdded = New Dictionary(Of String, String)
			IndexesDropped = New Dictionary(Of String, String)
			CurrentTableName = TextBox1.Text

			' Clear and rebuild the TableDefinition collection with the new changes.

			TableDefinition.Clear()
			For jj = 1 To FieldRowCount
				Panel2.Controls("F" & jj).Tag = jj ' Put original row numbers back in order.
				f = GetRowValues(jj)
				TableDefinition.Add(jj, f)
			Next jj

			' Clear and rebuild the IndexesDefinition collection with the new changes.

			IndexesDefinition.Clear()
			zx = ""
			IndexName = ""
			If IndexRowCount > 0 Then
				For jj = 1 To IndexRowCount
					If Panel4.Controls("I" & jj).Text <> "" And Panel4.Controls("I" & jj).Text <> IndexName Then
						IndexesDefinition.Add(IndexName, zx)
						zx = ""
						IndexName = Panel4.Controls("I" & jj).Text
					End If

					' Assemble the list of field names.

					If zx <> "" Then zx &= ","
					zx &= Panel4.Controls("F" & jj).Text
					If CType(Panel4.Controls("C" & jj), ComboBox).SelectedIndex = 0 Then zx &= " ASC" Else zx &= " DESC"
				Next jj

				' Add the last index information.

				IndexesDefinition.Add(IndexName, zx)
			End If

			' No matter what we've just done, create or alter a table, further actions must
			' of course be altering a table.

			TableEditMode = TableEditModeEnum.TableAlter

		Catch ex As Exception

			StatusLabel1.Text = "Script failed"
			MsgBox(ex.Message, MsgBoxStyle.Information, "Update SQL Script")
			Transaction.Rollback()
		End Try

		' Enable the timer.  This will cause the messsage to stay on the status label
		' for 5 seconds, then be cleared.

		Timer1.Enabled = True

	End Sub

	'**********************************************************************

	' The View in Script Editor menu option is selected.

	'**********************************************************************
	Private Sub mnuView_Click(sender As Object, e As EventArgs) Handles mnuView.Click

		' Select the next control to make sure any Field Leave event is processed.
		' Otherwise, some changes may not be reflected in the SQL script generated.

		Panel2.SelectNextControl(Panel2.Controls(0), True, True, False, True)

		' Open a new editor window with the assembled script.

		Dim se As New frmScriptEditor
		se.MdiParent = frmMain
		se.DefaultText = AssembleSQLScript(TableEditMode)
		se.mnuExecute.Enabled = False
		se.Show()

	End Sub
	'**********************************************************************

	' The timer has clicked.  Erase any current status message.

	'**********************************************************************
	Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

		StatusLabel1.Text = ""
		Timer1.Enabled = False

	End Sub

	'**********************************************************************

	' The Insert Row menu option is selected.

	'**********************************************************************
	Private Sub mnuInsert_Click(sender As Object, e As EventArgs) Handles mnuInsertRow.Click

		Dim ii As Integer
		Dim jj As Integer
		Dim RowNum As Integer
		Dim M1 As ContextMenuStrip
		Dim Ck As CheckBox
		Dim T1 As TextBox
		Dim P1 As Panel

		' Find the parent control of this menu, and then the panel in which it's contained.

		M1 = CType(sender, ToolStripMenuItem).Owner
		If TypeOf (M1.SourceControl) Is TextBox Then
			T1 = M1.SourceControl
			RowNum = Val(Mid(T1.Name, 2))
			P1 = T1.Parent
		Else
			Ck = M1.SourceControl
			RowNum = Val(Mid(Ck.Name, 4))
			P1 = Ck.Parent
		End If

		' Add a new index row.

		ii = AddIndexRow() ' Save number of new row

		' Move the contents of all the rows down.

		For jj = IndexRowCount To RowNum + 1 Step -1
			P1.Controls("I" & jj).Text = P1.Controls("I" & (jj - 1)).Text
			P1.Controls("I" & jj).Tag = P1.Controls("I" & (jj - 1)).Tag ' This contains the original row number.
			P1.Controls("F" & jj).Text = P1.Controls("F" & (jj - 1)).Text
			P1.Controls("C" & jj).Text = P1.Controls("C" & (jj - 1)).Text
			CType(P1.Controls("Lb" & jj), CheckedListBox).SetItemChecked(0, CType(P1.Controls("Lb" & (jj - 1)), CheckedListBox).GetItemChecked(0))
			CType(P1.Controls("Lb" & jj), CheckedListBox).SetItemChecked(1, CType(P1.Controls("Lb" & (jj - 1)), CheckedListBox).GetItemChecked(1))
		Next jj

		' Make the current row blank.

		P1.Controls("I" & RowNum).Text = ""
		P1.Controls("I" & RowNum).Tag = ii ' The emptied ("inserted") row takes on the new row number
		P1.Controls("F" & RowNum).Text = ""
		P1.Controls("C" & RowNum).Text = "Ascending"
		CType(P1.Controls("Lb" & jj), CheckedListBox).SetItemChecked(0, False)
		CType(P1.Controls("Lb" & jj), CheckedListBox).SetItemChecked(1, False)

		' Set the focus to the field name field of the empty line.
		' Lock the index name field: we may not
		' insert a new index between lines, only another field to
		' an existing index.

		CType(P1.Controls("I" & RowNum), TextBox).ReadOnly = True
		P1.Controls("F" & RowNum).Focus()

	End Sub

	'**********************************************************************

	' The Delete Row menu option is selected.

	'**********************************************************************
	Private Sub mnuDelete_Click(sender As Object, e As EventArgs) Handles mnuDeleteRow.Click, mnuDeleteField.Click

		Dim DeleteAllIndexRows As Boolean
		Dim ii As Integer
		Dim jj As Integer
		Dim RowNum As Integer
		Dim M1 As ContextMenuStrip
		Dim Ck As CheckBox
		Dim T1 As TextBox
		Dim P1 As Panel

		' Find the parent control of this menu, and then the panel in which it's contained.

		M1 = CType(sender, ToolStripMenuItem).Owner
		If TypeOf (M1.SourceControl) Is TextBox Then
			T1 = M1.SourceControl
			RowNum = Val(Mid(T1.Name, 2))
			P1 = T1.Parent
		Else
			Ck = M1.SourceControl
			RowNum = Val(Mid(Ck.Name, 4))
			P1 = Ck.Parent
		End If

		' Determine what panel invoked this: table fields, index or relationships

		If P1 Is Panel2 Then ' Table design

			' Move the contents of all the rows up into the deleted row.

			For jj = RowNum To FieldRowCount - 1
				P1.Controls("F" & jj).Text = P1.Controls("F" & (jj + 1)).Text
				P1.Controls("F" & jj).Tag = P1.Controls("F" & (jj + 1)).Tag ' This contains the original row number
				P1.Controls("C" & jj).Text = P1.Controls("C" & (jj + 1)).Text
				P1.Controls("L" & jj).Text = P1.Controls("L" & (jj + 1)).Text
				For ii = 0 To CType(P1.Controls("Lb" & jj), CheckedListBox).Items.Count - 1
					CType(P1.Controls("Lb" & jj), CheckedListBox).SetItemChecked(ii, CType(P1.Controls("Lb" & (jj + 1)), CheckedListBox).GetItemChecked(ii))
				Next ii
				P1.Controls("D" & jj).Text = P1.Controls("D" & (jj + 1)).Text
			Next jj

			' Dispose of the bottom row.

			P1.Controls("F" & FieldRowCount).Dispose()
			P1.Controls("C" & FieldRowCount).Dispose()
			P1.Controls("L" & FieldRowCount).Dispose()
			P1.Controls("Lb" & FieldRowCount).Dispose()
			P1.Controls("D" & FieldRowCount).Dispose()
			FieldRowCount -= 1
			LinkLabel1.Top -= TextBox1.Height

		ElseIf P1 Is Panel4 Then ' Index design

			' See if the row being deleted contains the name of an index.  If it does, then
			' deleting it deletes the entire index.

			If P1.Controls("I" & RowNum).Text <> "" Then
				DeleteAllIndexRows = True ' Delete all subsequent fields with no index name
			Else
				DeleteAllIndexRows = False ' Delete only the one field of the index
			End If

			' If the deleted row is the first row of a multi-field index, then delete all succeeding rows until
			' a new index name is found.  If the row to be deleted has no name, delete it only.

			Do
				For jj = RowNum To IndexRowCount - 1
					P1.Controls("I" & jj).Text = P1.Controls("I" & (jj + 1)).Text
					P1.Controls("I" & jj).Tag = P1.Controls("I" & (jj + 1)).Tag ' This contains the original row number.
					P1.Controls("F" & jj).Text = P1.Controls("F" & (jj + 1)).Text
					P1.Controls("C" & jj).Text = P1.Controls("C" & (jj + 1)).Text
					For ii = 0 To CType(P1.Controls("Lb" & jj), CheckedListBox).Items.Count - 1
						CType(P1.Controls("Lb" & jj), CheckedListBox).SetItemChecked(ii, CType(P1.Controls("Lb" & (jj + 1)), CheckedListBox).GetItemChecked(ii))
					Next ii
				Next jj
				P1.Controls("I" & IndexRowCount).Dispose()
				P1.Controls("F" & IndexRowCount).Dispose()
				P1.Controls("C" & IndexRowCount).Dispose()
				P1.Controls("Lb" & IndexRowCount).Dispose()
				LinkLabel2.Top = LinkLabel2.Top - TextBox1.Height
				IndexRowCount -= 1
			Loop While RowNum <= IndexRowCount AndAlso DeleteAllIndexRows AndAlso P1.Controls("I" & RowNum).Text = ""
		End If
	End Sub
	'**********************************************************************

	' The New SQL Script menu option is selected.

	'**********************************************************************

	Private Sub mnuNewSQL_Click(sender As Object, e As EventArgs) Handles mnuNewSQL.Click

		' Reset the designer by clearing all controls and putting it into add mode.

		Panel2.Controls.Clear()
		Panel4.Controls.Clear()

		' Re-initialize the panels

		InitializePanels()

		' Put the designer into CREATE TABLE mode.

		TableEditMode = TableEditModeEnum.TableCreate

		' Reset the table name.

		TextBox1.Text = "New Table Name"

	End Sub
	'**********************************************************************

	' Sub to initialize the table builder, index and Relationship builder panels

	'**********************************************************************
	Private Sub InitializePanels()

		' Declare variables

		Dim RowTop As Integer = 5

		' Labels for the table panel

		Dim L1 As Label
		Dim L2 As Label
		Dim L3 As Label
		Dim L4 As Label
		Dim L5 As Label

		' Labels for the index panel

		Dim L6 As Label
		Dim L7 As Label
		Dim L8 As Label
		Dim L9 As Label

		' Create a font for the labels

		Dim f As New Font("Arial", 8, FontStyle.Bold)

		' Clear out all the controls that may have been previously created (if we have left and are re-entering
		' this form.)

		Panel2.Controls.Clear()
		Panel4.Controls.Clear()

		' Add the primary link label, "Add New Field", which serves as the starting point for locating
		' all the other controls we add, to each of the two panels.

		LinkLabel1 = New LinkLabel
		LinkLabel1.Name = "LinkLabel1"
		LinkLabel1.Top = 40
		LinkLabel1.Left = 125
		LinkLabel1.Text = "Add New Field"
		LinkLabel1.LinkColor = Color.ForestGreen
		Panel2.Controls.Add(LinkLabel1)
		AddHandler LinkLabel1.Click, AddressOf LinkLabel1_Click

		LinkLabel2 = New LinkLabel
		LinkLabel2.Name = "LinkLabel2"
		LinkLabel2.Top = 40
		LinkLabel2.Left = 125
		LinkLabel2.Text = "Add New Field"
		LinkLabel2.LinkColor = Color.ForestGreen
		Panel4.Controls.Add(LinkLabel2)
		AddHandler LinkLabel2.Click, AddressOf LinkLabel2_Click

		' Create labels to provide titles for the table columns.

		L1 = New Label
		L1.Name = "LF"
		L1.Left = 5
		L1.Top = RowTop
		L1.Width = 120
		L1.TextAlign = ContentAlignment.MiddleLeft
		L1.Font = f
		L1.Text = "Field Name"
		L1.BackColor = SystemColors.ControlDark

		L2 = New Label
		L2.Left = 5 + L1.Width
		L2.Top = RowTop
		L2.Width = 110
		L2.TextAlign = ContentAlignment.MiddleLeft
		L2.Font = f
		L2.Text = "Data Type"
		L2.BackColor = SystemColors.ControlDark

		L3 = New Label
		L3.Top = RowTop
		L3.Left = 5 + L1.Width + L2.Width
		L3.Width = 50
		L3.TextAlign = ContentAlignment.MiddleLeft
		L3.Font = f
		L3.Text = "Length"
		L3.BackColor = SystemColors.ControlDark

		L4 = New Label
		L4.Top = RowTop
		L4.Left = 5 + L1.Width + L2.Width + L3.Width
		L4.Width = 150
		L4.TextAlign = ContentAlignment.MiddleLeft
		L4.Font = f
		L4.Text = "Attributes"
		L4.BackColor = SystemColors.ControlDark

		L5 = New Label
		L5.Top = RowTop
		L5.Left = 5 + L1.Width + L2.Width + L3.Width + L4.Width
		L5.Width = 90
		L5.TextAlign = ContentAlignment.MiddleLeft
		L5.Font = f
		L5.Text = "Default Value"
		L5.BackColor = SystemColors.ControlDark


		' Create labels to provide columns for the index panel

		L6 = New Label
		L6.Name = "LI"
		L6.Top = RowTop
		L6.Left = 5
		L6.Width = 120
		L6.TextAlign = ContentAlignment.MiddleLeft
		L6.Font = f
		L6.Text = "Index Name"
		L6.BackColor = SystemColors.ControlDark

		L7 = New Label
		L7.Top = RowTop
		L7.Left = 5 + L6.Width
		L7.Width = 120
		L7.TextAlign = ContentAlignment.MiddleLeft
		L7.Font = f
		L7.Text = "Field Name"
		L7.BackColor = SystemColors.ControlDark

		L8 = New Label
		L8.Top = RowTop
		L8.Left = 5 + L6.Width + L7.Width
		L8.Width = 110
		L8.TextAlign = ContentAlignment.MiddleLeft
		L8.Font = f
		L8.Text = "Sort Order"
		L8.BackColor = SystemColors.ControlDark

		L9 = New Label
		L9.Top = RowTop
		L9.Left = 5 + L6.Width + L7.Width + L8.Width
		L9.Width = 150
		L9.TextAlign = ContentAlignment.MiddleCenter
		L9.Font = f
		L9.Text = "Attributes"
		L9.BackColor = SystemColors.ControlDark

		' Add the labels to the table panel controls collection

		Panel2.Controls.Add(L1)
		Panel2.Controls.Add(L2)
		Panel2.Controls.Add(L3)
		Panel2.Controls.Add(L4)
		Panel2.Controls.Add(L4)
		Panel2.Controls.Add(L5)
		RowTop = RowTop + L1.Height + 1

		' Initialize the count of rows in each view.

		FieldRowCount = 0 ' Count of field rows
		IndexRowCount = 0

		' Move the pre-existing link label down  below the last row.

		LinkLabel1.Top = RowTop
		LinkLabel1.Left = 5

		' Add the labels to the index panel controls collection

		Panel4.Controls.Add(L6)
		Panel4.Controls.Add(L7)
		Panel4.Controls.Add(L8)
		Panel4.Controls.Add(L9)

		' Move the pre-existing link label down  below the last row.

		LinkLabel2.Top = RowTop
		LinkLabel2.Left = 5

		' Set the maximum value of the scroll bar control to the size of the inner panel.

		VScrollBar1.Maximum = Panel2.Height
		VScrollBar1.Value = 0

		' Indicate we've done this process

		PanelsInitialized = True

		' Dispose of items we created.

		f.Dispose()

	End Sub
	'**********************************************************************

	' Sub to add a new row to the list of table fields.  This may be called
	' with pre-existing values, if a table is being modified from the 
	' "Schema" property.  It takes an argument of type TableRowValues, but
	' is declared as object so it will receive the value of "nothing" if
	' no object is passed.

	'**********************************************************************
	Private Function AddTableRow(Optional ByVal ColumnValues As Object = Nothing) As Integer

		' Declare variables

		Dim jj As Integer
		Dim RowTop As Integer
		Dim NewRowNumber As Integer
		Dim T1 As TextBox
		Dim T2 As TextBox
		Dim T3 As TextBox
		Dim C1 As ComboBox
		Dim Lb As CheckedListBox
		Dim f As New Font("Microsoft Sans Serif", 9.75)

		' Determine the highest row number used so far, and increment it for the new row.

		For jj = 1 To FieldRowCount
			If Panel2.Controls("F" & jj).Tag > NewRowNumber Then NewRowNumber = Panel2.Controls("F" & jj).Tag
		Next jj
		NewRowNumber += 1

		' Get the position of this linklabel. That is where we'll begin adding the new controls.

		RowTop = LinkLabel1.Top

		' Increment the count of rows

		FieldRowCount += 1

		' Create controls in the panel for each line item.

		T1 = New TextBox
		T1.Name = "F" & FieldRowCount ' A field name box
		T1.BorderStyle = BorderStyle.FixedSingle
		T1.Top = RowTop
		T1.Left = 5
		T1.Width = 120
		T1.Text = "New Field Name"
		T1.Font = f
		T1.Tag = NewRowNumber ' Save the original row number in case the contents are shifted up a row by deleting a field.
		T1.ContextMenuStrip = ContextMenuStrip1
		If Not IsNothing(ColumnValues) Then T1.Text = ColumnValues.Field_Name
		AddHandler T1.MouseDown, AddressOf TextBox_MouseDown
		AddHandler T1.KeyDown, AddressOf TextBox_KeyDown

		C1 = New ComboBox
		C1.Name = "C" & FieldRowCount ' A data type box
		C1.DropDownStyle = ComboBoxStyle.DropDown
		C1.Top = RowTop
		C1.Left = 5 + T1.Width
		C1.Width = 110
		For jj = 1 To 31
			C1.Items.Add(Choose(jj, "BIGINT", "BINARY", "BIT", "CHAR", "DATE", "DATETIME", "DATETIME2", "DATETIMEOFFSET", "DECIMAL", "FLOAT", "DOUBLE", "INT", "MONEY", "NCHAR", "NTEXT", "NVARCHAR", "REAL", "SMALLDATETIME", "SMALLINT", "SMALLMONEY", "STRUCTURED", "TEXT", "TIME", "TIMESTAMP", "TINYINT", "UDT", "UNIQUEIDENTIFIER", "VARBINARY", "VARCHAR", "VARIANT", "XML"))
		Next jj
		C1.SelectedIndex = 0
		C1.Font = New Font("Microsoft Sans Serif", 8.75)
		C1.IntegralHeight = False
		C1.Height = T1.Height + 5
		C1.DropDownStyle = ComboBoxStyle.DropDown
		C1.AutoCompleteMode = AutoCompleteMode.Append
		C1.AutoCompleteSource = AutoCompleteSource.ListItems
		If Not IsNothing(ColumnValues) Then
			For jj = 0 To C1.Items.Count - 1
				If C1.Items(jj) = ColumnValues.Data_Type Then
					C1.SelectedIndex = jj
					Exit For
				End If
			Next jj
		End If
		AddHandler C1.MouseDown, AddressOf TextBox_MouseDown
		AddHandler C1.KeyDown, AddressOf TextBox_KeyDown

		T2 = New TextBox
		T2.Name = "L" & FieldRowCount ' A field length box
		T2.BorderStyle = BorderStyle.FixedSingle
		T2.Top = RowTop
		T2.Left = 5 + T1.Width + C1.Width
		T2.Width = 50
		T2.Text = ""
		T2.Font = f

		If Not IsNothing(ColumnValues) Then
			If Val(ColumnValues.Field_Length) > 0 Then T2.Text = ColumnValues.Field_Length
			If Val(ColumnValues.field_length) = -1 Then T2.Text = "MAX"
		End If
		AddHandler T2.MouseDown, AddressOf TextBox_MouseDown
		AddHandler T2.KeyDown, AddressOf TextBox_KeyDown

		Lb = New CheckedListBox
		Lb.Name = "Lb" & FieldRowCount ' A field attribute box
		Lb.Top = RowTop
		Lb.Left = 5 + T1.Width + C1.Width + T2.Width
		Lb.Height = C1.Height
		Lb.Width = 150
		Lb.IntegralHeight = False
		Lb.Height = T1.Height
		Lb.CheckOnClick = True
		For jj = 1 To 6
			Lb.Items.Add(Choose(jj, "PRIMARY KEY", "NULL", "IDENTITY", "DEFAULT", "UNIQUE", "UNICODE"))
		Next jj
		If Not IsNothing(ColumnValues) Then
			Lb.SetItemChecked(0, ColumnValues.Is_Primary_Key)
			Lb.SetItemChecked(1, ColumnValues.Is_Nullable)
			Lb.SetItemChecked(2, ColumnValues.Is_Identity)
			Lb.SetItemChecked(3, ColumnValues.Has_Default)
			Lb.SetItemChecked(4, ColumnValues.Is_Unique)
			Lb.SetItemChecked(5, ColumnValues.Is_Unicode)
		End If
		AddHandler Lb.SelectedIndexChanged, AddressOf FieldAttributesChanged

		T3 = New TextBox
		T3.Name = "D" & FieldRowCount ' A default value/indentity box
		T3.BorderStyle = BorderStyle.FixedSingle
		T3.Top = RowTop
		T3.Left = 5 + T1.Width + C1.Width + T2.Width + Lb.Width
		T3.Width = 90
		T3.Text = ""
		T3.Font = f
		If Not IsNothing(ColumnValues) Then T3.Text = ColumnValues.Default_Value
		AddHandler T3.MouseDown, AddressOf TextBox_MouseDown
		AddHandler T3.KeyDown, AddressOf TextBox_KeyDown

		' Add the new controls to the pane.

		Panel2.Controls.Add(T1)
		Panel2.Controls.Add(C1)
		Panel2.Controls.Add(T2)
		Panel2.Controls.Add(Lb)
		Panel2.Controls.Add(T3)

		' Increase the X position by the height of this new row

		RowTop += T1.Height

		' Move the pre-existing link label down to be below the last row

		LinkLabel1.Top = RowTop
		LinkLabel1.Left = 5

		' If the panel is too small for all the controls, enlarge it.

		If LinkLabel1.Top + LinkLabel1.Height > Panel2.Height Then Panel2.Height = LinkLabel1.Top + LinkLabel1.Height + 5

		' Set the maximum value of the scroll bar control to the size of the inner panel.

		VScrollBar1.Maximum = Panel2.Height + 100
		VScrollBar1.SmallChange = T1.Height
		VScrollBar1.LargeChange = 200


		' Return the new row value

		Return FieldRowCount

	End Function
	'**********************************************************************

	' Sub to add a new row to the index list.  This takes an argument of
	' type IndexValues, but is declared as object so that it will receive
	' the value of "nothing", which it won't get if it's declared as
	' IndexValues.

	'**********************************************************************
	Private Function AddIndexRow(Optional ByVal IndexRow As Object = Nothing) As Integer

		' Declare variables

		Dim jj As Integer
		Dim RowTop As Integer
		Dim NewRowNumber As Integer
		Dim T1 As TextBox
		Dim T2 As TextBox
		Dim C1 As ComboBox
		Dim Lb As CheckedListBox
		Dim f As New Font("Microsoft Sans Serif", 9.75)

		' Determine the highest row number used so far, and increment it for the new row.

		For jj = 1 To IndexRowCount
			If Panel4.Controls("I" & jj).Tag > NewRowNumber Then NewRowNumber = Panel4.Controls("I" & jj).Tag
		Next jj
		NewRowNumber += 1

		' Get the position of this linklabel. That is where we'll begin adding the new controls.

		RowTop = LinkLabel2.Top

		' Increment the count of rows

		IndexRowCount += 1

		' Create controls in the panel for each field in the index.

		T1 = New TextBox
		T1.Name = "I" & IndexRowCount
		T1.BorderStyle = BorderStyle.FixedSingle
		T1.Top = RowTop
		T1.Left = 5
		T1.Width = 120
		T1.Text = ""
		T1.Font = f
		T1.Tag = NewRowNumber  ' Save the original row number, in case the contents are shifted up or down a row by deleting or inserting a field.
		T1.ContextMenuStrip = ContextMenuStrip2
		If Not IsNothing(IndexRow) AndAlso IndexRow.Is_Primary_Field Then T1.Text = IndexRow.Index_Name
		AddHandler T1.MouseDown, AddressOf TextBox_MouseDown
		AddHandler T1.KeyDown, AddressOf TextBox_KeyDown

		T2 = New TextBox
		T2.Name = "F" & IndexRowCount
		T2.BorderStyle = BorderStyle.FixedSingle
		T2.Top = RowTop
		T2.Left = 5 + T1.Width
		T2.Width = 120
		T2.Text = ""
		T2.Font = f
		If Not IsNothing(IndexRow) Then T2.Text = IndexRow.Field_Name
		AddHandler T2.MouseDown, AddressOf TextBox_MouseDown
		AddHandler T2.KeyDown, AddressOf TextBox_KeyDown

		C1 = New ComboBox
		C1.Name = "C" & IndexRowCount
		C1.DropDownStyle = ComboBoxStyle.DropDown
		C1.Top = RowTop
		C1.Left = 5 + T1.Width + T2.Width
		C1.Width = 110
		For jj = 1 To 2
			C1.Items.Add(Choose(jj, "Ascending", "Descending"))
		Next jj
		C1.Font = New Font("Microsoft Sans Serif", 8.75)
		C1.IntegralHeight = False
		C1.Height = T1.Height + 5
		If Not IsNothing(IndexRow) Then C1.SelectedIndex = IndexRow.Sort Else C1.SelectedIndex = 0
		AddHandler C1.KeyDown, AddressOf TextBox_KeyDown

		Lb = New CheckedListBox
		Lb.Name = "Lb" & IndexRowCount
		Lb.Top = RowTop
		Lb.Left = 5 + T1.Width + C1.Width + T2.Width
		Lb.Height = C1.Height
		Lb.Width = 150
		Lb.IntegralHeight = False
		Lb.Height = T1.Height
		Lb.CheckOnClick = True
		For jj = 1 To 2
			Lb.Items.Add(Choose(jj, "CLUSTERED", "UNIQUE"))
		Next jj
		If Not IsNothing(IndexRow) Then
			Lb.SetItemChecked(0, IndexRow.is_Clustered)
			Lb.SetItemChecked(1, IndexRow.is_Unique)
		End If
		AddHandler Lb.SelectedIndexChanged, AddressOf IndexAttributesChanged

		Panel4.Controls.Add(T1)
		Panel4.Controls.Add(T2)
		Panel4.Controls.Add(C1)
		Panel4.Controls.Add(Lb)

		' Set the focus to the index name field and clear the FieldChanged flag.

		T1.Focus()

		' Increase the X position by the height of this new row

		RowTop += T1.Height

		' Move the pre-existing link label down to be below the last row

		LinkLabel2.Top = RowTop
		LinkLabel2.Left = 5

		' Set the maximum value of the scroll bar control to the size of the inner panel.

		VScrollBar1.Maximum = Panel4.Height
		VScrollBar1.LargeChange = T1.Height

		' Return the new index row number

		Return IndexRowCount

	End Function
	'**********************************************************************

	' Function to determine and return the name of the index to which
	' a field belongs.

	'**********************************************************************
	Private Function FindIndexName(ByVal Row As Integer) As String

		' Declare variables

		Dim ii As Integer

		ii = Row
		Do While ii > 1 AndAlso Panel4.Controls("I" & ii).Text = ""
			ii -= 1
		Loop
		Return Panel4.Controls("I" & ii).Text

	End Function

	'**********************************************************************

	' Function to tokenize an SQL script.

	'**********************************************************************
	Private Function TokenizeScript(ByVal SQL As String) As String

		' Declare variables

		Dim Bracketed As Boolean
		Dim ii As Integer
		Dim jj As Integer
		Dim StartPos As Integer
		Dim EndPos As Integer
		Dim xx As String
		Dim zx As String = Chr(0)
		Dim Word As String = ""
		Dim Keywords() As String = {"", "CREATE", "ALTER", "DROP", "TABLE", "INDEX", "CONSTRAINT", "FOREIGN KEY", "REFERENCES", "ON", "NOT NULL", "NULL", "PRIMARY KEY", "IDENTITY", "UNIQUE", "DEFAULT", "CLUSTERED", "NONCLUSTERED", "ASC", "DESC", "UPDATE CASCADE", "DELETE CASCADE"}

		' Begin replacing keywords with tokens. Keywords must begin with
		' a space, or be at the beginning of the line.

		ii = 1
		Word = ""
		Do While ii <= Len(SQL)

			' Get the current character of the text as we iterate through it.

			xx = Mid(SQL, ii, 1)

			' Watch for brackets.

			If xx = "[" Then
				Bracketed = True
				Word += xx
			ElseIf xx = "]" Then
				Word += xx
				Bracketed = False

				' If the character is alphanumeric, or an underscore (which is the only valid
				' non-alphanumeric character allowed in names), or bracketed (which encloses names),
				' accumulate it as a word.

			ElseIf (UCase(xx) >= "A" And UCase(xx) <= "Z") Or (xx >= "0" And xx <= "9") Or xx = "_" Or Bracketed Then
				If StartPos = 0 Then StartPos = ii
				Word += UCase(xx)

				' Any other character is a delimiter.  If we haven't set the endpoint of a word
				' yet, we may go ahead and check the completed word against the list.

			ElseIf EndPos = 0 And Word <> "" Then

				' Some keywords consist of two words: "FOREIGN KEY", "NOT NULL"  or "PRIMARY KEY".
				' We need to keep accumulating the "word" if we've found the first half of
				' any of those terms.

				If xx = " " And (Word = "FOREIGN" Or Word = "NOT" Or Word = "PRIMARY" Or Word = "UPDATE" Or Word = "DELETE") Then
					Word += xx

					' Otherwise, if the word is a keyword, replace it with a token and reset
					' our first/last pointers, reset the index through the string, and clear
					' out the word for the next one.

				Else
					EndPos = ii
					For jj = 1 To UBound(Keywords)
						If Word = Keywords(jj) Then
							If StartPos > 1 Then
								SQL = VB.Left(SQL, StartPos - 1) & zx & Chr(jj) & Mid(SQL, EndPos)
							Else
								SQL = zx & Chr(jj) & Mid(SQL, EndPos) ' Word is at first of line
							End If
							ii = StartPos + 2
							Exit For
						End If
					Next jj
					StartPos = 0
					EndPos = 0
					Word = ""
				End If
			End If
			ii += 1
		Loop

		' Return the tokenized script.

		Return SQL

	End Function
	'**********************************************************************

	' Sub to parse an SQL script.

	'**********************************************************************
	Private Sub ParseScript(ByVal SQL As String)

		' Declare variables

		Dim UniqueIndex As Boolean
		Dim StartPos As Integer
		Dim EndPos As Integer
		Dim ParenLevel As Integer
		Dim Token As Integer
		Dim ii As Integer
		Dim Action As Integer
		Dim IndexType As Integer
		Dim xx As String
		Dim TableName As String = ""
		Dim IndexName As String = ""

		' First tokenize the script.

		SQL = TokenizeScript(SQL)

		' Remove text in which we're not interested.  This includes the [dbo]. preface to the table name and the
		' GO keyword, which is not supported by ExecuteNonQuery.

		SQL = Trim(SRep(SRep(SQL, 1, "[dbo].", ""), 1, "GO" & vbCrLf, "", CompareMethod.Binary))

		' Begin parsing the script, looking for tokens. At this level of parsing,
		' we don't do anything except look for certain tokens.

		ii = 1
		ParenLevel = 0
		Do While ii <= Len(SQL)
			xx = Mid(SQL, ii, 1)

			' We have encountered a null code.  That means that the next value in the string is
			' a token we need to process.  The tokens in numerical order are:
			' "CREATE", "ALTER", "DROP", "TABLE", "INDEX", "CONSTRAINT", "FOREIGN KEY", "REFERENCES", "ON", "NOT NULL", "NULL", "PRIMARY KEY", "IDENTITY", "UNIQUE", "DEFAULT", "CLUSTERED", "NONCLUSTERED", "ASC", "DESC", "UPDATE CASCADE", "DELETE CASCADE"

			If xx = TokenFlag Then
				ii += 1
				Token = Asc(Mid(SQL, ii, 1))
				Select Case Token
					Case 1, 2, 3 ' Create, alter or drop command
						Action = Token
					Case 4 ' Table
						ii += 1 ' Advance past the "table" token.

						' Extract the table name, which will be enclosed in "[]" brackets.
						' Then extract the entire table definition, which will be enclosed
						' by the outermost parentheses.

						StartPos = 0
						EndPos = 0
						Do While ii <= Len(SQL)
							xx = Mid(SQL, ii, 1)
							Select Case xx
								Case "["
									If TableName = "" Then StartPos = ii + 1
								Case "]"
									If TableName = "" And StartPos > 0 Then
										EndPos = ii - 1
										TableName = Mid(SQL, StartPos, EndPos - StartPos + 1)
										TextBox1.Text = TableName

										' Remember the table name, in case it changes.

										CurrentTableName = TextBox1.Text
									End If
								Case "("
									If ParenLevel = 0 Then StartPos = ii ' We found the outer opening parenthesis.
									ParenLevel += 1
								Case ")"
									If ParenLevel = 1 Then ' We found the outer closing parenthesis.
										EndPos = ii
										ParseTableDefinition(Action, Mid(SQL, StartPos, EndPos - StartPos + 1))
										ParenLevel = 0
										Exit Do
									End If
									ParenLevel -= 1
								Case Else
							End Select
							ii += 1
						Loop

					Case 5 ' Index

						' Extract the index name, which will be enclosed in "[]" brackets.
						' Then extract the entire index definition, which will end with
						' an outermost closing parenthesis.

						StartPos = 0
						EndPos = 0
						Do While ii <= Len(SQL)
							xx = Mid(SQL, ii, 1)
							Select Case xx
								Case "["
									If IndexName = "" Then StartPos = ii
								Case "]"
									If IndexName = "" And StartPos > 0 Then
										EndPos = ii - 1
										IndexName = Mid(SQL, StartPos, EndPos - StartPos + 1)
									End If
								Case "("
									ParenLevel += 1
								Case ")"
									If ParenLevel = 1 Then ' We found the outer closing parenthesis.
										EndPos = ii
										ParseIndexDefinition(UniqueIndex, IndexType, Mid(SQL, StartPos, EndPos - StartPos + 1))
										UniqueIndex = False
										IndexType = 0
										ParenLevel = 0
										IndexName = ""
										Exit Do
									End If
									ParenLevel -= 1
								Case Else
							End Select
							ii += 1
						Loop
					Case 14 ' Unique keyword
						UniqueIndex = True
					Case 16, 17 ' Clustered or non-clustered keyword
						IndexType = Token
					Case Else
						MsgBox("Unrecognized or unexpected token encountered in SQL script.  Value=" & Token & ".", MsgBoxStyle.Exclamation, "Table Designer")
				End Select
			End If
			ii += 1
		Loop
	End Sub
	'**********************************************************************

	' Sub to parse the contents of a table definition into individual
	' field definition lines.

	'**********************************************************************
	Private Sub ParseTableDefinition(ByVal Action As Integer, ByVal SQL As String)

		' Declare variables

		Dim ii As Integer
		Dim ParenLevel As Integer
		Dim xx As String
		Dim zx As String

		' See what our action is: create or alter.

		If Action = 1 Then ' create table

			ii = 1
			ParenLevel = 0
			zx = ""
			Do While ii <= Len(SQL)
				xx = Mid(SQL, ii, 1)
				Select Case xx
					Case "("
						ParenLevel += 1
						If ParenLevel > 1 Then zx += xx
					Case ")" ' The last line will end with a close parenthesis
						ParenLevel -= 1
						If ParenLevel = 0 Then ParseFieldDefinition(zx)
						If ParenLevel > 0 Then zx += xx
					Case "," ' All but the last line ends with a comma
						If ParenLevel = 1 Then
							ParseFieldDefinition(zx)
							zx = ""
						Else
							zx += xx
						End If
					Case Else
						zx += xx
				End Select
				ii += 1
			Loop
		End If

	End Sub

	'**********************************************************************

	' Sub to parse a line of SQL code to extract the components of a
	' field definition.

	'**********************************************************************
	Private Sub ParseFieldDefinition(ByVal Line As String)

		' Declare variables

		Dim ii As Integer
		Dim jj As Integer
		Dim ParenLevel As Integer
		Dim Token As Integer
		Dim xx As String = ""
		Dim zx As String = ""

		Dim FieldName As String = ""
		Dim DataType As String = ""
		Dim FieldLen As String = ""
		Dim IdentityParameters As String = ""
		Dim DefaultVal As String = ""
		Dim ColumnValues As New TableRowValues

		' Trim white space from the beginning of the line

		For jj = 1 To Len(Line)
			xx = Mid(Line, jj, 1)
			If xx = TokenFlag Or xx > " " Then
				Line = Mid(Line, jj)
				Exit For
			End If
		Next jj

		' Begin parsing the line into a field name, a data type, a field len (maybe),
		' and field attributes (maybe). Lines may also be Constraint lines or Primary Key lines.
		' Check for those first. Also, a constraint may be a primary key, so we'll handle
		' all these variations as constraints.  (Usually a primary key is defined by an attribute
		' on the field line, but not always, and it may or may not be defined as a constraint.)

		If VB.Left(Line, 2) = TokenFlag & Chr(6) Or VB.Left(Line, 2) = TokenFlag & Chr(12) Then
			ParseConstraint(Line)

			' Once these exceptions are handled, other lines are (or should be) ordinary
			' field definitions.
		Else

			ii = 1
			zx = ""
			Do While ii <= Len(Line)
				xx = Mid(Line, ii, 1)
				Select Case xx
					Case " " ' A space will delimit each item in the definition.
						If FieldName <> "" And DataType = "" And zx <> "" Then
							For jj = 1 To Len(zx)
								If Mid(zx, jj, 1) = "(" Then
									FieldLen = CStr(Val(Mid(zx, jj + 1)))
									zx = VB.Left(zx, jj - 1)
									Exit For
								End If
							Next jj
							Select Case zx
								Case "BIGINT", "BINARY", "BIT", "CHAR", "DATETIME", "DATETIME2", "DATETIMEOFFSET", "DECIMAL", "FLOAT", "DOUBLE", "INT", "MONEY", "NCHAR", "NTEXT", "NVARCHAR", "REAL", "SMALLDATETIME", "SMALLINT", "SMALLMONEY", "STRUCTURED", "TEXT", "TIME", "TIMESTAMP", "TINYINT", "UDT", "UNIQUEIDENTIFIER", "VARBINARY", "VARCHAR", "VARIANT", "XML"
									DataType = zx
							End Select
						ElseIf FieldLen <> "" And DataType <> "" And Val(zx) > 0 Then
							FieldLen = zx
						End If
						zx = ""

					Case "]" ' Field names are enclosed in brackets.
						FieldName = Mid(zx, 2)
						zx = ""

					Case TokenFlag
						ii += 1
						Token = Asc(Mid(Line, ii, 1))
						Select Case Token
							Case 10 ' Not null
							Case 11 ' Null
								ColumnValues.Is_Nullable = True
							Case 12 ' Primary Key
								ColumnValues.Is_Primary_Key = True
							Case 13 ' Identity
								ColumnValues.Is_Identity = True
								zx = ""
								ii += 1
								Do While ii <= Len(Line)
									xx = Mid(Line, ii, 1)
									If xx <> " " Then zx += xx
									If xx = ")" Then
										IdentityParameters = zx
										zx = ""
									End If
									ii += 1
								Loop
							Case 14 ' Unique
								ColumnValues.Is_Unique = True
							Case 15 ' Default
								ColumnValues.Has_Default = True
								zx = ""
								ii += 1
								Do While ii <= Len(Line)
									xx = Mid(Line, ii, 1)
									If xx <> " " Then zx += xx
									If xx = "," Or xx = TokenFlag Or ii = Len(Line) Then
										DefaultVal = zx
										zx = ""
									End If
									ii += 1
								Loop
						End Select
					Case "("
						ParenLevel += 1
						zx += xx ' Don't accumulate outer level parenthesis
					Case ")"
						ParenLevel -= 1
						zx += xx ' Don't accumulate outer level parenthesis
					Case vbLf, vbCr, vbCrLf, vbTab
					Case Else
						zx += xx
				End Select
				ii = ii + 1
			Loop

			' Now move the parsed values into the ColumnAttributes array

			ColumnValues.Field_Name = FieldName
			ColumnValues.Data_Type = DataType
			ColumnValues.Field_Length = FieldLen
			ColumnValues.Default_Value = DefaultVal & IdentityParameters ' Both of these will not be non-blank

			' Add the line to the list of fields.

			AddTableRow(ColumnValues)
		End If
	End Sub
	'**********************************************************************

	' Sub to parse an index definition line.

	'**********************************************************************
	Private Sub ParseIndexDefinition(UniqueIndex As Boolean, IndexType As Integer, Line As String)

		' Declare variables

		Dim ii As Integer
		Dim xx As String
		Dim zx As String = ""
		Dim IndexName As String = ""
		Dim TableName As String = ""
		Dim IndexRow As New IndexRowValues

		' Token values that may be encountered in an index line are:
		' 9 = On
		' 12 = Primary
		' 14 = Unique
		' 16 = Clustered
		' 17 = Nonclustered (Default)
		' 18 = Ascending (Default)
		' 19 = Descending
		' Because the CLUSTERED/NONCLUSTERED option, if present, comes before the INDEX keyword, it was extracted
		' much earlier in the parsing process, and so it passed to this routine as IndexType.  Set that
		' value now.  The same is true of the UNIQUE keyword.

		IndexRow.Is_Unique = UniqueIndex
		If IndexType = 16 Then IndexRow.Is_Clustered = True

		' Now begin parsing the line.  First thing to extract is the index name.  If an index contains
		' more than one field, the name will be on the first line added only, with one subsequent line
		' with no index name for each additional field in the index.

		ii = 1
		Do While ii <= Len(Line)
			xx = Mid(Line, ii, 1)
			If (xx.ToUpper >= "A" And xx.ToUpper <= "Z") Or (xx >= "0" And xx <= "9") Or xx = "_" Then
				zx += xx
			ElseIf xx = Chr(9) Then ' ON keyword encountered: we can save the index name now.
				IndexName = zx
				IndexRow.Index_Name = IndexName
				zx = ""
			ElseIf xx = Chr(14) Then
				IndexRow.Is_Unique = True
			ElseIf xx = Chr(19) Then ' The DESC keyword is encountered.
				IndexRow.Sort = 1 ' 0=Ascending, 1=Descending
			ElseIf xx = "(" Then ' Beginning of field list: we can begin saving fields now.
				zx = ""
			ElseIf xx = "]" And IndexName <> "" And TableName = "" Then ' We We don't need the table name, but we save it as an indicator that we're ready to accumulate field names.
				TableName = zx
				zx = ""
			ElseIf (xx = "," Or xx = ")") And IndexName <> "" And TableName <> "" Then ' We have one field name.  Add a row, then clear out the index name in case another row is added for the same index
				IndexRow.Field_Name = zx
				AddIndexRow(IndexRow)
				IndexRow = New IndexRowValues
				zx = ""
			End If
			ii += 1
		Loop

	End Sub
	'**********************************************************************

	' Sub to parse a CONSTRAINT definition.

	'**********************************************************************
	Private Sub ParseConstraint(ByVal Line As String)

		' Declare variables

		Dim ii As Integer
		Dim jj As Integer
		Dim xx As String
		Dim FieldName As String
		Dim KeyValues As New ForeignKeyValues

		' There are three types of constraints: unique fields, foreign
		' keys (relationships) and primary keys. Token values are:
		' 7 = Foreign Key (not handled here)
		' 8 = References
		' 12 = Primary Key (declared as a constraint or not)
		' 14 = Unique

		' The constraint is a primary key

		If InStr(Line, Chr(12)) > 0 Then ' Primary key with or without constraint.
			FieldName = ""
			ii = InStr(Line, Chr(12)) + 1
			Do While ii <= Len(Line)
				xx = Mid(Line, ii, 1)
				If (UCase(xx) > "A" And UCase(xx) <= "Z") Or (xx >= "0" And xx <= "9") Then
					FieldName += xx
				End If
				ii += 1
			Loop

			' Now look for the specified field in the table field rows and when found,
			' check the "Primary Key"  attribute for that field.

			For jj = 0 To Panel2.Controls.Count - 1
				If TypeOf (Panel2.Controls(jj)) Is TextBox And VB.Left(Panel2.Controls(jj).Name, 1) = "F" Then ' Field name
					If LCase(Panel2.Controls(jj).Text) = LCase(FieldName) Then ' Field matches
						CType(Panel2.Controls(jj + 3), CheckedListBox).SetItemChecked(0, True) ' 4 is index of "Primary Key" attribute.
						Exit For
					End If
				End If
			Next

			' The constraint is a unique field constraint.

		Else

			FieldName = ""
			ii = InStr(Line, Chr(14)) + 1 ' Unique constraint
			Do While ii <= Len(Line)
				xx = Mid(Line, ii, 1)
				If (UCase(xx) > "A" And UCase(xx) <= "Z") Or (xx >= "0" And xx <= "9") Then
					FieldName += xx
				End If
				ii += 1
			Loop

			' Now look for the specified field in the table field rows and when found,
			' check the "Unique"  attribute for that field.

			For jj = 0 To Panel2.Controls.Count - 1
				If TypeOf (Panel2.Controls(jj)) Is TextBox And VB.Left(Panel2.Controls(jj).Name, 1) = "F" Then ' Field name
					If LCase(Panel2.Controls(jj).Text) = LCase(FieldName) Then ' Field matches
						CType(Panel2.Controls(jj + 3), CheckedListBox).SetItemChecked(4, True) ' 4 is index of "Unique" attribute.
						Exit For
					End If
				End If
			Next
		End If

	End Sub
	'**********************************************************************

	' Function to assemble the SQL script that defines a table or ohanges to
	' a table from the information entered into the table designer.

	'**********************************************************************
	Private Function AssembleSQLScript(ScriptType As TableEditModeEnum) As String

		' Declare variables.

		Dim ii As Integer
		Dim ww As String = ""
		Dim xx As String = ""
		Dim yy As String = ""
		Dim zx As String = ""
		Dim Script As String = ""
		Dim PrimaryKey As String = ""
		Dim IndexName As String = ""
		Dim CL As Collection
		Dim kvp1 As KeyValuePair(Of Integer, TableRowValues)
		Dim kvp2 As KeyValuePair(Of Integer, String)
		Dim kvp4 As KeyValuePair(Of String, String)
		Dim i As IndexRowValues
		Dim f As TableRowValues

		'******************************************************************************************

		' CREATE TABLE
		' Assemble the script for creating a new table.

		'******************************************************************************************

		If ScriptType = TableEditModeEnum.TableCreate Then

			' Get all the rows in the table.

			FieldsAdded.Clear()
			For jj = 1 To FieldRowCount

				' Get the values from one row.

				f = GetRowValues(jj)

				' Add the values for the row to the table definition.  We index the fields
				' only by row number, so we can later determine if a field is renamed.

				FieldsAdded.Add(jj, f)
			Next jj

			xx = "CREATE TABLE [dbo].[" & TextBox1.Text & "]" & vbCrLf & "(" & vbCrLf

			' Begin looping through all the defined fields, adding them to the text.

			For Each kvp1 In FieldsAdded
				If Script <> "" Then Script &= "," & vbCrLf
				Script &= "     [" & kvp1.Value.Field_Name & "]  "
				Script &= kvp1.Value.Data_Type & " "
				If kvp1.Value.Field_Length <> "" Then Script &= "(" & kvp1.Value.Field_Length & ")"
				If kvp1.Value.Is_Primary_Key And PrimaryKey = "" Then PrimaryKey = "     PRIMARY KEY CLUSTERED ([" & kvp1.Value.Field_Name & "] ASC)"
				If kvp1.Value.Is_Unicode Then Script &= " COLLATE Latin1_General_100_CI_AI_SC_UTF8"
				If kvp1.Value.Is_Nullable Then Script &= " NULL" Else Script &= " NOT NULL"
				If kvp1.Value.Has_Default Then
					If IsNumericDefault(kvp1.Value.Default_Value) Then
						Script &= " DEFAULT ((" & kvp1.Value.Default_Value & "))"
					Else
						zx = SRep(SRep(kvp1.Value.Default_Value, 1, """", "'"), 1, "'", "")
						Script &= " DEFAULT (('" & kvp1.Value.Default_Value & "'))"
					End If
				End If
				If kvp1.Value.Is_Unique Then Script &= " UNIQUE"
				If kvp1.Value.Is_Identity Then Script &= " IDENTITY " & kvp1.Value.Default_Value
			Next kvp1
			If PrimaryKey <> "" Then Script &= "," & vbCrLf & PrimaryKey
			Script &= vbCrLf & ")" & vbCrLf

			' Assemble the parts of the script so far.

			Script = xx & Script
			xx = ""

			' Add the script for any indexes.

			If IndexRowCount > 0 Then
				IndexName = ""
				For jj = 1 To IndexRowCount

					' See if the index name has changed, and if so, begin a new CREATE INDEX script.

					If Panel4.Controls("I" & jj).Text <> "" And Panel4.Controls("I" & jj).Text <> IndexName Then ' This will be the first (or only) field in the index.
						If IndexName <> "" Then Script &= ");" & vbCrLf
						IndexName = Panel4.Controls("I" & jj).Text
						Script &= "CREATE "
						If CType(Panel4.Controls("Lb" & jj), CheckedListBox).GetItemChecked(1) Then Script &= "UNIQUE "
						If CType(Panel4.Controls("Lb" & jj), CheckedListBox).GetItemChecked(0) Then Script &= "CLUSTERED " Else Script &= "NONCLUSTERED "
						Script &= "INDEX "
						Script &= IndexName & " ON [" & TextBox1.Text & "] (" & Panel4.Controls("F" & jj).Text
						If CType(Panel4.Controls("C" & jj), ComboBox).SelectedIndex = 0 Then Script &= " ASC" Else Script &= " DESC"
					Else ' These will be second and subsequent fields in the index.
						Script &= "," & Panel4.Controls("F" & jj).Text
						If CType(Panel4.Controls("C" & jj), ComboBox).SelectedIndex = 0 Then Script &= " ASC" Else Script &= " DESC"
					End If
				Next jj
				Script &= ");" & vbCrLf
			End If

			'******************************************************************************************

			' MODIFY TABLE
			'If we are modifying an existing table/index/foreign key, assemble the script.

			'******************************************************************************************
		Else

			' Get all the changes made to the table: additions, changes and deletions.
			' If any errors were detected while compiling changes, return no script.

			If Not GetChanges() Then Return ""

			' If the table name has changed, we need to create the script to rename it first.

			Script = ""
			If CurrentTableName <> TextBox1.Text And TextBox1.Text <> "" Then
				Script &= "EXEC sp_rename '[dbo]." & CurrentTableName & "','" & TextBox1.Text & "';" & vbCrLf
			End If

			' If we have any fields that have been renamed, we need to create the script to
			' do that next.

			If FieldNamesChanged.Count > 0 Then
				For Each kvp2 In FieldNamesChanged
					Script &= "EXEC sp_rename '" & TextBox1.Text & "." & TableDefinition(kvp2.Key).Field_Name & "','" & FieldNamesChanged(kvp2.Key) & "', 'COLUMN';" & vbCrLf
				Next kvp2
			End If

			' See if there are any indexes that have to be dropped before we drop a column.

			If FieldsDropped.Count > 0 Then
				For Each kvp2 In FieldsDropped
					CL = GetIndexes(TextBox1.Text, kvp2.Value)
					For ii = 1 To CL.Count
						Script &= "DROP INDEX " & CL(ii) & " ON [" & TextBox1.Text & "];" & vbCrLf
					Next ii
				Next kvp2
			End If

			' See if there were any indexes explicitly dropped.

			If IndexesDropped.Count > 0 Then
				For Each kvp4 In IndexesDropped
					Script &= "DROP INDEX " & kvp4.Value & " ON [" & TextBox1.Text & "];" & vbCrLf
				Next kvp4
			End If

			' If any fields need to be dropped, get any constraints attached to the fields to be dropped.

			If FieldsDropped.Count > 0 Then
				zx = ""
				For Each kvp2 In FieldsDropped
					CL = GetConstraints(TextBox1.Text, FieldsDropped(kvp2.Key))
					For ii = 1 To CL.Count
						xx = "       [" & CL(ii) & "]"
						If ii < CL.Count Then xx &= ","
						zx &= xx & vbCrLf
					Next ii
					If zx <> "" Then
						Script &= "ALTER TABLE [dbo].[" & TextBox1.Text & "]" & vbCrLf
						Script &= "   DROP" & vbCrLf & "      CONSTRAINT" & vbCrLf
						Script &= zx
					End If
				Next kvp2

				' Create script for any dropped fields.

				Script &= "ALTER TABLE [dbo].[" & TextBox1.Text & "]" & vbCrLf
				Script &= "   DROP" & vbCrLf
				zx = ""
				For Each kvp2 In FieldsDropped
					If zx <> "" Then zx &= "," & vbCrLf ' Comma after all but last field
					xx = "      COLUMN [" & FieldsDropped(kvp2.Key) & "]"
					zx &= xx
				Next kvp2
				Script &= zx & vbCrLf & ";" & vbCrLf
			End If

			' If the PRIMARY KEY, DEFAULT or UNIQUE attribute was removed from any field, create the
			' script to drop those constraints.

			For Each kvp1 In FieldsModified
				If kvp1.Value.Is_Primary_Key <> TableDefinition(kvp1.Key).Is_Primary_Key And TableDefinition(kvp1.Key).Is_Primary_Key Then
					Script &= "ALTER TABLE [dbo].[" & TextBox1.Text & "]" & vbCrLf
					Script &= "   DROP" & vbCrLf
					CL = GetConstraints(TextBox1.Text, kvp1.Value.Field_Name, 2) ' 2 retrieves primary key constraints
					For ii = 1 To CL.Count
						xx = "      CONSTRAINT [" & CL(ii) & "];"
						If ii < CL.Count Then xx += ","
						Script &= xx & vbCrLf
					Next ii
					Script &= ";" & vbCrLf
				End If

				' If a default value changed, create the script to drop the constraints.

				If (TableDefinition(kvp1.Key).Has_Default And Not kvp1.Value.Has_Default) Or (TableDefinition(kvp1.Key).Has_Default And kvp1.Value.Default_Value <> TableDefinition(kvp1.Key).Default_Value) Then
					Script &= "ALTER TABLE [dbo].[" & TextBox1.Text & "]" & vbCrLf
					Script &= "   DROP" & vbCrLf
					CL = GetConstraints(TextBox1.Text, TableDefinition(kvp1.Key).Field_Name, 1) ' 2 retrieves default constraints
					For ii = 1 To CL.Count
						xx = "      CONSTRAINT [" & CL(ii) & "];"
						If ii < CL.Count Then xx += ","
						Script &= xx & vbCrLf
					Next ii
					Script &= ";" & vbCrLf
				End If

				' If a unique attribute was dropped, create the script to drop the constraints.

				If TableDefinition(kvp1.Key).Is_Unique And Not kvp1.Value.Is_Unique Then
					Script &= "ALTER TABLE [dbo].[" & TextBox1.Text & "]" & vbCrLf
					Script &= "   DROP" & vbCrLf
					CL = GetConstraints(TextBox1.Text, TableDefinition(kvp1.Key).Field_Name, 3) ' 3 retrieves unique constraints
					For ii = 1 To CL.Count
						xx = "      CONSTRAINT [" & CL(ii) & "];"
						If ii < CL.Count Then xx += ","
						Script &= xx & vbCrLf
					Next ii
					Script &= ";" & vbCrLf
				End If

			Next kvp1

			' Now we can start the script to alter the table itself.
			' First create script for any added fields.

			If FieldsAdded.Count > 0 Then
				Script &= "ALTER TABLE [dbo].[" & TextBox1.Text & "]" & vbCrLf
				Script &= "   ADD" & vbCrLf
				PrimaryKey = ""
				xx = ""
				For Each kvp1 In FieldsAdded

					' Assemble the line.

					If xx <> "" Then xx &= "," & vbCrLf
					xx &= "     [" & kvp1.Value.Field_Name & "]  "
					xx &= kvp1.Value.Data_Type & " "
					If kvp1.Value.Field_Length <> "" Then xx &= "(" & kvp1.Value.Field_Length & ")" ' Field length must be in parentheses
					If kvp1.Value.Is_Primary_Key And PrimaryKey = "" Then PrimaryKey = "     PRIMARY KEY CLUSTERED ([" & kvp1.Value.Field_Name & "] ASC)"
					If kvp1.Value.Is_Nullable Then xx &= " NULL" Else xx &= " NOT NULL"
					If kvp1.Value.Has_Default Then xx &= " DEFAULT ((" & kvp1.Value.Default_Value & "))"
					If kvp1.Value.Is_Unique Then xx &= " UNIQUE"
					If kvp1.Value.Is_Identity Then xx &= " IDENTITY " & kvp1.Value.Default_Value
				Next kvp1
				Script &= xx & vbCrLf
				If PrimaryKey <> "" Then Script &= PrimaryKey
				Script &= ";" & vbCrLf
			End If

			' If data type, field length, NULL or default value was modified, create the script to change the field

			If FieldsModified.Count > 0 Then
				For Each kvp1 In FieldsModified
					If TableDefinition(kvp1.Key).Data_Type <> kvp1.Value.Data_Type Or TableDefinition(kvp1.Key).Field_Length <> kvp1.Value.Field_Length Or TableDefinition(kvp1.Key).Is_Nullable <> kvp1.Value.Is_Nullable Then

						' Get the number of the row affected.

						ii = Val(kvp1.Key)

						' Now we can change any fields whose data type or length has changed.  However, we cannot
						' restore a default value when altering a column. That will be done separaately.

						Script &= "ALTER TABLE [dbo].[" & TextBox1.Text & "]" & vbCrLf
						Script &= "   ALTER COLUMN [" & FieldsModified(ii).Field_Name & "] " & FieldsModified(ii).Data_Type
						If FieldsModified(ii).Field_Length <> "" Then Script &= " (" & FieldsModified(ii).Field_Length & ")"
						If Not FieldsModified(ii).Is_Nullable Then Script &= " NOT NULL" Else Script &= " NULL"
						If FieldsModified.Count > 2 Then Script &= ";" & vbCrLf
						Script &= vbCrLf
					End If

					' If a unique attribute was added, create the script to add that now.

					If kvp1.Value.Is_Unique Then
						Script &= "ALTER TABLE [dbo].[" & TextBox1.Text & "]" & vbCrLf & "   ADD UNIQUE NONCLUSTERED ([" & kvp1.Value.Field_Name & "] ASC);" & vbCrLf
					End If

					' If a default value was added, create the script to add that now.
					' The DefaultAdded collection will contain "field_name,rownum",
					' so we have to parse it.  Then, we retrieve the field's default
					' value from the "Dx" textbox, where x=the field row number.

					If (Not TableDefinition(kvp1.Key).Has_Default And kvp1.Value.Has_Default) Or (TableDefinition(kvp1.Key).Has_Default And kvp1.Value.Has_Default And TableDefinition(kvp1.Key).Default_Value <> kvp1.Value.Default_Value) Then
						zx = kvp1.Value.Default_Value
						If Not IsNumericDefault(zx) Then zx = "'" & SRep(SRep(zx, 1, """", ""), 1, "'", "") & "'"
						Script &= "ALTER TABLE [dbo].[" & TextBox1.Text & "]" & vbCrLf & "   ADD DEFAULT ((" & zx & ")) FOR [" & kvp1.Value.Field_Name & "];" & vbCrLf
					End If

					' If a primary key was added, create the script to add that now.  We only need to
					' look at element 1 of collection PKAdded, as a table may only have 1 primary key.

					If kvp1.Value.Is_Primary_Key Then
						Script &= "ALTER TABLE [dbo].[" & TextBox1.Text & "]" & vbCrLf & "    ADD PRIMARY KEY CLUSTERED ([" & kvp1.Value.Field_Name & "] ASC);" & vbCrLf
					End If

					' If the Unicode collation was selected, create the script to change that collation now.

					If kvp1.Value.Is_Unicode Then
						Script &= "ALTER TABLE [dbo].[" & TextBox1.Text & "]" & vbCrLf & "    ALTER COLUMN [" & kvp1.Value.Field_Name & "]" & kvp1.Value.Data_Type
						If kvp1.Value.Field_Length <> "" Then Script &= " (" & kvp1.Value.Field_Length & ")"
						Script &= " COLLATE Latin1_General_100_CI_AI_SC_UTF8"
						If kvp1.Value.Is_Nullable Then Script &= " NULL;" & vbCrLf Else Script &= " NOT NULL;" & vbCrLf
					End If

				Next kvp1
			End If

			' If any indexes were added or modified, recreate them now.

			If IndexesAdded.Count > 0 Then
				For Each kvp4 In IndexesAdded
					For jj = 1 To IndexRowCount
						i = GetIndexValues(jj)
						If i.Index_Name = kvp4.Value Then

							' See if the field is the primary field of the index, and if so,
							'  begin a new CREATE INDEX script.

							If i.Is_Primary_Field Then ' This will be the first (or only) field in the index.
								IndexName = i.Index_Name
								Script &= "CREATE "
								If i.Is_Unique Then Script &= "UNIQUE "
								If i.Is_Clustered Then Script &= "CLUSTERED " Else Script &= "NONCLUSTERED "
								Script &= "INDEX "
								Script &= i.Index_Name & " ON [" & TextBox1.Text & "] (" & i.Field_Name
								If i.Sort = 0 Then Script &= " ASC" Else Script &= " DESC"
							Else ' These will be second and subsequent fields in the index.
								Script &= "," & i.Field_Name
								If i.Sort = 0 Then Script &= " ASC" Else Script &= " DESC"
							End If
						End If
					Next jj
					Script &= ");" & vbCrLf
				Next kvp4
			End If
		End If


		' Return the assembled script

		Return Script

	End Function
	'**********************************************************************

	' Function to return a collection of index names that may
	' be attached to a field.

	'**********************************************************************
	Private Function GetIndexes(ByVal TableName As String, ByVal FieldName As String) As Collection

		' Declare variables

		Dim TableID As Integer
		Dim ColumnID As Integer
		Dim IndexID As Integer
		Dim IList As New Collection
		Dim Cmd As SqlCommand
		Dim DA As New SqlDataAdapter
		Dim DS As New DataSet
		Dim DS2 As New DataSet
		Dim DT1 As DataTable
		Dim DT2 As DataTable

		' Get the ID of the table, and the column ID of the column to be dropped.

		Cmd = New SqlCommand("SELECT sys.tables.object_id,sys.columns.column_id FROM sys.tables INNER JOIN sys.columns ON sys.tables.object_id=sys.columns.object_id WHERE sys.tables.name='" & TableName & "' AND sys.columns.name='" & FieldName & "'", DB)
		DA.SelectCommand = Cmd
		DA.Fill(DS, "Table")
		DT1 = DS.Tables("Table")
		If DT1.Rows.Count > 0 Then
			TableID = DT1.Rows(0)("object_id")
			ColumnID = DT1.Rows(0)("column_id")
			DS.Clear()

			' Now get any indexes attached to the field.

			DA.SelectCommand.CommandText = "SELECT index_id FROM sys.index_columns WHERE Object_id=" & TableID & " AND column_id=" & ColumnID
			DA.Fill(DS, "Table")
			DT1 = DS.Tables("Table")

			' If any indexes were found, look up each index.  We only return indexes that are not primary key
			' or unique value indexes.

			If DT1.Rows.Count > 0 Then
				For jj = 0 To DT1.Rows.Count - 1
					IndexID = DT1.Rows(jj)("index_id")
					DS2.Clear()
					DA.SelectCommand.CommandText = "SELECT * FROM sys.indexes WHERE Object_id=" & TableID & " AND index_id=" & IndexID
					DA.Fill(DS2, "Table")
					DT2 = DS2.Tables("Table")
					For ii = 0 To DT2.Rows.Count - 1
						If Not DT2.Rows(ii)("is_unique_constraint") And Not DT2.Rows(ii)("is_primary_key") Then IList.Add(DT2.Rows(ii)("name"))
					Next ii
				Next jj
			End If
			DS.Clear()
		End If


		' Dispose of objects we've created.

		DA.Dispose()
		DS.Dispose()
		DS2.Dispose()
		Cmd.Dispose()

		' Return the list of constraints attached to the field.

		Return IList

	End Function
	'**********************************************************************

	' Sub to check for changes to a table.

	'**********************************************************************
	Private Function GetChanges() As Boolean

		' Declare variables

		Dim RowModified As Boolean
		Dim RowDeleted As Boolean
		Dim IndexDeleted As Boolean
		Dim jj As Integer
		Dim RowNum As Integer
		Dim zx As String
		Dim IndexName As String
		Dim f As TableRowValues
		Dim kvp As KeyValuePair(Of Integer, TableRowValues)
		Dim kvp2 As KeyValuePair(Of String, String)
		Dim t As TextBox

		' If there is no current table, all the fields are new fields, so they'll be added
		' the the FieldsAdded collection.

		If TableDefinition.Count = 0 Then
			FieldsAdded.Clear()
			For jj = 1 To FieldRowCount

				' Get the values from one row.

				f = GetRowValues(jj)

				' Add the values for the row to the table definition.  We index the fields
				' only by row number, so we can later determine if a field is renamed.

				FieldsAdded.Add(jj, f)
			Next jj

			' If we have an existing table definition, we need to determine what rows have been
			' added, deleted, or modified.

		Else

			' Clear out the previous collections.

			FieldsAdded.Clear()
			FieldsDropped.Clear()
			FieldsModified.Clear()
			FieldNamesChanged.Clear()
			IndexesAdded.Clear()
			IndexesDropped.Clear()

			' First check for fields that have been modified.

			For jj = 1 To FieldRowCount
				RowNum = Panel2.Controls("F" & jj).Tag ' This contains the original row number
				RowModified = False
				If TableDefinition.ContainsKey(RowNum) Then

					' Get the current values of the row, to compare against the original values.

					f = GetRowValues(jj)

					' Make sure we got values for the row.  If the row was deleted, we'll have "Nothing".

					If Not IsNothing(f.Field_Name) Then
						If f.Data_Type <> TableDefinition(RowNum).Data_Type Then RowModified = True
						If f.Field_Length <> "MAX" And f.Field_Length <> TableDefinition(RowNum).Field_Length Then RowModified = True
						If f.Default_Value <> TableDefinition(RowNum).Default_Value Then RowModified = True
						If f.Has_Default <> TableDefinition(RowNum).Has_Default Then RowModified = True
						If f.Is_Primary_Key <> TableDefinition(RowNum).Is_Primary_Key Then RowModified = True
						If f.Is_Nullable <> TableDefinition(RowNum).Is_Nullable Then RowModified = True
						If f.Is_Unique <> TableDefinition(RowNum).Is_Unique Then RowModified = True
						If f.Is_Unicode <> TableDefinition(RowNum).Is_Unicode Then RowModified = True

						' Now see if the name only has been changed.

						If f.Field_Name <> TableDefinition(RowNum).Field_Name Then FieldNamesChanged.Add(RowNum, f.Field_Name)
						If RowModified Then FieldsModified.Add(RowNum, f)

						' If the row no longer exists, add it to the FieldsDropped collection.

					Else
						FieldsDropped.Add(RowNum, TableDefinition(RowNum).Field_Name)
					End If

					' If a row does not exist in the TableDefinition collection, add it to the 
					' list of fields added.

				Else

					' Next check for new fields that have been added.

					f = GetRowValues(jj)
					FieldsAdded.Add(RowNum, f)
				End If
			Next jj

			' Now check for fields that have been dropped.  If the name of the field in the same row
			' number matches, then the field was certainly not dropped.  If it seems to have been dropped,
			' because the name doesn't match, see if the row number is in the rowsmodified collection (it may have been renamed).
			' If neither situation is true, add the field to the list of fields to drop.

			For Each kvp In TableDefinition
				RowDeleted = True
				For jj = 1 To FieldRowCount
					If Panel2.Controls("F" & jj).Tag = kvp.Key Then
						RowDeleted = False
						Exit For
					End If
				Next jj
				If RowDeleted And Not FieldsModified.ContainsKey(kvp.Key) Then FieldsDropped.Add(kvp.Key, kvp.Value.Field_Name)
			Next kvp

			' Now check indexes.  Loop through all the controls in the indexes panel.  If any change
			' has been made to an index, it must be dropped and re-created.  We will also do this
			' if an index is renamed, even though one can rename an index, simply to avoid complications
			' generating the SQL script to effect changes.

			IndexName = ""
			zx = ""
			If IndexRowCount > 0 Then
				For jj = 1 To IndexRowCount

					' Get the index name field text control as a text box.

					t = Panel4.Controls("I" & jj) ' Get control as text box.

					' If the index name has changed, save it.

					If t.Text <> IndexName And IndexName <> "" And t.Text <> "" Then
						If Not IndexesDefinition.ContainsKey(IndexName) Then
							IndexesAdded.Add(IndexName, IndexName)
						ElseIf IndexesDefinition(IndexName) <> zx Then
							IndexesAdded.Add(IndexName, IndexName)
							IndexesDropped.Add(IndexName, IndexName)
						End If
						zx = ""
					End If

					' Save the index name

					If t.Text <> "" Then IndexName = t.Text

					' Assemble the list of fields for the index.

					If zx <> "" Then zx &= ","
					zx &= Panel4.Controls("F" & jj).Text
					If CType(Panel4.Controls("C" & jj), ComboBox).SelectedIndex = 0 Then zx &= " ASC" Else zx &= " DESC"
				Next jj

				' Save the last index information.

				If Not IndexesDefinition.ContainsKey(IndexName) Then
					IndexesAdded.Add(IndexName, IndexName)
				ElseIf IndexesDefinition(IndexName) <> zx Then
					IndexesAdded.Add(IndexName, IndexName)
					IndexesDropped.Add(IndexName, IndexName)
				End If
			End If

			' Check for indexes that have been dropped.

			For Each kvp2 In IndexesDefinition
				IndexDeleted = True
				For jj = 1 To IndexRowCount
					If Panel4.Controls("I" & jj).Text = kvp2.Key Then
						IndexDeleted = False
						Exit For
					End If
				Next jj

				' If the index was deleted, add it to the list of dropped indexes.

				If IndexDeleted Then IndexesDropped.Add(kvp2.Key, kvp2.Key)
			Next kvp2

		End If

		' If the table name has been changed, and anything else as well, alert the user that
		' the changes cannot be made simultaneously.

		If CurrentTableName <> "" And CurrentTableName <> "New Table Name" And CurrentTableName <> TextBox1.Text And (FieldsAdded.Count > 0 Or FieldsDropped.Count > 0 Or FieldsModified.Count > 0 Or IndexesAdded.Count > 0 Or IndexesDropped.Count > 0) Then
			MsgBox("The table name cannot be changed at the same time as other table changes are made.  Change the table name first, then upload the change and make other changes next.", MsgBoxStyle.Information, "Rename Table Failed")
			Return False
		End If

		Return True

	End Function
	'**********************************************************************

	' Sub to get the values from a specified row into a TableRowValues variable.

	'**********************************************************************
	Private Function GetRowValues(RowNum As Integer) As TableRowValues

		' Declare variables

		Dim f As New TableRowValues
		Dim lb As CheckedListBox

		' Make sure the specified row exists.  It may have been deleted, in which case
		' this function returns "nothing".

		If IsNothing(Panel2.Controls("F" & RowNum)) Then Return Nothing

		' Get the data from the specified row into the structure.

		f.Field_Name = Panel2.Controls("F" & RowNum).Text
		f.Data_Type = Panel2.Controls("C" & RowNum).Text
		f.Field_Length = Panel2.Controls("L" & RowNum).Text
		lb = CType(Panel2.Controls("Lb" & RowNum), CheckedListBox)
		f.Is_Primary_Key = lb.GetItemChecked(0)
		f.Is_Nullable = lb.GetItemChecked(1)
		f.Is_Identity = lb.GetItemChecked(2)
		f.Has_Default = lb.GetItemChecked(3)
		f.Is_Unique = lb.GetItemChecked(4)
		f.Is_Unicode = lb.GetItemChecked(5)
		If f.Is_Identity Then f.Default_Value = Panel2.Controls("D" & RowNum).Text
		If f.Has_Default Then f.Default_Value = Panel2.Controls("D" & RowNum).Text

		' Return the assembled information structure

		Return f

	End Function
	'**********************************************************************

	' Sub to get the values from a specified row into an IndexRowValues variable.

	'**********************************************************************
	Private Function GetIndexValues(RowNum As Integer) As IndexRowValues

		' Declare variables

		Dim jj As Integer
		Dim f As IndexRowValues = Nothing
		Dim lb As CheckedListBox

		' Make sure the specified row exists.  It may have been deleted, in which case
		' this function returns "nothing".

		If IsNothing(Panel4.Controls("I" & RowNum)) Then Return Nothing


		' Get the data from the specified row into the structure.

		f.Index_Name = Panel4.Controls("I" & RowNum).Text
		f.Field_Name = Panel4.Controls("F" & RowNum).Text
		If VB.Left(Panel4.Controls("C" & RowNum).Text, 1) = "A" Then f.Sort = 0 Else f.Sort = 1
		lb = CType(Panel4.Controls("Lb" & RowNum), CheckedListBox)
		f.Is_Clustered = lb.GetItemChecked(0)
		f.Is_Unique = lb.GetItemChecked(1)
		If f.Index_Name <> "" Then f.Is_Primary_Field = True ' The field with an index name is the first field.

		' If the index name is blank, work our way up the rows until we find the index name,
		' and add it to the structure.

		If f.Index_Name = "" Then
			For jj = RowNum - 1 To 1 Step -1
				If Not Panel4.Controls("I" & jj) Is Nothing AndAlso Panel4.Controls("I" & jj).Text <> "" Then
					f.Index_Name = Panel4.Controls("I" & jj).Text
					Exit For
				End If
			Next jj
		End If

		' Return the assembled information structure

		Return f

	End Function
	'**********************************************************************

	' The SQLScript property is the SQL script we've either build, or
	' are editing.

	'**********************************************************************
	Public Property SQLScript As String

		' Return the SQL Script that defines the CREATE TABLE process
		' for the table we're working on, whether altered or not.

		Get

			SQLScript = AssembleSQLScript(TableEditModeEnum.TableCreate)

		End Get

		' If an SQL Script is assigned to this property, parse the text and
		' fill in the current table design so that it may be edited if desired.

		Set(value As String)

			' Declare variables

			Dim zx As String

			' First make sure the script is a script to "CREATE TABLE".  No
			' other script can be processed here.

			zx = LTrim(value)
			If UCase(VB.Left(zx, 12)) <> "CREATE TABLE" And UCase(VB.Left(zx, 11)) <> "CREATETABLE" Then
				MsgBox("Script is not a table definition.", MsgBoxStyle.Exclamation, "Open SQL Script")
				Exit Property
			End If

			' Clear out all the panels and re-initialize them.

			Panel2.Controls.Clear()
			Panel4.Controls.Clear()
			InitializePanels()

			' Parse the SQL script passed to the property.

			ParseScript(zx)

			' Set the table edit mode to "create table"

			TableEditMode = TableEditModeEnum.TableCreate

		End Set
	End Property

	'**********************************************************************

	' The Schema property is data table that contains the list of fields
	' in a table, and their data types.  Assigning the schema table to
	' this property will fill the "Table" fields with the list of all
	' fields currently defined in the table.

	'**********************************************************************
	Public WriteOnly Property Schema As DataTable
		Set(SchemaTable As DataTable)

			' Declare variables

			Dim ii As Integer
			Dim TableID As Integer
			Dim zx As String
			Dim FieldDefinition As New TableRowValues
			Dim IndexFields As New IndexRowValues
			Dim DA As New SqlDataAdapter
			Dim DS1 As New DataSet
			Dim DT1 As DataTable
			Dim DS2 As New DataSet
			Dim DT2 As New DataTable
			Dim Cmd As SqlCommand

			' Clear out all the panels.

			Panel2.Controls.Clear()
			Panel4.Controls.Clear()

			' Initialize all the panels: tables, indexes and relationships

			InitializePanels()

			' Set the mode to alter table mode.

			TableEditMode = TableEditModeEnum.TableAlter

			' Put the name of the table into the table name text field.

			TextBox1.Text = SchemaTable.Rows(0)("TABLE_NAME")

			' Remember the table name, in case it changes.

			CurrentTableName = TextBox1.Text

			' Get the object_id of the table

			Cmd = New SqlCommand("SELECT object_id FROM sys.tables WHERE name='" & SchemaTable.Rows(0)("TABLE_NAME") & "'", DB)
			DA.SelectCommand = Cmd
			DA.Fill(DS1, "Table")
			DT1 = DS1.Tables("Table")
			TableID = DT1.Rows(0)("object_id")

			' The columns of the schema table have the following names:
			' TABLE_NAME,COLUMN_NAME,COLUMN_DEFAULT,IS_NULLABLE,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH,NUMERIC_PRECISION,DATETIME_PRECISION

			' Add a row to the tables panel for each field in the table schema.

			For ii = 0 To SchemaTable.Rows.Count - 1
				FieldDefinition = New TableRowValues
				FieldDefinition.Field_Name = SchemaTable.Rows(ii)("COLUMN_NAME")
				FieldDefinition.Data_Type = UCase(SchemaTable.Rows(ii)("DATA_TYPE"))
				If Val(GetR(SchemaTable.Rows(ii), "CHARACTER_MAXIMUM_LENGTH")) > 0 Then
					FieldDefinition.Field_Length = CStr(GetR(SchemaTable.Rows(ii), "CHARACTER_MAXIMUM_LENGTH"))
				Else
					If Val(GetR(SchemaTable.Rows(ii), "CHARACTER_MAXIMUM_LENGTH")) = -1 Then FieldDefinition.Field_Length = "-1"
				End If
				If GetR(SchemaTable.Rows(ii), "IS_NULLABLE") = "YES" Then FieldDefinition.Is_Nullable = True Else FieldDefinition.Is_Nullable = False
				FieldDefinition.Default_Value = SRep(SRep(GetR(SchemaTable.Rows(ii), "COLUMN_DEFAULT"), 1, ")", ""), 1, "(", "")
				If FieldDefinition.Default_Value <> "" Then FieldDefinition.Has_Default = True Else FieldDefinition.Has_Default = False

				' look to see if the column is an identity.

				FieldDefinition.Is_Identity = False
				DA.SelectCommand.CommandText = "SELECT * FROM sys.identity_columns WHERE object_id=" & TableID & " AND column_id=" & GetR(SchemaTable.Rows(ii), "ORDINAL_POSITION")
				DS2.Clear()
				DA.Fill(DS2, "Table")
				DT2 = DS2.Tables("Table")
				If DT2.Rows.Count > 0 Then
					FieldDefinition.Is_Identity = True
					FieldDefinition.Default_Value = "(" & DT2.Rows(0)("SEED_VALUE") & "," & DT2.Rows(0)("INCREMENT_VALUE") & ")"
				End If

				' Look to see if the column has a UNIQUE constraint.

				FieldDefinition.Is_Unique = False
				DA.SelectCommand.CommandText = "SELECT sys.indexes.name,sys.indexes.is_unique_constraint FROM sys.indexes INNER JOIN sys.index_columns ON sys.indexes.object_id=sys.index_columns.object_id AND sys.indexes.index_id=sys.index_columns.index_id WHERE sys.indexes.object_id=" & TableID & " AND sys.index_columns.column_id=" & GetR(SchemaTable.Rows(ii), "ORDINAL_POSITION") & " AND sys.indexes.is_unique_constraint=1"
				DS2.Clear()
				DA.Fill(DS2, "Table")
				DT2 = DS2.Tables("Table")
				If DT2.Rows.Count > 0 Then FieldDefinition.Is_Unique = GetR(DT2.Rows(0), "is_unique_constraint")

				' See what the column's collation sequence is.

				If GetR(SchemaTable.Rows(ii), "COLLATION_NAME") = "Latin1_General_100_CI_AI_SC_UTF8" Then FieldDefinition.Is_Unicode = True Else FieldDefinition.Is_Unicode = False

				' See if the column is a primary key.

				FieldDefinition.Is_Primary_Key = False
				DA.SelectCommand.CommandText = "SELECT DISTINCT sys.columns.name,sys.indexes.is_primary_key FROM (sys.columns LEFT JOIN sys.index_columns ON sys.columns.object_id=sys.index_columns.object_id AND sys.columns.column_id=sys.index_columns.column_id)LEFT JOIN sys.indexes ON sys.columns.object_id=sys.columns.object_id AND sys.index_columns.index_id=sys.indexes.index_id WHERE sys.columns.object_id=" & TableID & " AND sys.columns.column_id=" & GetR(SchemaTable.Rows(ii), "ORDINAL_POSITION") & " AND sys.indexes.is_primary_key=1"
				DS2.Clear()
				DA.Fill(DS2, "Table")
				DT2 = DS2.Tables("Table")
				If DT2.Rows.Count > 0 Then FieldDefinition.Is_Primary_Key = GetR(DT2.Rows(0), "is_primary_key")

				' Add the column information to the table list and to the TableDefinition collection.
				' The key will be the row number, as later that will be the only way to tell if a field
				' was only renamed, or otherwise altered.

				TableDefinition.Add(AddTableRow(FieldDefinition), FieldDefinition) ' Change from 0-based to 1-based.
			Next ii

			' Get the indexes attached to this table. Only retrieve indexes of type 2 which are not
			' unique field constraints; type 1 are primary key indexes.

			Cmd = New SqlCommand("SELECT object_id, name, index_id, type, is_unique FROM sys.indexes WHERE object_id=" & TableID & " AND type=2 AND is_unique_constraint=0", DB)
			DA.SelectCommand = Cmd
			DS1.Clear()
			DA.Fill(DS1, "Table")
			DT1 = DS1.Tables("Table")

			' Make sure there are any indexes.

			If DT1.Rows.Count > 0 Then

				' For each index, get the columns of which it is comprised and add each to the
				' indexes panel.

				For jj = 0 To DT1.Rows.Count - 1
					IndexFields = New IndexRowValues
					IndexFields.Index_Name = DT1.Rows(jj)("name")
					IndexFields.Is_Primary_Field = True
					If DT1.Rows(jj)("type") = 1 Then IndexFields.Is_Clustered = True
					If DT1.Rows(jj)("is_unique") Then IndexFields.Is_Unique = DT1.Rows(jj)("is_unique")

					' For each index, get the information for each field that comprises it.

					DA.SelectCommand.CommandText = "SELECT name,sys.index_columns.is_descending_key FROM sys.index_columns INNER JOIN sys.columns ON sys.index_columns.object_id=sys.columns.object_id AND sys.index_columns.column_id=sys.columns.column_id WHERE sys.index_columns.object_id=" & TableID & " AND sys.index_columns.index_id=" & DT1.Rows(jj)("index_id")
					DS2.Clear()
					DA.Fill(DS2, "Table")
					DT2 = DS2.Tables("Table")
					zx = ""
					For ii = 0 To DT2.Rows.Count - 1
						IndexFields.Field_Name = DT2.Rows(ii)("name")
						If DT2.Rows(ii)("is_descending_key") Then IndexFields.Sort = 1 ' 0= ASC, 1= DESC

						' Add the index row to the display.

						AddIndexRow(IndexFields)

						' Assemble the list of fields for this index.

						If zx <> "" Then zx &= ","
						zx &= DT2.Rows(ii)("name") & " " & Choose(IndexFields.Sort + 1, "ASC", "DESC")

						' For display purposes, clear out fields that belong only with the first row,
						' which specifies the index name and attributes.

						IndexFields.Is_Primary_Field = False ' All subsequent fields are false.
						IndexFields.Is_Clustered = False ' Clustered attribute only appears on line with index name.
						IndexFields.Is_Unique = False ' Unique attribute only appears on line with index name.
					Next ii

					' Add the index information to our current index configuration.

					IndexesDefinition.Add(DT1.Rows(jj)("name"), zx)
				Next jj
			End If

			' Set the initial view to "Table".

			optTable.Checked = True

			' Dispose of objects we've created.

			DA.Dispose()
			DS1.Dispose()
			DS2.Dispose()
			Cmd.Dispose()

		End Set
	End Property

End Class
