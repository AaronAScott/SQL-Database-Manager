Option Strict Off
Option Explicit On
Imports System.Data.SqlClient
Imports VB = Microsoft.VisualBasic

Public Class frmManageRelationships
	Inherits System.Windows.Forms.Form
	'**********************************************************************
	' SQL Data Manager Manage Table Relationships form
	' DM_MANAGERELATIONSHIPS.VB
	' Written: November 2018
	' Programmer: Aaron Scott
	' Copyright (C) 2017-2018 Sirius Software All Rights Reserved
	'**********************************************************************

	' Declare variables local to this module

	Private PanelInitialized As Boolean = False
	Private RelationRowCount As Integer
	Private SelectedRow As Integer
	Private LinkLabel1 As LinkLabel

	' Create some collections to keep track of changes made to a table, if we are modifying
	' an existing table.

	Private Relations As New Dictionary(Of String, ForeignKeyValues)
	Private FKAdded As New Dictionary(Of String, ForeignKeyValues)
	Private FKDropped As New Dictionary(Of String, ForeignKeyValues)

	Private Enum TableEditModeEnum
		TableCreate
		TableAlter
	End Enum
	Private TableEditMode As TableEditModeEnum = TableEditModeEnum.TableCreate

	'**********************************************************************

	' The form is loaded.

	'**********************************************************************
	Private Sub frmManageRelationships_Load(sender As Object, e As EventArgs) Handles MyBase.Load

		' Put up a message

		frmMain.StatusLabel.Text = "Loading Form.  Please Wait..."
		System.Windows.Forms.Application.DoEvents()

		' Initialize the panels

		If Not PanelInitialized Then InitializePanels()


		' Clear the message

		frmMain.StatusLabel.Text = ""
		System.Windows.Forms.Application.DoEvents()

	End Sub
	'**********************************************************************

	' The form is closed.

	'**********************************************************************
	Private Sub frmManageRelationships_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

		' Declare variables

		Dim c As Control

		' Loop through each control in the panel and remove the event handlers.

		For Each c In Panel2.Controls
			If TypeOf (c) Is TextBox Then
				RemoveHandler c.MouseDown, AddressOf TextBox_MouseDown
				RemoveHandler c.KeyDown, AddressOf TextBox_KeyDown
			End If
		Next c

		' Remove all controls

		Panel2.Controls.Clear()


	End Sub
	'**********************************************************************

	' The Scroll bar is clicked.

	'**********************************************************************
	Private Sub VScrollBar1_ValueChanged(sender As Object, e As EventArgs) Handles VScrollBar1.ValueChanged
		Panel2.Top = -VScrollBar1.Value
	End Sub
	'**********************************************************************

	' The Add New Relationship link label is clicked. This is bound
	' at run time.

	'**********************************************************************
	Private Sub LinkLabel1_Click(sender As Object, e As EventArgs)

		' Add a new, empty relationship.

		AddNewRelationship()

	End Sub

	'**********************************************************

	' Sub to handle KeyDown events for all the relation
	' text boxes.  This is bound at run time.

	'**********************************************************

	Private Sub TextBox_KeyDown(sender As Object, e As KeyEventArgs)

		' Declare variables and get the row number of the text box
		' in which the KeyDown event was triggered.

		Dim RowNum As Integer = Val(Mid(sender.name, 3))
		Dim zx As String = VB.Left(sender.name, 2)
		Dim T As TextBox

		' We only are interested in up and down arrows.

		If e.KeyCode = Keys.Up Or e.KeyCode = Keys.Down Then

			' Determine in which panel the textbox in which the key was pressed resides.

			If e.KeyCode = Keys.Down And RowNum + 1 <= RelationRowCount Then RowNum += 1
			If e.KeyCode = Keys.Up And RowNum - 1 >= 1 Then RowNum -= 1
			T = Panel2.Controls(zx & RowNum)
			T.Focus()
			T.SelectAll()

			' If the key was an up or down arrow, then indicate we've handled it here.

			e.Handled = True
		End If

	End Sub
	'**********************************************************************

	' The View in Script Editor menu option is selected.

	'**********************************************************************
	Private Sub mnuView_Click(sender As Object, e As EventArgs) Handles mnuViewScript.Click

		' Declare variables

		Dim zx As String

		' Get any changes made to the relationships.

		GetChanges()

		' Get the script to effect any changes.

		zx = SQLScript()

		' Open a new editor window with the assembled script.

		Dim se As New frmScriptEditor
		se.MdiParent = frmMain
		se.DefaultText = zx
		se.mnuExecute.Enabled = False
		se.Show()

	End Sub

	'**********************************************************************

	' The Upload to Database menu option is selected.

	'**********************************************************************
	Private Sub mnuUpload_Click(sender As Object, e As EventArgs) Handles mnuUpload.Click

		' Declare variables.

		Dim zx As String

		' Get any changes made to the relationships.

		GetChanges()

		' Get the script to effect any changes.

		zx = SQLScript()

		Dim Cmd As SqlCommand = New SqlCommand(zx, DB)

		' Execute the query and watch for errors.

		Try
			Cmd.ExecuteNonQuery()
			StatusLabel1.Text = "Script executed successfully"

			' Clear all the collections that recorded the changes.

			ClearCollections()

		Catch ex As Exception

			StatusLabel1.Text = "Script failed"
			MsgBox(ex.Message)

		End Try

		' Enable the timer.  This will cause the messsage to stay on the status label
		' for 5 seconds, then be cleared.

		Timer1.Enabled = True

		' Look for open windows that might be affected by changes made here,
		' and tell themm to refresh themselves.

		For Each f In My.Application.OpenForms
			If TypeOf (f) Is frmObjectList Then CType(f, frmObjectList).RefreshObjectList()
			If TypeOf (f) Is frmRelationships Then CType(f, frmRelationships).RefreshDisplay()
		Next f

	End Sub

	'**********************************************************************

	' The timer has clicked.  Erase any current status message.

	'**********************************************************************
	Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

		StatusLabel1.Text = ""
		Timer1.Enabled = False

	End Sub
	'**********************************************************************

	' Event handler for a textbox MouseDown event. This is bound at
	' run time.

	'**********************************************************************
	Private Sub TextBox_MouseDown(sender As Object, e As MouseEventArgs)

		' Declare variables

		Dim RowNum As Integer

		' Make sure the right mouse button is selected

		If e.Button = Windows.Forms.MouseButtons.Right Then

			' Get the row number of the text box.

			RowNum = Val(Mid(sender.Name, 3))

			' Save the row number as our selected row.

			SelectedRow = RowNum

		End If
	End Sub
	'**********************************************************************

	' The Delete Row menu option is selected.

	'**********************************************************************
	Private Sub mnuDelete_Click(sender As Object, e As EventArgs) Handles mnuDelete.Click

		Dim jj As Integer
		Dim RowNum As Integer

		' Get the current row number

		RowNum = SelectedRow

		' If the row is not already listed to be dropped, add it to the list of
		' relations to be dropped.

		For jj = RowNum To RelationRowCount - 1
			Panel2.Controls("Rn" & jj).Text = Panel2.Controls("Rn" & (jj + 1)).Text
			Panel2.Controls("Pt" & jj).Text = Panel2.Controls("Pt" & (jj + 1)).Text
			Panel2.Controls("Pf" & jj).Text = Panel2.Controls("Pf" & (jj + 1)).Text
			Panel2.Controls("Rt" & jj).Text = Panel2.Controls("Rt" & (jj + 1)).Text
			Panel2.Controls("Rf" & jj).Text = Panel2.Controls("Rf" & (jj + 1)).Text
			CType(Panel2.Controls("Lb" & jj), CheckedListBox).SetItemChecked(0, CType(Panel2.Controls("Lb" & (jj + 1)), CheckedListBox).GetItemChecked(0))
			CType(Panel2.Controls("Lb" & jj), CheckedListBox).SetItemChecked(1, CType(Panel2.Controls("Lb" & (jj + 1)), CheckedListBox).GetItemChecked(1))
		Next jj
		Panel2.Controls("Rn" & RelationRowCount).Dispose()
		Panel2.Controls("Pt" & RelationRowCount).Dispose()
		Panel2.Controls("Pf" & RelationRowCount).Dispose()
		Panel2.Controls("Rt" & RelationRowCount).Dispose()
		Panel2.Controls("Rf" & RelationRowCount).Dispose()
		Panel2.Controls("Lb" & RelationRowCount).Dispose()
		LinkLabel1.Top = LinkLabel1.Top - 22 ' Height of a text box.
		RelationRowCount -= 1

	End Sub
	'**********************************************************************

	' Sub to initialize the Relationship builder panels

	'**********************************************************************
	Private Sub InitializePanels()

		' Declare variables

		Dim RowTop As Integer = 5

		' Labels for the columns

		Dim L1 As Label
		Dim L2 As Label
		Dim L3 As Label
		Dim L4 As Label
		Dim L5 As Label
		Dim L6 As Label

		' Create a font for the labels

		Dim f As New Font("Arial", 8, FontStyle.Bold)

		' Clear out all the controls that may have been previously created (if we have left and are re-entering
		' this form.)

		Panel2.Controls.Clear()

		' Add the primary link label, "Add New Relationship", which serves as the starting point for locating
		' all the other controls we add, to each of the two panels.

		LinkLabel1 = New LinkLabel
		LinkLabel1.Name = "LinkLabel1"
		LinkLabel1.Top = 40
		LinkLabel1.Left = 125
		LinkLabel1.Text = "Add New Relation"
		LinkLabel1.LinkColor = Color.ForestGreen
		Panel2.Controls.Add(LinkLabel1)
		AddHandler LinkLabel1.Click, AddressOf LinkLabel1_Click

		' Create labels to provide columns for the relationships panel

		L1 = New Label
		L1.Name = "LR"
		L1.Top = RowTop
		L1.Left = 5
		L1.Width = 120
		L1.TextAlign = ContentAlignment.MiddleLeft
		L1.Font = f
		L1.Text = "Relationship Name"
		L1.BackColor = SystemColors.ControlDark

		L2 = New Label
		L2.Top = RowTop
		L2.Left = 5 + L1.Width
		L2.Width = 120
		L2.TextAlign = ContentAlignment.MiddleLeft
		L2.Font = f
		L2.Text = "Primary Table"
		L2.BackColor = SystemColors.ControlDark

		L3 = New Label
		L3.Top = RowTop
		L3.Left = 5 + L1.Width + L2.Width
		L3.Width = 120
		L3.TextAlign = ContentAlignment.MiddleLeft
		L3.Font = f
		L3.Text = "Primary Field"
		L3.BackColor = SystemColors.ControlDark

		L4 = New Label
		L4.Top = RowTop
		L4.Left = 5 + L1.Width + L2.Width + L3.Width
		L4.Width = 120
		L4.TextAlign = ContentAlignment.MiddleLeft
		L4.Font = f
		L4.Text = "Related Table"
		L4.BackColor = SystemColors.ControlDark

		L5 = New Label
		L5.Top = RowTop
		L5.Left = 5 + L1.Width + L2.Width + L3.Width + L4.Width
		L5.Width = 125
		L5.TextAlign = ContentAlignment.MiddleLeft
		L5.Font = f
		L5.Text = "Related Field"
		L5.BackColor = SystemColors.ControlDark

		L6 = New Label
		L6.Top = RowTop
		L6.Left = 5 + L1.Width + L2.Width + L3.Width + L4.Width + L5.Width
		L6.Width = 130
		L6.TextAlign = ContentAlignment.MiddleLeft
		L6.Font = f
		L6.Text = "Relationship"
		L6.BackColor = SystemColors.ControlDark

		' Add the labels to the relationships panel controls collection

		Panel2.Controls.Add(L1)
		Panel2.Controls.Add(L2)
		Panel2.Controls.Add(L3)
		Panel2.Controls.Add(L4)
		Panel2.Controls.Add(L4)
		Panel2.Controls.Add(L5)
		Panel2.Controls.Add(L6)
		RowTop = RowTop + L1.Height + 1

		' Initialize the count of rows.

		RelationRowCount = 0

		' Move the pre-existing link label down  below the last row.

		LinkLabel1.Top = RowTop
		LinkLabel1.Left = 5


		' Set the maximum value of the scroll bar control to the size of the inner panel.

		VScrollBar1.Maximum = Panel1.Height
		VScrollBar1.Value = 0

		' Indicate we've done this process

		PanelInitialized = True

		' Dispose of items we created.

		f.Dispose()

	End Sub
	'**********************************************************************

	' Routine to add a new row to the relationships list.  This takes an
	' argument of type ForeignKeyValues, but is declared as object, so it
	' will receive a value of "nothing" if no argument is passed.

	'**********************************************************************
	Private Sub AddNewRelationship(Optional ByVal ForeignKey As Object = Nothing)

		' Declare variables

		Dim RowTop As Integer
		Dim T1 As TextBox
		Dim T2 As TextBox
		Dim T3 As TextBox
		Dim T4 As TextBox
		Dim T5 As TextBox
		Dim Lb As CheckedListBox
		Dim f As New Font("Microsoft Sans Serif", 9.75)

		' Get the position of this linklabel. That is where we'll begin adding the new controls.

		RowTop = LinkLabel1.Top

		' Increment the count of rows

		RelationRowCount += 1

		' Create controls in the panel for each field in the index.

		T1 = New TextBox
		T1.Name = "Rn" & RelationRowCount
		T1.BorderStyle = BorderStyle.FixedSingle
		T1.Top = RowTop
		T1.Left = 5
		T1.Width = 120
		T1.Text = "New Relation Name"
		T1.Font = f
		T1.ContextMenuStrip = ContextMenuStrip1
		If Not IsNothing(ForeignKey) Then T1.Text = ForeignKey.RelationshipName
		AddHandler T1.MouseDown, AddressOf TextBox_MouseDown
		AddHandler T1.KeyDown, AddressOf TextBox_KeyDown

		T2 = New TextBox
		T2.Name = "Pt" & RelationRowCount
		T2.BorderStyle = BorderStyle.FixedSingle
		T2.Top = RowTop
		T2.Left = 5 + T1.Width
		T2.Width = 120
		T2.Text = ""
		T2.Font = f
		If Not IsNothing(ForeignKey) Then T2.Text = ForeignKey.PrimaryTable
		AddHandler T2.MouseDown, AddressOf TextBox_MouseDown
		AddHandler T2.KeyDown, AddressOf TextBox_KeyDown

		T3 = New TextBox
		T3.Name = "Pf" & RelationRowCount
		T3.BorderStyle = BorderStyle.FixedSingle
		T3.Top = RowTop
		T3.Left = 5 + T1.Width + T2.Width
		T3.Width = 120
		T3.Text = ""
		T3.Font = f
		If Not IsNothing(ForeignKey) Then T3.Text = ForeignKey.PrimaryField
		AddHandler T3.MouseDown, AddressOf TextBox_MouseDown
		AddHandler T3.KeyDown, AddressOf TextBox_KeyDown

		T4 = New TextBox
		T4.Name = "Rt" & RelationRowCount
		T4.BorderStyle = BorderStyle.FixedSingle
		T4.Top = RowTop
		T4.Left = 5 + T1.Width + T2.Width + T3.Width
		T4.Width = 120
		T4.Text = ""
		T4.Font = f
		If Not IsNothing(ForeignKey) Then T4.Text = ForeignKey.RelatedTable
		AddHandler T4.MouseDown, AddressOf TextBox_MouseDown
		AddHandler T4.KeyDown, AddressOf TextBox_KeyDown

		T5 = New TextBox
		T5.Name = "Rf" & RelationRowCount
		T5.BorderStyle = BorderStyle.FixedSingle
		T5.Top = RowTop
		T5.Left = 5 + T1.Width + T2.Width + T3.Width + T4.Width
		T5.Width = 120
		T5.Text = ""
		T5.Font = f
		If Not IsNothing(ForeignKey) Then T5.Text = ForeignKey.RelatedField
		AddHandler T5.MouseDown, AddressOf TextBox_MouseDown
		AddHandler T5.KeyDown, AddressOf TextBox_KeyDown

		Lb = New CheckedListBox
		Lb.Name = "Lb" & RelationRowCount
		Lb.Top = RowTop
		Lb.Left = 5 + T1.Width + T2.Width + T3.Width + T4.Width + T5.Width
		Lb.Height = T1.Height
		Lb.Width = 134
		Lb.IntegralHeight = False
		Lb.Height = T1.Height
		Lb.CheckOnClick = True
		For jj = 1 To 2
			Lb.Items.Add(Choose(jj, "Cascade Update", "Cascade Delete"))
		Next jj
		If Not IsNothing(ForeignKey) Then
			Lb.SetItemChecked(0, ForeignKey.CascadeUpdate)
			Lb.SetItemChecked(1, ForeignKey.CascadeDelete)
		End If

		Panel2.Controls.Add(T1)
		Panel2.Controls.Add(T2)
		Panel2.Controls.Add(T3)
		Panel2.Controls.Add(T4)
		Panel2.Controls.Add(T5)
		Panel2.Controls.Add(Lb)


		' Increase the X position by the height of this new row

		RowTop = RowTop + T1.Height

		' Move the pre-existing link label down to be below the last row

		LinkLabel1.Top = RowTop
		LinkLabel1.Left = 5

		' Set the maximum value of the scroll bar control to the size of the inner panel.

		VScrollBar1.Maximum = Panel1.Height
		VScrollBar1.LargeChange = T1.Height

	End Sub
	'**********************************************************************

	' Function to assemble the SQL script that defines a relationship or ohanges to
	' a relationship from the information entered into the relationship manager.

	'**********************************************************************
	Public Function SQLScript() As String

		' Declare variables.

		Dim xx As String = ""
		Dim Script As String = ""
		Dim fk As KeyValuePair(Of String, ForeignKeyValues)

		' If a relationship was removed from the table, or changed
		' in any way, create the script to drop those constraints.
		' A modified relationship will be re-created later.

		For Each fk In FKDropped
			Script &= "ALTER TABLE [dbo].[" & fk.Value.PrimaryTable & "]" & vbCrLf
			Script &= "   DROP" & vbCrLf
			Script &= "      CONSTRAINT " & fk.Value.RelationshipName & ";" & vbCrLf
		Next fk

		' If a relation was added, create the script to add that now.

		If FKAdded.Count > 0 Then
			For Each fk In FKAdded
				Script &= "ALTER TABLE [dbo].[" & fk.Value.PrimaryTable & "] WITH NOCHECK" & vbCrLf
				Script &= "ADD CONSTRAINT [" & fk.Value.RelationshipName & "] FOREIGN KEY (["
				Script &= fk.Value.PrimaryField & "])"
				Script &= " REFERENCES [dbo].[" & fk.Value.RelatedTable & "]"
				Script &= " ([" & fk.Value.RelatedField & "])"
				If fk.Value.CascadeUpdate Then Script &= " ON UPDATE CASCADE"
				If fk.Value.CascadeDelete Then Script &= " ON DELETE CASCADE"
				Script &= ";" & vbCrLf
				'Script &= "ALTER TABLE [dbo].[" & fk.Value.PrimaryTable & "] WITH CHECK CHECK CONSTRAINT [" & fk.Value.RelationshipName & "];"
			Next fk
		End If

		' Return the assembled script

		Return Script
	End Function

	'******************************************************************************************

	' Sub to get changes to the relationship collection.

	'******************************************************************************************
	Private Sub GetChanges()

		' Declare variables.

		Dim RelationFound As Boolean
		Dim RelationChanged As Boolean
		Dim jj As Integer
		Dim RowNum As Integer
		Dim r As ForeignKeyValues
		Dim fk As KeyValuePair(Of String, ForeignKeyValues)

		' Clear our records of relations added or dropped.

		FKAdded.Clear()
		FKDropped.Clear()

		' Loop through all the rows, looking for additions or changes.

		For RowNum = 1 To RelationRowCount

			' Save the data for the current row in our current relation variable.

			r = GetRowValues(RowNum)

			' If the row does not exist in the Relations collection, add it to the FKAdded collection.

			If Not Relations.ContainsKey(r.RelationshipName) Then FKAdded.Add(r.RelationshipName, r)
		Next RowNum

		' Now go through existing relations and see if they still exist, and if so, see if they have changed.

		For Each fk In Relations
			RelationFound = False
			For jj = 1 To RelationRowCount
				If Panel2.Controls("Rn" & jj).Text = fk.Key Then
					RowNum = jj
					RelationFound = True
					Exit For
				End If
			Next jj

			' If a relation does not still exist, then add it to the list of dropped relations.

			If Not RelationFound Then
				FKDropped.Add(fk.Key, fk.Value)

				' If a relation still exists, see if it has changed.

			Else
				RelationChanged = False
				r = GetRowValues(RowNum)
				If fk.Value.PrimaryTable <> r.PrimaryTable Then RelationChanged = True
				If fk.Value.PrimaryField <> r.PrimaryField Then RelationChanged = True
				If fk.Value.RelatedTable <> r.RelatedTable Then RelationChanged = True
				If fk.Value.RelatedField <> r.RelatedField Then RelationChanged = True
				If fk.Value.CascadeUpdate <> r.CascadeUpdate Then RelationChanged = True
				If fk.Value.CascadeDelete <> r.CascadeDelete Then RelationChanged = True

				' If a relation has changed, it needs to be both dropped and re-created, so add it to
				' both collections.

				If RelationChanged Then
					FKDropped.Add(fk.Key, fk.Value)
					FKAdded.Add(fk.Key, r)
				End If
			End If
		Next fk

	End Sub
	'******************************************************************************************

	' Function to return the relationship information from a specified row.

	'******************************************************************************************
	Private Function GetRowValues(RowNum As Integer) As ForeignKeyValues

		' Declare variables

		Dim r As New ForeignKeyValues

		' Get the information from the selected row.

		r.RelationshipName = Panel2.Controls("Rn" & RowNum).Text
		r.PrimaryTable = Panel2.Controls("Pt" & RowNum).Text
		r.PrimaryField = Panel2.Controls("Pf" & RowNum).Text
		r.RelatedTable = Panel2.Controls("Rt" & RowNum).Text
		r.RelatedField = Panel2.Controls("Rf" & RowNum).Text
		r.CascadeUpdate = CType(Panel2.Controls("Lb" & RowNum), CheckedListBox).GetItemChecked(0)
		r.CascadeDelete = CType(Panel2.Controls("Lb" & RowNum), CheckedListBox).GetItemChecked(1)

		' Return the row values.

		GetRowValues = r

	End Function
	'******************************************************************************************

	' Sub to clear all the lists of relations added and dropped.

	'******************************************************************************************
	Public Sub ClearCollections()
		FKAdded.Clear()
		FKDropped.Clear()
	End Sub
	'******************************************************************************************

	' Declare the sub to display all existing relationships for a specified table.

	'******************************************************************************************
	Public Sub ShowRelationships(TableName As String)

		' Declare variables

		Dim ForeignKeys As New Dictionary(Of String, ForeignKeyValues)
		Dim fk As KeyValuePair(Of String, ForeignKeyValues)

		' Initialize the relationships panel

		InitializePanels()

		' Get the foreign keys attached to this table.

		ForeignKeys = GetForeignKeys(TableName)

		' Loop through the foreign keys collection add each to the relationships pane.

		For Each fk In ForeignKeys

			' Add the relationship to the panel

			AddNewRelationship(fk.Value)

			' Add the relationship to our Relations collection.

			Relations.Add(fk.Key, fk.Value)
		Next fk

		System.Windows.Forms.Application.DoEvents()

		' Set the add/edit mode to edit.

		TableEditMode = TableEditModeEnum.TableAlter

	End Sub
	'******************************************************************************************

	' Declare the sub to display all existing relationships.

	'******************************************************************************************
	Public Sub ShowAllRelationships()

		' Declare variables

		Dim x As DataTable
		Dim ForeignKeys As New Dictionary(Of String, ForeignKeyValues)
		Dim fk As KeyValuePair(Of String, ForeignKeyValues)

		' Initialize the relationships panel

		InitializePanels()

		' Get a list of all the tables in the database.

		x = DB.GetSchema("Tables")
		For jj = 0 To x.Rows.Count - 1

			' Make sure the "table" is type BASE TABLE and not VIEW.  A view looks like a table
			' except for the table type, and we list views separately.

			If x.Rows(jj).Item("table_type") = "BASE TABLE" Then

				' If we encounter one of our settings tables, ignore it.

				If VB.Left(x.Rows(jj).Item("table_name"), 3) <> "db_" Then
					' Get the foreign keys attached to this table.

					ForeignKeys = GetForeignKeys(x.Rows(jj).Item("table_name"))

					' Loop through the foreign keys collection add each to the relationships pane.

					For Each fk In ForeignKeys

						' Add the relationship to the panel

						AddNewRelationship(fk.Value)

						' Add the relationship to our Relations collection.

						Relations.Add(fk.Key, fk.Value)
					Next fk

					System.Windows.Forms.Application.DoEvents()
				End If
			End If

		Next jj

		' Set the add/edit mode to edit.

		TableEditMode = TableEditModeEnum.TableAlter

	End Sub
End Class
