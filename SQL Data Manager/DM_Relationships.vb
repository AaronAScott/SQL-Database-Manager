Imports System.Data.SqlClient
Imports VB = Microsoft.VisualBasic
Public Class frmRelationships
	Inherits System.Windows.Forms.Form
	'**********************************************************
	' SQL Data Manager Relationships Viewer (frmRelationships)
	' DM_RELATIONSHIPS.VB
	' Written: November 2018
	' Programmer: Aaron Scott
	' Copyright (C) 2018 Sirius Software All Rights Reserved
	'**********************************************************

	' Declare variables that are the public properties of this module.

	Public LayoutChanged As Boolean

	' Declare variables local to this module

	Private StartX As Integer = 25 ' This specifies where we begin laying out new table objects.
	Private StartY As Integer = 100
	Private TableCount As Integer
	Private DragFromPoint As Point
	Private Tables As New Dictionary(Of String, FKTable)
	Private Relations As New Dictionary(Of String, ForeignKeyValues)

	' Create an Enum to indicate the various things a window might do in the mousemove event.

	Private Enum MouseMoveModes
		None
		DragWindow
		SizeWindow
	End Enum
	Private WindowMode As MouseMoveModes

	'**********************************************************

	' The form is loaded.

	'**********************************************************
	Private Sub frmRelationships_Load(sender As Object, e As EventArgs) Handles Me.Load

		' Display a startup status.

		frmMain.StatusLabel.Text = "Loading form.  Please wait..."
		System.Windows.Forms.Application.DoEvents()

		' Set the background color

		frmTheme.SetBackgroundColor(Me)

		' Get all the relationships in the database

		GetRelationships()

		' Get any saved layout.

		GetSavedLayout()

		' Clear the status message.

		frmMain.StatusLabel.Text = ""
		System.Windows.Forms.Application.DoEvents()
	End Sub

	'**********************************************************

	' The form is closed.

	'**********************************************************

	Private Sub frmRelationships_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

		' Declare variables

		Dim xx As Integer
		Dim t As KeyValuePair(Of String, FKTable)
		Dim fk As FKTable
		Dim Command As SqlCommand
		Dim DA As New SqlDataAdapter
		Dim DS As New DataSet
		Dim DT As DataTable
		Dim dr As DataRow
		Dim CmdBld As New SqlCommandBuilder

		' If the layout has changed, see if the user wants to save the layout.

		If LayoutChanged Then
			If MsgBox("Do you want to save changes to the layout of ""Table Relationships""?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, "Table Relationshiips") = MsgBoxResult.Yes Then

				' Try to access the table to save the relationships.  The first time the layout is saved, we will have
				' create the table.


				Command = New SqlCommand("SELECT * FROM db_info_relationships ORDER BY TableName", DB)
				DA.SelectCommand = Command

				' Use an SQLCommandBuilder to generate the other 3 types of commands for the data adapter.

				CmdBld.DataAdapter = DA

				' Watch for errors here. Under some circumstances, the command builder is unable
				' to generate update commands. If that happens, the table cannot be updated, only viewed,

				On Error Resume Next
				DA.UpdateCommand = CmdBld.GetUpdateCommand
				DA.InsertCommand = CmdBld.GetInsertCommand
				DA.DeleteCommand = CmdBld.GetDeleteCommand

				' Now try to access the layout table.

				On Error GoTo LayoutTableNotFound
				DA.Fill(DS, "Table")
				DT = DS.Tables("Table")

				' Save the size and position of each table object.

				On Error GoTo SaveLayoutError
				For Each t In Tables
					xx = Find(DT, "TableName='" & t.Value.TableName & "'")
					If xx <> NOMATCH Then
						dr = DT.Rows(xx)
					Else
						dr = DT.NewRow
						dr("TableName") = t.Value.TableName
						dr("RelationName") = t.Value.RelationName
					End If
					dr("X") = t.Value.Left
					dr("Y") = t.Value.Top
					dr("Width") = t.Value.Width
					dr("Height") = t.Value.Height
					If xx = NOMATCH Then DT.Rows.Add(dr)
					DA.Update(DT)
				Next t

				' Now save the size and position of the relationships window.


				xx = Find(DT, "TableName='Window'")
				If xx <> NOMATCH Then
					dr = DT.Rows(xx)
				Else
					dr = DT.NewRow
					dr("TableName") = "Window"
					dr("RelationName") = ""
				End If
				dr("X") = Me.Left
				dr("Y") = Me.Top
				dr("Width") = Me.Width
				dr("Height") = Me.Height
				If xx = NOMATCH Then DT.Rows.Add(dr)
				DA.Update(DT)

				' Dispose of our objects

				DA.Dispose()
				DS.Dispose()
				Command.Dispose()

			End If
		End If
		On Error GoTo 0

		Exit Sub

		' Trap for errors saving the layout.

SaveLayoutError:
		MsgBox("Failed to save relationship layout." & vbCrLf & Err.Description, MsgBoxStyle.Exclamation, "Save Relationships Layout")
		Exit Sub

		' Trap if the layout table has not been created yet.

LayoutTableNotFound:

		' Create the table and resume.

		Command = New SqlCommand("CREATE TABLE [dbo].[db_info_relationships]( [ID]  INT NOT NULL IDENTITY (1,1),  [TableName]  NVARCHAR(20) NOT NULL, [RelationName]  NVARCHAR(20) NOT NULL,  [X]  INT NOT NULL, [Y]  INT NOT NULL, [Width]  INT NOT NULL, [Height]  INT NOT NULL, PRIMARY KEY CLUSTERED ([ID] ASC))", DB)
		Command.ExecuteNonQuery()
		Resume

	End Sub

	'**********************************************************

	' The window has been resized or moved.

	'**********************************************************
	Private Sub frmRelationships_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd, Me.LocationChanged
		LayoutChanged = True
	End Sub
	'**********************************************************

	' The window must be repainted.

	'**********************************************************
	Private Sub frmRelationships_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint


		Dim ii As Integer
		Dim jj As Integer
		Dim pPrimary As Point
		Dim pRelated As Point
		Dim g As Graphics = e.Graphics
		Dim fk As KeyValuePair(Of String, ForeignKeyValues)

		' Redraw all the lines that connect the fields and tables together.

		For Each fk In Relations

			' Before trying to show a relationship, make sure both the primary and related
			' tables have been added to the tables collection.  When the form is loaded,
			' that will have happened before the first paint event is triggered, but during
			' a refresh, this event can be called before both tables have been added.

			If Tables.ContainsKey(fk.Value.PrimaryTable) And Tables.ContainsKey(fk.Value.RelatedTable) Then

				' Get the index in the list of fields of the field in the primary table that
				' is part of the relation.

				jj = Tables(fk.Value.PrimaryTable).Fields.FindStringExact(fk.Value.PrimaryField)

				' Make the field the top field in the list box of fields.

				Tables(fk.Value.PrimaryTable).Fields.TopIndex = jj

				' Get the index in the list of fields of the field in the referenced table that
				' is part of the relation.

				ii = Tables(fk.Value.RelatedTable).Fields.FindStringExact(fk.Value.RelatedField)

				' Make the field the top field in the list box of fields.

				Tables(fk.Value.RelatedTable).Fields.TopIndex = ii

				' Get the locations of the selected field of each table.

				pPrimary = Tables(fk.Value.PrimaryTable).Fields.GetItemRectangle(jj).Location
				pRelated = Tables(fk.Value.RelatedTable).Fields.GetItemRectangle(ii).Location

				' If the primary table is above or to the right of the related table,
				' draw a line from the right of the primary table to the left of the
				' related table.

				pPrimary.X += Tables(fk.Value.PrimaryTable).Location.X
				pPrimary.Y += Tables(fk.Value.PrimaryTable).Location.Y + 22 + Tables(fk.Value.PrimaryTable).Fields.GetItemHeight(jj) / 2 ' Allow for height of caption and half the height of the item.

				pRelated.X += Tables(fk.Value.RelatedTable).Location.X
				pRelated.Y += Tables(fk.Value.RelatedTable).Location.Y + 22 + Tables(fk.Value.RelatedTable).Fields.GetItemHeight(ii) / 2

				If pPrimary.X > pRelated.X Then
					pRelated.X += Tables(fk.Value.RelatedTable).Width
					g.DrawLine(Pens.Black, pPrimary, pRelated)
				Else
					g.DrawLine(Pens.Black, pPrimary, pRelated)
				End If
			End If
		Next fk

	End Sub

	'**********************************************************

	' The Edit Relationship menu option is clicked.

	'**********************************************************
	Private Sub mnuEditRelationship_Click(sender As Object, e As EventArgs) Handles mnuEditRelationship.Click

		Dim r = frmManageRelationships
		r.ShowAllRelationships()
		r.MdiParent = frmMain
		r.Show()

	End Sub

	'**********************************************************

	' The Add Relationship menu option is clicked.

	'**********************************************************
	Private Sub mnuAddRelationship_Click(sender As Object, e As EventArgs) Handles mnuAddRelationship.Click

		Dim r = frmManageRelationships
		r.MdiParent = frmMain
		r.Show()
	End Sub

	'**********************************************************

	' The Add Table menu option is clicked.

	'**********************************************************
	Private Sub mnuAddTable_Click(sender As Object, e As EventArgs) Handles mnuAddTable.Click
		frmAddTable.ShowDialog()
		Me.Invalidate()
	End Sub
	'**********************************************************

	' The Remove Table menu option is selected.

	'**********************************************************
	Private Sub mnuRemoveTable_Click(sender As Object, e As EventArgs) Handles mnuRemoveTable.Click

		' Declare variables

		Dim TableName As String = ""
		Dim c As Control
		Dim t As KeyValuePair(Of String, FKTable)

		' Loop through all the ables looking for the selected table and remove it.

		For Each t In Tables
			If t.Value.Active Then
				TableName = t.Key
				Tables.Remove(t.Key)
				Exit For
			End If
		Next t

		' Loop through all the window controls and remove the selected control.

		For Each c In Me.Controls
			If TypeOf (c) Is FKTable AndAlso CType(c, FKTable).TableName = TableName Then c.Dispose()
		Next c

		' Force the display to be redrawn.

		Me.Invalidate()

		' Indicate that the layout has changed.

		LayoutChanged = True

	End Sub
	'**********************************************************

	' Event handler for the table objects' "MouseDown" event.
	' This event is late-bound.
	' By declaring "sender" as FKTable instead of "object",
	' we don't have to cast it to FKTable to access that
	' objects properties and methods.

	'**********************************************************
	Private Sub FKTable_MouseDown(sender As FKTable, e As MouseEventArgs)

		' Clear the move/resize mode by default

		WindowMode = MouseMoveModes.None

		' Make sure the left mouse button is clicked.

		If e.Button = Windows.Forms.MouseButtons.Left Then

			' If the mouse is in the caption area of the table object, then
			' enter drag mode.

			If e.Y > sender.Top + 5 And e.Y <= sender.Top + 20 Then ' Make sure mouse in in title bar area
				WindowMode = MouseMoveModes.DragWindow

				' If we have a sizing cursor assigned to the table object,
				' then enter sizing mode.

			ElseIf sender.Cursor = Cursors.SizeNS Or sender.Cursor = Cursors.SizeWE Then
				WindowMode = MouseMoveModes.SizeWindow
			End If

			' Remember the starting point of our drag/size operation.

			DragFromPoint = e.Location
		End If

	End Sub

	'**********************************************************

	' Event handler for the table objects' "MouseMove" event.
	' This event is late-bound.
	' By declaring "sender" as FKTable instead of "object",
	' we don't have to cast it to FKTable to access that
	' objects properties and methods.

	'**********************************************************
	Private Sub FKTable_MouseMove(sender As FKTable, e As MouseEventArgs)

		' Declare variables

		Dim Tolerance As Integer = 5 ' This defines how much space we have around the border to react to size events
		Dim RelativeX As Integer
		Dim RelativeY As Integer
		Dim NewX As Integer
		Dim NewY As Integer

		' Do nothing unless the left button is down.

		If e.Button = Windows.Forms.MouseButtons.Left Then

			' Clear the display of the lines connecting the tables.

			Me.CreateGraphics.Clear(Me.BackColor)

			' Whatever we do now will change the layout.

			LayoutChanged = True

			' See if we are in window drag mode.

			If WindowMode = MouseMoveModes.DragWindow Then

				' Determine the position of the mouse pointer relative to the table.

				RelativeX = e.X - sender.Left
				RelativeY = e.Y - sender.Top

				' Determine how much the table has moved, based on the previous location.

				NewX = sender.Left
				NewY = sender.Top
				If e.X > DragFromPoint.X Then NewX += (e.X - DragFromPoint.X) Else NewX -= (DragFromPoint.X - e.X)
				If e.Y > DragFromPoint.Y Then NewY += (e.Y - DragFromPoint.Y) Else NewY -= (DragFromPoint.Y - e.Y)
				sender.Location = New Point(NewX, NewY)

				DragFromPoint = e.Location


				' Determine if we are in a window sizing mode.

			ElseIf WindowMode = MouseMoveModes.SizeWindow Then

				' See if we are sizing the window width.

				If sender.Cursor = Cursors.SizeWE Then

					' Determine if the mouse is on the left or right border of the table object.

					If e.X < sender.Left + Tolerance Then
						If e.X < DragFromPoint.X Then ' We're stretching the left side left
							sender.Left = e.X
							sender.Width += (DragFromPoint.X - e.X)
						Else
							sender.Left = e.X ' We're shrinking the left side right
							sender.Width -= (e.X - DragFromPoint.X)
						End If
					ElseIf e.X > sender.Left + sender.Width - Tolerance Then
						If e.X < DragFromPoint.X Then ' We're stretching the left side left
							sender.Width -= (DragFromPoint.X - e.X)
						Else
							sender.Width += (e.X - DragFromPoint.X) ' We're shrinking the left side right
						End If
					End If

					' Determine if we are in a height sizing mode.

				ElseIf sender.Cursor = Cursors.SizeNS Then

					' Determine if the mouse is at the top or bottom of the table object.

					If e.Y < sender.Top + Tolerance Then
						If e.Y < DragFromPoint.Y Then ' We're stretching the top up
							sender.Top = e.Y
							sender.Height += (DragFromPoint.Y - e.Y)
						Else
							sender.Top = e.Y ' We're shrinking the top down
							sender.Height -= (e.Y - DragFromPoint.Y)
						End If
					ElseIf e.Y > sender.Top + sender.Height - Tolerance Then ' We're stretching the bottom down
						If e.Y < DragFromPoint.Y Then
							sender.Height -= (DragFromPoint.Y - e.Y)
						Else
							sender.Height += (e.Y - DragFromPoint.Y) ' We're shrinking the bottom up
						End If
					End If
				End If

				' Remember the starting point of our drag/size operation.

				DragFromPoint = e.Location
				System.Windows.Forms.Application.DoEvents()
			End If

		End If


	End Sub

	'**********************************************************

	' Event handler for the table objects' "MouseUp" event.
	' This event is late-bound.
	' By declaring "sender" as FKTable instead of "object",
	' we don't have to cast it to FKTable to access that
	' objects properties and methods.

	'**********************************************************
	Private Sub FKTable_MouseUp(sender As FKTable, e As MouseEventArgs)

		' Clear the starting point for drag/size operations.

		DragFromPoint = Nothing

		' Force the display to redraw the relationships.

		Me.Invalidate()

	End Sub
	'**********************************************************

	' Event handler for the table objects SelectionChanged
	' event.

	'**********************************************************
	Private Sub FKTable_SelectionChanged(sender As FKTable, e As EventArgs)

		' Declare variables

		Dim t As KeyValuePair(Of String, FKTable)

		' Loop through all the table objects.  The current active table
		' will be deactivated (the new table has already been activated).

		For Each t In Tables
			If t.Value.Active And Not t.Value Is sender Then t.Value.Active = False
		Next t
	End Sub
	'**********************************************************

	' The timer has ticked.  Clear the status label.

	'**********************************************************

	Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
		StatusLabel.Text = ""
		Timer1.Enabled = False
	End Sub
	'**********************************************************

	' Sub to fill the Relations dictionary with all existing relationships in the database.

	'**********************************************************
	Private Sub GetRelationships()

		Dim x As DataTable
		Dim Command As SqlCommand
		Dim DA As New SqlDataAdapter
		Dim DS As New DataSet
		Dim DT As DataTable
		Dim fk As KeyValuePair(Of String, ForeignKeyValues)
		Dim fkd As New Dictionary(Of String, ForeignKeyValues)

		' Get a list of all relationships in this database, by relationship name.  This
		' will contain any new ones, but not any deleted ones.

		x = DB.GetSchema("ForeignKeys") ' Gets relation names
		For jj = 0 To x.Rows.Count - 1
			TableCount = jj + 1

			' Create a command to get the name of the primary table and the referenced table of the current relationship.

			Command = New SqlCommand("SELECT sys.tables.name AS TableName, sys.foreign_keys.name as RelationName, sys.tables.object_id, sys.foreign_keys.parent_object_id AS RelationName, sys.foreign_keys.parent_object_id, sys.tables.object_id FROM (sys.foreign_key_columns INNER JOIN sys.foreign_keys ON sys.foreign_key_columns.constraint_object_id=sys.foreign_keys.object_id ) INNER JOIN sys.tables ON (sys.foreign_key_columns.parent_object_id=sys.tables.object_id OR sys.foreign_key_columns.referenced_object_id=sys.tables.object_id) WHERE sys.foreign_keys.name='" & x.Rows(jj)("constraint_name") & "';", DB)
			DA.SelectCommand = Command
			DS.Clear()
			DA.Fill(DS, "table")

			' Our data table will have two rows: one for the parent table, another for the referenced table.  Add
			' a table object for each table that does not already exist in the tables collection.

			DT = DS.Tables("table")
			For ii = 0 To DT.Rows.Count - 1
				If Not Tables.ContainsKey(DT.Rows(ii)("TableName")) Then AddTable(DT.Rows(ii)("TableName"), DT.Rows(ii)("RelationName"))

				' Get all the foreign keys belonging to the table.  A relationship consists of a foreign
				' key that contains a primary table, a referenced table, the names of the fields in each
				' which are connected, and whether the relation is cascaded by an update or delete.

				fkd = GetForeignKeys(DT.Rows(ii)("TableName"))
				For Each fk In fkd
					If Not Relations.ContainsKey(fk.Value.RelationshipName) Then Relations.Add(fk.Value.RelationshipName, fk.Value)
					System.Windows.Forms.Application.DoEvents()
				Next fk

			Next ii

			' Dispose of our objects.

			Command.Dispose()
			DA.Dispose()
			DS.Dispose()

		Next jj


	End Sub

	'**********************************************************

	' Sub to restore table objects' saved sizes and positions.

	'**********************************************************
	Private Sub GetSavedLayout()

		' Declare variables

		Dim ii As Integer
		Dim Command As SqlCommand
		Dim DA As New SqlDataAdapter
		Dim DS As New DataSet
		Dim DT As DataTable
		Dim t As FKTable

		' Now see if we have a saved layout.

		Try

			Command = New SqlCommand("SELECT * FROM [dbo].[db_info_relationships] ORDER BY TableName", DB)
			DA.SelectCommand = Command
			DS.Clear()
			DA.Fill(DS, "table")
			DT = DS.Tables("table")

			' If any of the saved tables still has a match in this window, restore its size
			' and position.

			For ii = 0 To DT.Rows.Count - 1
				If Tables.ContainsKey(DT.Rows(ii)("TableName")) Then
					t = Tables(DT.Rows(ii)("TableName"))
					t.Left = DT.Rows(ii)("X")
					t.Top = DT.Rows(ii)("Y")
					t.Width = DT.Rows(ii)("Width")
					t.Height = DT.Rows(ii)("Height")
				End If
				System.Windows.Forms.Application.DoEvents()
			Next ii

			' If the relationship window has a saved size and position, restore that.

			ii = Find(DT, "TableName='Window'")
			If ii <> NOMATCH Then
				Me.Left = DT.Rows(ii)("X")
				Me.Top = DT.Rows(ii)("Y")
				Me.Width = DT.Rows(ii)("Width")
				Me.Height = DT.Rows(ii)("Height")
			End If

			' Ignore errors--we may not have a saved layout.

		Catch ex As Exception

		End Try

		' Indicate the layout has not changed.

		LayoutChanged = False

	End Sub

	'**********************************************************

	' Sub to add a new table to the relationships window.

	'**********************************************************
	Public Sub AddTable(TableName As String, Optional RelationName As String = "")

		' Declare variables

		Dim ii As Integer
		Dim t As FKTable
		Dim Cmd As SqlCommand
		Dim DA As New SqlDataAdapter
		Dim DS As New DataSet
		Dim DT As DataTable

		' Create a new table object.  This is a custom panel, actually.

		t = New FKTable
		t.TableName = TableName
		t.RelationName = RelationName
		t.Location = New Point(StartX, StartY)
		t.Width = 100
		t.Height = 120

		' Determine where to place it, initially.  When we fill a row, start a new row.

		StartX += 120
		If StartX + t.Width > Me.Width Then
			StartX = 30
			StartY += 140
		End If

		' Get the list of all fields (columns) in the table.

		Cmd = New SqlCommand("SELECT * FROM information_schema.columns WHERE table_name = '" & TableName & "' ORDER BY ordinal_position", DB)
		DA.SelectCommand = Cmd
		DA.Fill(DS, "Table")
		DT = DS.Tables("Table")

		' The rows of the schema table have the following names:
		' TABLE_NAME,COLUMN_NAME,COLUMN_DEFAULT,IS_NULLABLE,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH,NUMERIC_PRECISION,DATETIME_PRECISION

		' Add a row to the tables panel for each field in the table schema.

		For ii = 0 To DT.Rows.Count - 1
			t.Fields.Items.Add(DT.Rows(ii)("COLUMN_NAME"))
		Next ii

		' Attach handlers to the table object for all mouse events we need.

		AddHandler t.MouseDownEvent, AddressOf FKTable_MouseDown
		AddHandler t.MouseUpEvent, AddressOf FKTable_MouseUp
		AddHandler t.MouseMoveEvent, AddressOf FKTable_MouseMove
		AddHandler t.SelectionChanged, AddressOf FKTable_SelectionChanged

		' Now add the table to the tables collection, as well as our
		' controls collection.  A table must be added to the tables collection
		' first, or the paint event will fail when the control is added.

		Tables.Add(TableName, t)
		Me.Controls.Add(t)

	End Sub
	'**********************************************************

	' Sub to rebuild the display of relationships.  This can be called from outside
	' this form, by the ManageRelationships form, after it is done adding
	' or modifying relationships.

	'**********************************************************
	Public Sub RefreshDisplay()

		' Declare variables

		Dim jj As Integer

		' Reset our layout positions.

		StartX = 25 ' This specifies where we begin laying out new table objects.
		StartY = 100

		' Clear the tables collection before we re-create it.

		Tables.Clear()

		' Clear the relations collection before we re-create it.

		Relations.Clear()

		' Remove all tables from the controls collection of the form.

		For jj = Me.Controls.Count - 1 To 0 Step -1
			If TypeOf (Me.Controls(jj)) Is FKTable Then Me.Controls(jj).Dispose()
		Next jj

		' Get a new list of relationships

		GetRelationships()

		' Get any saved layout.

		GetSavedLayout()

		' Force the window to redraw itself.

		Me.Invalidate()

	End Sub
	'**********************************************************

	' Property to determine if a table is already part of the displayed tables collection.

	'**********************************************************
	Public ReadOnly Property ContainsTable(TableName As String) As Boolean
		Get
			ContainsTable = Tables.ContainsKey(TableName)
		End Get
	End Property
End Class