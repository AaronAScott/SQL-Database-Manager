Imports System.Data.SqlClient
Imports VB = Microsoft.VisualBasic
Public Class frmObjectList
	Inherits System.Windows.Forms.Form
	'**********************************************************
	' SQL Data Manager Object List form (frmObjectList)
	' DM_OBJECTLIST.VB
	' Written: October 2017
	' Programmer: Aaron Scott
	' Copyright (C) 2017 Sirius Software All Rights Reserved
	'**********************************************************

	' Declare variables local to this module.

	Private FormSize As Size
	Private FormLocation As Point

	'**********************************************************

	' The form is loaded.

	'**********************************************************
	Private Sub frmObjectList_Load(sender As Object, e As EventArgs) Handles Me.Load

		' Declare variables

		Dim zx As String

		' Put up a message

		frmMain.StatusLabel.Text = "Loading Table.  Please Wait..."
		System.Windows.Forms.Application.DoEvents()

		' Fill the object list box

		RefreshObjectList()

		' Set the size of the object list window.

		zx = GetSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "DBObjectsWindow", "Size")
		If zx <> "" Then
			FormSize = New Size(Me.Width, Me.Height)
			Me.Height = Val(ParseString(zx))
			Me.Width = Val(zx)
		End If

		' Clear message

		frmMain.StatusLabel.Text = ""
		System.Windows.Forms.Application.DoEvents()


	End Sub

	'**********************************************************

	' The form is resized.

	'**********************************************************
	Private Sub frmObjectList_Resize(sender As Object, e As EventArgs) Handles Me.Resize

		' The following prevents the form from being maximized whenever the parent
		' form is moved or resized.

		If DontAllowResize Then Me.Size = FormSize Else FormSize = Me.Size

		' Resize the TreeView control to occupy the entire form.

		TreeView1.Bounds = Me.ClientRectangle

	End Sub
	'**********************************************************

	' The form is moved.

	'**********************************************************
	Private Sub frmObjectList_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged

		' The following prevents the form from being moved whenever the parent
		' form is moved or resized.

		If DontAllowResize Then Me.Location = FormLocation Else FormLocation = Me.Location

	End Sub

	'**********************************************************

	' The form is closed.

	'**********************************************************
	Private Sub frmObjectList_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

		' Declare variables

		Dim zx As String

		' When this form is closed, close the database.

		CloseDatabase()

		' Save the window size and position

		If Me.Left >= 0 And Me.Top >= 0 Then
			zx = Me.Height & "," & Me.Width
			SaveSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "DBObjectsWindow", "Size", zx)
		End If

		' Make sure the flag that prevents MDI child forms moving or resizing is cleared.

		DontAllowResize = False

	End Sub

	'**********************************************************

	' A label has been edited.

	'**********************************************************
	Private Sub TreeView1_AfterLabelEdit(sender As Object, e As NodeLabelEditEventArgs) Handles TreeView1.AfterLabelEdit

		' Declare variables

		Dim x1 As Integer
		Dim x2 As Integer
		Dim TableName As String
		Dim zx1 As String
		Dim zx2 As String
		Dim Command As SqlCommand = Nothing

		' See if the edit actually took place.

		If e.Label Is Nothing Then
			e.CancelEdit = True
		ElseIf e.Label <> e.Node.Text Then
			Try

				' Determine what kind of item was selected to edit.

				Select Case e.Node.Parent.Name
					Case "Tables", "Queries"
						Command = New SqlCommand("EXEC sp_rename '[dbo]." & e.Node.Text & "','" & e.Label & "'", DB)
						Command.ExecuteNonQuery()
						Command.Dispose()
					Case "Indexes"

						' Get the name of the table to which the index belongs

						x1 = InStr(1, e.Node.Text, "-")
						TableName = VB.Left(e.Node.Text, x1 - 1)

						' Find the index name inside the parentheses.

						x1 = InStr(1, e.Node.Text, "(")
						x2 = InStr(x1 + 1, e.Node.Text, ")")
						zx1 = Mid(e.Node.Text, x1 + 1, x2 - x1 - 1)

						' The new name may or may not be of the same format.  If it contains
						' parentheses, extract the contents; otherwise use the entire text.

						x1 = InStr(1, e.Label, "(")
						x2 = InStr(x1 + 1, e.Label, ")")
						If x1 > 0 And x2 > x1 Then zx2 = Mid(e.Label, x1 + 1, x2 - x1 - 1) Else zx2 = e.Label

						' Make sure we don't get any extra parentheses.

						zx2 = SRep(SRep(zx2, 1, "(", ""), 1, ")", "")

						' Assemble the command to rename the index and execute it.

						Command = New SqlCommand("EXEC sp_rename @objname = N'[dbo].[" & TableName & "].[" & zx1 & "]', @newname = N'" & zx2 & "', @objtype = N'INDEX';", DB)
						Command.ExecuteNonQuery()
						Command.Dispose()

						' Redisplay the index with the new name and cancel the edit.  This
						' is the only way to re-display the new name in the proper format,
						' including the table name.

						e.Node.Text = TableName & " - (" & zx2 & ")"
						e.CancelEdit = True

					Case "ForeignKeys"

						' Get the name of the table to which the index belongs

						x1 = InStr(1, e.Node.Text, "-")
						TableName = VB.Left(e.Node.Text, x1 - 1)

						' Find the foreign key name inside the parentheses.

						x1 = InStr(1, e.Node.Text, "(")
						x2 = InStr(x1 + 1, e.Node.Text, ")")
						zx1 = Mid(e.Node.Text, x1 + 1, x2 - x1 - 1)

						' The new name may or may not be of the same format.  If it contains
						' parentheses, extract the contents; otherwise use the entire text.

						x1 = InStr(1, e.Label, "(")
						x2 = InStr(x1 + 1, e.Label, ")")
						If x1 > 0 And x2 > x1 Then zx2 = Mid(e.Label, x1 + 1, x2 - x1 - 1) Else zx2 = e.Label

						' Make sure we don't get any extra parentheses.

						zx2 = SRep(SRep(zx2, 1, "(", ""), 1, ")", "")

						' Replace any brackets with double brackets, which is the only way to indicate a bracket.

						zx1 = SRep(SRep(zx1, 1, "[", "[["), 1, "]", "]]")
						zx2 = SRep(SRep(zx2, 1, "[", "[["), 1, "]", "]]")

						' Assemble the command to rename the index and execute it.

						Command = New SqlCommand("EXEC sp_rename @objname = N'[dbo].[" & zx1 & "]', @newname = N'" & zx2 & "', @objtype = N'OBJECT';", DB)
						Command.ExecuteNonQuery()
						Command.Dispose()

						' Redisplay the object with the new name, and cancel the edit.

						e.Node.Text = TableName & " - (" & zx2 & ")"
						e.CancelEdit = True

						' Other items (workspace settings, system tables) cannot be edited, so cancel
						' the edit.

					Case Else
						e.CancelEdit = True
				End Select
			Catch ex As Exception
				MsgBox("Rename of item " & e.Node.Text & " failed." & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Rename Database Object")
				e.CancelEdit = True
			End Try
		End If

	End Sub

	'**********************************************************

	' If a  node is right-clicked, to invoke the context menu,
	' make it the selected node just as if it were left-clicked.

	'**********************************************************

	Private Sub TreeView1_MouseDown(sender As Object, e As MouseEventArgs) Handles TreeView1.MouseDown

		' Check for the right mouse button.

		If e.Button = Windows.Forms.MouseButtons.Right Then
			Dim x As TreeViewHitTestInfo = TreeView1.HitTest(e.X, e.Y)
			TreeView1.SelectedNode = x.Node

			' Change the value for the context menu to reflect the object it refers to

			If TreeView1.SelectedNode.Text <> "Table" Then mnuAddNew.Text = "Add New " & TreeView1.SelectedNode.Text

			' Unless the node belongs to "Table", then
		End If
	End Sub

	'**********************************************************

	' A node in the treeview is clicked.

	'**********************************************************
	Private Sub TreeView1_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseDoubleClick

		' Declare variables

		Dim DA As New SqlDataAdapter
		Dim DS As New DataSet
		Dim DT As DataTable
		Dim Cmd As SqlCommand
		Dim dv As New frmDataView

		' Verify that the node clicked on is valid for opening something.

		If e.Node.Parent.Name = "Tables" Or e.Node.Parent.Name = "Queries" Or e.Node.Parent.Name = "Views" Or e.Node.Parent.Name = "Settings" Or e.Node.Parent.Name = "SystemCatalog" Then

			' Make sure form can be resized.

			DontAllowResize = False

			' If the clicked node is a table name, show the data for that table.

			dv.MdiParent = frmMain
			If e.Node.Parent.Name = "Tables" Or e.Node.Parent.Name = "Settings" Then
				dv.TableName = e.Node.Text
				dv.Show()
			End If

			' If the clicked node is a view name, show the data for that view.

			If e.Node.Parent.Name = "Views" Then
				dv.ViewName = e.Node.Text
				dv.Show()
			End If

			' If the clicked node is a system catalog, show the catalog data.

			If e.Node.Parent.Name = "SystemCatalog" Then
				dv.CatalogName = e.Node.Text
				dv.Show()
			End If

			' If the clicked node is a query, retrieve the SQL script, and determine if it
			' is a SELECT query or an action query.

			If e.Node.Parent.Name = "Queries" Then

				' Open the system table that contains the SQL text of the selected query.

				Cmd = New SqlCommand("SELECT object_definition(object_id('[dbo].[" & e.Node.Text & "]')) AS SQL_Text FROM sys.sql_modules", DB)
				DA.SelectCommand = Cmd
				DA.Fill(DS, "Table")
				DT = DS.Tables("Table")

				' Determine the query type.

				If DT.Rows.Count > 0 Then

					If InStr(DT.Rows(0)("SQL_Text"), "SELECT ", CompareMethod.Text) > 0 Then

						' If a select query, bring up the DataView form.

						dv.QueryName = e.Node.Text
						dv.Show()

						' For an action query, just execute the query.

					Else
						Try
							Cmd = New SqlCommand(e.Node.Text, DB)
							Cmd.CommandType = CommandType.StoredProcedure
							Cmd.ExecuteNonQuery()
							MsgBox("Query executed successfully.", MsgBoxStyle.Information, "Execute Query")
						Catch ex As Exception
							MsgBox("Query failed." & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Execute Query")
						End Try
					End If
				End If

				' Dispose of objects.

				DS.Dispose()
				DA.Dispose()

			End If

			' Arrange all the windows according to the arrange option

			If frmMain.mnuCascade.Checked Then frmMain.LayoutMdi(MdiLayout.Cascade)
			If frmMain.mnuTile.Checked Then frmMain.LayoutMdi(MdiLayout.TileHorizontal)
		End If
	End Sub

	'**********************************************************

	' The user has selected to remove a data table.

	'**********************************************************
	Private Sub mnuRemoveTable_Click(sender As Object, e As EventArgs) Handles mnuRemoveTable.Click, mnuRemoveSettings.Click

		' Declare variables

		Dim zx As String
		Dim Cmd As SqlCommand

		' Get the node from which the context menu was brought up

		Dim x As ContextMenuStrip = CType(sender.owner, ContextMenuStrip)
		Dim y As TreeNode = CType(x.SourceControl, TreeView).SelectedNode

		' This routine handles both removal of data tables and removal of system tables, but
		' we want a different warning message for each.

		If sender Is mnuRemoveTable Then
			zx = "Are you sure you want to remove the table """ & y.Name & """?" & vbCrLf & vbCrLf & "If you remove a table, you will lose all data it contains.  This action cannot be un-done."
		Else
			zx = "Are you sure you want to remove settings """ & y.Name & """?" & vbCrLf & vbCrLf & "If you remove this, all settings it contains will revert to their default values. This action cannot be un-done."
		End If

		' Make sure the user wants to proceed.

		If MsgBox(zx, MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Remove Data Table") = MsgBoxResult.Yes Then
			Cmd = New SqlCommand
			Cmd.CommandText = "DROP TABLE [dbo].[" & y.Name & "]"
			Cmd.Connection = DB

			' Attempt to execute the query to drop the table.

			Try
				Cmd.ExecuteNonQuery()
				RefreshObjectList()

				' Catch errors

			Catch ex As Exception
				MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Remove Table")
			End Try
		End If
	End Sub
	'**********************************************************

	' The user has selected to remove a query.

	'**********************************************************
	Private Sub mnuRemoveQuery_Click(sender As Object, e As EventArgs) Handles mnuRemoveQuery.Click

		' Declare variables

		Dim Cmd As SqlCommand

		' Get the node from which the context menu was brought up

		Dim x As ContextMenuStrip = CType(sender.owner, ContextMenuStrip)
		Dim y As TreeNode = CType(x.SourceControl, TreeView).SelectedNode

		' Make sure the user wants to proceed.

		If MsgBox("Are you sure you want to remove the query """ & y.Name & """?" & vbCrLf & vbCrLf & "If you remove a query, you may lose database function.", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Remove Query") = MsgBoxResult.Yes Then
			Cmd = New SqlCommand
			Cmd.CommandText = "DROP PROCEDURE [dbo].[" & y.Name & "]"
			Cmd.Connection = DB

			' Attempt to execute the query to drop the table.

			Try
				Cmd.ExecuteNonQuery()
				RefreshObjectList()

				' Catch errors

			Catch ex As Exception
				MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Remove Query")
			End Try
		End If
	End Sub
	'**********************************************************

	' The user has selected to remove a view.

	'**********************************************************
	Private Sub mnuRemoveView_Click(sender As Object, e As EventArgs) Handles mnuRemoveView.Click

		' Declare variables

		Dim Cmd As SqlCommand

		' Get the node from which the context menu was brought up

		Dim x As ContextMenuStrip = CType(sender.owner, ContextMenuStrip)
		Dim y As TreeNode = CType(x.SourceControl, TreeView).SelectedNode

		' Make sure the user wants to proceed.

		If MsgBox("Are you sure you want to remove the view """ & y.Name & """?" & vbCrLf & vbCrLf & "If you remove a view, you may lose database function.", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Remove View") = MsgBoxResult.Yes Then
			Cmd = New SqlCommand
			Cmd.CommandText = "DROP VIEW [dbo].[" & y.Name & "]"
			Cmd.Connection = DB

			' Attempt to execute the query to drop the table.

			Try
				Cmd.ExecuteNonQuery()
				RefreshObjectList()

				' Catch errors

			Catch ex As Exception
				MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Remove View")
			End Try
		End If
	End Sub
	'**********************************************************

	' The user has selected to remove an index.

	'**********************************************************
	Private Sub mnuRemoveIndex_Click(sender As Object, e As EventArgs) Handles mnuRemoveIndex.Click

		' Declare variables

		Dim Table As String
		Dim Index As String
		Dim zx As String
		Dim Cmd As SqlCommand

		' Get the node from which the context menu was brought up

		Dim x As ContextMenuStrip = CType(sender.owner, ContextMenuStrip)
		Dim y As TreeNode = CType(x.SourceControl, TreeView).SelectedNode

		' Parse the node text to extract the name of the table and the name of the key.

		zx = y.Text
		Table = ParseString(zx, " - (")
		Index = ParseString(zx, ")")

		' Make sure the user wants to proceed.

		If MsgBox("Are you sure you want to remove the index """ & y.Name & """?" & vbCrLf & vbCrLf & "If you remove and index, you may lose database function.", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Remove Index") = MsgBoxResult.Yes Then
			Cmd = New SqlCommand
			Cmd.CommandText = "DROP INDEX [" & Index & "] ON [" & Table & "];"
			Cmd.Connection = DB

			' Attempt to execute the query to drop the table.

			Try
				Cmd.ExecuteNonQuery()
				RefreshObjectList()

				' Catch errors

			Catch ex As Exception

				' If we got a message that the index could not be deleted because it was a primary key
				' then we need to alter the table to drop the constraint, in order to remove the index.

				If InStr(ex.Message, "PRIMARY KEY", CompareMethod.Text) > 0 Then
					Try
						Cmd.CommandText = "ALTER TABLE [dbo].[" & Table & "] DROP CONSTRAINT " & Index & ";"
						Cmd.ExecuteNonQuery()
						RefreshObjectList()

					Catch ex2 As Exception
						MsgBox(ex2.Message, MsgBoxStyle.Exclamation, "Remove Index")
					End Try
				Else
					MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Remove Index")
				End If
			End Try
		End If
	End Sub
	'**********************************************************

	' The user has selected to remove a Foreign Key.

	'**********************************************************
	Private Sub mnuModifyFK_Click(sender As Object, e As EventArgs) Handles mnuModifyFK.Click

		' Declare variables

		Dim zx As String
		Dim f As frmManageRelationships

		' Find the tree node to which this context menu is attached.

		Dim x As ContextMenuStrip = CType(sender.owner, ContextMenuStrip)
		Dim y As TreeNode = CType(x.SourceControl, TreeView).SelectedNode
		zx = VB.Left(y.Text, InStr(1, y.Text, " -") - 1) ' This will retrieve the name of the table

		' Show the relationships management form.

		f = New frmManageRelationships
		f.MdiParent = frmMain
		f.ShowRelationships(zx)
		f.Show()

	End Sub

	'**********************************************************

	' The user has selected to remove a Foreign Key.

	'**********************************************************
	Private Sub mnuRemoveFK_Click(sender As Object, e As EventArgs) Handles mnuRemoveFK.Click

		' Declare variables

		Dim Table As String
		Dim FK As String
		Dim zx As String
		Dim Cmd As SqlCommand

		' Get the node from which the context menu was brought up

		Dim x As ContextMenuStrip = CType(sender.owner, ContextMenuStrip)
		Dim y As TreeNode = CType(x.SourceControl, TreeView).SelectedNode

		' Parse the node text to extract the name of the table and the name of the key.

		zx = y.Text
		Table = ParseString(zx, " - (")
		FK = ParseString(zx, ")")

		' Make sure the user wants to proceed.

		If MsgBox("Are you sure you want to remove the foreign key """ & y.Name & """?" & vbCrLf & vbCrLf & "If you remove a foreign key, you may lose database function.", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Remove Foreign Key") = MsgBoxResult.Yes Then
			Cmd = New SqlCommand
			Cmd.CommandText = "ALTER TABLE [dbo].[" & Table & "] DROP CONSTRAINT " & FK & ";"
			Cmd.Connection = DB

			' Attempt to execute the query to drop the table.

			Try
				Cmd.ExecuteNonQuery()
				RefreshObjectList()

				' Catch errors

			Catch ex As Exception
				MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Remove Foreign Key")
			End Try
		End If
	End Sub

	'**********************************************************

	' The user has selected to build a new table using the
	' script builder.

	'**********************************************************
	Private Sub mnuScriptBuilder_Click(sender As Object, e As EventArgs) Handles mnuScriptBuilder.Click

		' Declare variables

		Dim zx As String

		' Assemble an SQL script that defines the ID field for a new table

		zx = "CREATE TABLE [dbo].[New Table Name]([ID]  INT  NOT NULL IDENTITY (1,1),PRIMARY KEY CLUSTERED ([ID] ASC))"

		' Display the script builder

		frmTableDesigner.MdiParent = frmMain
		frmTableDesigner.SQLScript = zx
		frmTableDesigner.Show()
		frmMain.LaunchProcess("ArrangeWindows")

	End Sub

	'**********************************************************

	' The user has selected to build a new object using the
	' script editor.

	'**********************************************************
	Private Sub mnuScriptEditor_Click(sender As Object, e As EventArgs) Handles mnuScriptEditor.Click


		' Declare variables

		Dim zx As String = ""
		Dim se As New frmScriptEditor

		' Determine the type of object to be built.

		Select Case TreeView1.SelectedNode.Text
			Case "Tables"
				zx = "CREATE TABLE [dbo].[NewTableName]" & vbCrLf & "(" & vbCrLf & "   [ID]  INT  NOT NULL IDENTITY (1,1)," & vbCrLf & "-- Field definitions must be inserted here" & vbCrLf & "   PRIMARY KEY CLUSTERED ([ID] ASC))" & vbCrLf & ")" & vbCrLf
			Case "Indexes"
				zx = "CREATE INDEX [dbo].[NewIndexName]" & vbCrLf & "(" & vbCrLf & "-- Field definitions must be inserted here" & vbCrLf & ")" & vbCrLf
			Case "Views"
				zx = "CREATE VIEW [dbo].[NewViewName]" & vbCrLf & "AS" & vbCrLf & "SELECT" & vbCrLf & "-- Field definitions must be inserted here" & vbCrLf & ")" & vbCrLf
			Case "Queries"
				zx = "CREATE PROCEDURE [dbo].[NewQueryName]" & vbCrLf & "AS" & vbCrLf & "SET NOCOUNT ON;" & vbCrLf & "-- Field definitions must be inserted here" & vbCrLf & ";" & vbCrLf

		End Select

		' Display the script editor

		se.DefaultText = zx
		se.MdiParent = frmMain
		se.Show()
		frmMain.LaunchProcess("ArrangeWindows")

	End Sub
	'**********************************************************

	' The View Schema menu option is selected.

	'**********************************************************
	Private Sub mnuViewSchema_Click(sender As Object, e As EventArgs) Handles mnuViewSchema.Click

		' Declare variables

		Dim zx As String
		Dim dv As New frmDataView

		' Find the tree node to which this context menu is attached.

		Dim x As ContextMenuStrip = CType(sender.owner, ContextMenuStrip)
		Dim y As TreeNode = CType(x.SourceControl, TreeView).SelectedNode
		zx = y.Text ' This will retrieve the name of the table

		' Open the system table that contains the names of the columns in the selected table, and 
		' the data type of each column.

		dv.SQLText = "SELECT * FROM information_schema.columns WHERE table_name = '" & zx & "' ORDER BY ordinal_position"
		dv.MdiParent = frmMain
		dv.Show()

	End Sub
	'**********************************************************

	' The Modify Table menu option was selected.

	'**********************************************************
	Private Sub mnuModifyTable_Click(sender As Object, e As EventArgs) Handles mnuModifyTable.Click

		' Declare variables

		Dim zx As String
		Dim Cmd As SqlCommand
		Dim DA As New SqlDataAdapter
		Dim DS As New DataSet
		Dim DT As DataTable

		' Find the tree node to which this context menu is attached.

		Dim x As ContextMenuStrip = CType(sender.owner, ContextMenuStrip)
		Dim y As TreeNode = CType(x.SourceControl, TreeView).SelectedNode
		zx = y.Text ' This will retrieve the name of the table

		' Open the system table that contains the names of the columns in the selected table, and 
		' the data type of each column.

		Cmd = New SqlCommand("SELECT * FROM information_schema.columns WHERE table_name = '" & zx & "' ORDER BY ordinal_position", DB)
		DA.SelectCommand = Cmd
		DA.Fill(DS, "Table")
		DT = DS.Tables("Table")


		' The system table exposes the following columns:
		' TABLE_NAME,COLUMN_NAME,COLUMN_DEFAULT,IS_NULLABLE,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH,NUMERIC_PRECISION,DATETIME_PRECISION

		' Pass this table to the Table Designer form for editing.

		Dim td As New frmTableDesigner
		td.Schema = DT
		td.MdiParent = frmMain
		td.Show()

		' Dispose of objects

		DA.Dispose()
		DS.Dispose()


	End Sub
	'**********************************************************

	' The Modify Index menu option was selected.

	'**********************************************************
	Private Sub mnuModifyIndex_Click(sender As Object, e As EventArgs) Handles mnuModifyIndex.Click

		' Declare variables

		Dim zx As String
		Dim Cmd As SqlCommand
		Dim DA As New SqlDataAdapter
		Dim DS As New DataSet
		Dim DT As DataTable

		' Find the tree node to which this context menu is attached.

		Dim x As ContextMenuStrip = CType(sender.owner, ContextMenuStrip)
		Dim y As TreeNode = CType(x.SourceControl, TreeView).SelectedNode
		zx = VB.Left(y.Text, InStr(1, y.Text, " -") - 1) ' This will retrieve the name of the table

		' Open the system table that contains the names of the columns in the selected table, and 
		' the data type of each column.

		Cmd = New SqlCommand("SELECT * FROM information_schema.columns WHERE table_name = '" & zx & "' ORDER BY ordinal_position", DB)
		DA.SelectCommand = Cmd
		DA.Fill(DS, "Table")
		DT = DS.Tables("Table")


		' The system table exposes the following columns:
		' TABLE_NAME,COLUMN_NAME,COLUMN_DEFAULT,IS_NULLABLE,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH,NUMERIC_PRECISION,DATETIME_PRECISION

		' Pass this table to the Table Designer form for editing.

		Dim td As New frmTableDesigner
		td.Schema = DT
		td.MdiParent = frmMain
		td.optIndex.Checked = True
		td.Show()

		' Dispose of objects

		DA.Dispose()
		DS.Dispose()


	End Sub
	'**********************************************************

	' The Modify Query menu option was selected.

	'**********************************************************

	Private Sub mnuModifyQuery_Click(sender As Object, e As EventArgs) Handles mnuModifyQuery.Click

		' Declare variables

		Dim zx As String
		Dim Cmd As SqlCommand
		Dim DA As New SqlDataAdapter
		Dim DS As New DataSet
		Dim DT As DataTable
		Dim se As New frmScriptEditor

		' Find the tree node to which this context menu is attached.

		Dim x As ContextMenuStrip = CType(sender.owner, ContextMenuStrip)
		Dim y As TreeNode = CType(x.SourceControl, TreeView).SelectedNode
		zx = y.Text ' This will retrieve the name of the query

		' Retrieving the SQL definition is slow.  Put on a wait cursor.

		Me.Cursor = Cursors.WaitCursor
		System.Windows.Forms.Application.DoEvents()

		' Open the system table that contains the SQL text of the selected query.

		Cmd = New SqlCommand("SELECT object_definition(object_id('[dbo].[" & zx & "]')) AS SQL_Text FROM sys.sql_modules", DB)
		DA.SelectCommand = Cmd
		DA.Fill(DS, "Table")
		DT = DS.Tables("Table")

		' Show the text of the query, replacing the command "create" with "alter"

		se.DefaultText = SRep(DT.Rows(0)("SQL_Text"), 1, "CREATE PROCEDURE", "ALTER PROCEDURE", CompareMethod.Text)
		se.MdiParent = frmMain
		se.Show()

		' Restore the normal cursor

		Me.Cursor = Cursors.Default
		System.Windows.Forms.Application.DoEvents()
	End Sub


	'**********************************************************

	' The "add  new" context menu is selected for a query, view
	' or index.  These can only be done with the script builder
	' and not the table designer, so we just pass this event
	' on to the event handler for the script builder menu.

	'**********************************************************
	Private Sub mnuAddNew_Click(sender As Object, e As EventArgs) Handles mnuAddNew.Click
		mnuScriptEditor_Click(mnuScriptEditor, New EventArgs)
	End Sub
	'**********************************************************

	' Sub to fill the object list with the list of tables and
	' other objects.

	'**********************************************************
	Public Sub RefreshObjectList()

		' Declare variables

		Dim jj As Integer
		Dim zx As String = ""
		Dim x As DataTable
		Dim n As TreeNode

		' Add a "tables" , "queries" and other nodes beneath the database.

		TreeView1.Nodes.Clear()
		TreeView1.Nodes.Add("Database", Databasename, "Database")
		TreeView1.Nodes("Database").Nodes.Add("Tables", "Tables", "DataTable")
		TreeView1.Nodes("Database").Nodes("Tables").ContextMenuStrip = ContextMenuStrip1a
		TreeView1.Nodes("Database").Nodes.Add("Queries", "Queries", "SQLQuery")
		TreeView1.Nodes("Database").Nodes("Queries").ContextMenuStrip = ContextMenuStrip1b
		TreeView1.Nodes("Database").Nodes.Add("Views", "Views", "View")
		TreeView1.Nodes("Database").Nodes("Views").ContextMenuStrip = ContextMenuStrip1b
		TreeView1.Nodes("Database").Nodes.Add("Indexes", "Indexes", "Index")
		TreeView1.Nodes("Database").Nodes("Indexes").ContextMenuStrip = ContextMenuStrip1b
		TreeView1.Nodes("Database").Nodes.Add("ForeignKeys", "Foreign Keys", "Key")
		TreeView1.Nodes("Database").Nodes.Add("Settings", "Workspace Settings", "Settings")
		TreeView1.Nodes("Database").Nodes.Add("SystemCatalog", "System Catalog", "Gears")

		' Now add all the tables beneath the database.  Available schema names are:
		' Databases,Procedures,Tables,Views,Columns,ForeignKeys,Indexes

		x = DB.GetSchema("Tables")
		For jj = 0 To x.Rows.Count - 1

			' Make sure the "table" is type BASE TABLE and not VIEW.  A view looks like a table
			' except for the table type, and we list views separately.

			If x.Rows(jj).Item("table_type") = "BASE TABLE" Then

				' If we encounter one of our settings tables, place them under the workspace settings,
				' even though they are really just ordinary tables.  That keeps them away from the
				' data tables.

				If VB.Left(x.Rows(jj).Item("table_name"), 3) = "db_" Then
					n = TreeView1.Nodes("Database").Nodes("Settings").Nodes.Add(x.Rows(jj).Item("table_name"), x.Rows(jj).Item("table_name"), "Settings")
					n.ContextMenuStrip = ContextMenuStrip7
				Else
					n = TreeView1.Nodes("Database").Nodes("Tables").Nodes.Add(x.Rows(jj).Item("table_name"), x.Rows(jj).Item("table_name"), "DataTable") ' Item index 2 is where the name is stored.
					n.ContextMenuStrip = ContextMenuStrip2
				End If
			End If
			System.Windows.Forms.Application.DoEvents()
		Next jj

		' Add all the queries beneath the database.

		x = DB.GetSchema("Procedures")
		For jj = 0 To x.Rows.Count - 1
			n = TreeView1.Nodes("Database").Nodes("Queries").Nodes.Add(x.Rows(jj).Item("specific_name"), x.Rows(jj).Item("specific_name"), "SQLQuery") ' Item index 2 is where the name is stored.
			n.ContextMenuStrip = ContextMenuStrip3
			System.Windows.Forms.Application.DoEvents()
		Next jj

		' Add all the views beneath the database.

		x = DB.GetSchema("Views")
		For jj = 0 To x.Rows.Count - 1
			n = TreeView1.Nodes("Database").Nodes("Views").Nodes.Add(x.Rows(jj).Item("table_name"), x.Rows(jj).Item("table_name"), "View") ' Item index 2 is where the name is stored.
			n.ContextMenuStrip = ContextMenuStrip4
			System.Windows.Forms.Application.DoEvents()
		Next jj

		' Add all the indexes beneath the database.  Ignore Unique constraints.

		x = DB.GetSchema("Indexes")
		For jj = 0 To x.Rows.Count - 1
			If VB.Left(x.Rows(jj).Item("index_name"), 2) <> "UQ" Then
				n = TreeView1.Nodes("Database").Nodes("Indexes").Nodes.Add(x.Rows(jj).Item("constraint_name"), x.Rows(jj).Item("table_name") & " - (" & x.Rows(jj).Item("constraint_name") & ")", "Index") ' Item index 2 is where the name is stored.
				n.ContextMenuStrip = ContextMenuStrip5
			End If
			System.Windows.Forms.Application.DoEvents()
		Next jj

		' Add all the Foreign Keys beneath the database.

		x = DB.GetSchema("ForeignKeys")
		For jj = 0 To x.Rows.Count - 1
			n = TreeView1.Nodes("Database").Nodes("ForeignKeys").Nodes.Add(x.Rows(jj).Item("constraint_name"), x.Rows(jj).Item("table_name") & " - (" & x.Rows(jj).Item("constraint_name") & ")", "Key") ' Item index 2 is where the name is stored.
			n.ContextMenuStrip = ContextMenuStrip6
			System.Windows.Forms.Application.DoEvents()
		Next jj

		' Add the System Catalog tables

		For jj = 1 To 13
			zx = Choose(jj, "sys.check_constraints", "sys.columns", "sys.default_constraints", "sys.foreign_keys", "sys.foreign_key_columns", "sys.identity_columns", "sys.index_columns", "sys.indexes", "sys.key_constraints", "sys.tables", "information_schema.columns", "sys.sql_modules", "sys.all_sql_modules")
			TreeView1.Nodes("Database").Nodes("SystemCatalog").Nodes.Add(zx, zx, "Gears")
		Next jj

		' Expand the database and table nodes to begin with.

		TreeView1.Nodes("Database").Expand()
		TreeView1.Nodes("Database").Nodes("Tables").Expand()

	End Sub

End Class
