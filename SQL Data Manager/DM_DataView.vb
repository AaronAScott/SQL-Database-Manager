Imports System.Data.SqlClient
Imports VB = Microsoft.VisualBasic
Imports System.IO

Public Class frmDataView
	'**********************************************************
	' SQL Data Manager Data View form (frmDataView)
	' DM_DATAVIEW.VB
	' Written: October 2017
	' Programmer: Aaron Scott
	' Copyright (C) 2017-2019 Sirius Software All Rights Reserved
	'**********************************************************

	' Declare variables local to this module

	Private LayoutChanged As Boolean = False
	Private SortedColumn As Integer = 0
	Private lObjectName As String
	Private lTableName As String
	Private lViewName As String
	Private lQueryName As String
	Private lCatalogName As String
	Private lSQLText As String
	Private DA As New SqlDataAdapter
	Private DS As New DataSet
	Private DT As DataTable
	Private FormSize As Size
	Private FormLocation As Point

	'**********************************************************

	' This sub overrides the default WndProc procedure, to
	' interecept a message that causes the MDI children windows
	' to get moved and resized whenever the main form is moved
	' or resized, which is REALLY annoying!

	'**********************************************************
	Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

		' Return control to base message handler.

		MyBase.WndProc(m)

		' If the window message was to move/resize the window, restore it to where
		' and now it was.

		If m.Msg = &H47 Then GetSavedLayout(lObjectName) ' WM_WINDOWPOSCHANGED

	End Sub
	'**********************************************************

	' The form is loaded.

	'**********************************************************
	Private Sub frmDataView_Load(sender As Object, e As EventArgs) Handles MyBase.Load

		' Set the window theme color.

		frmTheme.SetBackgroundColor(Me)

		' Remember the size

		FormSize = Me.Size

		' Get any saved layout for this object.

		GetSavedLayout(lObjectName)

		' Get any saved cell formatting information.

		DataGridView1.DefaultCellStyle.Font = New Font("Microsoft Sans Serif", 10)
		GetSavedFormat(lObjectName)

	End Sub

	'**********************************************************

	' The form is closed.

	'**********************************************************
	Private Sub frmDataView_FormClosed(sender As Object, e As EventArgs) Handles Me.FormClosed

		' Declare variables

		Dim xx As Integer
		Dim f As Form
		Dim Command As SqlCommand
		Dim CmdBld As New SqlCommandBuilder
		Dim dr As DataRow

		' Look for any open search or replace forms belonging to this form
		' and close them.

		For Each f In Application.OpenForms
			If f.Owner Is Me Then f.Close()
		Next f

		' Try to access the table to save the window settings.  The first time the layout is saved, we will have
		' create the table.

		If LayoutChanged Then
			If MsgBox("The layout of this window has changed.  Do you want to save the changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Save Window Layout") = MsgBoxResult.Yes Then
				Command = New SqlCommand("SELECT * FROM db_info_tables ORDER BY ObjectName", DB)
				DA.SelectCommand = Command

				' Use an SQLCommandBuilder to generate the other 3 types of commands for the data adapter.

				CmdBld.DataAdapter = DA

				' Watch for errors here. Under some circumstances, the command builder is unable+
				' to generate update commands. If that happens, the table cannot be updated, only viewed,

				On Error Resume Next
				DA.UpdateCommand = CmdBld.GetUpdateCommand
				DA.InsertCommand = CmdBld.GetInsertCommand
				DA.DeleteCommand = CmdBld.GetDeleteCommand

				' Now try to access the layout table.

				On Error GoTo LayoutTableNotFound
				DA.Fill(DS, "Table")
				DT = DS.Tables("Table")

				' Save the size, position and sort column of the object.

				On Error GoTo SaveLayoutError
				xx = Find2(DT, "ObjectName='" & lObjectName & "'")
				If xx <> NOMATCH Then
					dr = DT.Rows(xx)
				Else
					dr = DT.NewRow
					dr("ObjectName") = lObjectName
				End If

				dr("X") = Me.Left
				dr("Y") = Me.Top
				dr("Width") = Me.Width
				dr("Height") = Me.Height
				dr("SortColumn") = DataGridView1.SortedColumn.Index
				If xx = NOMATCH Then DT.Rows.Add(dr)
				DA.Update(DT)


				' Dispose of our objects

				DA.Dispose()
				DS.Dispose()
				Command.Dispose()

				On Error GoTo 0
			End If
		End If

		Exit Sub

		' Trap for errors saving the layout.

SaveLayoutError:
		MsgBox("Failed to save window layout." & vbCrLf & Err.Description, MsgBoxStyle.Exclamation, "Save Data View Layout")
		Exit Sub

		' Trap if the layout table has not been created yet.

LayoutTableNotFound:

		' Create the table and resume.

		Command = New SqlCommand("CREATE TABLE [dbo].[db_info_tables]( [ID]  INT NOT NULL IDENTITY (1,1),  [ObjectName]  NVARCHAR(50) NOT NULL,  [X]  INT NOT NULL, [Y]  INT NOT NULL, [Width]  INT NOT NULL, [Height]  INT NOT NULL,  [SortColumn]  INT NOT NULL, PRIMARY KEY CLUSTERED ([ID] ASC))", DB)
		Command.ExecuteNonQuery()
		Resume

	End Sub
	'**********************************************************

	' The form is resized

	'**********************************************************
	Private Sub frmDataView_Resize(sender As Object, e As EventArgs) Handles Me.Resize

		' Declare variables.

		Dim r As Rectangle

		' The following prevents the form from being maximized whenever the parent
		' form is moved or resized.

		If DontAllowResize Then Me.Size = FormSize Else FormSize = Me.Size

		' Resize the datagridview control to fill the resized window.

		r = New Rectangle(Me.ClientRectangle.Location.X, Me.ClientRectangle.Location.Y + MenuStrip1.Height + 1, Me.ClientRectangle.Width, Me.ClientRectangle.Height - MenuStrip1.Height - ToolStrip1.Height)
		DataGridView1.Bounds = r

	End Sub
	'**********************************************************

	' The form is moved.

	'**********************************************************
	Private Sub frmDataView_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged

		' The following prevents the form from being moved whenever the parent
		' form is moved or resized.

		If DontAllowResize Then Me.Location = FormLocation Else FormLocation = Me.Location

	End Sub

	'**********************************************************

	' The window has been resized or moved.

	'**********************************************************
	Private Sub frmDataView_LayoutChanged(sender As Object, e As EventArgs) Handles Me.ResizeEnd, Me.LocationChanged, DataGridView1.Sorted
		LayoutChanged = True
	End Sub
	'**********************************************************

	' The Search and Replace menu option is selected.

	'**********************************************************
	Private Sub mnuReplace_Click(sender As Object, e As EventArgs) Handles mnuReplace.Click

		' Declare variables

		Dim f As frmSearchReplace

		' Create a new search and replace form and set this form as its owner.

		f = New frmSearchReplace
		CType(f, frmSearchReplace).Table = DataGridView1
		f.Owner = Me
		f.Text = "Search and Replace " & lTableName
		f.Show()

	End Sub
	'**********************************************************

	' The Search menu option is selected.

	'**********************************************************
	Private Sub mnuSearch_Click(sender As Object, e As EventArgs) Handles mnuSearch.Click

		' Declare variables.

		Dim f As frmSearch

		' Create a new search form and set this form as its owner.

		f = New frmSearch
		f.Table = DataGridView1
		f.Owner = Me
		f.Text = "Search " & lTableName
		f.Show()

	End Sub
	'**********************************************************

	' A key is pressed in the form.

	'**********************************************************
	Private Sub frmDataView_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

		' Declare variables.

		Dim ii As Integer
		Dim jj As Integer
		Dim ww As String = ""
		Dim xx As String = ""
		Dim zx As String = ""
		Dim dr As DataRow = Nothing

		' See if the user is trying to paste data, with shift-insert.

		If e.Shift And e.KeyCode = Keys.Insert Then

			' See if the clipboard contains Comma Separated Value data, which could come from copying records
			' from another database via this program, or from Microsoft Access.

			If Clipboard.ContainsData("CSV") Then

				' If there was CSV data, get it as text.  This will, however, change all the commas to tabs.

				xx = Clipboard.GetData("Text")

				' The first line of the data will be column names, if it's valid table information.  Parse this and
				' see if it matches the columns in the displayed table.

				'ii = 0
				'DataMatchesTable = True
				'zx = ParseString(xx, vbCrLf)
				'Do While zx <> ""
				'	If ParseString(zx, vbTab) <> DataGridView1.Columns(ii).Name Then
				'		DataMatchesTable = False
				'		Exit Do
				'	End If
				'	ii += 1
				'Loop

				' If the data matches, we can attempt to paste it into the displayed dataset.
				Try
					jj = 0
					Do While xx <> ""
						zx = ParseString(xx, vbCrLf).TrimStart(vbTab)
						ii = 0
						dr = DT.NewRow
						Do While zx <> ""
							ww = ParseString(zx, vbTab)
							If ii > 0 Then ' Ignore first cell, which is always ID
								If ww <> "" Then
									dr(ii) = ww
								Else
									If Not DT.Columns(ii).AllowDBNull Then dr(ii) = ""
								End If
							End If
							ii += 1
							System.Windows.Forms.Application.DoEvents()
						Loop
						DT.Rows.Add(dr)
						jj += 1
					Loop
					DA.Update(DT)

				Catch ex As Exception
					MsgBox("Failed to insert data from clipboard." & vbCrLf & ex.Message, MsgBoxStyle.Information, "Paste Data Failure")

				End Try
			End If
		End If

	End Sub


	Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

		e.Cancel = True

	End Sub

	'**********************************************************

	' A data grid row is entered.

	'**********************************************************
	Private Sub DataGridView1_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.RowEnter
		If DT.Rows.Count > 0 Then StatusLabel.Text = "Record " & (e.RowIndex + 1) & " of " & DT.Rows.Count Else StatusLabel.Text = "Table is empty"
	End Sub

	'**********************************************************

	' The Move to First Record button is clicked.

	'**********************************************************
	Private Sub btnFirst_Click(sender As Object, e As EventArgs) Handles btnFirst.Click
		DataGridView1.FirstDisplayedScrollingRowIndex = 0
		StatusLabel.Text = "Record 1 of " & DT.Rows.Count
		DataGridView1.Rows(0).Cells(0).Selected = True
		DataGridView1.CurrentCell = DataGridView1.Rows(0).Cells(0)
	End Sub
	'**********************************************************

	' The Move to Last Record button is clicked.

	'**********************************************************
	Private Sub btnLast_Click(sender As Object, e As EventArgs) Handles btnLast.Click
		Dim xx As Integer = DataGridView1.DisplayedRowCount(True)
		DataGridView1.FirstDisplayedScrollingRowIndex = DT.Rows.Count - xx
		StatusLabel.Text = "Record " & DT.Rows.Count & " of " & DT.Rows.Count
		DataGridView1.Rows(DT.Rows.Count - 1).Cells(0).Selected = True
		DataGridView1.CurrentCell = DataGridView1.Rows(DT.Rows.Count - 1).Cells(0)
	End Sub


	'**********************************************************

	' A row of data has been validated.

	'**********************************************************
	Private Sub DataGridView1_RowValidated(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.RowValidated

		' Try to update the dataset.

		Try
			DA.Update(DT)

			' Display any errors.

		Catch ex As Exception
			DisplayMessage("Row not updated")
			DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
			Dim r As Rectangle = DataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True)
			ToolTip1.Show(ex.Message, Me, r.X, r.Y - DataGridView1.ColumnHeadersHeight * 2.5, 3000)
			DT.RejectChanges()
		End Try

		' Allow processes to complete

		System.Windows.Forms.Application.DoEvents()

	End Sub
	'**********************************************************

	' The font menu item is clicked.

	'**********************************************************
	Private Sub mnuFontSize_Click(sender As Object, e As EventArgs) Handles mnuFontSize.Click

		' Declare variables

		Dim f As Font

		' Show the font dialog box.

		f = DataGridView1.DefaultCellStyle.Font
		frmMain.FontDialog1.ShowDialog()
		f = frmMain.FontDialog1.Font
		DataGridView1.DefaultCellStyle.Font = f

		' Save the new font

		SaveCellFormat(lObjectName)
	End Sub
	'**********************************************************

	' The timer has ticked.  Clear any message and turn
	' timer back off.

	'**********************************************************
	Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
		MessageLabel.Text = ""
		MessageLabel.Image = Nothing
		Timer1.Enabled = False
	End Sub

	'**********************************************************

	' Sub to display a message on the status bar and
	' start the timer, which will clear it after 5 seconds.

	'**********************************************************
	Private Sub DisplayMessage(ByVal Msg As String)
		MessageLabel.Text = VB.Left(Msg, 80)
		MessageLabel.Image = My.Resources.Exclamation
		Timer1.Enabled = True
	End Sub
	'**********************************************************

	' Sub to restore the object's saved sizes and positions.

	'**********************************************************
	Private Sub GetSavedLayout(ObjectName As String)

		' Declare variables

		Dim xx As Integer
		Dim Command As SqlCommand
		Dim DA As New SqlDataAdapter
		Dim DS As New DataSet
		Dim DT As DataTable

		' Save the name of the table, query, view, etc.

		lObjectName = ObjectName

		' Now see if we have a saved layout.

		Try

			Command = New SqlCommand("SELECT * FROM [dbo].[db_info_tables] ORDER BY ObjectName", DB)
			DA.SelectCommand = Command
			DS.Clear()
			DA.Fill(DS, "table")
			DT = DS.Tables("table")

			' See if we have a saved layout for this object.

			xx = Find2(DT, "ObjectName='" & ObjectName & "'")
			If xx <> NOMATCH Then
				Me.Left = DT.Rows(xx)("X")
				Me.Top = DT.Rows(xx)("Y")
				Me.Width = DT.Rows(xx)("Width")
				Me.Height = DT.Rows(xx)("Height")
				SortedColumn = DT.Rows(xx)("SortColumn")
				DataGridView1.Sort(DataGridView1.Columns(SortedColumn), System.ComponentModel.ListSortDirection.Ascending)
			End If

			' Ignore errors--we may not have a saved layout.

		Catch ex As Exception

		End Try

		' Indicate the layout has not changed.

		LayoutChanged = False

	End Sub
	'**********************************************************

	' Sub to restore the object's saved cell formatting information.

	'**********************************************************
	Private Sub GetSavedFormat(ObjectName As String)

		' Declare variables

		Dim xx As Integer
		Dim Command As SqlCommand
		Dim DA As New SqlDataAdapter
		Dim DS As New DataSet
		Dim DT As DataTable

		' Save the name of the table, query, view, etc.

		lObjectName = ObjectName

		' Now see if we have a saved layout.

		Try

			Command = New SqlCommand("SELECT * FROM [dbo].[db_info_format] ORDER BY ObjectName", DB)
			DA.SelectCommand = Command
			DS.Clear()
			DA.Fill(DS, "table")
			DT = DS.Tables("table")

			' See if we have a saved layout for this object.

			xx = Find2(DT, "ObjectName='" & ObjectName & "'")
			If xx <> NOMATCH Then
				DataGridView1.DefaultCellStyle.Font = New Font(CStr(DT.Rows(xx)("FontName")), CSng(DT.Rows(xx)("FontSize")))
				DataGridView1.DefaultCellStyle.ForeColor = Color.FromArgb(CInt(DT.Rows(xx)("Color")))
			End If

			' Ignore errors--we may not have saved format information.

		Catch ex As Exception

		End Try

	End Sub
	'**********************************************************

	' Sub to save changed cell format information.

	'**********************************************************

	Private Sub SaveCellFormat(ObjectName As String)

		' Declare variables

		Dim xx As Integer
		Dim Command As SqlCommand
		Dim CmdBld As New SqlCommandBuilder
		Dim DA As New SqlDataAdapter
		Dim DS As New DataSet
		Dim DT As DataTable
		Dim dr As DataRow

		' Try to access the table to save the window settings.  The first time the layout is saved, we will have
		' create the table.

		Command = New SqlCommand("SELECT * FROM db_info_format ORDER BY ObjectName", DB)
		DA.SelectCommand = Command

		' Use an SQLCommandBuilder to generate the other 3 types of commands for the data adapter.

		CmdBld.DataAdapter = DA

		' Watch for errors here. Under some circumstances, the command builder is unable+
		' to generate update commands. If that happens, the table cannot be updated, only viewed,

		On Error Resume Next
		DA.UpdateCommand = CmdBld.GetUpdateCommand
		DA.InsertCommand = CmdBld.GetInsertCommand
		DA.DeleteCommand = CmdBld.GetDeleteCommand

		' Now try to access the format table.

		On Error GoTo LayoutTableNotFound
		DA.Fill(DS, "Table")
		DT = DS.Tables("Table")

		' Save the font name and size of the object.

		On Error GoTo SaveFormatError
		xx = Find2(DT, "ObjectName='" & lObjectName & "'")
			If xx <> NOMATCH Then
				dr = DT.Rows(xx)
			Else
				dr = DT.NewRow
				dr("ObjectName") = lObjectName
			End If

		dr("FontName") = DataGridView1.DefaultCellStyle.Font.Name
		dr("FontSize") = DataGridView1.DefaultCellStyle.Font.SizeInPoints
		dr("Color") = Color.Black.ToArgb
		If xx = NOMATCH Then DT.Rows.Add(dr)
		DA.Update(DT)


		' Dispose of our objects

		DA.Dispose()
		DS.Dispose()
		Command.Dispose()

		On Error GoTo 0

		Exit Sub

		' Trap for errors saving the layout.

SaveFormatError:
		MsgBox("Failed to save cell format information." & vbCrLf & Err.Description, MsgBoxStyle.Exclamation, "Save Cell Format")
		Exit Sub

		' Trap if the layout table has not been created yet.

LayoutTableNotFound:

		' Create the table and resume.

		Command = New SqlCommand("CREATE TABLE [dbo].[db_info_Format]( [ID]  INT NOT NULL IDENTITY (1,1),  [ObjectName]  NVARCHAR(50) NOT NULL,  [FontName]  NVARCHAR(25) NOT NULL, [FontSize]  FLOAT NOT NULL, [Color] INT DEFAULT ((0)), PRIMARY KEY CLUSTERED ([ID] ASC))", DB)
		Command.ExecuteNonQuery()
		Resume


	End Sub

	'**********************************************************

	' The TableName Property

	'**********************************************************
	Public Property TableName
		Get
			TableName = lTableName
		End Get

		Set(value)

			' Declare variables

			Dim Message As String = ""
			Dim CmdBld As New SqlCommandBuilder

			' Save the name of the table in the object name.

			lObjectName = value

			' Save the table name in the local variable, and set the caption of the form

			lTableName = value
			Me.Text = "Table: " & lTableName

			' Build an SQL command to select all the data in the specified table

			Dim Command = New SqlCommand("SELECT * FROM [dbo].[" & lTableName & "]", DB)
			DA.SelectCommand = Command

			' Use an SQLCommandBuilder to generate the other 3 types of commands for the data adapter.

			CmdBld.DataAdapter = DA

			' Watch for errors here. Under some circumstances, the command builder is unable
			' to generate update commands. If that happens, the table cannot be updated, only viewed,

			Try
				DA.UpdateCommand = CmdBld.GetUpdateCommand
				DA.InsertCommand = CmdBld.GetInsertCommand
				DA.DeleteCommand = CmdBld.GetDeleteCommand


				' Fill the dataset and point to the table, then set the
				' table as the data source of the control.

				DA.Fill(DS, "Table")
				DT = DS.Tables("Table")
				DataGridView1.DataSource = DT

				' The initial view will sort by the first field of each table, which is always the record ID.
				' The puts records into the order in which they were entered.  But if a layout was saved,
				' it will sort in the order of the saved layout.

				DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)

				' Show the record count on the status strip

				StatusLabel.Text = DT.Rows.Count & " records"

			Catch ex As Exception
				MsgBox(ex.Message, MsgBoxStyle.Information)
			End Try
		End Set
	End Property

	'**********************************************************

	' The ViewName Property

	'**********************************************************
	Public Property ViewName
		Get
			ViewName = lViewName
		End Get

		Set(value)

			' Declare variables

			Dim zx1 As String

			' Save the view name in the local variable, and set the caption of the form

			lViewName = value
			Me.Text = "View: " & lViewName

			' Build an SQL command to select all the data in the specified table

			zx1 = "SELECT [dbo].* FROM " & lViewName
			Dim Command = New SqlCommand("SELECT * FROM [dbo].[" & lViewName & "]", DB)
			DA.SelectCommand = Command

			' Fill the dataset and point to the table, then set the
			' table as the data source of the control.

			DA.Fill(DS, "Table")
			DT = DS.Tables("Table")
			DataGridView1.DataSource = DT

			' Show the record count on the status strip

			StatusLabel.Text = DT.Rows.Count & " records"

		End Set
	End Property

	'**********************************************************

	' The QueryName Property

	'**********************************************************
	Public Property QueryName
		Get
			QueryName = lQueryName
		End Get

		Set(value)

			' Declare variables

			Dim QueryScript As String
			Dim CmdBld As New SqlCommandBuilder

			' Save the query name in the local variable, and set the caption of the form

			lQueryName = value
			Me.Text = "Query: " & lQueryName

			' Open the system table that contains the SQL text of the selected query.

			Dim Command = New SqlCommand("SELECT object_definition(object_id('[dbo].[" & lQueryName & "]')) AS SQL_Text FROM sys.sql_modules", DB)
			DA.SelectCommand = Command
			DA.Fill(DS, "Table")
			DT = DS.Tables("Table")

			' Make sure we got a valid return of the query definition.

			If DT.Rows.Count > 0 Then

				' Determine the query type.

				' The query script will begin with CREATE PROCEDURE or ALTER PROCEDURE.  We need to strip out
				' just the query itself to execute it.

				QueryScript = Mid(DT.Rows(0)("SQL_Text"), InStr(1, DT.Rows(0)("SQL_Text"), "SELECT ", CompareMethod.Text))

				' Build an SQL command containing the stored procedure.

				Command = New SqlCommand(QueryScript, DB)
				Command.CommandType = CommandType.Text
				DA = New SqlDataAdapter
				DS = New DataSet
				DA.SelectCommand = Command

				' Use an SQLCommandBuilder to generate the other 3 types of commands for the data adapter.

				CmdBld.DataAdapter = DA

				' Watch for errors here. Under some circumstances, the command builder is unable
				' to generate update commands. If that happens, the table cannot be updated, only viewed,

				Try
					DA.UpdateCommand = CmdBld.GetUpdateCommand
					DA.InsertCommand = CmdBld.GetInsertCommand
					DA.DeleteCommand = CmdBld.GetDeleteCommand
				Catch ex As Exception
				End Try


				' Fill the dataset and point to the table, then set the
				' table as the data source of the control.

				Try
					DA.Fill(DS, "Table")
					DT = DS.Tables("Table")
					DataGridView1.DataSource = DT

					' Show the record count on the status strip

					StatusLabel.Text = DT.Rows.Count & " records"

				Catch ex As Exception
					MsgBox("Failed to retrieve recordds." & vbCrLf & ex.Message, MsgBoxStyle.Information, "Query Failed")
				End Try
			End If
			DA.Dispose()
			DS.Dispose()

		End Set
	End Property
	'**********************************************************

	' The CatalogName Property.  This allows us to look at
	' system catalogs.

	'**********************************************************
	Public Property CatalogName
		Get
			CatalogName = lCatalogName
		End Get

		Set(value)

			' Declare variables

			Dim Message As String = ""

			' Save the table name in the local variable, and set the caption of the form

			lCatalogName = value
			Me.Text = "System Catalog: " & CatalogName

			' Build an SQL command to select all the data in the specified table

			Dim Command = New SqlCommand("SELECT * FROM " & lCatalogName, DB)
			DA.SelectCommand = Command

			' Watch for errors here. We may not modify system objects: the table cannot be updated, only viewed,

			Try

				' Fill the dataset and point to the table, then set the
				' table as the data source of the control.

				DA.Fill(DS, "Table")
				DT = DS.Tables("Table")
				DataGridView1.DataSource = DT

				' Show the record count on the status strip

				StatusLabel.Text = DT.Rows.Count & " records"

			Catch ex As Exception
				MsgBox(ex.Message, MsgBoxStyle.Information)
			End Try
		End Set
	End Property
	'**********************************************************

	' The SQL Text Property. This property allows us to
	' send a text query to the form for viewing.

	'**********************************************************
	Public Property SQLText
		Get
			SQLText = lSQLText
		End Get

		Set(value)

			' Save the view name in the local variable, and set the caption of the form

			lSQLText = value
			Me.Text = "View SQL Command Output"


			' Build an SQL command to select all the data in the specified table

			Dim Command = New SqlCommand(lSQLText, DB)
			DA.SelectCommand = Command

			' Fill the dataset and point to the table, then set the
			' table as the data source of the control.

			Try
				DA.Fill(DS, "Table")
				DT = DS.Tables("Table")
				DataGridView1.DataSource = DT

				' Show the record count on the status strip

				StatusLabel.Text = DT.Rows.Count & " records"

			Catch ex As Exception
				MsgBox(ex.Message, MsgBoxStyle.Exclamation, "View Query")
			End Try

		End Set
	End Property


End Class