Imports VB = Microsoft.VisualBasic
Public Class frmAddTable
	'**********************************************************************
	' SQL Data Manager Add Table to Relationships form
	' DM_ADDTABLE.VB
	' Written: November 2018
	' Programmer: Aaron Scott
	' Copyright (C) 2017-2018 Sirius Software All Rights Reserved
	'**********************************************************************


	'**********************************************************************

	' The form is loaded.

	'**********************************************************************
	Private Sub frmAddTable_Load(sender As Object, e As EventArgs) Handles Me.Load

		' Declare variables

		Dim jj As Integer
		Dim x As DataTable

		' Loop through all the tables in the database, checking to see if they
		' are already displayed in the relationships window, and adding them
		' to the list box of tables available to add if they are not.

		ListBox1.Items.Clear()
		x = DB.GetSchema("Tables")
		For jj = 0 To x.Rows.Count - 1

			' Make sure the "table" is type BASE TABLE and not VIEW.  A view looks like a table
			' except for the table type, and we list views separately.

			If x.Rows(jj).Item("table_type") = "BASE TABLE" Then

				' If we encounter one of our settings tables, ignore it.

				If VB.Left(x.Rows(jj).Item("table_name"), 3) <> "db_" Then

					' See if the relationships form already contains the table name.

					If Not frmRelationships.ContainsTable(x.Rows(jj).Item("table_name")) Then
						ListBox1.Items.Add(x.Rows(jj).Item("table_name"))
					End If
				End If
			End If
			System.Windows.Forms.Application.DoEvents()
		Next jj
	End Sub

	'**********************************************************************

	' The Add Table button is clicked.

	'**********************************************************************
	Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click

		' Declare variables

		Dim zx As String

		' Get the name of the selected table.

		zx = ListBox1.Items(ListBox1.SelectedIndex)

		' Add the table to the relationships window.

		frmRelationships.AddTable(zx)

		' Remove the table from the list box.

		ListBox1.Items.Remove(zx)

		' Disable the add button until the next selection.

		btnAdd.Enabled = False


	End Sub

	'**********************************************************************

	' The Close button is clicked.

	'**********************************************************************
	Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

		Me.Close()

	End Sub

	'**********************************************************************

	' A table name is selected in the list box.  Enable the Add button.
	'**********************************************************************
	Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
		btnAdd.Enabled = True
	End Sub

	'**********************************************************************

	' A table name is double-clicked.  Add it to the relationships window.

	'**********************************************************************
	Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
		btnAdd_Click(btnAdd, New EventArgs)
	End Sub
End Class