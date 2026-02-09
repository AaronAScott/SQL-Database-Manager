Imports VB = Microsoft.VisualBasic
Imports System.Data.SqlClient
Public Class frmSearchReplace
	'**********************************************************
	' SQL Data Manager Search Data and Replace form.
	' DM_SEARCHREPLACE.VB
	' Written: December 2018
	' Programmer: Aaron Scott
	' Copyright (C) 2017-2018 Sirius Software All Rights Reserved
	'**********************************************************

	' Declare variables that are the properties of this module.

	Public Table As DataGridView

	' Declare variables local to this module

	Private FoundFirst As Boolean
	Private CurrentRowIndex As Integer
	Private CurrentColumnIndex As Integer

	' Declare a type for saving undo information for each cell replaced.

	Private Structure CellInfo
		Dim Row As Integer
		Dim Column As Integer
		Dim FormattedValue As Object
	End Structure

	' Create a collection of information for all cells changed, so they can be undone.

	Private UndoInfo As New Collection

	'**********************************************************

	' The form is loaded.

	'**********************************************************
	Private Sub frmSearchReplace_Load(sender As Object, e As EventArgs) Handles Me.Load

		' Clear the flag that indicates we found at least one match.

		FoundFirst = False
		btnReplace.Text = "&Find"
		CurrentColumnIndex = -1
		CurrentRowIndex = 0
	End Sub
	'**********************************************************

	' A key is pressed in the form.

	'**********************************************************
	Private Sub frmSearchReplace_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown


		' If ctrl-z is pressed, cancel changes.

		If e.Control And e.KeyValue = Keys.Z Then
			btnCancel_Click(btnUndo, New EventArgs)
		End If
	End Sub
	'**********************************************************

	' The Replace button is clicked.

	'**********************************************************
	Private Sub btnReplace_Click(sender As Object, e As EventArgs) Handles btnReplace.Click

		' Declare variables

		Dim Found As Boolean
		Dim ii As Integer
		Dim jj As Integer
		Dim kk As Integer
		Dim zx As String
		Dim t As String
		Dim ci As CellInfo
		Dim cell As DataGridViewCell
		Dim comp As CompareMethod = CompareMethod.Text

		' Disable all buttons while we search to prevent reentrancy problems.

		btnReplace.Enabled = False
		btnReplaceAll.Enabled = False
		btnClose.Enabled = False

		' Determine if we are doing a case-sensitive search.

		If CheckBox1.Checked Then comp = CompareMethod.Binary

		' Clear the flag that indicates we've found a match.

		Found = False

		' If we've already found a match, then make the replacement before we look
		' for the next match.

		If FoundFirst Then

			' Save current cell information so we can undo changes if necessary.

			ci = New CellInfo
			ci.Row = Table.CurrentCell.RowIndex
			ci.Column = Table.CurrentCell.ColumnIndex
			ci.FormattedValue = Table.CurrentCell.FormattedValue
			UndoInfo.Add(ci)

			' Make the replacement.

			Table.CurrentCell.Value = SRep(Table.CurrentCell.FormattedValue.ToString, 1, TextBox1.Text, TextBox2.Text, comp)

			' Enable the Undo button.

			btnUndo.Enabled = True
		End If

		' Begin searching in the next column from the column where the
		' last match was found.

		kk = CurrentColumnIndex + 1

		' Begin searching from the next column of the current row.

		For ii = CurrentRowIndex To Table.Rows.Count - 1
			For jj = kk To Table.ColumnCount - 1
				cell = Table.Rows(ii).Cells(jj)

				' Get the cell contents as a string, for comparison.

				If CheckBox1.Checked Then
					t = TextBox1.Text
					zx = cell.FormattedValue.ToString
				Else
					t = TextBox1.Text.ToLower
					zx = cell.FormattedValue.ToString.ToLower
				End If

				' Check for a match based on the selected option.  If we find a match, we make
				' that the current cell, which automatically selects it in the table.

				If optStart.Checked Then ' Match start of field.
					If VB.Left(zx, t.Length) = t Then
						Found = True
						Table.CurrentCell = cell
						Exit For
					End If
				ElseIf optEntire.Checked Then ' Match entire field.
					If zx = t Then
						Found = True
						Table.CurrentCell = cell
						Exit For
					End If
				Else ' Match any part of field.

					If InStr(1, zx, t, comp) > 0 Then
						Found = True
						Table.CurrentCell = cell
						Exit For
					End If
				End If
			Next jj

			' If we found a match, exit the loop.

			If Found Then Exit For
			kk = 0
		Next ii


		' If we found no match, there were no, or no more, matches.

		If Not Found Then
			If Not FoundFirst Then
				MsgBox("Search string not found.", MsgBoxStyle.Information, "Search and Replace")
			Else
				MsgBox("No more occurrences.", MsgBoxStyle.Information, "Search and Replace")
			End If

			' Otherwise, remember the current row and column.

		Else
			CurrentColumnIndex = jj
			CurrentRowIndex = ii

			' If we've found at least one match, change the button text to "Replace"

			btnReplace.Text = "&Replace"
			FoundFirst = True
		End If

		' Renable all the buttons.

		btnReplace.Enabled = True
		btnReplaceAll.Enabled = True
		btnClose.Enabled = True

	End Sub
	'**********************************************************

	' The Replace All button is clicked.

	'**********************************************************
	Private Sub btnReplaceAll_Click(sender As Object, e As EventArgs) Handles btnReplaceAll.Click

		' Declare variables

		Dim Found As Boolean
		Dim ii As Integer
		Dim jj As Integer
		Dim ReplacedCount As Integer
		Dim zx As String
		Dim t As String
		Dim ci As CellInfo
		Dim cell As DataGridViewCell
		Dim comp As CompareMethod = CompareMethod.Text

		' Disable all buttons while we search to prevent reentrancy problems.

		btnReplace.Enabled = False
		btnReplaceAll.Enabled = False
		btnClose.Enabled = False

		' Determine if we are doing a case-sensitive search.

		If CheckBox1.Checked Then comp = CompareMethod.Binary

		' Clear the flag that indicates we've found a match.

		Found = False

		' Start a new collection of undo information.

		UndoInfo = New Collection

		' Begin searching from the next column of the current row.

		For ii = 0 To Table.Rows.Count - 1
			For jj = 0 To Table.ColumnCount - 1
				cell = Table.Rows(ii).Cells(jj)

				' Get the cell contents as a string, for comparison.

				If CheckBox1.Checked Then
					t = TextBox1.Text
					zx = cell.FormattedValue.ToString
				Else
					t = TextBox1.Text.ToLower
					zx = cell.FormattedValue.ToString.ToLower
				End If

				' Check for a match based on the selected option.  If we find a match, we make
				' that the current cell, which automatically selects it in the table.

				If optStart.Checked Then ' Match start of field.
					If VB.Left(zx, t.Length) = t Then
						Found = True
						Table.CurrentCell = cell
						Exit For
					End If
				ElseIf optEntire.Checked Then ' Match entire field.
					If zx = t Then
						Found = True
						Table.CurrentCell = cell
						Exit For
					End If
				Else ' Match any part of field.

					If InStr(1, zx, t, comp) > 0 Then
						Found = True
						Table.CurrentCell = cell
						Exit For
					End If
				End If
			Next jj

			' If we found a match, make the replacement.

			If Found Then

				' Save current cell information so we can undo changes if necessary.

				ci = New CellInfo
				ci.Row = Table.CurrentCell.RowIndex
				ci.Column = Table.CurrentCell.ColumnIndex
				ci.FormattedValue = Table.CurrentCell.FormattedValue
				UndoInfo.Add(ci)

				' Make the replacement.

				Table.CurrentCell.Value = SRep(Table.CurrentCell.FormattedValue.ToString, 1, TextBox1.Text, TextBox2.Text, comp)

				' Count the replacements made.

				ReplacedCount += 1

				' Enable the Undo button.

				btnUndo.Enabled = True

			End If
			Found = False
		Next ii

		' Report the number of replacements made.

		MsgBox("Search completed. " & ReplacedCount & " replacment(s) made.", MsgBoxStyle.Information, "Search and Replace")
		ReplacedCount = 0

		' Renable all the buttons.

		btnReplace.Enabled = True
		btnReplaceAll.Enabled = True
		btnClose.Enabled = True

	End Sub
	'**********************************************************

	' The close button is clicked.

	'**********************************************************
	Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

		' Close the form.

		Me.Close()
	End Sub

	'**********************************************************

	' The cancel button is clicked.

	'**********************************************************
	Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnUndo.Click
		UndoChanges()
		Table.Refresh()
		Me.Close()
	End Sub
	'**********************************************************

	' The search-for text has changed.

	'**********************************************************
	Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

		' If we have undo information already, clear it out: it's
		' committed once we begin another search/replace operation.

		If UndoInfo.Count > 0 Then
			UndoInfo = New Collection
			btnUndo.Enabled = False
		End If
	End Sub
	'**********************************************************

	' Sub to undo any changes made to a table.

	'**********************************************************
	Private Sub UndoChanges()

		' Declare variables.

		Dim c As CellInfo

		' Loop through the collection of all changed cells and restore
		' their original value.

		For Each c In UndoInfo
			Table.Rows(c.Row).Cells(c.Column).Value = c.FormattedValue
		Next c

		' Clear out the Undo information collection.

		UndoInfo = New Collection

		' Disable the Undo button.

	End Sub


End Class