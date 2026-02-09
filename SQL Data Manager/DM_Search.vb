Imports VB = Microsoft.VisualBasic
Public Class frmSearch
    '**********************************************************
    ' SQL Data Manager Search Data form.
    ' DM_SEARCH.VB
    ' Written: December 2018
    ' Programmer: Aaron Scott
    ' Copyright (C) 2017-2018 Sirius Software All Rights Reserved
    '**********************************************************

    ' Declare variables that are the properties of this module.

    Public Table As DataGridView

    ' Declare variables local to this module

    Private CurrentRowIndex As Integer
    Private CurrentColumnIndex As Integer
    Private StartColumn As Integer
    Private EndColumn As Integer

    '**********************************************************

    ' The form is loaded.

    '**********************************************************
    Private Sub frmSearch_Load(sender As Object, e As EventArgs) Handles Me.Load

        ' Declare variables.

        Dim jj As Integer

        ' Fill the combo box with the list of columns.  Automatically select
        ' the current column as the default search column.

        ComboBox1.Items.Add("Entire Table")
        For jj = 0 To Table.ColumnCount - 1
            ComboBox1.Items.Add(Table.Columns(jj).Name)
            If jj = Table.CurrentCell.ColumnIndex Then ComboBox1.SelectedIndex = jj + 1
        Next jj

    End Sub

    '**********************************************************

    ' The Find button is clicked.

    '**********************************************************
    Private Sub btnFind_Click(sender As Object, e As EventArgs) Handles btnFind.Click

        ' Declare variables

        Dim Found As Boolean
        Dim ii As Integer
        Dim jj As Integer
        Dim zx As String
        Dim cell As DataGridViewCell

        ' Disable all buttons while we search to prevent reentrancy problems.

        btnFind.Enabled = False
        btnFindNext.Enabled = False
        btnClose.Enabled = False

        ' Clear the flag that indicates we've found a match.

        Found = False

        ' Begin searching from the first column of the first row.

        For ii = 0 To Table.Rows.Count - 1
            For jj = StartColumn To EndColumn
                cell = Table.Rows(ii).Cells(jj)

                ' Get the cell contents as a string, for comparison.

                zx = cell.FormattedValue.ToString

                ' Check for a match based on the selected option.  If we find a match, we make
                ' that the current cell, which automatically selects it in the table.

                If optStart.Checked Then ' Match start of field.
                    If VB.Left(zx.ToLower, Len(TextBox1.Text)) = TextBox1.Text.ToLower Then
                        Found = True
                        Table.CurrentCell = cell
                        Exit For
                    End If
                ElseIf optEntire.Checked Then ' Match entire field.
                    If zx.ToLower = TextBox1.Text.ToLower Then
                        Found = True
                        Table.CurrentCell = cell
                        Exit For
                    End If
                Else
                    ' Match any part of field.

                    If InStr(1, zx, TextBox1.Text, CompareMethod.Text) > 0 Then
                        Found = True
                        Table.CurrentCell = cell
                        Exit For
                    End If
                End If
            Next jj

            ' If we found a match exit the loop.

            If Found Then Exit For
        Next ii


        ' If we found no match, there were no matches.

        If Not Found Then
            MsgBox("Search string not found", MsgBoxStyle.Information, "Search Table")

            ' Otherwise, remember the current row and column.

        Else
            CurrentColumnIndex = jj
            CurrentRowIndex = ii
        End If

        ' Renable all the buttons.

        btnFind.Enabled = True
        btnFindNext.Enabled = True
        btnClose.Enabled = True

    End Sub
    '**********************************************************

    ' The Find Next button is clicked.

    '**********************************************************
    Private Sub btnFindNext_Click(sender As Object, e As EventArgs) Handles btnFindNext.Click

        Dim Found As Boolean
        Dim ii As Integer
        Dim jj As Integer
        Dim kk As Integer
        Dim zx As String
        Dim cell As DataGridViewCell

        ' Disable all buttons while we search to prevent reentrancy problems.

        btnFind.Enabled = False
        btnFindNext.Enabled = False
        btnClose.Enabled = False

        ' Clear the flag that indicates we've found a match.

        Found = False

        ' If we're searching the entire table, then our search column
        ' must begin in the next column from the column where the
        ' last match was found.  Otherwise, me must start in the
        ' next row.

        If StartColumn <> EndColumn Then
            kk = CurrentColumnIndex + 1
            If kk >= Table.ColumnCount Then kk = 0
        Else
            kk = StartColumn
            CurrentRowIndex += 1
        End If

        ' Begin searching from the next column of the current row.

        For ii = CurrentRowIndex To Table.Rows.Count - 1
            For jj = kk To EndColumn
                cell = Table.Rows(ii).Cells(jj)

                ' Get the cell contents as a string, for comparison.

                zx = cell.FormattedValue.ToString

                ' Check for a match based on the selected option.  If we find a match, we make
                ' that the current cell, which automatically selects it in the table.

                If optStart.Checked Then ' Match start of field.
                    If VB.Left(zx.ToLower, Len(TextBox1.Text)) = TextBox1.Text.ToLower Then
                        Found = True
                        Table.CurrentCell = cell
                        Exit For
                    End If
                ElseIf optEntire.Checked Then ' Match entire field.
                    If zx.ToLower = TextBox1.Text.ToLower Then
                        Found = True
                        Table.CurrentCell = cell
                        Exit For
                    End If
                Else ' Match any part of field.

                    If InStr(1, zx, TextBox1.Text, CompareMethod.Text) > 0 Then
                        Found = True
                        Table.CurrentCell = cell
                        Exit For
                    End If
                End If
            Next jj

            ' If we found a match, exit the loop.

            If Found Then Exit For
        Next ii

        ' If nothing found, we have found all occurences.

        If Not Found Then
            MsgBox("No more occurrences found.", MsgBoxStyle.Information, "Search Table")

            ' Otherwise, remember our current row and column.

        Else
            CurrentColumnIndex = jj
            CurrentRowIndex = ii
        End If

        ' Re-enable all buttons.

        btnFind.Enabled = True
        btnFindNext.Enabled = True
        btnClose.Enabled = True
    End Sub
    '**********************************************************

    ' The close button is clicked.

    '**********************************************************
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    '**********************************************************

    ' The search field has changed.

    '**********************************************************
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        If ComboBox1.SelectedIndex = 0 Then
            StartColumn = 0
            EndColumn = Table.ColumnCount - 1
        Else
            StartColumn = ComboBox1.SelectedIndex - 1
            EndColumn = StartColumn
        End If

    End Sub
End Class