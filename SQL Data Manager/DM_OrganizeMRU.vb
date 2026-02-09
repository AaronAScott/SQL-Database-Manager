Public Class frmOrganizeMRU

	'**********************************************************
	' SQL Data Manager Organize MRU list
	' DM_ORGANIZEMRU.VB
	' Written: October 2017
	' Programmer: Aaron Scott
	' Copyright 2017 Sirius Software all rights reserved
	'**********************************************************


	'**********************************************************

	' The form is loaded.

	'**********************************************************
	Private Sub frmOrganizeMRU_Load(sender As Object, e As EventArgs) Handles Me.Load

		' Declare variables

		Dim jj As Integer
		Dim zx As String

		' Fill the list box with the current database history.

		CheckedListBox1.Items.Clear()
		For jj = 1 To 8
			zx = GetSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", CStr(jj), "")
			If zx <> "" Then CheckedListBox1.Items.Add(zx)
		Next jj
		System.Windows.Forms.Application.DoEvents()

	End Sub
	'**********************************************************

	' The Okay button is clicked.

	'**********************************************************
	Private Sub cmdOkay_Click(sender As Object, e As EventArgs) Handles cmdOkay.Click

		' Declare variables

		Dim ii As Integer
		Dim jj As Integer

		' Delete the current MRU list

		If GetSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", "1") <> "" Then DeleteSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList")
		frmMain.MRU1.Text = ""
		frmMain.MRU1.Visible = False
		frmMain.MRU2.Text = ""
		frmMain.MRU2.Visible = False
		frmMain.MRU3.Text = ""
		frmMain.MRU3.Visible = False
		frmMain.MRU4.Text = ""
		frmMain.MRU4.Visible = False
		frmMain.MRU5.Text = ""
		frmMain.MRU5.Visible = False
		frmMain.MRU6.Text = ""
		frmMain.MRU6.Visible = False
		frmMain.MRU7.Text = ""
		frmMain.MRU7.Visible = False
		frmMain.MRU8.Text = ""
		frmMain.MRU8.Visible = False
		frmMain.MRUSeparator.Visible = False

		' Now add back any that were not selected.

		ii = 1
		For jj = 0 To CheckedListBox1.Items.Count - 1
			If Not CheckedListBox1.GetItemChecked(jj) Then
				SaveSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", CStr(ii), CheckedListBox1.Items(jj))
				ii += 1
			End If
		Next jj

		' Refill the MRU list

		FillMRU()

		' Return the dialog result and close the form.

		DialogResult = Windows.Forms.DialogResult.OK
		Me.Close()
	End Sub

	'**********************************************************

	' The Cancel button is clicked.

	'**********************************************************
	Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
		DialogResult = Windows.Forms.DialogResult.Cancel
		Me.Close()
	End Sub

End Class