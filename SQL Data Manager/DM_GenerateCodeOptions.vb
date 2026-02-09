Public Class frmGenerateCodeOptions
	'******************************************************************
	' SQL Data Manager Generate Data Access Command Code options
	' DM_GENERATECODEOPTIONS.VB
	' Written: October 2018
	' Programmer: Aaron Scott
	' Copyright (C) 2017-2018 Sirius Software All Rights Reserved
	'******************************************************************

	'******************************************************************
	'
	' The form is loaded
	'
	'******************************************************************
	Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load


	End Sub


	'******************************************************************
	'
	' A key is pressed in the abbreviations field.  Uppercase it.
	'
	'******************************************************************
	Private Sub txtAbbrev_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAbbrev.KeyPress

		If e.KeyChar >= "a" And e.KeyChar <= "z" Then e.KeyChar = UCase(e.KeyChar)

	End Sub

	'******************************************************************
	'
	' The okay button is clicked.
	'
	'******************************************************************
	Private Sub cmdOkay_Click(sender As Object, e As EventArgs) Handles cmdOkay.Click

		Me.DialogResult = Windows.Forms.DialogResult.OK

	End Sub

	'******************************************************************
	'
	' The cancel buton is clicked.
	'
	'******************************************************************
	Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

		Me.DialogResult = Windows.Forms.DialogResult.Cancel

	End Sub

	'******************************************************************
	'
	' A text box is entered.
	'
	'******************************************************************
	Private Sub Text_Enter(sender As Object, e As EventArgs) Handles txtName.Enter, txtAbbrev.Enter

		' Select all the text.

		CType(sender, TextBox).SelectAll()

	End Sub
End Class