Public Class FKTable
	Inherits Panel
	'**********************************************************
	' SQL Data Manager Table class for Relationships editor (a custom panel class)
	' FKTABLE.VB
	' Written: November 2018
	' Programmer: Aaron Scott
	' Copyright (C) 2018 Sirius Software All Rights Reserved
	'**********************************************************

	' Declare the mouse events that we will pass upwards to the parent form.

	Public Event MouseDownEvent(sender As Object, e As MouseEventArgs)
	Public Event MouseUpEvent(sender As Object, e As MouseEventArgs)
	Public Event MouseMoveEvent(sender As Object, e As MouseEventArgs)
	Public Event SelectionChanged(sender As Object, e As EventArgs)

	' Declare the static properties of this class

	Public TableName As String
	Public RelationName As String

	' Declare variables local to this module

	Private mActive As Boolean
	Private mFields As ListBox
	Private DragFromPoint As Point

	'**********************************************************

	' Declare the Fields property (returns the list box with
	' the name of the table's fields).

	'**********************************************************
	Public ReadOnly Property Fields As ListBox
		Get
			Fields = mFields
		End Get
	End Property

	'**********************************************************

	' Declare the Active property.  This property determines
	' how the object caption will be displayed.

	'**********************************************************
	Public Property Active As Boolean
		Get
			Active = mActive
		End Get
		Set(value As Boolean)
			mActive = value

			' Invalidate the object so it will be redrawn in its new state

			Me.Invalidate()

		End Set
	End Property

	'**********************************************************

	' Event handler for the MouseDown event.

	'**********************************************************
	Private Sub FKTable_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown

		' Create a new MouseMoveEventArgs that will pass the location of the mouse
		' relative to the parent form.

		Dim NewArgs As New MouseEventArgs(e.Button, e.Clicks, e.X + Me.Left, e.Y + Me.Top, e.Delta)

		' Pass this event to the parent form.

		RaiseEvent MouseDownEvent(sender, NewArgs)

		' Whenever the  mouse is pressed on this object, it becaomes active.

		Active = True

		' Raise the SelectionChanged event in the parent form.

		RaiseEvent SelectionChanged(Me, New EventArgs)

	End Sub

	'**********************************************************

	' Event handler for the MouseUp event.

	'**********************************************************
	Private Sub FKTable_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp

		' Create a new MouseMoveEventArgs that will pass the location of the mouse
		' relative to the parent form.

		Dim NewArgs As New MouseEventArgs(e.Button, e.Clicks, e.X + Me.Left, e.Y + Me.Top, e.Delta)

		' Pass this event to the parent form.

		RaiseEvent MouseUpEvent(sender, NewArgs)

		' Restore the default cursor, which may have changed to a sizing
		' cursor when the mouse button was pressed.

		Me.Cursor = Cursors.Default

	End Sub

	'**********************************************************

	' Event handler for the MouseMove event.

	'**********************************************************
	Public Sub FKTable_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove

		' Create a new MouseMoveEventArgs that will pass the location of the mouse
		' relative to the parent form.

		Dim NewArgs As New MouseEventArgs(e.Button, e.Clicks, e.X + Me.Left, e.Y + Me.Top, e.Delta)

		' Declare variables

		Dim Tolerance As Integer = 5  ' This determines how close we have to be to the object's borders to change to sizing mode.

		' See if no button is pressed.

		If e.Button = Windows.Forms.MouseButtons.None Then

			' See if we need to change to window sizing cursor

			If (e.X >= 0 And e.X <= Tolerance) Or (e.X >= Me.Width - Tolerance And e.X <= Me.Width) Then
				Me.Cursor = Cursors.SizeWE
				Exit Sub
			ElseIf (e.Y >= 0 And e.Y <= Tolerance) Or (e.Y >= Me.Height - Tolerance And e.Y <= Me.Height) Then
				Me.Cursor = Cursors.SizeNS
				Exit Sub
			Else
				Me.Cursor = Cursors.Default
			End If

			System.Windows.Forms.Application.DoEvents()

			' If any button is pressed, pass this event to the parent form.
		Else
			RaiseEvent MouseMoveEvent(sender, NewArgs)
		End If



	End Sub

	'**********************************************************

	' The object is resized.

	'**********************************************************
	Private Sub FKTable_Resize(sender As Object, e As EventArgs) Handles Me.Resize

		' Resize the fields list box to match the new object size.

		Fields.Width = Me.Width - 6
		Fields.Height = Me.Height - 24

	End Sub

	'**********************************************************

	' Event handler for the object's paint event.

	'**********************************************************
	Private Sub FKTable_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint

		' Declare variables

		Dim ArcSize As Integer = 10
		Dim g As Graphics = e.Graphics
		Dim r As Rectangle
		Dim fN As New Font("Arial", 10)
		Dim p As Drawing2D.GraphicsPath
		Dim F As New StringFormat

		' Set the text trimming mode for the object's caption.

		F.Trimming = StringTrimming.EllipsisCharacter

		' Define the caption area and draw it.  We use a graphics path, as this is a complex shape with
		' rounded corners on the top left and right.

		p = New Drawing2D.GraphicsPath
		p.AddLine(0, 20, 0, ArcSize)
		p.AddArc(0, 0, ArcSize, ArcSize, 180, 90)
		p.AddLine(ArcSize, 0, Me.Width - ArcSize, 0)
		p.AddArc(Me.Width - ArcSize, 0, ArcSize, ArcSize, 270, 90)
		p.AddLine(Me.Width, ArcSize, Me.Width, 20)
		p.AddLine(Me.Width, 20, 0, 20)
		g.DrawPath(SystemPens.ActiveBorder, p)

		' If the object is active, fill the caption dark.

		If Active Then
			g.FillPath(SystemBrushes.Highlight, p)
			r = New Rectangle(0, 21, Me.Width, Me.Height - 20)
			g.FillRectangle(SystemBrushes.ActiveBorder, r)
			r = New Rectangle(ArcSize, 2, Me.Width, 20)
			g.DrawString(TableName, fN, SystemBrushes.HighlightText, r, F)

			' Otherwise, fill the caption with the normal color.

		Else
			g.FillPath(SystemBrushes.InactiveCaption, p)
			r = New Rectangle(0, 21, Me.Width, Me.Height - 20)
			g.FillRectangle(SystemBrushes.ActiveBorder, r)
			r = New Rectangle(ArcSize, 2, Me.Width, 20)
			g.DrawString(TableName, fN, SystemBrushes.GrayText, r, F)
		End If

		' Dispose of our graphics path.

		p.Dispose()

	End Sub

	'**********************************************************

	' The object is created.

	'**********************************************************
	Public Sub New()

		' When the class is instantiated, add a listbox to it.

		mFields = New ListBox
		mFields.IntegralHeight = False
		mFields.Location = New Point(3, 21)
		Me.Controls.Add(mFields)

		' Set the default properties of this class.

		Me.Width = 100
		Me.Height = 120
		Me.BackColor = Color.Transparent  ' Panel will be transparent where we don't draw on it.

	End Sub

	'**********************************************************

	' The object is destroyed.

	'**********************************************************
	Protected Overrides Sub Finalize()

		' When the class is destroyed, remove the list box.

		mFields.Dispose()

		MyBase.Finalize()

	End Sub

End Class
