Imports System.ComponentModel
Imports System.Drawing
Imports System.Text
Imports System.Text.RegularExpressions

<ToolboxItem(True)>
Public Class MarkdownViewer
	Inherits Panel
	'*******************************************************************
	' Markdown viewer for Visual Basic programs.
	' MARKDOWNVIEWER.VB
	' Written: March 2026
	' Programmer: Aaron Scott
	' Copyright 2026 Sirius Software All Rights Reserved
	'*******************************************************************


	'*******************************************************************

	' Rendering structures

	'*******************************************************************

	' This class represents a portion of a line to be represented in
	' a fashion incidated by the MD tags.

	Private Class InlineRun
		Public Property Text As String
		Public Property Style As FontStyle

		Public Sub New(text As String, style As FontStyle)
			Me.Text = text
			Me.Style = style
		End Sub
	End Class

	' This class represents one line of the MD file.

	Private Class RenderLine
		Public Property Text As String
		Public Property Style As RenderStyle
		Public Property Runs As List(Of InlineRun)
		Public Property Height As Integer
		Public Property IsContinuation As Boolean
		Public Property AnchorName As String          ' for headers
		Public Property TargetAnchorName As String    ' for TOC links

		Public Sub New(text As String, style As RenderStyle)
			Me.Text = text
			Me.Style = style
			Me.Runs = New List(Of InlineRun) From {
			    New InlineRun(text, FontStyle.Regular)
			}
		End Sub

		Public Sub New(runs As List(Of InlineRun), style As RenderStyle)
			Me.Runs = runs
			Me.Style = style
			Me.Text = String.Join("", runs.Select(Function(r) r.Text))
		End Sub
	End Class

	' Enum for the types of RenderLines a file may contain.
	Private Enum RenderStyle
		Header1
		Header2
		Header3
		Body
		Rule
		Bullet
		AnchorLink
	End Enum

	' Declare variables local to this class.

	Private _parsedLines As New List(Of RenderLine)   ' source (unwrapped)
	Private _lines As New List(Of RenderLine)         ' wrapped, rendered
	Private mRawText As String

	' static fonts based on style
	Private _fontBody As Font
	Private _fontHeader1 As Font
	Private _fontHeader2 As Font
	Private _fontHeader3 As Font
	Private _fontBullet As Font
	Private _fontAnchor As Font

	'*******************************************************************

	' Constructor.

	'*******************************************************************
	Public Sub New()
		Me.DoubleBuffered = True
		Me.AutoScroll = True
		Me.Font = MyBase.Font
	End Sub
	'*******************************************************************

	' The font property.

	'*******************************************************************
	Public Overrides Property Font As Font
		Get
			Return MyBase.Font
		End Get
		Set(value As Font)

			Dim baseFont = value
			_fontBody = baseFont
			_fontBullet = baseFont
			_fontHeader1 = New Font(baseFont.FontFamily, baseFont.Size + 6, FontStyle.Bold)
			_fontHeader2 = New Font(baseFont.FontFamily, baseFont.Size + 3, FontStyle.Bold)
			_fontHeader3 = New Font(baseFont.FontFamily, baseFont.Size + 1, FontStyle.Bold)
			_fontAnchor = New Font(baseFont, FontStyle.Underline)
			RecalculateLayout()
		End Set
	End Property
	'*******************************************************************

	' Function to return the font for displaying a line.

	'*******************************************************************
	Private Function GetLineFont(style As RenderStyle) As Font
		Select Case style
			Case RenderStyle.Header1 : Return _fontHeader1
			Case RenderStyle.Header2 : Return _fontHeader2
			Case RenderStyle.Header3 : Return _fontHeader3
			Case RenderStyle.Bullet : Return _fontBullet
			Case RenderStyle.AnchorLink : Return _fontAnchor
			Case Else : Return _fontBody
		End Select
	End Function

	'*******************************************************************

	' Function to return the font for a run, for example italic or bold.

	'*******************************************************************
	Private Function GetRunFont(lineStyle As RenderStyle, runStyle As FontStyle) As Font
		Dim base = GetLineFont(lineStyle)
		Return New Font(base, base.Style Or runStyle)
	End Function
	'*******************************************************************

	' Sub to load an .md file and display it.

	'*******************************************************************
	Public Sub LoadFile(path As String)
		Try
			Dim text = IO.File.ReadAllText(path, Encoding.Default)
			mRawText = text
			ParseMarkdown(text)
			RecalculateLayout()
			Me.Invalidate()
		Catch ex As Exception
		End Try
	End Sub

	'*******************************************************************

	' Property to set or retrieve the actual .md file text.

	'*******************************************************************
	Public Property RawText As String
		Get
			Return mRawText
		End Get
		Set(value As String)
			mRawText = value
			If Not DesignMode AndAlso Not String.IsNullOrEmpty(value) Then
				ParseMarkdown(value)
				RecalculateLayout()
				Me.Invalidate()
			End If
		End Set
	End Property

	'*******************************************************************

	' Markdown parsing

	'*******************************************************************
	Private Sub ParseMarkdown(md As String)

		_parsedLines.Clear()
		_lines.Clear()

		If String.IsNullOrEmpty(md) Then
			Return
		End If

		Dim rawLines = md.Replace(vbCrLf, vbLf).Split(vbLf)

		For Each line In rawLines

			Dim trimmed = line.TrimStart()

			' Look for heading codes.

			If line.StartsWith("# ") Then ' Heading 1
				Dim text = line.Substring(2).Trim()
				Dim rl = New RenderLine(text, RenderStyle.Header1)
				rl.AnchorName = GenerateAnchorName(text)
				_parsedLines.Add(rl)

			ElseIf line.StartsWith("## ") Then ' Heading 2
				Dim text = line.Substring(3).Trim()
				Dim rl = New RenderLine(text, RenderStyle.Header2)
				rl.AnchorName = GenerateAnchorName(text)
				_parsedLines.Add(rl)

			ElseIf line.StartsWith("### ") Then ' Heading 3
				Dim text = line.Substring(4).Trim()
				Dim rl = New RenderLine(text, RenderStyle.Header3)
				rl.AnchorName = GenerateAnchorName(text)
				_parsedLines.Add(rl)

			ElseIf line.Trim() = "---" Then ' Rule
				_parsedLines.Add(New RenderLine("", RenderStyle.Rule))

				' TOC-style anchor links: - [Text](#anchor)
			ElseIf trimmed.StartsWith("- [") AndAlso trimmed.Contains("](") AndAlso trimmed.EndsWith(")") Then
				Dim rl = ParseAnchorLink(trimmed)
				If rl IsNot Nothing Then
					_parsedLines.Add(rl)
				Else
					' Fallback to normal bullet if parsing fails
					Dim bulletText = trimmed.Substring(2).Trim()
					_parsedLines.Add(New RenderLine(ParseInline(bulletText), RenderStyle.Bullet))
				End If

				' Look for bullet codes.
			ElseIf trimmed.StartsWith("- ") OrElse
				  trimmed.StartsWith("* ") OrElse
				  trimmed.StartsWith("+ ") Then

				Dim bulletText = trimmed.Substring(2).Trim()
				_parsedLines.Add(New RenderLine(ParseInline(bulletText), RenderStyle.Bullet))

				' Otherwise, just plain text.
			Else
				_parsedLines.Add(New RenderLine(ParseInline(line), RenderStyle.Body))
			End If
		Next
	End Sub

	'*******************************************************************

	' Function to return an anchor name, if one exists.

	'*******************************************************************
	Private Function GenerateAnchorName(headerText As String) As String

		Dim lower = headerText.ToLowerInvariant()
		' Replace non-alphanumeric with spaces
		lower = Regex.Replace(lower, "[^a-z0-9\s-]", "")
		' Replace whitespace with hyphens
		lower = Regex.Replace(lower, "\s+", "-")
		' Collapse multiple hyphens
		lower = Regex.Replace(lower, "-{2,}", "-")
		Return lower.Trim("-"c)

	End Function
	'*******************************************************************

	' Function to return the name of a subheading or heading to which
	' a link can jump.

	'*******************************************************************
	Private Function ParseAnchorLink(trimmed As String) As RenderLine

		' Expect: - [Text](#anchor)
		Dim m = Regex.Match(trimmed, "^- \[(?<text>[^\]]+)\]\(#(?<anchor>[^)]+)\)")
		If Not m.Success Then Return Nothing

		Dim displayText = m.Groups("text").Value
		Dim anchor = m.Groups("anchor").Value

		Dim runs As New List(Of InlineRun) From {
		    New InlineRun(displayText, FontStyle.Regular)
		}

		Dim rl As New RenderLine(runs, RenderStyle.AnchorLink)
		rl.TargetAnchorName = anchor
		Return rl
	End Function

	'*******************************************************************

	' Inline parsing.  This breaks a line into various "runs" of
	' different rendering appearence, as in italic or bold.

	'*******************************************************************
	Private Function ParseInline(text As String) As List(Of InlineRun)

		Dim runs As New List(Of InlineRun)
		Dim pattern = "(\*\*.*?\*\*|\*.*?\*)"
		Dim parts = Regex.Split(text, pattern)

		For Each part In parts
			If part.StartsWith("**") AndAlso part.EndsWith("**") Then
				runs.Add(New InlineRun(part.Substring(2, part.Length - 4), FontStyle.Bold))

			ElseIf part.StartsWith("*") AndAlso part.EndsWith("*") Then
				runs.Add(New InlineRun(part.Substring(1, part.Length - 2), FontStyle.Italic))

			Else
				runs.Add(New InlineRun(part, FontStyle.Regular))
			End If
		Next part

		Return runs

	End Function

	'*******************************************************************

	' Layout calculation.

	'*******************************************************************
	Private Sub RecalculateLayout()

		If _parsedLines Is Nothing OrElse _parsedLines.Count = 0 Then Exit Sub
		If ClientSize.Width <= 0 Then Exit Sub

		Dim newLines As New List(Of RenderLine)

		Using g As Graphics = Me.CreateGraphics()
			For Each line In _parsedLines
				If line.Style = RenderStyle.Body OrElse
				   line.Style = RenderStyle.Bullet OrElse
				   line.Style = RenderStyle.AnchorLink Then

					Dim wrapped = WrapLine(line, g, Me.ClientSize.Width - 10)
					newLines.AddRange(wrapped)
				Else
					newLines.Add(line)
				End If
			Next

			_lines = newLines

			Dim total As Integer = 0
			For Each line In _lines
				Select Case line.Style
					Case RenderStyle.Rule
						line.Height = 12
					Case Else
						line.Height = CInt(GetLineFont(line.Style).GetHeight(g)) + 2
				End Select
				total += line.Height
			Next

			AutoScrollMinSize = New Size(ClientSize.Width, total)
		End Using
	End Sub

	'*******************************************************************

	' Function to return a portion of a line that must be wrapped around.

	'*******************************************************************
	Private Function WrapLine(line As RenderLine, g As Graphics, maxWidth As Integer) As List(Of RenderLine)

		Dim wrapped As New List(Of RenderLine)
		Dim currentRuns As New List(Of InlineRun)
		Dim currentWidth As Integer = 0

		' Preserve blank lines
		If line.Text.Trim() = "" Then
			Dim blank = New RenderLine(New List(Of InlineRun) From {
			    New InlineRun("", FontStyle.Regular)
			}, line.Style)
			blank.AnchorName = line.AnchorName
			blank.TargetAnchorName = line.TargetAnchorName
			Return New List(Of RenderLine) From {blank}
		End If

		For Each run In line.Runs
			Dim runFont = GetRunFont(line.Style, run.Style)

			' Split into words + spaces
			Dim tokens = Regex.Matches(run.Text, "\S+|\s+").
				 Cast(Of Match)().
				 Select(Function(m) m.Value)

			For Each token In tokens
				Dim tokenWidth = CInt(g.MeasureString(token, runFont).Width)

				' Wrap if needed
				If currentWidth + tokenWidth > maxWidth AndAlso currentRuns.Count > 0 Then
					Dim rl = New RenderLine(New List(Of InlineRun)(currentRuns), line.Style)
					rl.AnchorName = line.AnchorName
					rl.TargetAnchorName = line.TargetAnchorName
					rl.IsContinuation = (wrapped.Count > 0 AndAlso line.Style = RenderStyle.Bullet)
					wrapped.Add(rl)
					currentRuns.Clear()
					currentWidth = 0
				End If

				currentRuns.Add(New InlineRun(token, run.Style))
				currentWidth += tokenWidth
			Next token
		Next run

		' Final line
		If currentRuns.Count > 0 Then
			Dim rl = New RenderLine(currentRuns, line.Style)
			rl.AnchorName = line.AnchorName
			rl.TargetAnchorName = line.TargetAnchorName
			rl.IsContinuation = (wrapped.Count > 0 AndAlso line.Style = RenderStyle.Bullet)
			wrapped.Add(rl)
		End If

		Return wrapped
	End Function
	'*******************************************************************

	' In case the control is resized, we must recalculate the layout.

	'*******************************************************************
	Protected Overrides Sub OnResize(e As EventArgs)
		MyBase.OnResize(e)

		If Not DesignMode AndAlso IsHandleCreated AndAlso _parsedLines IsNot Nothing AndAlso _parsedLines.Count > 0 Then

			BeginInvoke(Sub()
						  RecalculateLayout()
						  Invalidate()
					  End Sub)
		End If
	End Sub
	'*******************************************************************

	' Hit testing and scrolling

	'*******************************************************************
	Private Function GetLineIndexAtPoint(p As Point) As Integer
		If _lines Is Nothing OrElse _lines.Count = 0 Then Return -1

		Dim y As Integer = AutoScrollPosition.Y
		For i As Integer = 0 To _lines.Count - 1
			Dim line = _lines(i)
			Dim top = y
			Dim bottom = y + line.Height
			If p.Y >= top AndAlso p.Y < bottom Then
				Return i
			End If
			y += line.Height
		Next

		Return -1
	End Function

	'*******************************************************************

	' Sub to scroll the page to a line selected by a link.

	'*******************************************************************
	Private Sub ScrollToLine(index As Integer)
		If index < 0 OrElse index >= _lines.Count Then Return

		Dim y As Integer = 0
		For i As Integer = 0 To index - 1
			y += _lines(i).Height
		Next

		Me.AutoScrollPosition = New Point(0, y)
		Me.Invalidate()
	End Sub
	'*******************************************************************

	' Handler for the mouse click event.

	'*******************************************************************

	Protected Overrides Sub OnMouseClick(e As MouseEventArgs)

		MyBase.OnMouseClick(e)

		If e.Button <> MouseButtons.Left Then Return
		If _lines Is Nothing OrElse _lines.Count = 0 Then Return

		Dim idx = GetLineIndexAtPoint(e.Location)
		If idx < 0 Then Return

		Dim line = _lines(idx)
		If line.Style <> RenderStyle.AnchorLink Then Return
		If String.IsNullOrEmpty(line.TargetAnchorName) Then Return

		' Find target header by AnchorName
		Dim targetIndex As Integer = -1
		For i As Integer = 0 To _lines.Count - 1
			Dim l = _lines(i)
			If Not String.IsNullOrEmpty(l.AnchorName) AndAlso
			   String.Equals(l.AnchorName, line.TargetAnchorName, StringComparison.OrdinalIgnoreCase) Then
				targetIndex = i
				Exit For
			End If
		Next i

		If targetIndex >= 0 Then
			ScrollToLine(targetIndex)
		End If
	End Sub

	'*******************************************************************

	' Paint the display: the actual rendering of the .md text takes
	' place here.

	'*******************************************************************
	Protected Overrides Sub OnPaint(e As PaintEventArgs)

		MyBase.OnPaint(e)

		If DesignMode OrElse _lines Is Nothing OrElse _lines.Count = 0 Then Return

		Dim g = e.Graphics
		Dim y As Integer = AutoScrollPosition.Y

		For Each line In _lines
			Dim lineFont = GetLineFont(line.Style)

			Select Case line.Style

				Case RenderStyle.Rule
					g.DrawLine(Pens.Gray, 0, y + 4, ClientSize.Width, y + 4)

				Case RenderStyle.Header3, RenderStyle.Header2, RenderStyle.Header1
					g.DrawString(line.Text, lineFont, Brushes.Black, 0, y)

				Case RenderStyle.Bullet
					Dim bullet = "• "
					Dim indent = 20

					If Not line.IsContinuation Then
						' First visual line of bullet
						g.DrawString(bullet, lineFont, Brushes.Black, 0, y)
					End If

					Dim x As Integer = If(line.IsContinuation, 0, indent)

					For Each run In line.Runs
						Dim runFont = GetRunFont(line.Style, run.Style)
						g.DrawString(run.Text, runFont, Brushes.Black, x, y)
						x += CInt(g.MeasureString(run.Text, runFont).Width)
					Next

				Case RenderStyle.AnchorLink
					Dim x As Integer = 0
					For Each run In line.Runs
						Dim runFont = GetRunFont(line.Style, run.Style)
						Using linkBrush As New SolidBrush(Color.Blue)
							g.DrawString(run.Text, runFont, linkBrush, x, y)
						End Using
						x += CInt(g.MeasureString(run.Text, runFont).Width)
					Next

				Case RenderStyle.Body
					Dim x As Integer = 0
					For Each run In line.Runs
						Dim runFont = GetRunFont(line.Style, run.Style)
						g.DrawString(run.Text, runFont, Brushes.Black, x, y)
						x += CInt(g.MeasureString(run.Text, runFont).Width)
					Next

				Case Else
					g.DrawString(line.Text, lineFont, Brushes.Black, 0, y)

			End Select

			y += line.Height
		Next
	End Sub

End Class