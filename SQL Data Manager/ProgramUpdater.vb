Module ProgramUpdater
	'*******************************************************************
	' Program Updater function for Visual Basic programs.
	' PROGRAMUPDATER.VB
	' Written: May 2022
	' Updated: December 2023
	' Programmer: Aaron Scott
	' Copyright 2022-2023 Sirius Software All Rights Reserved
	'*******************************************************************

	' Declare variables local to this form.

	Private SelectedName As String = ProgramName
	Private Result As MsgBoxResult
	Private f As Form
	Private fNotes As Form
	Private DestinationFolder As String = GetPath(Application.ExecutablePath)
	Private UpdateFileList As String = ""

	' Create the controls we'll need on the update message.  We need
	' to create them now, so they can be referenced in late-bound routines.

	Private btnYes As New Button
	Private btnNo As New Button
	Private btnNotes As New Button
	Private Panel1 As New Panel
	Private Label1 As New Label
	Private lblHeader_0 As New Label
	Private lblHeader_1 As New Label
	Private lblHeader_2 As New Label
	Private ListBox1 As New ListBox
	Private PictureBox1 As New PictureBox


	'*******************************************************************

	' Sub to check for available program and dependency updates.

	'*******************************************************************
	Public Sub CheckForUpdates()

		' Declare variables.

		Dim Version As Single = Val(My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & My.Application.Info.Version.Build & My.Application.Info.Version.MinorRevision)
		Dim UpdaterVersion As String = "1.000"
		Dim UpdateFiles As IReadOnlyCollection(Of String)
		Dim UpdateList As String
		Dim DisplayList As String
		Dim InstalledVersion As String
		Dim CurrentVersion As String
		Dim wx As String
		Dim xx As String
		Dim d As Dependency
		Dim fi As FileVersionInfo

		' First check to see if the update utility itself requires updating.

		CheckUpdateUtility(UpdaterVersion)

		' See if there's an update available on OneDrive, and, if so, copy it into the executable directory.  

		CheckProgramVersion(Version)

		' Check if any program dependencies need updating.

		CheckProgramDependencies()

		' See if there are any update files in the executable folder, and if so, accumulate a list of them.
		' This is compatible with the old manual-update approach.

		Try
			DisplayList = ""
			UpdateList = ""
			UpdateFiles = My.Computer.FileSystem.GetFiles(DestinationFolder, FileIO.SearchOption.SearchTopLevelOnly, "*.update")
			For Each wx In UpdateFiles

				' Get the version of the update file.

				fi = FileVersionInfo.GetVersionInfo(wx)
				CurrentVersion = fi.FileMajorPart & "." & fi.FileMinorPart & fi.FileBuildPart & fi.FilePrivatePart

				' Get the version of the installed file.

				Try
					If My.Computer.FileSystem.FileExists(wx.Replace(".update", ".exe")) Then
						fi = FileVersionInfo.GetVersionInfo(wx.Replace(".update", ".exe"))
					Else
						fi = FileVersionInfo.GetVersionInfo(wx.Replace(".update", ".dll"))
					End If
					InstalledVersion = fi.FileMajorPart & "." & fi.FileMinorPart & fi.FileBuildPart & fi.FilePrivatePart
				Catch ex As Exception
					InstalledVersion = "0"
				End Try

				' Assemble the entry for the list box.

				xx = FileNameNoPath(wx)
				xx = xx.Substring(0, xx.IndexOf("."))
				If xx = ProgramName Then
					If DisplayList <> "" Then DisplayList &= ","
					DisplayList &= xx & "." & "exe"
					If UpdateList <> "" Then UpdateList &= ","
					UpdateList &= xx & "," & "exe"
				Else
					If Dependencies.Contains(xx) Then
						d = Dependencies(xx)
						If DisplayList <> "" Then DisplayList &= ","
						DisplayList &= d.Name & "." & d.ObjectType
						If UpdateList <> "" Then UpdateList &= ","
						UpdateList &= d.Name & "," & d.ObjectType
					End If
				End If
				DisplayList &= "," & InstalledVersion & "," & CurrentVersion
			Next wx

			' See if the user wants to update now.

			If DisplayList <> "" Then
				If UpdateMsgBox(DisplayList) = MsgBoxResult.Yes Then

					' Start the update utility, passing it a switch (to prove validity), and the name of the update file(s).

					Shell(DestinationFolder & "SiriusUpdateUtility.exe " & "/update," & GetPath(Application.ExecutablePath) & "," & FileNameNoPath(Application.ExecutablePath) & "," & UpdateList)

					' End the program.

					End
				End If
			End If

		Catch ex As Exception
			MsgBox("Update failed." & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, ProgramName & " Update")
		End Try

	End Sub

	'*******************************************************************

	' Sub to check for the correct version of the updater utility,
	' and update it if necessary.

	'*******************************************************************
	Private Sub CheckUpdateUtility(UpdaterVersion As String)

		' Declare variables

		Dim zx As String
		Dim fi As FileVersionInfo
		Dim v As Single

		' See if the updater exists.

		If Not My.Computer.FileSystem.FileExists(DestinationFolder & "SiriusUpdateUtility.exe") Then
			v = 0 ' Setting the version to zero will cause the file to be copied over.
		Else

			' Get the version of the updater.

			fi = FileVersionInfo.GetVersionInfo(DestinationFolder & "SiriusUpdateUtility.exe")
			v = CSng(fi.FileMajorPart & "." & fi.FileMinorPart & fi.FileBuildPart & fi.FilePrivatePart)
		End If

		' Get the version number of the current updater.

		Try
			zx = My.Computer.FileSystem.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\OneDrive\ProgramUpdates\SiriusUpDateUtility\currentversion.txt")
			If zx.IndexOf(vbCrLf) > 0 Then zx = zx.Substring(0, zx.IndexOf(vbCrLf))

			' If the latest version is more recent, get the update.

			If CSng(zx) > v Then

				' Copy the .update file into the folder with the executable.

				My.Computer.FileSystem.CopyFile(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\OneDrive\ProgramUpdates\SiriusUpdateUtility\SiriusUpdateUtility.update", DestinationFolder & "SiriusUpdateUtility.exe", True)
			End If

		Catch ex As Exception ' Do nothing: there is no update available.
		End Try
	End Sub

	'*******************************************************************

	' Sub to check the program for the current version.

	'*******************************************************************
	Private Sub CheckProgramVersion(Version As String)

		' Declare variables

		Dim zx As String

		Try
			' Get the latest version number, to compare against the program's version.

			zx = My.Computer.FileSystem.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\OneDrive\ProgramUpdates\" & SRep(ProgramName, 1, " ", "") & "\currentversion.txt")
			If zx.IndexOf(vbCrLf) > 0 Then zx = zx.Substring(0, zx.IndexOf(vbCrLf))

			' If the latest version is more recent, get the update.

			If CSng(zx) > CSng(Version) Then

				' Copy the .update file into the folder with the executable.

				My.Computer.FileSystem.CopyFile(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\OneDrive\ProgramUpdates\" & SRep(ProgramName, 1, " ", "") & "\" & ProgramName & ".update", DestinationFolder & ProgramName & ".update", True)
			End If

		Catch ex As Exception ' Do nothing: there is no update available.
		End Try

	End Sub
	'*******************************************************************

	' Sub to check the versions of all program dependencies.

	'*******************************************************************
	Private Sub CheckProgramDependencies()

		' Declare variables.

		Dim zx As String
		Dim d As Dependency
		Dim fi As FileVersionInfo
		Dim cf1 As IO.FileInfo
		Dim cf2 As IO.FileInfo
		Dim v As Single


		' Loop through the list of dependencies, if any.

		For Each d In Dependencies

			Try
				' If the object type is an .exe or a .dll file.

				Select Case d.ObjectType
					Case "exe", "dll"
						' Get the latest version number, to compare against the program's version.

						zx = My.Computer.FileSystem.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\OneDrive\ProgramUpdates\" & d.Name & "\currentversion.txt")
						If zx.IndexOf(vbCrLf) > 0 Then zx = zx.Substring(0, zx.IndexOf(vbCrLf))

						' See if the dependent file exists.  If it's missing, we'll copy it over as though
						' it were out of date.

						If Not My.Computer.FileSystem.FileExists(DestinationFolder & d.Name & "." & d.ObjectType) Then
							v = 0 ' Setting the version to 0 will cause it to get copied over.

							' Get the version of the currently-existing file.

						Else
							fi = FileVersionInfo.GetVersionInfo(DestinationFolder & d.Name & "." & d.ObjectType)
							v = CSng(fi.FileMajorPart & "." & fi.FileMinorPart & fi.FileBuildPart & fi.FilePrivatePart)
						End If

						' If the latest version is more recent, get the update.

						If CSng(zx) > v Then

							' Copy the .update file into the folder with the executable.

							My.Computer.FileSystem.CopyFile(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\OneDrive\ProgramUpdates\" & d.Name & "\" & d.Name & ".update", DestinationFolder & d.Name & ".update", True)
						End If

					' If the object type is a file to be copied.

					Case "file"

						' If the file doesn't exist, just copy it over.

						If Not My.Computer.FileSystem.FileExists(DestinationFolder & d.FileToCopy) Then
							My.Computer.FileSystem.CopyFile(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\OneDrive\ProgramUpdates\" & d.Name & "\" & d.FileToCopy, DestinationFolder & d.FileToCopy)
						Else

							' If the file DOES exist, compare the date and time of the current copy and the
							' installed copy.

							cf1 = My.Computer.FileSystem.GetFileInfo(DestinationFolder & d.FileToCopy)
							cf2 = My.Computer.FileSystem.GetFileInfo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\OneDrive\ProgramUpdates\" & d.Name & "\" & d.FileToCopy)

							' If the current version has a newer date than the installed version, copy it over.

							If cf2.LastWriteTime > cf1.LastWriteTime Then My.Computer.FileSystem.CopyFile(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\OneDrive\ProgramUpdates\" & d.Name & "\" & d.FileToCopy, DestinationFolder & d.FileToCopy, True)

						End If
				End Select

			Catch ex As Exception ' Do nothing: there is no update available.
			End Try
		Next d
	End Sub
	'*******************************************************************

	' The UpdateMessageBox function.  This is called to display the message box
	' and return its result.

	'*******************************************************************
	Public Function UpdateMsgBox(UpdateList As String) As DialogResult

		' Declare variables

		Dim BackgroundColorNumber As Integer
		Dim TextColorNumber As Integer
		Dim FontNumber As Integer
		Dim SelectedName As String = ProgramName
		Dim zx As String
		Dim fS As New Font("MS Sans Serif", 8.25, FontStyle.Regular)
		Dim fL As New Font("Arial", 12, FontStyle.Regular)

		' Create the form and add the panel to it.

		f = New Form
		With f
			.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
			.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			.ClientSize = New System.Drawing.Size(455, 230)
			.FormBorderStyle = FormBorderStyle.FixedDialog
			.ControlBox = False
			.MaximizeBox = False
			.MinimizeBox = False
			.Name = "frmUpdateMessage"
			.ShowIcon = False
			.ShowInTaskbar = False
			.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
			.Text = "Program Update Available"
		End With

		' Set the properties for the Yes button.

		btnYes.Location = New System.Drawing.Point(286, 12)
		btnYes.Name = "btnYes"
		btnYes.Size = New System.Drawing.Size(75, 23)
		btnYes.TabIndex = 2
		btnYes.Text = "Yes"
		btnYes.UseVisualStyleBackColor = True
		AddHandler btnYes.Click, AddressOf btnYes_Click

		' Set the properties for the no button.

		btnNo.Location = New System.Drawing.Point(368, 12)
		btnNo.Name = "btnNo"
		btnNo.Size = New System.Drawing.Size(75, 23)
		btnNo.TabIndex = 3
		btnNo.Text = "No"
		btnNo.UseVisualStyleBackColor = True
		AddHandler btnNo.Click, AddressOf btnNo_Click

		' Set the properties for the View Release Notes button.

		btnNotes.Location = New System.Drawing.Point(12, 12)
		btnNotes.Name = "btnNotes"
		btnNotes.Size = New System.Drawing.Size(123, 23)
		btnNotes.TabIndex = 4
		btnNotes.Text = "View Release Notes"
		btnNotes.UseVisualStyleBackColor = True
		AddHandler btnNotes.Click, AddressOf btnNotes_Click

		' Create the panel and add the buttons to it.

		Panel1.Controls.Add(btnNotes)
		Panel1.Controls.Add(btnNo)
		Panel1.Controls.Add(btnYes)
		Panel1.Location = New System.Drawing.Point(2, 180)
		Panel1.Name = "Panel1"
		Panel1.Size = New System.Drawing.Size(450, 49)
		Panel1.TabIndex = 5

		' Create the picture box.

		PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
		PictureBox1.Location = New System.Drawing.Point(14, 22)
		PictureBox1.Name = "PictureBox1"
		PictureBox1.Size = New System.Drawing.Size(64, 64)
		PictureBox1.TabIndex = 6
		PictureBox1.TabStop = False
		AddHandler PictureBox1.Paint, AddressOf Picturebox1_Paint

		' Create the information label.

		Label1.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Label1.Location = New System.Drawing.Point(113, 22)
		Label1.Name = "Label1"
		Label1.Size = New System.Drawing.Size(332, 70)
		Label1.TabIndex = 7
		Label1.Text = ""

		' Create the 3 list box header labels.

		lblHeader_0.Text = "Module Name"
		lblHeader_0.Location = New System.Drawing.Point(113, 95)
		lblHeader_0.Size = New System.Drawing.Point(152, 20)
		lblHeader_0.BorderStyle = BorderStyle.FixedSingle
		lblHeader_0.TabIndex = 8

		lblHeader_1.Text = "Installed Version"
		lblHeader_1.Location = New System.Drawing.Point(lblHeader_0.Left + lblHeader_0.Width - 1, lblHeader_0.Top)
		lblHeader_1.Size = New System.Drawing.Point(90, 20)
		lblHeader_1.BorderStyle = BorderStyle.FixedSingle
		lblHeader_1.TabIndex = 9

		lblHeader_2.Text = "Current Version"
		lblHeader_2.Location = New System.Drawing.Point(lblHeader_1.Left + lblHeader_1.Width - 1, lblHeader_0.Top)
		lblHeader_2.Size = New System.Drawing.Point(90, 20)
		lblHeader_2.BorderStyle = BorderStyle.FixedSingle
		lblHeader_2.TabIndex = 10

		' Create the list box.

		ListBox1.Location = New System.Drawing.Point(lblHeader_0.Left, lblHeader_0.Top + lblHeader_0.Height)
		ListBox1.Name = "ListBox1"
		ListBox1.Size = New System.Drawing.Point(lblHeader_0.Width + lblHeader_1.Width + lblHeader_2.Width - 2, f.Height - Panel1.Top - lblHeader_0.Height - 5)
		ListBox1.DrawMode = DrawMode.OwnerDrawFixed
		ListBox1.TabIndex = 11

		Do While UpdateList <> ""
			zx = ParseString(UpdateList) & vbTab
			zx &= ParseString(UpdateList) & vbTab
			zx &= ParseString(UpdateList)
			ListBox1.Items.Add(zx)
		Loop
		AddHandler ListBox1.SelectedIndexChanged, AddressOf Listbox1_SelectedIndexChanged
		AddHandler ListBox1.DrawItem, AddressOf ListBox1_DrawItem


		' Add all the controls to the form.

		f.Controls.Add(PictureBox1)
		f.Controls.Add(Label1)
		f.Controls.Add(lblHeader_0)
		f.Controls.Add(lblHeader_1)
		f.Controls.Add(lblHeader_2)
		f.Controls.Add(ListBox1)
		f.Controls.Add(Panel1)

		' Get the theme number, and separate out the font number and the color
		' number.  The theme number is a hex value from 0x0111 to 0x0222. 
		' The background color number is the 2^2 value, the
		' color number is the 2^1 value and the font number the 2^0
		' column.  There are two background colors, three font colors
		' and two text sizes, for a total of 12 variations in the theme.

		zx = ProgramColorTheme
		BackgroundColorNumber = Val(Mid(zx, 4, 1))
		TextColorNumber = Val(Mid(zx, 5, 1))
		FontNumber = Val(Mid(zx, 6, 1))

		' Set the color and font and message text.

		f.BackColor = Choose(BackgroundColorNumber, SystemColors.Window, Color.FromArgb(255, 182, 189, 255), Color.Tan, UserDefinedColor)
		Panel1.BackColor = SystemColors.Window
		Label1.Font = Choose(FontNumber, fS, fL)
		Label1.ForeColor = Choose(TextColorNumber, Color.Black, Color.DarkBlue, Color.DarkGoldenrod, Color.White)
		Label1.Text = "A new version of " & ProgramName & ", or one of its dependencies, is available. Do you want to update now?"

		' Show the form.

		PlaySound("SystemQuestion", 0, SND_ALIAS + SND_ASYNC)
		f.ShowDialog()

		' Remove handlers.

		RemoveHandler Panel1.Controls("btnYes").Click, AddressOf btnYes_Click
		RemoveHandler Panel1.Controls("btnNo").Click, AddressOf btnNo_Click
		RemoveHandler Panel1.Controls("btnNotes").Click, AddressOf btnNotes_Click
		RemoveHandler PictureBox1.Paint, AddressOf Picturebox1_Paint
		RemoveHandler ListBox1.SelectedIndexChanged, AddressOf Listbox1_SelectedIndexChanged
		RemoveHandler ListBox1.DrawItem, AddressOf ListBox1_DrawItem

		' Return the dialog result.

		Return Result
	End Function
	'*******************************************************************

	' The Yes button is clicked.  This is bound at run time.

	'*******************************************************************
	Private Sub btnYes_Click(sender As Object, e As EventArgs)
		Result = MsgBoxResult.Yes
		f.Close()
	End Sub
	'*******************************************************************

	' The No button is clicked.  This is bound at run time.

	'*******************************************************************
	Private Sub btnNo_Click(sender As Object, e As EventArgs)
		Result = MsgBoxResult.No
		f.Close()
	End Sub
	'*******************************************************************

	' The View Release Notes button is clicked.  This is bound at run time.

	'*******************************************************************
	Private Sub btnNotes_Click(sender As Object, e As EventArgs)

		' Create a form for viewing the release notes.

		fNotes = New Form
		With fNotes
			.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
			.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			.ClientSize = New System.Drawing.Size(600, 500)
			.FormBorderStyle = FormBorderStyle.FixedDialog
			.ControlBox = True
			.MaximizeBox = False
			.MinimizeBox = False
			.Name = "frmReleaseNotes"
			.ShowIcon = False
			.ShowInTaskbar = False
			.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
			.Text = SelectedName & " Release Notes"
		End With

		' Add the text box.

		Dim t As New TextBox
		t.Multiline = True
		t.Dock = DockStyle.Fill
		t.Location = New Point(0, 0)
		t.ScrollBars = ScrollBars.Vertical
		t.ReadOnly = True
		t.Font = New Font("Arial", 10)
		Try
			t.Text = My.Computer.FileSystem.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\OneDrive\ProgramUpdates\" & SelectedName.Replace(" ", "") & "\currentversion.txt")
		Catch ex As Exception
			MsgBox("Cannot open release notes for " & SelectedName & vbCrLf & ex.Message, MsgBoxStyle.Information, "View release Notes")
		End Try
		t.SelectionStart = t.Text.Length
		fNotes.Controls.Add(t)

		' Show the notes.

		fNotes.Show()

	End Sub
	'*******************************************************************

	' Event handler for drawing the icon. This is bound at run time.

	'*******************************************************************
	Private Sub Picturebox1_Paint(sender As Object, e As PaintEventArgs)

		' Declare variables

		Dim ImageAttr As New System.Drawing.Imaging.ImageAttributes
		Dim g As Graphics = e.Graphics
		Dim p As PictureBox = CType(sender, PictureBox)
		Dim r As New Rectangle(0, 0, p.Width, p.Height)

		' Set the transparent color to white.

		ImageAttr.SetColorKey(Color.White, Color.White)

		' Draw the image.

		g.DrawImage(My.Resources.Question, r, 0, 0, My.Resources.Question.Width, My.Resources.Question.Height, GraphicsUnit.Pixel, ImageAttr)


	End Sub
	'*******************************************************************

	' Event Handler for the list box.  This is bound at runtime.

	'*******************************************************************
	Private Sub Listbox1_SelectedIndexChanged(sender As Object, e As EventArgs)

		' Declare variables

		Dim l As ListBox = CType(sender, ListBox)
		Dim zx As String

		If l.SelectedIndex >= 0 Then
			zx = l.Items(l.SelectedIndex)
			SelectedName = ParseString(zx, ".")
		End If
	End Sub

	'**********************************************************

	' Event Handler for the "Draw Item" event of the update
	' module list box.  This is bound at runtime.

	'**********************************************************
	Private Sub ListBox1_DrawItem(sender As Object, e As DrawItemEventArgs)

		' Declare variables

		Dim jj As Integer
		Dim INDENT As Single
		Dim xx As String
		Dim zx As String
		Dim Rect As Rectangle
		Dim g As Graphics = e.Graphics
		Dim f As Font = e.Font
		Dim bC As New SolidBrush(Color.Red)
		Dim bN As New SolidBrush(e.ForeColor)
		Dim l As Label

		' Calculate the size of an indent

		INDENT = g.MeasureString(" ", f).Width

		' First draw the background.

		e.DrawBackground()

		' Unless there is an item to draw, do nothing.

		If e.Index >= 0 Then

			' Get the line to be displayed.

			xx = sender.items(e.Index)

			' Calculate the rectangles for module name, installed version and current version

			For jj = 1 To 3
				l = Choose(jj, lblHeader_0, lblHeader_1, lblHeader_2)
				Rect = l.Bounds
				Rect.X -= sender.left
				Rect.Y = e.Bounds.Y

				' Get the data for the current column from the line to be displayed.

				zx = ParseString(xx, vbTab)

				' The second and third columns are numeric, so calculate the position to right-justify
				' them beneath the column header label, unless the version is missing (=0).

				If jj > 1 And zx <> "0" Then Rect.X = Rect.X + Rect.Width - g.MeasureString(zx, f).Width - 2 * INDENT

				' Display the data in the column.

				If jj = 2 And zx = "0" Then
					g.DrawString("Missing", f, bC, Rect)
				Else
					g.DrawString(zx, f, bN, Rect)
				End If

			Next jj
		End If

		' Dispose of objects we've created.

		bN.Dispose()
		bC.Dispose()

	End Sub
End Module