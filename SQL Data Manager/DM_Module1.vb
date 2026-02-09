Imports System.Data.SqlClient
Imports VB = Microsoft.VisualBasic

Module DM_Module1
	'*******************************************************************
	' SQL Data Manager Module 1
	' DM_MODULE1.VB
	' Written: October 2017
	' Programmer: Aaron Scott
	' Copyright 2017 Sirius Software All Rights Reserved
	'*******************************************************************


	' Declare variables available throughout the program

	Public ServerName As String
	Public DB As New SqlConnection

	' Create an Enum for the 4 components of a color theme.

	Public Enum ThemeItem
		All
		MainWindow
		BackgroundColorNumber
		TextColorNumber
		FontNumber
	End Enum

	' Declare variables local to this module

	' Data access objects for the control and events log tables.


	'*******************************************************************

	' Sub to create a new database.  The SQL database template is
	' copied over to the new database.

	'*******************************************************************
	Public Sub CreateNewDatabase()

		' Declare variables

		Dim jj As Integer
		Dim xx As Integer
		Dim zx As String = ""
		Dim NewDatabase As String

		' Get the name of the new database.

		frmMain.SaveFileDialog1.AddExtension = True
		frmMain.SaveFileDialog1.Filter = "Microsoft SQL Files (*.mdf)|*.mdf"
		frmMain.SaveFileDialog1.DefaultExt = "mdf"
		frmMain.SaveFileDialog1.InitialDirectory = System.IO.Directory.GetCurrentDirectory
		frmMain.SaveFileDialog1.OverwritePrompt = True
		frmMain.SaveFileDialog1.SupportMultiDottedExtensions = True
		xx = frmMain.SaveFileDialog1.ShowDialog()

		' If the dialog returns a valid file name, proceed.

		If xx = DialogResult.OK Then

			' Get the name of the database, not including its path.

			NewDatabase = frmMain.SaveFileDialog1.FileName
			For jj = Len(NewDatabase) To 1 Step -1
				If Mid(NewDatabase, jj, 1) = "\" Then
					zx = Mid(NewDatabase, jj + 1)
					zx = VB.Left(zx, Len(zx) - 4)
					Exit For
				End If
			Next jj

			' Create a new connection

			Try
				Databasename = NewDatabase
				Dim TempConnection As New SqlConnection("Data Source=" & ServerName & ";database='';integrated security=true")
				TempConnection.Open()

				' Set up the command to create the database, and execute it.

				Dim Command As New SqlCommand("CREATE DATABASE [" & zx & "] ON (NAME='" & zx & "', FILENAME='" & Databasename & "')", TempConnection)
				Command.ExecuteNonQuery()
				Command.Dispose()

				' Inform the user how to log on to the new database

				MsgBox("Database " & Databasename & " has been successfully created.", MsgBoxStyle.Information, "Create New Database")

				OpenADatabase(NewDatabase)
			Catch e As Exception
				MsgBox("Failed to create new database." & vbCrLf & e.Message, MsgBoxStyle.Exclamation, ProgramName)

			End Try
		End If

	End Sub

	'*******************************************************************

	' Function to open a database and return the open status.
	' In fact, we only open a connection to the database for the sole
	' purpose of determining it can be accessed.  All of the datasets
	' access their tables regardless of this connection.

	'*******************************************************************
	Public Function OpenADatabase(Optional OpenDatabaseName As String = "", Optional ReOpen As Boolean = False) As Boolean

		' Declare variables

		Dim o As Boolean

		If DB.State = ConnectionState.Open Then CloseDatabase()

		' If no database is specified, use the last one accessed.

		If OpenDatabaseName = "" Then My.Settings.LastDatabaseUsed = OpenDatabaseName

		' Remember the datapath and the database name.

		Datapath = GetPath(OpenDatabaseName)
		Databasename = OpenDatabaseName

		' Now try to open a connection.  We'll return True or False depending on whether
		' we succeed.

		o = True
		Try
			If DB.State = ConnectionState.Open Then
				DB.Close()
			End If

			DB.ConnectionString = MyConnectionString()
			DB.Open()

			' Display the list of tables

			frmObjectList.MdiParent = frmMain
			frmObjectList.Show()

			' Display the database name on the form caption

			frmMain.Text = ProgramName & " (" & LCase(Databasename) & ")"

		Catch ex As Exception
			MsgBox("Cannot open database." & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, ProgramName)
			o = False

		End Try

		' Enable menu items

		EnableMenus(o)

		'Enable toolbar buttons which depend on an open database

		EnableButtons(o)

		' If we succeeded, do things that require an open database.

		If o Then

			' Update the Most Recently Used list in the File
			' menu.

			UpdateMRU(Databasename)
		End If

		Return o

	End Function
	'*******************************************************************

	' Function to return the path of a file name.

	'*******************************************************************
	Public Function GetPath(ByVal FileName As String) As String

		' Declare variables

		Dim jj As Integer
		Dim Datapath As String = ""

		If InStr(1, FileName, "\") > 0 Then
			For jj = Len(FileName) To 1 Step -1
				If Mid(FileName, jj, 1) = "\" Then
					Datapath = Left(FileName, jj)
					Exit For
				End If
			Next jj
		End If

		Return Datapath


	End Function
	'*******************************************************************

	' Sub to close the database.

	'*******************************************************************
	Public Sub CloseDatabase()


		' Close any open forms

		Do While My.Application.OpenForms.Count > 1
			My.Application.OpenForms(1).Close()
		Loop

		' Close the database.

		DB.Close()

		' Indicate the database is now closed.

		DbOpen = False

		' Disable the menu options which depend on an open database

		EnableMenus(False)

		'Disable toolbar buttons which depend on an open database

		EnableButtons(False)

		' Display the program name on the form caption

		frmMain.Text = ProgramName

	End Sub

	'**************************************************
	'
	' Sub to fill the MRU slots in the FILE menu with
	' the list of Most Recently Used databases
	'
	'**************************************************
	Public Sub FillMRU()

		' Declare variables

		Dim jj As Integer
		Dim zx As String
		Dim m As ToolStripMenuItem

		' See if any of the 8 MRU lines is defined.

		For jj = 1 To 8
			zx = GetSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", CStr(jj), "")
			m = Choose(jj, frmMain.MRU1, frmMain.MRU2, frmMain.MRU3, frmMain.MRU4, frmMain.MRU5, frmMain.MRU6, frmMain.MRU7, frmMain.MRU8)
			If zx <> "" Then
				m.Text = "&" & CStr(jj) & " " & zx
				m.Visible = True
				frmMain.MRUSeparator.Visible = True
			Else
				m.Visible = False
			End If
		Next jj
		System.Windows.Forms.Application.DoEvents()


	End Sub

	'**************************************************
	'
	' Sub to update and refill the Most Recently Used
	' database slots of the FILE menu
	'
	'**************************************************
	Public Sub UpdateMRU(ByRef DBName As String)

		' Declare variables

		Dim DbInList As Boolean
		Dim jj As Integer
		Dim j1 As Integer
		Dim zx As String

		' Always make the name initial-cap

		DBName = Capitalize(DBName)

		' See if this database is already among the last 4
		' MRU databases and, if so, move the ones heretofor
		' more recently used down.

		DbInList = False
		For jj = 1 To 8
			zx = GetSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", CStr(jj), "")
			If Trim(zx.ToLower) = LCase(DBName) Then
				DbInList = True
				If jj > 1 Then
					For j1 = jj - 1 To 1 Step -1
						zx = GetSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", CStr(j1), "")
						SaveSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", CStr(j1 + 1), zx)
					Next j1
					SaveSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", "1", LCase(DBName))
				End If
			End If
		Next jj

		' Shuffle any already-defined MRUs down and drop
		' off number 8 (least recently used) if the
		' database just opened was not in the last 8 MRUs

		If DbInList = False Then
			zx = GetSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", "1", "")
			If LCase(DBName) <> LCase(zx) And zx <> "" Then
				For jj = 8 To 2 Step -1
					zx = GetSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", CStr(jj - 1), "")
					If Trim(zx) <> "" Then
						SaveSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", CStr(jj), zx)
					End If
				Next jj
			End If
			SaveSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", "1", LCase(DBName))
		End If

		' Now refill the MRU slots

		FillMRU()

	End Sub
	'**************************************************
	'
	' Sub to delete and refill the Most Recently Used
	' database slots of the FILE menu
	'
	'**************************************************
	Public Sub DeleteMRU(ByRef Index As Integer)

		' Declare variables

		Dim jj As Integer
		Dim zx As String


		' See if this database is among the last 4
		' MRU databases and, if so, move the ones less
		' recently used up.

		For jj = Index To 8
			zx = GetSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", CStr(jj + 1), "")
			If zx <> "" Then
				SaveSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", CStr(jj), zx)
			Else
				If GetSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", CStr(jj), "") <> "" Then DeleteSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "MRUList", CStr(jj))
			End If
		Next jj


		' Now refill the MRU slots

		FillMRU()

	End Sub
	'**************************************************
	'
	' Sub to enable or disable menu items based upon
	' the State flag passed (State or State).
	'
	'**************************************************
	Public Sub EnableMenus(ByRef State As Boolean)

		frmMain.mnuTableDesigner.Enabled = State
		frmMain.mnuScriptEditor.Enabled = State
		frmMain.mnuCloseDatabase.Enabled = State
		frmMain.mnuRelationships.Enabled = State
		frmMain.mnuGenerateCommands.Enabled = State
		frmMain.mnuCompactDB.Enabled = State
		System.Windows.Forms.Application.DoEvents() ' Allow system to update display

	End Sub

	'**************************************************
	'
	' Sub to enable or disable toolbar buttons depending
	' upon the state flag passed (State or false).
	'
	'**************************************************
	Public Sub EnableButtons(ByRef State As Boolean)

		System.Windows.Forms.Application.DoEvents() ' Allow system to update display

	End Sub
	'**************************************************

	' Function to find a row in a datatable.
	' Input :  Table    = Datatable
	'          Criteria = The criteria to be met.
	' Output:  The index of the first record to meet
	'          the criteria.

	'**************************************************
	Public Function Find(Table As DataTable, Criteria As String, Optional IgnoreBlanks As Boolean = False)

		' Declare variables

		Dim jj As Integer
		Dim FieldToSearch As String = ""
		Dim xx As String = ""
		Dim zx As String = ""
		Dim Compare As String = ""
		Dim CompareValue As Object = Nothing

		' Begin parsing the string

		For jj = 1 To Len(Criteria)

			' Once we encounter a comparison string, the field to search is the value before that position.

			zx = Mid(Criteria, jj, 1)
			If zx = "=" Or zx = "<" Or zx = ">" And FieldToSearch = "" Then
				FieldToSearch = Trim(VB.Left(Criteria, jj - 1))

				' Now look for the comparison operator

				xx = Mid(Criteria, jj, 2)
				If xx <> "<>" And xx <> "<=" And xx <> ">=" Then
					xx = Mid(Criteria, jj, 1)
					Compare = Mid(Criteria, jj + 1)
				Else
					Compare = Mid(Criteria, jj + 2)
				End If

				Exit For
			End If
		Next jj

		' Now convert the comparison value too a string or a value.

		If Left(Compare, 1) = "'" Then
			CompareValue = LCase(SRep(Compare, 1, "'", ""))
		Else
			CompareValue = Val(Compare)
		End If

		' Begin searching the data table. If any errors occur, return a NOMATCH.

		Try
			For jj = 0 To Table.Rows.Count - 1
				zx = LCase(GetR(Table.Rows(jj), FieldToSearch)) ' Lower case everything for a case-insensitive search
				If IgnoreBlanks Then zx = Trim(zx)
				Select Case xx
					Case "="
						If zx = CompareValue Then Return jj
					Case "<>"
						If zx <> CompareValue Then Return jj
					Case "<"
						If zx < CompareValue Then Return jj
					Case ">"
						If zx > CompareValue Then Return jj
					Case ">="
						If zx >= CompareValue Then Return jj
					Case "<="
						If zx <= CompareValue Then Return jj
				End Select
			Next jj
		Catch
			Return -1 ' NOMATCH
		End Try

		' If we didn't find anything, return -1, also defined as the constant "NOMATCH"
		' in the Global.vb module.

		Return -1 ' NOMATCH

	End Function
	'**************************************************

	' Function to find a row in a datatable.  This version
	' uses a binary search for speed, but finds ONLY
	' an equals condition.  The table must be sorted
	' by the search column.
	' Input :  Table    = Datatable
	'          Compare = The value to search for.
	' Output:  The index of the first record that equals
	'          the searched-for value.

	'**************************************************
	Public Function Find2(Table As DataTable, Criteria As String, Optional IgnoreBlanks As Boolean = False)

		' Declare variables

		Dim ii As Integer
		Dim jj As Integer
		Dim Top As Integer
		Dim Bottom As Integer
		Dim Midl As Integer
		Dim FieldToSearch As String = ""
		Dim xx As String = ""
		Dim zx As String = ""
		Dim Compare As String = ""
		Dim CompareValue As Object = Nothing

		' Begin parsing the string

		For jj = 1 To Len(Criteria)

			' Once we encounter a comparison string, the field to search is the value before that position.

			zx = Mid(Criteria, jj, 1)
			If zx = "=" Or zx = "<" Or zx = ">" And FieldToSearch = "" Then
				FieldToSearch = Trim(VB.Left(Criteria, jj - 1))

				' Now look for the comparison operator and extract
				' the searched-for value that follows it.

				xx = Mid(Criteria, jj, 2)
				If xx <> "<>" And xx <> "<=" And xx <> ">=" Then
					xx = Mid(Criteria, jj, 1)
					Compare = Mid(Criteria, jj + 1)
				Else
					Compare = Mid(Criteria, jj + 2)
				End If

				Exit For
			End If
		Next jj

		' Now convert the comparison value to a string or a value.

		If Left(Compare, 1) = "'" Then
			CompareValue = LCase(SRep(Compare, 1, "'", ""))
		Else
			CompareValue = Val(Compare)
		End If

		' Begin searching the data table. If any errors occur, return a NOMATCH.

		Try
			Bottom = 0
			Top = Table.Rows.Count - 1
			Midl = Int(Bottom + (Top - Bottom + 0.1) / 2)
			Do
				xx = LCase(GetR(Table.Rows(Midl), FieldToSearch)) ' Lower case everything for a case-insensitive search
				If IgnoreBlanks Then xx = Trim(zx)
				If CompareValue = xx Then
					Return Midl
				ElseIf CompareValue > xx Then
					If Bottom = Midl Then Return -1
					Bottom = Midl
					If Top - Midl > 1 Then ii = Int((Top - Midl) / 2) Else ii = 1
					Midl = Midl + ii
				ElseIf CompareValue < xx Then
					If Top = Midl Then Return -1
					Top = Midl
					If Midl - Bottom > 1 Then ii = Int((Midl - Bottom) / 2) Else ii = 1
					Midl = Midl - ii
				End If
			Loop While ii > 0 And Midl >= 0 And Midl < Table.Rows.Count
		Catch
			Return -1 ' NOMATCH
		End Try

		' If we didn't find anything, return -1, also defined as the constant "NOMATCH"
		' in the Global.vb module.

		Return -1 ' NOMATCH

	End Function
	'*******************************************************************

	' Function to return the name of the a file without the path.

	'*******************************************************************
	Public Function FileNameNoPath(d As String) As String

		' Declare variables

		Dim jj As Integer
		Dim zx As String = ""

		' Peel off the path information and return just the database name.

		For jj = d.Length - 1 To 0 Step -1
			If d.Substring(jj, 1) = "\" Then
				zx = d.Substring(jj + 1)
				Return zx
			End If
		Next jj

		Return d

	End Function
	'*******************************************************************

	' Sub to generate the Visual Basic module containing all the
	' data access commands (SELECT,UPDATE,INSERT and DELETE)
	' for each table in a database.

	'*******************************************************************
	Public Sub GenerateDataAccessCommandModule(Optional SaveFileName As String = "")

		' Declare variables

		Dim ii As Integer
		Dim jj As Integer
		Dim xx As String
		Dim zx As String
		Dim VBText As String = ""
		Dim DA As New SqlDataAdapter
		Dim DS As New DataSet
		Dim DT As DataTable
		Dim x As DataTable
		Dim Cmd As SqlCommand

		' First assemble the header information.

		VBText = "Module " & frmGenerateCodeOptions.txtAbbrev.Text & "_DataAdapterQueries" & vbCrLf
		VBText &= vbTab & "'*******************************************************************" & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "' " & frmGenerateCodeOptions.txtName.Text & " Data Adapter functions." & vbCrLf
		VBText &= vbTab & "' " & frmGenerateCodeOptions.txtAbbrev.Text & "_DATAADAPTERQUERIES.VB" & vbCrLf
		VBText &= vbTab & "' Auto-generated by Sirius SQL Database Manager: " & VB.Format(Now, "MMMM yyyy") & vbCrLf
		VBText &= vbTab & "' Copyright " & Now.Year & " Sirius Software All Rights Reserved" & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "'*******************************************************************" & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "' Declare the global data adapters" & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "Public ServerName As String" & vbCrLf

		' Now add the code to create the data adapters for all the tables beneath the database.

		x = DB.GetSchema("Tables")
		For jj = 0 To x.Rows.Count - 1

			' Make sure the "table" is type BASE TABLE and not VIEW.  A view looks like a table
			' except for the table type, and we list views separately.  Ignore any
			' table that begins with "db_" as those are the settings used by this program
			' and not the database user.

			If x.Rows(jj).Item("table_type") = "BASE TABLE" And VB.Left(x.Rows(jj).Item("table_name"), 3) <> "db_" Then VBText &= vbTab & "Public " & x.Rows(jj).Item("table_name") & "DA As SQLDataAdapter" & vbCrLf
			System.Windows.Forms.Application.DoEvents()
		Next jj
		VBText &= vbCrLf

		' Now add the function to create the connection string.

		VBText &= vbTab & "'*******************************************************************" & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "' Function to return the connection string to the current database" & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "'*******************************************************************" & vbCrLf
		VBText &= vbTab & "Public Function MyConnectionString() As String" & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "' For security purposes, and to ensure the connection string is" & vbCrLf
		VBText &= vbTab & "' correct, use a connection string builder." & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "Dim cs As New SqlConnectionStringBuilder" & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "'Get the server name.  The default is the old 2012 version of SQL Server Express." & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "ServerName = ""(LocalDB)\"" & GetSetting(""SiriusSoftwareGlobal"", ""SQLServer"", ""InstanceName"", ""v11.0"")" & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "cs.Add(""Data Source"", ServerName)" & vbCrLf
		VBText &= vbTab & "cs.Add(""AttachDbFilename"", Databasename)" & vbCrLf
		VBText &= vbTab & "cs.Add(""Integrated Security"", True)" & vbCrLf
		VBText &= vbTab & "cs.Add(""Connect Timeout"", ""60"")" & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "Return cs.ConnectionString" & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "End Function" & vbCrLf & vbCrLf


		' Now loop through all the tables, creating SELECT, INSERT, UPDATE and DELETE command code.

		For jj = 0 To x.Rows.Count - 1

			' Make sure the "table" is type BASE TABLE and not VIEW.  A view looks like a table
			' except for the table type.

			If x.Rows(jj).Item("table_type") = "BASE TABLE" And VB.Left(x.Rows(jj).Item("table_name"), 3) <> "db_" Then

				' Open the system table that contains the names of the columns in the selected table, and 
				' the data type of each column.

				' The columns of the schema table have the following names:
				' TABLE_NAME,COLUMN_NAME,COLUMN_DEFAULT,IS_NULLABLE,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH,NUMERIC_PRECISION,DATETIME_PRECISION

				Cmd = New SqlCommand("SELECT * FROM information_schema.columns WHERE table_name = '" & x.Rows(jj).Item("table_name") & "' ORDER BY ordinal_position", DB)
				DA.SelectCommand = Cmd
				DS.Clear()
				DA.Fill(DS, "Table")
				DT = DS.Tables("Table")


				' Create the code to build the SELECT command.

				VBText &= vbTab & "'*******************************************************************" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "' Function to return the SELECT SQL command to select records" & vbCrLf
				VBText &= vbTab & "' from the " & x.Rows(jj).Item("table_name") & " table." & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "'*******************************************************************" & vbCrLf
				VBText &= vbTab & "Public Function " & x.Rows(jj).Item("table_name") & "SelectCommand() As SqlCommand" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "' Declare variables" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "Dim Command = New SqlCommand(""SELECT * FROM [dbo].[" & x.Rows(jj).Item("table_name") & "]"",DB)" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "' Set command type" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "Command.CommandType = CommandType.Text" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "Return Command" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "End Function" & vbCrLf & vbCrLf

				' Create the code to build the INSERT command.

				VBText &= vbTab & "'*******************************************************************" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "' Function to return the INSERT SQL command to insert a record" & vbCrLf
				VBText &= vbTab & "' into the " & x.Rows(jj).Item("table_name") & " table." & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "'*******************************************************************" & vbCrLf
				VBText &= vbTab & "Public Function " & x.Rows(jj).Item("table_name") & "InsertCommand() As SqlCommand" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "' Declare variables" & vbCrLf
				VBText &= vbCrLf

				zx = ""
				For ii = 0 To DT.Rows.Count - 1
					If UCase(DT.Rows(ii)("COLUMN_NAME")) <> "ID" Then
						If zx <> "" Then zx &= ","
						zx &= "[" & DT.Rows(ii)("COLUMN_NAME") & "]"
					End If
				Next ii
				VBText &= vbTab & "Dim Command = New SQLCommand (""INSERT INTO [dbo].[" & x.Rows(jj).Item("table_name") & "] (" & zx & ") VALUES ("

				zx = ""
				For ii = 0 To DT.Rows.Count - 1
					If UCase(DT.Rows(ii)("COLUMN_NAME")) <> "ID" Then
						If zx <> "" Then zx &= ","
						zx &= "@" & DT.Rows(ii)("COLUMN_NAME")
					End If
				Next ii
				VBText &= zx & ");"" & vbCrLf & """


				zx = ""
				For ii = 0 To DT.Rows.Count - 1
					If zx <> "" Then zx &= ","
					zx &= DT.Rows(ii)("COLUMN_NAME")
				Next ii
				VBText &= " SELECT " & zx & " FROM " & x.Rows(jj).Item("table_name") & " WHERE (Id = SCOPE_IDENTITY())"", DB)" & vbCrLf

				VBText &= vbCrLf
				VBText &= vbTab & "' Set command type" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "Command.CommandType = CommandType.Text" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "' Add the parameters for the insert query." & vbCrLf
				VBText &= vbCrLf

				zx = ""
				xx = ""
				For ii = 0 To DT.Rows.Count - 1
					If GetR(DT.Rows(ii), "CHARACTER_MAXIMUM_LENGTH") = 0 Then
						If DT.Rows(ii)("DATA_TYPE") = "bit" Then xx = "1"
						If DT.Rows(ii)("DATA_TYPE") = "int" Then xx = "4"
					Else
						xx = DT.Rows(ii)("CHARACTER_MAXIMUM_LENGTH")
					End If
					If UCase(DT.Rows(ii)("COLUMN_NAME")) <> "ID" Then zx &= vbTab & "Command.Parameters.Add(""@" & DT.Rows(ii)("COLUMN_NAME") & """, SqlDbType." & DT.Rows(ii)("DATA_TYPE") & ", " & xx & ", """ & DT.Rows(ii)("COLUMN_NAME") & """)" & vbCrLf
				Next ii
				VBText &= zx & vbCrLf

				VBText &= vbTab & "Return Command" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "End Function" & vbCrLf & vbCrLf

				' Create the code to build the UPDATE command.

				VBText &= vbTab & "'*******************************************************************" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "' Function to return the UPDATE SQL command to update a record" & vbCrLf
				VBText &= vbTab & "' in the " & x.Rows(jj).Item("table_name") & " table." & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "'*******************************************************************" & vbCrLf
				VBText &= vbTab & "Public Function " & x.Rows(jj).Item("table_name") & "UpdateCommand() As SqlCommand" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "' Declare variables" & vbCrLf
				VBText &= vbCrLf

				zx = ""
				For ii = 0 To DT.Rows.Count - 1
					If UCase(DT.Rows(ii)("COLUMN_NAME")) <> "ID" Then
						If zx <> "" Then zx &= ","
						zx &= "[" & DT.Rows(ii)("COLUMN_NAME") & "]=@" & DT.Rows(ii)("COLUMN_NAME")
					End If
				Next ii
				VBText &= vbTab & "Dim Command = New SqlCommand(""UPDATE [dbo].[" & x.Rows(jj).Item("table_name") & "] SET " & zx & " WHERE [Id]=@Id;"" & vbCrLf & """

				zx = ""
				For ii = 0 To DT.Rows.Count - 1
					If zx <> "" Then zx &= ","
					zx &= DT.Rows(ii)("COLUMN_NAME")
				Next ii

				VBText &= " SELECT " & zx & " FROM " & x.Rows(jj).Item("table_name") & " WHERE [Id]=@Original_Id"", DB)" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "' Set command type" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "Command.CommandType = CommandType.Text" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "' Add the parameters for the update query." & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "Command.Parameters.Add(""@Id"", SqlDbType.Int, 4, ""ID"")" & vbCrLf
				VBText &= vbTab & "Command.Parameters.Add(""@Original_Id"", SqlDbType.Int, 4, ""ID"")" & vbCrLf

				zx = ""
				For ii = 0 To DT.Rows.Count - 1
					If GetR(DT.Rows(ii), "CHARACTER_MAXIMUM_LENGTH") = 0 Then
						If DT.Rows(ii)("DATA_TYPE") = "bit" Then xx = "1"
						If DT.Rows(ii)("DATA_TYPE") = "int" Then xx = "4"
					Else
						xx = DT.Rows(ii)("CHARACTER_MAXIMUM_LENGTH")
					End If
					If UCase(DT.Rows(ii)("COLUMN_NAME")) <> "ID" Then zx &= vbTab & "Command.Parameters.Add(""@" & DT.Rows(ii)("COLUMN_NAME") & """, SqlDbType." & DT.Rows(ii)("DATA_TYPE") & ", " & xx & ", """ & DT.Rows(ii)("COLUMN_NAME") & """)" & vbCrLf
				Next ii
				VBText &= zx & vbCrLf

				VBText &= vbCrLf
				VBText &= vbTab & "Return Command" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "End Function" & vbCrLf & vbCrLf

				' Create the code to build the DELETE command.

				VBText &= vbTab & "'*******************************************************************" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "' Function to return the DELETE SQL command to delete a record" & vbCrLf
				VBText &= vbTab & "' from the " & x.Rows(jj).Item("table_name") & " table." & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "'*******************************************************************" & vbCrLf
				VBText &= vbTab & "Public Function " & x.Rows(jj).Item("table_name") & "DeleteCommand() As SqlCommand" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "' Declare variables" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "Dim Command = New SqlCommand(""DELETE FROM [dbo].[" & x.Rows(jj).Item("table_name") & "] WHERE [Id] = @Original_ID"",DB)" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "' Set command type" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "Command.CommandType = CommandType.Text" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "' Add the parameters for the delete query." & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "Command.Parameters.Add(""@Original_ID"", SqlDbType.Int, 4, ""ID"")" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "Return Command" & vbCrLf
				VBText &= vbCrLf
				VBText &= vbTab & "End Function" & vbCrLf & vbCrLf
			End If
		Next jj

		' Now add the code to create the "InitializeDataAdapters" sub.


		VBText &= vbTab & "'*******************************************************************" & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "' Sub to initialize the global data adapters." & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "'*******************************************************************" & vbCrLf
		VBText &= vbTab & "Public Sub InitializeDataAdapters()" & vbCrLf & vbCrLf

		' Loop through all the tables.

		For jj = 0 To x.Rows.Count - 1

			' Make sure the "table" is type BASE TABLE and not VIEW.  A view looks like a table
			' except for the table type. Ignore tables that begin with "db_" as these are settings
			' tables used by this program and not the database user.

			If x.Rows(jj).Item("table_type") = "BASE TABLE" And VB.Left(x.Rows(jj).Item("table_name"), 3) <> "db_" Then
				VBText &= vbTab & x.Rows(jj)("table_name") & "DA = New SqlDataAdapter" & vbCrLf
				VBText &= vbTab & x.Rows(jj)("table_name") & "DA.SelectCommand = " & x.Rows(jj)("table_name") & "SelectCommand()" & vbCrLf
				VBText &= vbTab & x.Rows(jj)("table_name") & "DA.InsertCommand = " & x.Rows(jj)("table_name") & "InsertCommand()" & vbCrLf
				VBText &= vbTab & x.Rows(jj)("table_name") & "DA.UpdateCommand = " & x.Rows(jj)("table_name") & "UpdateCommand()" & vbCrLf
				VBText &= vbTab & x.Rows(jj)("table_name") & "DA.DeleteCommand = " & x.Rows(jj)("table_name") & "DeleteCommand()" & vbCrLf
				VBText &= vbTab & x.Rows(jj)("table_name") & "DA.TableMappings.Add(""" & x.Rows(jj)("table_name") & """, ""Table"")" & vbCrLf
				VBText &= vbCrLf
				System.Windows.Forms.Application.DoEvents()
			End If
		Next jj
		VBText &= vbTab & "End Sub" & vbCrLf & vbCrLf

		' Add the code to create the "DetachDataAdapters" sub.

		VBText &= vbTab & "'*******************************************************************" & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "' Sub to detach the global data adapters from the database." & vbCrLf
		VBText &= vbCrLf
		VBText &= vbTab & "'*******************************************************************" & vbCrLf
		VBText &= vbTab & "Public Sub DetachDataAdapters()" & vbCrLf & vbCrLf

		For jj = 0 To x.Rows.Count - 1

			' Make sure the "table" is type BASE TABLE and not VIEW.  A view looks like a table
			' except for the table type.  Ignore tables that begin with "db_" as these are settings
			' used by this program and not the database user.

			If x.Rows(jj).Item("table_type") = "BASE TABLE" And VB.Left(x.Rows(jj).Item("table_name"), 3) <> "db_" Then
				VBText &= vbTab & x.Rows(jj)("table_name") & "DA.SelectCommand.Connection = Nothing" & vbCrLf
				VBText &= vbTab & x.Rows(jj)("table_name") & "DA.InsertCommand.Connection = Nothing" & vbCrLf
				VBText &= vbTab & x.Rows(jj)("table_name") & "DA.UpdateCommand.Connection = Nothing" & vbCrLf
				VBText &= vbTab & x.Rows(jj)("table_name") & "DA.DeleteCommand.Connection = Nothing" & vbCrLf
				VBText &= vbTab & x.Rows(jj)("table_name") & "DA.TableMappings.Clear" & vbCrLf
				VBText &= vbTab & x.Rows(jj)("table_name") & "DA.Dispose()" & vbCrLf
				VBText &= vbTab & x.Rows(jj)("table_name") & "DS.Dispose()" & vbCrLf
				VBText &= vbCrLf
				System.Windows.Forms.Application.DoEvents()
			End If
		Next jj
		VBText &= vbTab & "End Sub" & vbCrLf
		VBText &= "End Module" & vbCrLf

		' Save the assembled code to the output file.

		If SaveFileName <> "" Then
			My.Computer.FileSystem.WriteAllText(SaveFileName, VBText, False)
		Else
			Clipboard.SetText(VBText)
		End If


	End Sub
	'**************************************************

	' Sub to compact the database.

	'**************************************************
	Public Sub CompactDatabase()

		' Declare variables

		Dim idx As DataTable
		Dim jj As Integer
		Dim TableName As String
		Dim IndexName As String
		Dim Command As SqlCommand

		' Put up a status message.

		frmMain.StatusLabel.Text = "Compacting database..."
		System.Windows.Forms.Application.DoEvents()

		' Create a command to compact the database, leaving 10% free space.

		Command = New SqlCommand("DBCC SHRINKDATABASE('" & Databasename & "', 10)", DB)
		Command.CommandTimeout = 0 ' No timeout: this can take a lot of time to compact a db
		Command.ExecuteNonQuery()
		Command.Dispose()

		' Get a list of all indexes in the database.

		idx = DB.GetSchema("Indexes")

		' Rebuild all the indexes.

		For jj = 0 To idx.Rows.Count - 1
			IndexName = idx.Rows(jj).Item("constraint_name")
			TableName = idx.Rows(jj).Item("table_name")
			frmMain.StatusLabel.Text = "Rebuilding Index " & IndexName & " on table " & TableName & "..."
			System.Windows.Forms.Application.DoEvents()
			Command = New SqlCommand("ALTER INDEX " & IndexName & " ON " & TableName & " REBUILD;", DB)
			Command.CommandTimeout = 0 ' No timeout: this can take a lot of time to rebuild an index
			Command.ExecuteNonQuery()
			Command.Dispose()
		Next jj

		' Clear status message.

		frmMain.StatusLabel.Text = ""
		System.Windows.Forms.Application.DoEvents()

	End Sub
	'*******************************************************************

	' Function to return the name of the current database without
	' the path.

	'*******************************************************************
	Public Function DatabaseNameNoPath() As String

		' Declare variables

		Dim jj As Integer
		Dim zx As String = ""

		' Peel off the path information and return just the database name.

		For jj = Len(Databasename) To 1 Step -1
			If Mid(Databasename, jj, 1) = "\" Then
				zx = Mid(Databasename, jj + 1)
				Exit For
			End If
		Next jj

		Return zx

	End Function

	'*******************************************************************

	' Function to return the connection string to the current database

	'*******************************************************************
	Public Function MyConnectionString() As String

		' For security purposes, and to ensure the connection string is
		' correct, use a connection string builder.

		Dim cs As New SqlConnectionStringBuilder

		' Build the connection string.

		cs.Add("Data Source", ServerName)
		cs.Add("AttachDbFilename", Databasename)
		cs.Add("Integrated Security", True)
		cs.Add("Connect Timeout", "60")

		Return cs.ConnectionString

	End Function
	'*******************************************************************

	' Function to return a collection of foreign keys attached to a table
	' when passed the table name.

	'*******************************************************************
	Public Function GetForeignKeys(TableName As String) As Dictionary(Of String, ForeignKeyValues)

		Dim TableID As Integer
		Dim DA As New SqlDataAdapter
		Dim DS As New DataSet
		Dim DT As DataTable

		' Get the ID of the table from the supplied name.

		Try

			DA.SelectCommand = New SqlCommand("SELECT object_id FROM sys.tables WHERE name='" & TableName & "'", DB)
			DS.Clear()
			DA.Fill(DS, "Table")
			DT = DS.Tables("Table")
			If DT.Rows.Count > 0 Then TableID = DT.Rows(0)("object_id")

			' Call the other version of this routine, and return the collection it returns.

			Return GetForeignKeys(TableID)

		Catch ex As Exception
			MsgBox("Error retrieving foreign keys." & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "GetForeignKeys(String) Function")
		End Try

		' If we failed here, return an empty collection.

		Return New Dictionary(Of String, ForeignKeyValues)

	End Function
	'**********************************************************************

	' Function to return a collection of constraint names that may
	' be attached to a field.
	' Constraints may be specified:
	'  Type = 0 = All
	'         1 = Default constraints only
	'         2 = Primary Key constraints only
	'         3 = Unique constraints only

	'**********************************************************************
	Public Function GetConstraints(ByVal TableName As String, ByVal FieldName As String, Optional Type As Integer = 0) As Collection

		' Declare variables

		Dim TableID As Integer
		Dim ColumnID As Integer
		Dim CList As New Collection
		Dim Cmd As SqlCommand
		Dim DA As New SqlDataAdapter
		Dim DS As New DataSet
		Dim DT As DataTable

		' Get the ID of the table, and the column ID of the column to be dropped.

		Cmd = New SqlCommand("SELECT sys.tables.object_id,sys.columns.column_id FROM sys.tables INNER JOIN sys.columns ON sys.tables.object_id=sys.columns.object_id WHERE sys.tables.name='" & TableName & "' AND sys.columns.name='" & FieldName & "'", DB)
		DA.SelectCommand = Cmd
		DA.Fill(DS, "Table")
		DT = DS.Tables("Table")
		If DT.Rows.Count > 0 Then
			TableID = DT.Rows(0)("object_id")
			ColumnID = DT.Rows(0)("column_id")
			DS.Clear()

			' Now get any default constraints attached to the field.

			If Type = 0 Or Type = 1 Then
				DA.SelectCommand.CommandText = "SELECT name FROM sys.default_constraints WHERE parent_object_id=" & TableID & " AND parent_column_id=" & ColumnID
				DA.Fill(DS, "Table")
				DT = DS.Tables("Table")
				If DT.Rows.Count > 0 Then
					For jj = 0 To DT.Rows.Count - 1
						CList.Add(DT.Rows(jj)("name"))
					Next jj
				End If
				DS.Clear()
			End If

			' Get any primary key constraints attached to the field

			If Type = 0 Or Type = 2 Then
				DA.SelectCommand.CommandText = "SELECT sys.key_constraints.name FROM (sys.columns INNER JOIN sys.index_columns ON sys.columns.object_id=sys.index_columns.object_id AND sys.columns.column_id=sys.index_columns.column_id) INNER JOIN sys.key_constraints ON sys.index_columns.object_id=sys.key_constraints.parent_object_id AND sys.key_constraints.unique_index_id=sys.index_columns.index_id WHERE sys.key_constraints.type='PK' AND sys.columns.name='" & FieldName & "' AND sys.columns.object_id=" & TableID
				DA.Fill(DS, "Table")
				DT = DS.Tables("Table")
				If DT.Rows.Count > 0 Then
					For jj = 0 To DT.Rows.Count - 1
						CList.Add(DT.Rows(jj)("name"))
					Next jj
				End If
				DS.Clear()
			End If

			' Get any unique constraints attched to the field

			If Type = 0 Or Type = 3 Then
				DA.SelectCommand.CommandText = "SELECT sys.key_constraints.name FROM (sys.columns INNER JOIN sys.index_columns ON sys.columns.object_id=sys.index_columns.object_id AND sys.columns.column_id=sys.index_columns.column_id) INNER JOIN sys.key_constraints ON sys.index_columns.object_id=sys.key_constraints.parent_object_id AND sys.key_constraints.unique_index_id=sys.index_columns.index_id WHERE sys.key_constraints.type='UQ' AND sys.columns.name='" & FieldName & "' AND sys.columns.object_id=" & TableID
				DA.Fill(DS, "Table")
				DT = DS.Tables("Table")
				If DT.Rows.Count > 0 Then
					For jj = 0 To DT.Rows.Count - 1
						CList.Add(DT.Rows(jj)("name"))
					Next jj
				End If
			End If
			DS.Clear()
		End If

		' Dispose of objects we've created.

		DA.Dispose()
		DS.Dispose()
		Cmd.Dispose()

		' Return the list of constraints attached to the field.

		Return CList

	End Function

	'*******************************************************************

	' Function to indicate if a string is a numeric default value or not.

	'*******************************************************************
	Public Function IsNumericDefault(s As String) As Boolean

		' Declare variables

		Dim jj As Integer
		Dim zx As Char

		' Check each character.  If the string is numeric, it may be
		' enclosed by parentheses, or not.  Ignore those.

		For jj = 1 To s.Length
			zx = Mid(s, jj, 1)
			If zx <> "(" And zx <> ")" And (zx < "0" Or zx > "9") Then Return False
		Next jj

		Return True

	End Function
	'*******************************************************************

	' Function to return a collection of foreign keys attached to a table
	' when passed the table object id.

	'*******************************************************************
	Public Function GetForeignKeys(TableObjectID As Integer) As Dictionary(Of String, ForeignKeyValues)

		' Declare variables

		Dim jj As Integer
		Dim TableName As String = ""
		Dim DA As New SqlDataAdapter
		Dim DS1 As New DataSet
		Dim DS2 As New DataSet
		Dim DT1 As DataTable
		Dim DT2 As DataTable
		Dim FKValues As ForeignKeyValues
		Dim FKeys As New Dictionary(Of String, ForeignKeyValues)

		' Get the name of the table from the supplied ID.

		Try

			DA.SelectCommand = New SqlCommand("SELECT name FROM sys.tables WHERE object_id=" & TableObjectID, DB)
			DS1.Clear()
			DA.Fill(DS1, "Table")
			DT1 = DS1.Tables("Table")
			If DT1.Rows.Count > 0 Then TableName = DT1.Rows(0)("name")

			' Get the foreign keys attached to this table.

			DA.SelectCommand = New SqlCommand("SELECT name,delete_referential_action, update_referential_action, sys.foreign_key_columns.* FROM sys.foreign_keys INNER JOIN sys.foreign_key_columns ON sys.foreign_keys.object_id=sys.foreign_key_columns.constraint_object_id WHERE sys.foreign_keys.parent_object_id=" & TableObjectID, DB)
			DS1.Clear()
			DA.Fill(DS1, "Table")
			DT1 = DS1.Tables("Table")

			' For each foreign key, get the name of the field in the table, the name of the field in the
			' related table, and the name of the related table.

			For jj = 0 To DT1.Rows.Count - 1
				FKValues = New ForeignKeyValues
				FKValues.PrimaryTable = TableName
				FKValues.RelationshipName = DT1.Rows(jj)("name")
				FKValues.CascadeDelete = CBool(DT1.Rows(jj)("delete_referential_action")) ' 0 = NO_ACTION, 1 = CASCADE
				FKValues.CascadeUpdate = CBool(DT1.Rows(jj)("update_referential_action")) ' 0 = NO_ACTION, 1 = CASCADE

				' Get the name of the related table

				DS2.Clear()
				DA.SelectCommand = New SqlCommand("SELECT name FROM sys.tables WHERE object_id=" & DT1.Rows(jj)("referenced_object_id"), DB)
				DA.Fill(DS2, "Table")
				DT2 = DS2.Tables("Table")
				FKValues.RelatedTable = DT2.Rows(0)("name")

				' Get the name of the primary field

				DS2.Clear()
				DA.SelectCommand = New SqlCommand("SELECT name FROM sys.columns WHERE object_id=" & TableObjectID & " AND column_id=" & DT1.Rows(jj)("parent_column_id"), DB)
				DA.Fill(DS2, "Table")
				DT2 = DS2.Tables("Table")
				FKValues.PrimaryField = DT2.Rows(0)("name")

				' Get the name of the related field in the related table

				DS2.Clear()
				DA.SelectCommand = New SqlCommand("SELECT name FROM sys.columns WHERE object_id=" & DT1.Rows(jj)("referenced_object_id") & " and column_id=" & DT1.Rows(jj)("referenced_column_id"), DB)
				DA.Fill(DS2, "Table")
				DT2 = DS2.Tables("Table")
				FKValues.RelatedField = DT2.Rows(0)("name")

				' Add the relationship to the panel

				FKeys.Add(FKValues.RelationshipName, FKValues)
			Next jj

			' Trap errors

		Catch ex As Exception
			MsgBox("Error retrieving foreign keys." & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "GetForeignKeys(Integer) Function")

		End Try

		' Return the collection

		GetForeignKeys = FKeys

	End Function
	'**************************************************
	'
	' Property to get or set the message box theme
	'
	'**************************************************
	Public Property ProgramColorTheme(Optional Component As ThemeItem = ThemeItem.All) As String
		Get

			' Declare variables

			Dim zx As String

			' Get the theme string.

			zx = GetSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "ProgramColorTheme", "ThemeStyle", "0x1111")

			' Determine what portion of the theme to return.

			Select Case Component
				Case ThemeItem.All
					ProgramColorTheme = zx
				Case ThemeItem.MainWindow
					ProgramColorTheme = Mid(zx, 3, 1)
				Case ThemeItem.BackgroundColorNumber
					ProgramColorTheme = Mid(zx, 4, 1)
				Case ThemeItem.TextColorNumber
					ProgramColorTheme = Mid(zx, 5, 1)
				Case ThemeItem.FontNumber
					ProgramColorTheme = Mid(zx, 6, 1)
				Case Else
					ProgramColorTheme = zx
			End Select
		End Get
		Set(value As String)

			' Declare variables.

			Dim x1 As String
			Dim x2 As String
			Dim x3 As String
			Dim x4 As String

			' Extract the portions of the theme to test them.

			x1 = Mid(value, 3, 1)
			x2 = Mid(value, 4, 1)
			x3 = Mid(value, 5, 1)
			x4 = Mid(value, 6, 1)

			' Now validate each value.

			If (x1 < "1" Or x1 > "4") Or (x2 < "1" Or x2 > "3") Or (x3 < "1" Or x3 > "4") Or (x4 < "1" Or x4 > "2") Then

				' We will use the windows message box to display errors, to prevent reentrancy 
				' issues with the routine calling itself.

				VB.MsgBox("Invalid theme value: """ & value & """.  Theme not saved.", MsgBoxStyle.Exclamation, "Save Theme Property")
			Else

				' Save the new theme. 

				SaveSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "ProgramColorTheme", "ThemeStyle", value)
			End If
		End Set
	End Property
End Module
