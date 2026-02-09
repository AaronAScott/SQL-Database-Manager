Imports System.Data.SqlClient
Imports VB = Microsoft.VisualBasic
Public Class frmScriptEditor
	Inherits System.Windows.Forms.Form
	'**********************************************************
	' SQL Data Manager SQL Editor (frmScriptEditor)
	' DM_SCRIPTEDITOR.VB
	' Written: October 2017
	' Programmer: Aaron Scott
	' Copyright (C) 2017 Sirius Software All Rights Reserved
	'**********************************************************

	' Declare variables local to this module

	Private WordListsSorted As Boolean
	Private PreviousCaretPosition As Integer = 0
	Private ScriptSaved As Boolean = True
	Private ScriptFileName As String = ""

	' Define the list of reserved SQL keywords.  These words will be highlighted blue, if supported.

	Private ReservedWords() As String = {"ADD", "EXTERNAL", "PROCEDURE", "ALL", "FETCH", "PUBLIC", "ALTER", "FILE", "RAISERROR", "AND", "FILLFACTOR", "READ", "ANY", "FOR", "READTEXT", "AS", "FOREIGN", "RECONFIGURE", "ASC", "FREETEXT", "REFERENCES", "AUTHORIZATION", "FREETEXTTABLE", "REPLICATION", "BACKUP", "FROM", "RESTORE", "BEGIN", "FULL", "RESTRICT", "BETWEEN", "FUNCTION", "RETURN", "BREAK", "GOTO", "REVERT", "BROWSE", "GRANT", "REVOKE", "BULK", "GROUP", "RIGHT", "BY", "HAVING", "ROLLBACK", "CASCADE", "HOLDLOCK", "ROWCOUNT", "CASE", "IDENTITY", "ROWGUIDCOL", "CHECK", "IDENTITY_INSERT", "RULE", "CHECKPOINT", "IDENTITYCOL", "SAVE", "CLOSE", "IF", "SCHEMA", "CLUSTERED", "IN", "SECURITYAUDIT", "COALESCE", "INDEX", "SELECT", "COLLATE", "INNER", "SEMANTICKEYPHRASETABLE", "COLUMN", "INSERT", "SEMANTICSIMILARITYDETAILSTABLE", "COMMIT", "INTERSECT", "SEMANTICSIMILARITYTABLE", "COMPUTE", "INTO", "SESSION_USER", "CONSTRAINT", "IS", "SET", "CONTAINS", "JOIN", "SETUSER", "CONTAINSTABLE", "KEY", "SHUTDOWN", "CONTINUE", "KILL", "SOME", "CONVERT", "LEFT", "STATISTICS", "CREATE", "LIKE", "SYSTEM_USER", "CROSS", "LINENO", "TABLE", "CURRENT", "LOAD", "TABLESAMPLE", "CURRENT_DATE", "MERGE", "TEXTSIZE", "CURRENT_TIME", "NATIONAL", "THEN", "CURRENT_TIMESTAMP", "NOCHECK", "TO", "CURRENT_USER", "NONCLUSTERED", "TOP", "CURSOR", "NOT", "TRAN", "DATABASE", "NULL", "TRANSACTION", "DBCC", "NULLIF", "TRIGGER", "DEALLOCATE", "OF", "TRUNCATE", "DECLARE", "OFF", "TRY_CONVERT", "DEFAULT", "OFFSETS", "TSEQUAL", "DELETE", "ON", "UNION", "DENY", "OPEN", "UNIQUE", "DESC", "OPENDATASOURCE", "UNPIVOT", "DISK", "OPENQUERY", "UPDATE", "DISTINCT", "OPENROWSET", "UPDATETEXT", "DISTRIBUTED", "OPENXML", "USE", "DOUBLE", "OPTION", "USER", "DROP", "OR", "VALUES", "DUMP", "ORDER", "VARYING", "ELSE", "OUTER", "VIEW", "END", "OVER", "WAITFOR", "ERRLVL", "PERCENT", "WHEN", "ESCAPE", "PIVOT", "WHERE", "EXCEPT", "PLAN", "WHILE", "EXEC", "PRECISION", "WITH", "EXECUTE", "PRIMARY", "WITHINGROUP", "EXISTS", "PRINT", "WRITETEXT", "EXIT", "PROC", "'odbckeywords", "ABSOLUTE", "EXEC", "OVERLAPS", "ACTION", "EXECUTE", "PAD", "ADA", "EXISTS", "PARTIAL", "ADD", "EXTERNAL", "PASCAL", "ALL", "EXTRACT", "POSITION", "ALLOCATE", "FALSE", "PRECISION", "ALTER", "FETCH", "PREPARE", "AND", "FIRST", "PRESERVE", "ANY", "FLOAT", "DOUBLE", "PRIMARY", "ARE", "FOR", "PRIOR", "AS", "FOREIGN", "PRIVILEGES", "ASC", "FORTRAN", "PROCEDURE", "ASSERTION", "FOUND", "PUBLIC", "AT", "FROM", "READ", "AUTHORIZATION", "FULL", "REAL", "AVG", "GET", "REFERENCES", "BEGIN", "GLOBAL", "RELATIVE", "BETWEEN", "GO", "RESTRICT", "BIT", "GOTO", "REVOKE", "BIT_LENGTH", "GRANT", "RIGHT", "BOTH", "GROUP", "ROLLBACK", "BY", "HAVING", "ROWS", "CASCADE", "HOUR", "SCHEMA", "CASCADED", "IDENTITY", "SCROLL", "CASE", "IMMEDIATE", "SECOND", "CAST", "IN", "SECTION", "CATALOG", "INCLUDE", "SELECT", "CHAR", "INDEX", "SESSION", "CHAR_LENGTH", "INDICATOR", "SESSION_USER", "CHARACTER", "INITIALLY", "SET", "CHARACTER_LENGTH", "INNER", "SIZE", "CHECK", "INPUT", "SMALLINT", "CLOSE", "INSENSITIVE", "SOME", "COALESCE", "INSERT", "SPACE", "COLLATE", "INT", "SQL", "COLLATION", "INTEGER", "SQLCA", "COLUMN", "INTERSECT", "SQLCODE", "COMMIT", "INTERVAL", "SQLERROR", "CONNECT", "INTO", "SQLSTATE", "CONNECTION", "IS", "SQLWARNING", "CONSTRAINT", "ISOLATION", "SUBSTRING", "CONSTRAINTS", "JOIN", "SUM", "CONTINUE", "KEY", "SYSTEM_USER", "CONVERT", "LANGUAGE", "TABLE", "CORRESPONDING", "LAST", "TEMPORARY", "COUNT", "LEADING", "THEN", "CREATE", "LEFT", "TIME", "CROSS", "LEVEL", "TIMESTAMP", "CURRENT", "LIKE", "TIMEZONE_HOUR", "CURRENT_DATE", "LOCAL", "TIMEZONE_MINUTE", "CURRENT_TIME", "LOWER", "TO", "CURRENT_TIMESTAMP", "MATCH", "TRAILING", "CURRENT_USER", "MAX", "TRANSACTION", "CURSOR", "MIN", "TRANSLATE", "DATE", "MINUTE", "TRANSLATION", "DAY", "MODULE", "TRIM", "DEALLOCATE", "MONTH", "TRUE", "DEC", "NAMES", "UNION", "DECIMAL", "NATIONAL", "UNIQUE", "DECLARE", "NATURAL", "UNKNOWN", "DEFAULT", "NCHAR", "UPDATE", "DEFERRABLE", "NEXT", "UPPER", "DEFERRED", "NO", "USAGE", "DELETE", "NONE", "USER", "DESC", "NOT", "USING", "DESCRIBE", "NULL", "VALUE", "DESCRIPTOR", "NULLIF", "VALUES", "DIAGNOSTICS", "NUMERIC", "VARCHAR", "DISCONNECT", "OCTET_LENGTH", "VARYING", "DISTINCT", "OF", "VIEW", "DOMAIN", "ON", "WHEN", "DOUBLE", "ONLY", "WHENEVER", "DROP", "OPEN", "WHERE", "ELSE", "OPTION", "WITH", "END", "OR", "WORK", "END-EXEC", "ORDER", "WRITE", "ESCAPE", "OUTER", "YEAR", "EXCEPT", "OUTPUT", "ZONE", "EXCEPTION", "'FutureKeywords", "ABSOLUTE", "HOST", "RELATIVE", "ACTION", "HOUR", "RELEASE", "ADMIN", "IGNORE", "RESULT", "AFTER", "IMMEDIATE", "RETURNS", "AGGREGATE", "INDICATOR", "ROLE", "ALIAS", "INITIALIZE", "ROLLUP", "ALLOCATE", "INITIALLY", "ROUTINE", "ARE", "INOUT", "ROW", "ARRAY", "INPUT", "ROWS", "ASENSITIVE", "INT", "SAVEPOINT", "ASSERTION", "INTEGER", "SCROLL", "ASYMMETRIC", "INTERSECTION", "SCOPE", "AT", "INTERVAL", "SEARCH", "ATOMIC", "ISOLATION", "SECOND", "BEFORE", "ITERATE", "SECTION", "BINARY", "LANGUAGE", "SENSITIVE", "BIT", "LARGE", "SEQUENCE", "BLOB", "LAST", "SESSION", "BOOLEAN", "LATERAL", "SETS", "BOTH", "LEADING", "SIMILAR", "BREADTH", "LESS", "SIZE", "CALL", "LEVEL", "SMALLINT", "CALLED", "LIKE_REGEX", "SPACE", "CARDINALITY", "LIMIT", "SPECIFIC", "CASCADED", "LN", "SPECIFICTYPE", "CAST", "LOCAL", "SQL", "CATALOG", "LOCALTIME", "SQLEXCEPTION", "CHAR", "LOCALTIMESTAMP", "SQLSTATE", "CHARACTER", "LOCATOR", "SQLWARNING", "CLASS", "MAP", "START", "CLOB", "MATCH", "STATE", "COLLATION", "MEMBER", "STATEMENT", "COLLECT", "METHOD", "STATIC", "COMPLETION", "MINUTE", "STDDEV_POP", "CONDITION", "MOD", "STDDEV_SAMP", "CONNECT", "MODIFIES", "STRUCTURE", "CONNECTION", "MODIFY", "SUBMULTISET", "CONSTRAINTS", "MODULE", "SUBSTRING_REGEX", "CONSTRUCTOR", "MONTH", "SYMMETRIC", "CORR", "MULTISET", "SYSTEM", "CORRESPONDING", "NAMES", "TEMPORARY", "COVAR_POP", "NATURAL", "TERMINATE", "COVAR_SAMP", "NCHAR", "THAN", "CUBE", "NCLOB", "TIME", "CUME_DIST", "NEW", "TIMESTAMP", "CURRENT_CATALOG", "NEXT", "TIMEZONE_HOUR", "CURRENT_DEFAULT_TRANSFORM_GROUP", "NO", "TIMEZONE_MINUTE", "CURRENT_PATH", "NONE", "TRAILING", "CURRENT_ROLE", "NORMALIZE", "TRANSLATE_REGEX", "CURRENT_SCHEMA", "NUMERIC", "TRANSLATION", "CURRENT_TRANSFORM_GROUP_FOR_TYPE", "OBJECT", "TREAT", "CYCLE", "OCCURRENCES_REGEX", "TRUE", "DATA", "OLD", "UESCAPE", "DATE", "ONLY", "UNDER", "DAY", "OPERATION", "UNKNOWN", "DEC", "ORDINALITY", "UNNEST", "DECIMAL", "OUT", "USAGE", "DEFERRABLE", "OVERLAY", "USING", "DEFERRED", "OUTPUT", "VALUE", "DEPTH", "PAD", "VAR_POP", "DEREF", "PARAMETER", "VAR_SAMP", "DESCRIBE", "PARAMETERS", "VARCHAR", "DESCRIPTOR", "PARTIAL", "VARIABLE", "DESTROY", "PARTITION", "WHENEVER", "DESTRUCTOR", "PATH", "WIDTH_BUCKET", "DETERMINISTIC", "POSTFIX", "WITHOUT", "DICTIONARY", "PREFIX", "WINDOW", "DIAGNOSTICS", "PREORDER", "WITHIN", "DISCONNECT", "PREPARE", "WORK", "DOMAIN", "PERCENT_RANK", "WRITE", "DYNAMIC", "PERCENTILE_CONT", "XMLAGG", "EACH", "PERCENTILE_DISC", "XMLATTRIBUTES", "ELEMENT", "POSITION_REGEX", "XMLBINARY", "END-EXEC", "PRESERVE", "XMLCAST", "EQUALS", "PRIOR", "XMLCOMMENT", "EVERY", "PRIVILEGES", "XMLCONCAT", "EXCEPTION", "RANGE", "XMLDOCUMENT", "FALSE", "READS", "XMLELEMENT", "FILTER", "REAL", "XMLEXISTS", "FIRST", "RECURSIVE", "XMLFOREST", "FLOAT", "REF", "XMLITERATE", "FOUND", "REFERENCING", "XMLNAMESPACES", "FREE", "REGR_AVGX", "XMLPARSE", "FULLTEXTTABLE", "REGR_AVGY", "XMLPI", "FUSION", "REGR_COUNT", "XMLQUERY", "GENERAL", "REGR_INTERCEPT", "XMLSERIALIZE", "GET", "REGR_R2", "XMLTABLE", "GLOBAL", "REGR_SLOPE", "XMLTEXT", "GO", "REGR_SXX", "XMLVALIDATE", "GROUPING", "REGR_SXY", "YEAR", "HOLD", "REGR_SYY", "ZONE", "NOCOUNT"}

	' Define the list of non-supported SQL keywords.  These words will be highlighted red.

	Private NotSupportedWords() As String = {"FutureKeywords", "ABSOLUTE", "HOST", "RELATIVE", "ACTION", "HOUR", "RELEASE", "ADMIN", "IGNORE", "RESULT", "AFTER", "IMMEDIATE", "RETURNS", "AGGREGATE", "INDICATOR", "ROLE", "ALIAS", "INITIALIZE", "ROLLUP", "ALLOCATE", "INITIALLY", "ROUTINE", "ARE", "INOUT", "ROW", "ARRAY", "INPUT", "ROWS", "ASENSITIVE", "SAVEPOINT", "ASSERTION", "INTEGER", "SCROLL", "ASYMMETRIC", "INTERSECTION", "SCOPE", "AT", "INTERVAL", "SEARCH", "ATOMIC", "ISOLATION", "SECOND", "BEFORE", "ITERATE", "SECTION", "BINARY", "LANGUAGE", "SENSITIVE", "LARGE", "SEQUENCE", "BLOB", "LAST", "SESSION", "BOOLEAN", "LATERAL", "SETS", "BOTH", "LEADING", "SIMILAR", "BREADTH", "LESS", "SIZE", "CALL", "LEVEL", "SMALLINT", "CALLED", "LIKE_REGEX", "SPACE", "CARDINALITY", "LIMIT", "SPECIFIC", "CASCADED", "LN", "SPECIFICTYPE", "CAST", "LOCAL", "SQL", "CATALOG", "LOCALTIME", "SQLEXCEPTION", "CHAR", "LOCALTIMESTAMP", "SQLSTATE", "CHARACTER", "LOCATOR", "SQLWARNING", "CLASS", "MAP", "START", "CLOB", "MATCH", "STATE", "COLLATION", "MEMBER", "STATEMENT", "COLLECT", "METHOD", "STATIC", "COMPLETION", "MINUTE", "STDDEV_POP", "CONDITION", "MOD", "STDDEV_SAMP", "CONNECT", "MODIFIES", "STRUCTURE", "CONNECTION", "MODIFY", "SUBMULTISET", "CONSTRAINTS", "MODULE", "SUBSTRING_REGEX", "CONSTRUCTOR", "MONTH", "SYMMETRIC", "CORR", "MULTISET", "SYSTEM", "CORRESPONDING", "NAMES", "TEMPORARY", "COVAR_POP", "NATURAL", "TERMINATE", "COVAR_SAMP", "THAN", "CUBE", "NCLOB", "TIME", "CUME_DIST", "NEW", "TIMESTAMP", "CURRENT_CATALOG", "NEXT", "TIMEZONE_HOUR", "CURRENT_DEFAULT_TRANSFORM_GROUP", "NO", "TIMEZONE_MINUTE", "CURRENT_PATH", "NONE", "TRAILING", "CURRENT_ROLE", "NORMALIZE", "TRANSLATE_REGEX", "CURRENT_SCHEMA", "NUMERIC", "TRANSLATION", "CURRENT_TRANSFORM_GROUP_FOR_TYPE", "OBJECT", "TREAT", "CYCLE", "OCCURRENCES_REGEX", "TRUE", "DATA", "OLD", "UESCAPE", "DATE", "ONLY", "UNDER", "DAY", "OPERATION", "UNKNOWN", "DEC", "ORDINALITY", "UNNEST", "DECIMAL", "OUT", "USAGE", "DEFERRABLE", "OVERLAY", "USING", "DEFERRED", "OUTPUT", "VALUE", "DEPTH", "PAD", "VAR_POP", "DEREF", "PARAMETER", "VAR_SAMP", "DESCRIBE", "PARAMETERS", "VARCHAR", "DESCRIPTOR", "PARTIAL", "VARIABLE", "DESTROY", "PARTITION", "WHENEVER", "DESTRUCTOR", "PATH", "WIDTH_BUCKET", "DETERMINISTIC", "POSTFIX", "WITHOUT", "DICTIONARY", "PREFIX", "WINDOW", "DIAGNOSTICS", "PREORDER", "WITHIN", "DISCONNECT", "PREPARE", "WORK", "DOMAIN", "PERCENT_RANK", "WRITE", "DYNAMIC", "PERCENTILE_CONT", "XMLAGG", "EACH", "PERCENTILE_DISC", "XMLATTRIBUTES", "ELEMENT", "POSITION_REGEX", "XMLBINARY", "END-EXEC", "PRESERVE", "XMLCAST", "EQUALS", "PRIOR", "XMLCOMMENT", "EVERY", "PRIVILEGES", "XMLCONCAT", "EXCEPTION", "RANGE", "XMLDOCUMENT", "FALSE", "READS", "XMLELEMENT", "FILTER", "REAL", "XMLEXISTS", "FIRST", "RECURSIVE", "XMLFOREST", "FLOAT", "REF", "XMLITERATE", "FOUND", "REFERENCING", "XMLNAMESPACES", "FREE", "REGR_AVGX", "XMLPARSE", "FULLTEXTTABLE", "REGR_AVGY", "XMLPI", "FUSION", "REGR_COUNT", "XMLQUERY", "GENERAL", "REGR_INTERCEPT", "XMLSERIALIZE", "GET", "REGR_R2", "XMLTABLE", "GLOBAL", "REGR_SLOPE", "XMLTEXT", "GO", "REGR_SXX", "XMLVALIDATE", "GROUPING", "REGR_SXY", "YEAR", "HOLD", "REGR_SYY", "ZONE"}

	'**********************************************************

	' The form is loaded.

	'**********************************************************
	Private Sub frmScriptEditor_Load(sender As Object, e As EventArgs) Handles Me.Load

		' Declare variables

		Dim zx As String

		' Put up a message

		frmMain.StatusLabel.Text = "Loading Form.  Please Wait..."
		System.Windows.Forms.Application.DoEvents()

		' Sort the reserved words list so it can be searched with a binary search.
		' If they have already been sorted by the DefaultText property, do not
		' re-sort them.

		If Not WordListsSorted Then
			Sort(ReservedWords)

			' Sort the non-supported words list

			Sort(NotSupportedWords)
		End If

		' Set the position of the main form and its size

		zx = GetSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "ScriptWindow", "Size")
		If zx <> "" Then
			Me.Top = Val(ParseString(zx))
			Me.Left = Val(ParseString(zx))
			Me.Height = Val(ParseString(zx))
			Me.Width = Val(zx)
		End If

		' Clear message

		frmMain.StatusLabel.Text = ""
		System.Windows.Forms.Application.DoEvents()

	End Sub
	'**********************************************************

	' The form is about to close.

	'**********************************************************
	Private Sub frmScriptEditor_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

		' Declare variables

		Dim xx As Integer

		' Ask to save the script before exiting

		If Not ScriptSaved Then
			xx = MsgBox("The current SQL script has not been saved.  Do you want to save it before closing the editor?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Close Script Editor")
			If xx = MsgBoxResult.Yes Then mnuSave_Click(mnuSave, New System.EventArgs)
			If xx = MsgBoxResult.Cancel Then e.Cancel = True
		End If

	End Sub
	'**********************************************************

	' The form is closed.

	'**********************************************************
	Private Sub frmScriptEditor_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

		' Declare variables

		Dim zx As String

		' Save the window size and position

		If Me.Left > 0 And Me.Top > 0 Then
			zx = Me.Top & "," & Me.Left & "," & Me.Height & "," & Me.Width
			SaveSetting("Sirius" & SRep(ProgramName, 1, " ", ""), "ScriptWindow", "Size", zx)
		End If

	End Sub


	'**********************************************************

	' Because of KeyPreview, this form will see any keystrokes
	' beforel the rich text box will see them.  So if an up
	' or down key is pressed, which will have changed the
	' insertion point before the RTB can look at the current
	' text, we will save that insertion point for it to look
	' at, if necessary.

	'**********************************************************
	Private Sub frmScriptEditor_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
		PreviousCaretPosition = Rtb1.SelectionStart
	End Sub

	'**********************************************************

	' The form is resized.

	'**********************************************************
	Private Sub frmScriptEditor_Resize(sender As Object, e As EventArgs) Handles Me.Resize

		' Make the Rich Text Box fill the entire window

		Dim Rect As Rectangle
		Rect = Me.ClientRectangle
		Rect.Height = Rect.Height - StatusStrip1.Height
		Rtb1.Bounds = Rect

	End Sub

	'**********************************************************

	' Key up and key down keys do not pass on to the KeyPress
	' event, yet we need to search for reserved words in any
	' text typed before such a key is pressed.  So we'll call
	' the keypress event handler ourselves, passing it null
	' code, which that routine is designed to recognize.

	'**********************************************************
	Private Sub Rtb1_KeyDown(sender As Object, e As KeyEventArgs) Handles Rtb1.KeyDown
		If e.KeyCode = Keys.Up Or e.KeyCode = Keys.Down Then RichTextBox1_KeyPress(Rtb1, New KeyPressEventArgs(Chr(0)))
	End Sub
	'**********************************************************

	' Mouse click events do not pass on to the KeyPress
	' event, yet we need to search for reserved words in any
	' text typed before the mouse was clicked.  So we'll call
	' the keypress event handler ourselves, passing it null
	' code, which that routine is designed to recognize.

	'**********************************************************
	Private Sub Rtb1_MouseClick(sender As Object, e As MouseEventArgs) Handles Rtb1.MouseClick

		' Get the current caret position (the point where the mouse was clicked.)
		' We'll need to move the caret back to that position when we're done.

		Dim xx As Integer = Rtb1.SelectionStart

		' Because moving the caret with a mouse didn't save the former caret position
		' via key preview in the main form, it will point to the last character pressed,
		' rather than the space after it.  So we need to increment it.

		PreviousCaretPosition += 1

		' Make sure no text has been selected, and the mouse was just moved..

		If Rtb1.SelectedText = "" Then

			' Call the key press routine now to check if a reserved word was typed.

			RichTextBox1_KeyPress(Rtb1, New KeyPressEventArgs(Chr(0)))

			' Restore the insertion point to where the mouse was clicked.

			Rtb1.SelectionStart = xx
		End If

	End Sub


	'**********************************************************

	' A key is pressed in the RichTextBox contol.

	'**********************************************************
	Private Sub RichTextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Rtb1.KeyPress

		' Declare variables

		Dim ii As Integer
		Dim WordLastChar As Integer
		Dim SaveCaretPos As Integer
		Dim xx As String
		Dim Word As String = ""

		' Get the position of the caret, which is our starting point for searching
		' for a word.  However, that position may have changed before this routine was
		' called, by a mouse click or an up/down arrow key.  If that happens, the
		' cursor position is captured by the form, using KeyPreview, before it can reach
		' the RichTextBox, and the event handler has remembered the caret position
		' before it changed.  If that happens, we'll change the value of ii to the
		' former caret position in a couple of code lines farther down.

		ii = Rtb1.SelectionStart ' This point will move backward as we search for the start of the word

		' Ignore any key press that's not a printable character, since every single key on the keyboard
		' will trigger this event.

		If Not IsPrintableChar(e.KeyChar) And Not Asc(e.KeyChar) = 13 Then Exit Sub

		' If we got a null character, which will be sent by the form because of KeyPreview,
		' then an up or down arrow was pressed and the current insertion point changed
		' before this routine was called, so use the position from before it changed
		' as the starting point for our word search.

		If Asc(e.KeyChar) = 0 Then ii = PreviousCaretPosition

		' Indicate that the file has changed since it was saved.

		ScriptSaved = False

		' Get the location of the caret in the text. We must first make sure
		' no text is selected, because that is the only time the SelectionStart
		' property indicates the caret position.

		Rtb1.SelectedText = ""
		WordLastChar = ii ' This stays the same to indicate the end of the word
		SaveCaretPos = ii

		' if the character pressed was a space, extract the word from the text.

		If e.KeyChar = " " Or Asc(e.KeyChar) = 0 Then

			' If the character pressed was ENTER (Ascii value 13), then begin searching one 
			' character earlier, since that will be the last character of the line/word.

			If Asc(e.KeyChar) = 13 Then
				If ii > 1 Then ii -= 1
				WordLastChar -= 1 ' When we've terminated a word with CrLf, we have to mark the end of the word one character previous
			End If

			' Get the character prior to the character just typed, to be tested.

			If ii > 0 Then xx = UCase(Mid(Rtb1.Text, ii, 1)) Else xx = "" ' We may be at the beginning of the text.

			' Keep searching backwards as long as the character is a letter or number, and
			' we haven't yet reached the start of the line.

			Do While ii > 0 And ((xx >= "A" And xx <= "Z") Or (xx >= "0" And xx <= "9"))
				ii = ii - 1
				If ii > 0 Then xx = UCase(Mid(Rtb1.Text, ii, 1))
			Loop

			' Now that we have found a delimiter for the current word,
			' extract the word. We not HAVE a delimiter if we've reached
			' the beginning of the file, so we must extract the word
			' differently.

			If WordLastChar > 1 Then Word = Mid(Rtb1.Text, ii + 1, WordLastChar - ii)

			' See if it's a reserved word.

			If IsReservedWord(Word) Then

				' Select the word

				Rtb1.Select(ii, WordLastChar - ii)

				' See if it's an unsupported word

				If IsNotSupported(Word) Then

					' Color non-supported words red

					Rtb1.SelectionColor = Color.Red

				Else

					' Color supported words blue

					Rtb1.SelectionColor = Color.Blue

				End If

				' Capitalize the word, whether or not supported.

				Rtb1.SelectedText = UCase(Rtb1.SelectedText)

			End If


		End If

		' Always de-select any text, so the SelectionStart property will
		' return the caret position.

		Rtb1.SelectedText = ""

		' Restore the caret position

		Rtb1.SelectionStart = SaveCaretPos

		' Always reset the selection color to black

		Rtb1.SelectionColor = Color.Black

	End Sub

	'**********************************************************

	' The Open Script menu option is selected.

	'**********************************************************
	Private Sub mnuOpenSQL_Click(sender As Object, e As EventArgs) Handles mnuOpenSQL.Click

		' Declare variables

		Dim zx0 As String

		' Get the current directory (the common dialog box may
		' change it).

		zx0 = FileSystem.CurDir

		' Use the common dialog open box to get the file nafrmMain.

		frmMain.OpenFileDialog1.FileName = ""
		frmMain.OpenFileDialog1.Filter = "SQL Script Files (*.SQL)|*.SQL|All Files (*.*)|*.*"
		frmMain.OpenFileDialog1.FilterIndex = 1
		frmMain.OpenFileDialog1.DefaultExt = "SQl"
		frmMain.OpenFileDialog1.Title = "Open SQL Script File"
		If frmMain.OpenFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then

			' Save the file name

			ScriptFileName = frmMain.OpenFileDialog1.FileName

			' Restore the current directory

			FileSystem.ChDir(zx0)

			' Open the specified script file

			zx0 = My.Computer.FileSystem.ReadAllText(ScriptFileName)

			' Clear the text box and put the text into the rich text box.

			Rtb1.Clear()
			Rtb1.Text = zx0

			' Now process the text, highlighting all reserved SQL words.

			HighlightReservedWords()

			' Indicate the script hasn't changed.

			ScriptSaved = True

			' Enable the close SQL menu option

			mnuCloseSQL.Enabled = True


		End If
	End Sub

	'**********************************************************

	' The Execute Script menu option is selected.

	'**********************************************************
	Private Sub mnuExecute_Click(sender As Object, e As EventArgs) Handles mnuExecute.Click

		' Declare variables.

		Dim Cmd As New SqlCommand
		Dim Transaction As SqlTransaction

		' Create a transaction.

		Transaction = DB.BeginTransaction

		' Create an SQL command from the contents of the editor.

		Cmd.Connection = DB
		Cmd.CommandType = CommandType.Text
		Cmd.CommandText = Rtb1.Text
		Cmd.Transaction = Transaction

		' Execute the query and watch for errors.

		Try
			Cmd.ExecuteNonQuery()
			StatusLabel1.Text = "Script executed successfully"

			' Indicate that the script has been saved.

			ScriptSaved = True

			' Commit the transaction

			Transaction.Commit()

			' Refresh the object list.

			frmObjectList.RefreshObjectList()
		Catch ex As Exception


			StatusLabel1.Text = "Script failed"
			MsgBox(ex.Message, MsgBoxStyle.Information, "Upload SQL Script")
			Transaction.Rollback()
		End Try

		' Enable the timer.  This will cause any message to be cleared automatically after 5 seconds.

		Timer1.Enabled = True

	End Sub
	'**********************************************************

	' The new SQL script menu option is selected.

	'**********************************************************
	Private Sub mnuNewSQL_Click(sender As Object, e As EventArgs) Handles mnuNewSQL.Click

		' Prepare for a new script.

		Rtb1.Clear()
		ScriptFileName = ""
		ScriptSaved = True
		mnuCloseSQL.Enabled = True

	End Sub
	'**********************************************************

	' The save menu option is selected

	'**********************************************************
	Private Sub mnuSave_Click(sender As Object, e As EventArgs) Handles mnuSave.Click

		' If this is a new file, and there is no file name, perform a "save as" instead.

		If ScriptFileName = "" Then
			mnuSaveAs_Click(mnuSaveAs, New System.EventArgs)
		Else
			My.Computer.FileSystem.WriteAllText(ScriptFileName, Rtb1.Text, False)
			ScriptSaved = True
		End If


	End Sub
	'**********************************************************

	' The save as menu option is selected.

	'**********************************************************
	Private Sub mnuSaveAs_Click(sender As Object, e As EventArgs) Handles mnuSaveAs.Click

		' Declare variables

		Dim zx As String = "New SQL Script.sql"

		' If there is already a file oopen, use it as the default save name.

		If ScriptFileName <> "" Then zx = ScriptFileName

		' Use the common dialog open box to get the file nafrmMain.

		frmMain.SaveFileDialog1.FileName = zx
		frmMain.SaveFileDialog1.AddExtension = True
		frmMain.SaveFileDialog1.OverwritePrompt = True
		frmMain.SaveFileDialog1.Filter = "SQL Script Files (*.SQL)|*.SQL|All Files (*.*)|*.*"
		frmMain.SaveFileDialog1.FilterIndex = 1
		frmMain.SaveFileDialog1.DefaultExt = "SQl"
		frmMain.SaveFileDialog1.Title = "Save SQL Script File"
		If frmMain.SaveFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then

			' Save the script file name

			ScriptFileName = frmMain.SaveFileDialog1.FileName

			' Save the script

			My.Computer.FileSystem.WriteAllText(ScriptFileName, Rtb1.Text, False)

			' Indicate the script has been saved.

			ScriptSaved = True

		End If
	End Sub

	'**********************************************************

	' Sub to close an SQL file

	'**********************************************************
	Private Sub mnuCloseSQL_Click(sender As Object, e As EventArgs) Handles mnuCloseSQL.Click
		mnuCloseSQL.Enabled = False
		Rtb1.Clear()
	End Sub
	'**********************************************************

	' A new zoom level is selected.

	'**********************************************************
	Private Sub mnuZoom_Click(sender As Object, e As EventArgs) Handles mnu25.Click, mnu50.Click, mnu75.Click, mnu100.Click, mnu125.Click, mnu150.Click, mnu175.Click, mnu200.Click

		' Declare variables

		Dim Zoom As Single
		Dim zx As String

		' Get the name of the menu item.

		zx = sender.Name

		' Get the zoom factor

		Zoom = Val(Mid(zx, 4)) / 100

		' Set the RTF box zoom factor

		Rtb1.ZoomFactor = Zoom


	End Sub
	'**********************************************************

	' The close menu item is selected.

	'**********************************************************
	Private Sub mnuClose_Click(sender As Object, e As EventArgs) Handles mnuClose.Click
		Me.Close()
	End Sub

	'**********************************************************************

	' The timer has clicked.  Erase any current status message.

	'**********************************************************************
	Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

		StatusLabel1.Text = ""
		Timer1.Enabled = False

	End Sub   '**********************************************************

	' Sub to highlight all reserved words in an SQL script
	' when it is loaded from a file.

	'**********************************************************
	Private Sub HighlightReservedWords()

		' Declare variaables

		Dim Bracketed As Boolean
		Dim ii As Integer
		Dim WordFirstChar As Integer
		Dim WordLastChar As Integer
		Dim Word As String = ""
		Dim zx As String

		' Iterate through the text of the RTF box.

		For ii = 1 To Rtb1.TextLength

			' Upper-case the current character for comparison.

			zx = UCase(Mid(Rtb1.Text, ii, 1))
			Word = "" ' We have no word yet

			' If we encounter a bracketed word, we will not change it.

			If zx = "[" Then Bracketed = True
			If zx = "]" Then Bracketed = False

			' Determine where a word starts and stops.  A word must begin with a letter, and
			' end with a letter or number.  Any other character delimits the word.

			If WordFirstChar = 0 AndAlso (zx >= "A" And zx <= "Z") And Not Bracketed Then WordFirstChar = ii
			If WordFirstChar > 0 AndAlso ((zx < "A" Or zx > "Z") And Not (zx >= "0" And zx <= "9")) And zx <> "_" And Not Bracketed Then WordLastChar = ii - 1

			' When we find the start and stop of the word, extract it from the text.

			If WordFirstChar > 0 And WordLastChar > 0 Then Word = Mid(Rtb1.Text, WordFirstChar, WordLastChar - WordFirstChar + 1)

			' Once we have a word, see if it's a reserved word.

			If Word <> "" Then
				If IsReservedWord(Word) Then

					' If the word is reserved but not supported, color it red. Otherwise,
					' color it blue, but capitalize it in either case.

					Rtb1.Select(WordFirstChar - 1, WordLastChar - WordFirstChar + 1)
					If IsNotSupported(Word) Then
						Rtb1.SelectionColor = Color.Red
					Else
						Rtb1.SelectionColor = Color.Blue
					End If
					Rtb1.SelectedText = UCase(Rtb1.SelectedText)
				End If

				' Reset the pointers for finding the next word.

				WordFirstChar = 0
				WordLastChar = 0
				Word = ""
			End If
		Next ii

		' Always reselect the selection color to black when we're done highlighting words.

		Rtb1.SelectionColor = Color.Black

	End Sub
	'**********************************************************

	' Function to indicate if a word is a reserved SQL
	' keyword or not.

	'**********************************************************
	Private Function IsReservedWord(ByVal Word As String) As Boolean

		' Declare variables

		Dim ii As Integer

		' Upper case the supplied word for comparison

		Word = UCase(Word)

		' Look for the word in the reserved words list

		ii = Search(Word, ReservedWords)

		Return ii >= 0


	End Function
	'**********************************************************

	' Function to indicate if a reserved SQL keyword is
	' supported or not.

	'**********************************************************
	Private Function IsNotSupported(ByVal Word As String) As Boolean

		' Declare variables

		Dim ii As Integer

		' Upper case the supplied word for comparison

		Word = UCase(Word)

		' Look for the word in the reserved words list

		ii = Search(Word, NotSupportedWords)

		Return ii > 0

	End Function
	'**********************************************************

	' Function to indicate if a key that was pressed is that
	' of a printable character and not a function key of
	' some sort.

	'**********************************************************
	Private Function IsPrintableChar(C As String) As Boolean

		If InStr(Chr(0) & " !@#$%^&*()_+1234567890-={}|[]\:"";'<>?,./ABCDEFGHIJKLMNOPQRSTUVWXYZ", C, CompareMethod.Text) > 0 Then Return True Else Return False

	End Function

	'**********************************************************

	' The write-only property "DefaultText"

	'**********************************************************
	Public WriteOnly Property DefaultText As String
		Set(value As String)

			' Sort the reserved words list

			Sort(ReservedWords)

			' Sort the non-supported words list

			Sort(NotSupportedWords)

			' Indicate we have sorted the word lists.

			WordListsSorted = True

			' Set the default text value and process it to highlight
			' the reserved words.

			Rtb1.Text = value
			HighlightReservedWords()

			' Indicate the text hasn't changed.

			ScriptSaved = True

		End Set
	End Property
End Class