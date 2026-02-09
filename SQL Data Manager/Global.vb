Option Strict Off
Option Explicit On
Module GlobalModule
	'**********************************************************
	' Public Constants for Visual Basic Programs
	' GLOBAL.VB
	' Written: April 2007
	' Updated: July 2025
	' Programmer: Aaron Scott

	' Copyright (C) 1999-2025 Sirius Software
	' All Rights Reserved

	' Required modules: none
	'**********************************************************


	''''''''''''''''''''''''''''
	' Visual Basic Public constant file. This file can be loaded
	' into a code module.
	'
	' Some constants are commented out because they have
	' duplicates (e.g., NONE appears several places).
	'
	' If you are updating a Visual Basic application written with
	' an older version, you should replace your Public constants
	' with the constants in this file.
	'
	''''''''''''''''''''''''''''


	Public Structure NETRESOURCE
		Dim Scope As Integer
		Dim Type As Integer
		Dim DisplayType As Integer
		Dim Usage As Integer
		Dim LocalName As String
		Dim RemoteName As String
		Dim Comment As String
		Dim Provider As String
	End Structure

	Public Structure NETRESOURCELONG
		Dim Scope As Integer
		Dim Type As Integer
		Dim DisplayType As Integer
		Dim Usage As Integer
		Dim LocalName As Integer
		Dim RemoteName As Integer
		Dim Comment As Integer
		Dim Provider As Integer
	End Structure

	' Declare windows API functions which we will use
	Declare Function GetProfileString Lib "Kernel32.dll" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
	Declare Function WriteProfileString Lib "Kernel32.dll" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Integer
	Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
	Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Integer, ByVal nBkMode As Integer) As Integer
	Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
	Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
	Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
	Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
	Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData As Integer) As Integer
	Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
	Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Integer, ByVal nIndex As Integer) As Integer
	Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Long) As Integer
	Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByRef lpInitData As Integer) As Integer
	Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Integer) As Integer
	Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
	Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Integer, ByVal dwFlags As Integer) As Integer
	Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Integer, ByVal dwType As Integer, ByVal dwUsage As Integer, ByRef lpNetResource As NETRESOURCE, ByRef lphEnum As Integer) As Integer
	Declare Function WNetOpenEnumRoot Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Integer, ByVal dwType As Integer, ByVal dwUsage As Integer, ByVal lpNetResource As Integer, ByRef lphEnum As Integer) As Integer
	Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Integer, ByRef lpcCount As Integer, ByRef lpBuffer As Byte, ByRef lpBufferSize As Integer) As Integer
	Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Integer) As Integer
	Declare Function GetTempFileName Lib "Kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Integer
	Declare Function GetLongPathName Lib "Kernel32" Alias "GetLongPathNameA" (ByVal lpShortPath As String, ByVal lpLongPath As String, ByVal nSize As Integer) As Integer



	Public Const MF_BYPOSITION As Integer = &H400

	' Define types used by system function calls

	Public Structure OVERLAPPED
		Dim Internal As Integer
		Dim InternalHigh As Integer
		Dim offset As Integer
		Dim OffsetHigh As Integer
		Dim hEvent As Integer
	End Structure

	' Declare user-defined types which we will
	' need to access some registry functions.

	Public Structure FILETIME
		Dim dwLowDateTime As Integer
		Dim dwHighDateTime As Integer
	End Structure

	Public Structure SECURITY_ATTRIBUTES
		Dim nLength As Integer
		Dim lpSecurityDescriptor As Integer
		Dim bInheritHandle As Integer
	End Structure

	' Declare the registry functions we will use

	' Note: the RegCreateKeyEx function must be redefined for
	' Win95(as opposed to NT).  We do not pass the SECURITY_ATTRIBUTES type, so
	' we redefine the call so we can pass a "0" which will appear
	' as a "Null" when passed byval.

	Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByVal Reserved As Integer, ByVal lpClass As String, ByVal dwOptions As Integer, ByVal samDesired As Integer, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Integer, ByRef lpdwDisposition As Integer) As Integer
	'Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
	Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByRef phkResult As Integer) As Integer
	Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As String, ByRef lpcbData As Integer) As Integer
	Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal Reserved As Integer, ByVal dwType As Integer, ByVal lpData As String, ByVal cbData As Integer) As Integer
	Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByVal dwType As Integer, ByVal lpData As String, ByVal cbData As Integer) As Integer
	Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
	Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Integer, ByVal lpSubKey As String) As Integer
	Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Integer, ByVal lpValueName As String) As Integer
	Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Integer, ByVal dwIndex As Integer, ByVal lpName As String, ByRef lpcbName As Integer, ByVal lpReserved As Integer, ByVal lpClass As String, ByRef lpcbClass As Integer, ByRef lpftLastWriteTime As FILETIME) As Integer
	Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Integer, ByVal dwIndex As Integer, ByVal lpValueName As String, ByRef lpcbValueName As Integer, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Byte, ByRef lpcbData As Integer) As Integer

	' Declare the values of predefined registry keys

	Public Const HKEY_CLASSES_ROOT As Integer = &H80000000
	Public Const HKEY_CURRENT_CONFIG As Integer = &H80000005
	Public Const HKEY_CURRENT_USER As Integer = &H80000001
	Public Const HKEY_DYN_DATA As Integer = &H80000006
	Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Public Const HKEY_PERFORMANCE_DATA As Integer = &H80000004
	Public Const HKEY_USERS As Integer = &H80000003

	Public Const READ_CONTROL As Integer = &H20000
	Public Const READ_WRITE As Short = 2
	Public Const STANDARD_RIGHTS_ALL As Integer = &H1F0000
	Public Const STANDARD_RIGHTS_EXECUTE As Integer = (READ_CONTROL)
	Public Const STANDARD_RIGHTS_READ As Integer = (READ_CONTROL)
	Public Const STANDARD_RIGHTS_REQUIRED As Integer = &HF0000
	Public Const STANDARD_RIGHTS_WRITE As Integer = (READ_CONTROL)

	Public Const KEY_QUERY_VALUE As Short = &H1S
	Public Const KEY_SET_VALUE As Short = &H2S
	Public Const KEY_CREATE_SUB_KEY As Short = &H4S
	Public Const KEY_CREATE_LINK As Short = &H20S
	Public Const KEY_ENUMERATE_SUB_KEYS As Short = &H8S
	Public Const KEY_EVENT As Short = &H1S
	Public Const KEY_NOTIFY As Short = &H10S
	Public Const SYNCHRONIZE As Integer = &H100000
	Public Const KEY_ALL_ACCESS As Integer = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
	Public Const KEY_READ As Integer = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
	Public Const KEY_EXECUTE As Integer = ((KEY_READ) And (Not SYNCHRONIZE))
	Public Const KEY_WRITE As Integer = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

	Public Const REG_BINARY As Short = 3
	Public Const REG_DWORD As Short = 4
	Public Const REG_DWORD_BIG_ENDIAN As Short = 5
	Public Const REG_DWORD_LITTLE_ENDIAN As Short = 4
	Public Const REG_EXPAND_SZ As Short = 2
	Public Const REG_LINK As Short = 6
	Public Const REG_MULTI_SZ As Short = 7
	Public Const REG_NONE As Short = 0
	Public Const REG_SZ As Short = 1

	' Create a public example of a SECURITY_ATTRIBUTE structure

	Public SecAtt As SECURITY_ATTRIBUTES

	' Declare other functions we will use

	Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Integer) As Integer

	' Miscellaneous Public variables and constants

	Public CancelOperation As Boolean
	Public UnhandledExceptionTriggered As Boolean
	Public ProgramPath As String = ""
	Public Datapath As String = ""
	Public BackupPath As String = ""
	Public Databasename As String = ""
	Public LocalDatabase As String = ""
	Public ProgramName As String = ""
	Public Version As String = ""
	Public DBVersion As String = ""
	Public LicenseInfo As String = ""
	Public DbOpen As Boolean
	Public ProgramInUse As Boolean
	Public LoginName As String = ""
	Public UserName As String = ""
	Public UserIsSupervisor As Boolean
	Public LeftMargin As Single
	Public RightMargin As Single

	' Define the checksum values for the keys of each of the four basic programs.
	Public Const HA_Checksum = 378
	Public Const SEF_Checksum = 419
	Public Const SC_Checksum = 297
	Public Const DM_Checksum = 513

	' Define the licensing status conditions
	Public Enum LicensingStatus
		Expired
		Evaluation
		Licensed
	End Enum

	Public CustomColor1 As Color = Color.FromArgb(255, 128, 100, 56)
	Public CustomColor2 As Color = Color.FromArgb(255, 182, 189, 255)
	Public CustomColor3 As Color = Color.FromArgb(255, 228, 225, 112)
	Public CustomColor4 As Color = Color.FromArgb(255, 196, 219, 181)
	Public UserDefinedColor As Color = SystemColors.Window

	Public Dependencies As New Collection

	Public Const ADDRECORD As Short = 0
	Public Const MODIFYRECORD As Short = 1
	Public Const CHANGEID As Short = 2
	Public Const NOMATCH = -1
	Public Const e As Double = 2.71828182846
	Public Const pi As Double = 3.14159

	Public Class Dependency
		Public Name As String
		Public ObjectType As String
		Public FileToCopy As String
		Public Sub New(NameValue As String, Type As String, Optional FileName As String = "")
			Name = NameValue
			ObjectType = Type
			FileToCopy = FileName
		End Sub
	End Class
End Module