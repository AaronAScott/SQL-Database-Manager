Imports vb = Microsoft.VisualBasic
Imports System
Imports System.Security.Permissions
Imports Microsoft.Win32

Public Class frmRegistryEditor
    '**************************************************************
    ' Registry Editor for Visual Basic SQL programs
    ' REGISTRYEDITOR.VB
    ' Written: March 2018
    ' Programmer: Aaron Scott
    ' Copyright (C) 2003-2018 Sirius Software All Rights Reserved
    '**************************************************************



    '**************************************************************

    ' The form is loaded.

    '**************************************************************
    Private Sub frmRegistryEditor_Load(sender As Object, e As EventArgs) Handles Me.Load

        ' Declare variables

        Dim ii As Integer
        Dim jj As Integer
        Dim kk As Integer
        Dim Keys() As String
        Dim Subkeys() As String
        Dim Values() As String
        Dim n As TreeNode
        Dim c As Color
        Dim r1 As Microsoft.Win32.RegistryKey = Registry.CurrentUser.OpenSubKey("Software")
        Dim r2 As Microsoft.Win32.RegistryKey = r1.OpenSubKey("VB and VBA Program Settings")
        Dim r3 As Microsoft.Win32.RegistryKey
        Dim r4 As Microsoft.Win32.RegistryKey

        ' Clear the treeview control

        TreeView1.Nodes.Clear()

        ' Get all "Sirius" keys, regardless of which program they belong to.

        Keys = r2.GetSubKeyNames
        For ii = 0 To UBound(Keys)
            If vb.Left(Keys(ii), 6) = "Sirius" Then
                n = TreeView1.Nodes.Add(Keys(ii), Keys(ii))

                ' If the current node is for this program, it will be displayed in black.  If for
                ' some other Sirius program, it will be displayed in gray.

                If n.Name = "Sirius" & SRep(ProgramName, 1, " ", "") Then c = Color.Black Else c = Color.Gray
                n.ForeColor = c

                ' Get all the subkeys for the current node.

                r3 = r2.OpenSubKey(Keys(ii))
                Subkeys = r3.GetSubKeyNames
                For jj = 0 To UBound(Subkeys)
                    n = TreeView1.Nodes(Keys(ii)).Nodes.Add(Subkeys(jj), Subkeys(jj))
                    n.ForeColor = c

                    ' Get all the values for the various settings under this subkey.

                    r4 = r3.OpenSubKey(Subkeys(jj))
                    Values = r4.GetValueNames
                    For kk = 0 To UBound(Values)
                        n = TreeView1.Nodes(Keys(ii)).Nodes(Subkeys(jj)).Nodes.Add(Values(kk), Values(kk))
                        n.ForeColor = c
                    Next kk
                    r4.Dispose()
                Next jj
                r3.Dispose()
            End If
        Next ii

        ' Final cleanup

        r2.Dispose()
        r1.Dispose()


    End Sub

    '**************************************************************

    ' A node is clicked on the Treeview control.

    '**************************************************************
    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick

        ' Declare variables

        Dim r1 As Microsoft.Win32.RegistryKey = Registry.CurrentUser.OpenSubKey("Software")
        Dim r2 As Microsoft.Win32.RegistryKey = r1.OpenSubKey("VB and VBA Program Settings")
        Dim r3 As Microsoft.Win32.RegistryKey
        Dim r4 As Microsoft.Win32.RegistryKey

        ' Determine what level the clicked node is.

        Select Case e.Node.Level
            Case 0 ' The main node for a Sirius program
                If e.Node.Name = "Sirius" & SRep(ProgramName, 1, " ", "") Then btnReset.Enabled = True Else btnReset.Enabled = False
            Case 2 ' A group of settings for a program
                If e.Node.Parent.Parent.Name = "Sirius" & SRep(ProgramName, 1, " ", "") Then btnReset.Enabled = True Else btnReset.Enabled = False
                r3 = r2.OpenSubKey(e.Node.Parent.Parent.Name)
                r4 = r3.OpenSubKey(e.Node.Parent.Name)
                TextBox1.Text = r4.GetValue(e.Node.Name)
                r4.Dispose()
                r3.Dispose()
            Case Else ' All other nodes and values
                btnReset.Enabled = False
                TextBox1.Text = ""
        End Select

        ' Final cleanup

        r2.Dispose()
        r1.Dispose()

    End Sub

    '**************************************************************

    ' The Close button is clicked.

    '**************************************************************
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    '**************************************************************

    ' The Reset button is clicked.

    '**************************************************************
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click

        ' Declare variables

        Dim r1 As Microsoft.Win32.RegistryKey = Registry.CurrentUser.OpenSubKey("Software")
        Dim r2 As Microsoft.Win32.RegistryKey = r1.OpenSubKey("VB and VBA Program Settings", True)
        Dim r3 As Microsoft.Win32.RegistryKey
        Dim r4 As Microsoft.Win32.RegistryKey

        ' Determine the level of the clicked node.

        Select Case TreeView1.SelectedNode.Level

            ' A program node is reset.  This removes all settings for the program, restoring everything
            ' to its default value.

            Case 0 ' Program node
                If MsgBox("Resetting this setting will result in restoring ALL program settings to their default value(s)." & vbCrLf & "Are you sure you want to reset this setting?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Reset Registry Setting") = MsgBoxResult.Yes Then
                    r2.DeleteSubKeyTree(TreeView1.SelectedNode.Name)
                    TreeView1.Nodes(TreeView1.SelectedNode.Name).Remove()
                End If

                ' A specific setting is reset.  This removes only the selected setting, restoring it to its default value.

            Case 2 ' Setting node
                r3 = r2.OpenSubKey(TreeView1.SelectedNode.Parent.Parent.Name)
                r4 = r3.OpenSubKey(TreeView1.SelectedNode.Parent.Name, True)
                r4.DeleteValue(TreeView1.SelectedNode.Name)
                TreeView1.SelectedNode.Remove()
                r4.Dispose()
                r3.Dispose()

            Case Else ' All other nodes
                TextBox1.Text = ""
        End Select

        ' After a reset, disable the reset button

        btnReset.Enabled = False

        ' Final cleanup.

        r2.Dispose()
        r1.Dispose()
    End Sub
End Class