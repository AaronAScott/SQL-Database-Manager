<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
		Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
		Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuNewDatabase = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuOpenDatabase = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuCloseDatabase = New System.Windows.Forms.ToolStripMenuItem()
		Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripSeparator()
		Me.mnuTableDesigner = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuScriptEditor = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuRelationships = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuGenerateCommands = New System.Windows.Forms.ToolStripMenuItem()
		Me.MRUSeparator = New System.Windows.Forms.ToolStripSeparator()
		Me.MRU1 = New System.Windows.Forms.ToolStripMenuItem()
		Me.MRU2 = New System.Windows.Forms.ToolStripMenuItem()
		Me.MRU3 = New System.Windows.Forms.ToolStripMenuItem()
		Me.MRU4 = New System.Windows.Forms.ToolStripMenuItem()
		Me.MRU5 = New System.Windows.Forms.ToolStripMenuItem()
		Me.MRU6 = New System.Windows.Forms.ToolStripMenuItem()
		Me.MRU7 = New System.Windows.Forms.ToolStripMenuItem()
		Me.MRU8 = New System.Windows.Forms.ToolStripMenuItem()
		Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
		Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuUtilities = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuCompactDB = New System.Windows.Forms.ToolStripMenuItem()
		Me.ToolStripMenuItem4 = New System.Windows.Forms.ToolStripSeparator()
		Me.mnuOrganizeDB = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuResetSettings = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuTheme = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuWindow = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuNoArrange = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuTile = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuCascade = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuAbout = New System.Windows.Forms.ToolStripMenuItem()
		Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
		Me.StatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
		Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
		Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
		Me.FontDialog1 = New System.Windows.Forms.FontDialog()
		Me.mnuViewReadMe = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuViewLicense = New System.Windows.Forms.ToolStripMenuItem()
		Me.MenuStrip1.SuspendLayout()
		Me.StatusStrip1.SuspendLayout()
		Me.SuspendLayout()
		'
		'MenuStrip1
		'
		Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuUtilities, Me.mnuWindow, Me.mnuHelp})
		Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
		Me.MenuStrip1.Name = "MenuStrip1"
		Me.MenuStrip1.Size = New System.Drawing.Size(777, 24)
		Me.MenuStrip1.TabIndex = 1
		Me.MenuStrip1.Tag = "NoColor"
		Me.MenuStrip1.Text = "MenuStrip1"
		'
		'mnuFile
		'
		Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewDatabase, Me.mnuOpenDatabase, Me.mnuCloseDatabase, Me.ToolStripMenuItem2, Me.mnuTableDesigner, Me.mnuScriptEditor, Me.mnuRelationships, Me.mnuGenerateCommands, Me.MRUSeparator, Me.MRU1, Me.MRU2, Me.MRU3, Me.MRU4, Me.MRU5, Me.MRU6, Me.MRU7, Me.MRU8, Me.ToolStripMenuItem1, Me.mnuExit})
		Me.mnuFile.Name = "mnuFile"
		Me.mnuFile.Size = New System.Drawing.Size(37, 20)
		Me.mnuFile.Text = "&File"
		'
		'mnuNewDatabase
		'
		Me.mnuNewDatabase.Name = "mnuNewDatabase"
		Me.mnuNewDatabase.Size = New System.Drawing.Size(271, 22)
		Me.mnuNewDatabase.Text = "&New SQL Database"
		'
		'mnuOpenDatabase
		'
		Me.mnuOpenDatabase.Name = "mnuOpenDatabase"
		Me.mnuOpenDatabase.Size = New System.Drawing.Size(271, 22)
		Me.mnuOpenDatabase.Text = "&Open SQL Database..."
		'
		'mnuCloseDatabase
		'
		Me.mnuCloseDatabase.Enabled = False
		Me.mnuCloseDatabase.Name = "mnuCloseDatabase"
		Me.mnuCloseDatabase.Size = New System.Drawing.Size(271, 22)
		Me.mnuCloseDatabase.Text = "&Close Database"
		'
		'ToolStripMenuItem2
		'
		Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
		Me.ToolStripMenuItem2.Size = New System.Drawing.Size(268, 6)
		'
		'mnuTableDesigner
		'
		Me.mnuTableDesigner.Enabled = False
		Me.mnuTableDesigner.Name = "mnuTableDesigner"
		Me.mnuTableDesigner.Size = New System.Drawing.Size(271, 22)
		Me.mnuTableDesigner.Text = "Table &Designer"
		'
		'mnuScriptEditor
		'
		Me.mnuScriptEditor.Enabled = False
		Me.mnuScriptEditor.Name = "mnuScriptEditor"
		Me.mnuScriptEditor.Size = New System.Drawing.Size(271, 22)
		Me.mnuScriptEditor.Text = "SQL Script &Editor"
		'
		'mnuRelationships
		'
		Me.mnuRelationships.Enabled = False
		Me.mnuRelationships.Name = "mnuRelationships"
		Me.mnuRelationships.Size = New System.Drawing.Size(271, 22)
		Me.mnuRelationships.Text = "Table &Relationships"
		'
		'mnuGenerateCommands
		'
		Me.mnuGenerateCommands.Enabled = False
		Me.mnuGenerateCommands.Name = "mnuGenerateCommands"
		Me.mnuGenerateCommands.Size = New System.Drawing.Size(271, 22)
		Me.mnuGenerateCommands.Text = "&Generate Data Access Code Module..."
		'
		'MRUSeparator
		'
		Me.MRUSeparator.Name = "MRUSeparator"
		Me.MRUSeparator.Size = New System.Drawing.Size(268, 6)
		Me.MRUSeparator.Visible = False
		'
		'MRU1
		'
		Me.MRU1.Name = "MRU1"
		Me.MRU1.Size = New System.Drawing.Size(271, 22)
		Me.MRU1.Text = "&1"
		Me.MRU1.Visible = False
		'
		'MRU2
		'
		Me.MRU2.Name = "MRU2"
		Me.MRU2.Size = New System.Drawing.Size(271, 22)
		Me.MRU2.Text = "&2"
		Me.MRU2.Visible = False
		'
		'MRU3
		'
		Me.MRU3.Name = "MRU3"
		Me.MRU3.Size = New System.Drawing.Size(271, 22)
		Me.MRU3.Text = "&3"
		Me.MRU3.Visible = False
		'
		'MRU4
		'
		Me.MRU4.Name = "MRU4"
		Me.MRU4.Size = New System.Drawing.Size(271, 22)
		Me.MRU4.Text = "&4"
		Me.MRU4.Visible = False
		'
		'MRU5
		'
		Me.MRU5.Name = "MRU5"
		Me.MRU5.Size = New System.Drawing.Size(271, 22)
		Me.MRU5.Text = "&5"
		Me.MRU5.Visible = False
		'
		'MRU6
		'
		Me.MRU6.Name = "MRU6"
		Me.MRU6.Size = New System.Drawing.Size(271, 22)
		Me.MRU6.Text = "&6"
		Me.MRU6.Visible = False
		'
		'MRU7
		'
		Me.MRU7.Name = "MRU7"
		Me.MRU7.Size = New System.Drawing.Size(271, 22)
		Me.MRU7.Text = "&7"
		Me.MRU7.Visible = False
		'
		'MRU8
		'
		Me.MRU8.Name = "MRU8"
		Me.MRU8.Size = New System.Drawing.Size(271, 22)
		Me.MRU8.Text = "&8"
		Me.MRU8.Visible = False
		'
		'ToolStripMenuItem1
		'
		Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
		Me.ToolStripMenuItem1.Size = New System.Drawing.Size(268, 6)
		'
		'mnuExit
		'
		Me.mnuExit.Name = "mnuExit"
		Me.mnuExit.Size = New System.Drawing.Size(271, 22)
		Me.mnuExit.Text = "E&xit"
		'
		'mnuUtilities
		'
		Me.mnuUtilities.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCompactDB, Me.ToolStripMenuItem4, Me.mnuOrganizeDB, Me.mnuResetSettings, Me.mnuTheme})
		Me.mnuUtilities.Name = "mnuUtilities"
		Me.mnuUtilities.Size = New System.Drawing.Size(58, 20)
		Me.mnuUtilities.Text = "&Utilities"
		'
		'mnuCompactDB
		'
		Me.mnuCompactDB.Enabled = False
		Me.mnuCompactDB.Name = "mnuCompactDB"
		Me.mnuCompactDB.Size = New System.Drawing.Size(193, 22)
		Me.mnuCompactDB.Text = "&Compact Database"
		'
		'ToolStripMenuItem4
		'
		Me.ToolStripMenuItem4.Name = "ToolStripMenuItem4"
		Me.ToolStripMenuItem4.Size = New System.Drawing.Size(190, 6)
		'
		'mnuOrganizeDB
		'
		Me.mnuOrganizeDB.Name = "mnuOrganizeDB"
		Me.mnuOrganizeDB.Size = New System.Drawing.Size(193, 22)
		Me.mnuOrganizeDB.Text = "&Organize Database List"
		'
		'mnuResetSettings
		'
		Me.mnuResetSettings.Name = "mnuResetSettings"
		Me.mnuResetSettings.Size = New System.Drawing.Size(193, 22)
		Me.mnuResetSettings.Text = "Reset User Settings"
		'
		'mnuTheme
		'
		Me.mnuTheme.Name = "mnuTheme"
		Me.mnuTheme.Size = New System.Drawing.Size(193, 22)
		Me.mnuTheme.Text = "Change Color &Theme"
		'
		'mnuWindow
		'
		Me.mnuWindow.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNoArrange, Me.mnuTile, Me.mnuCascade})
		Me.mnuWindow.Name = "mnuWindow"
		Me.mnuWindow.Size = New System.Drawing.Size(63, 20)
		Me.mnuWindow.Text = "&Window"
		'
		'mnuNoArrange
		'
		Me.mnuNoArrange.Checked = True
		Me.mnuNoArrange.CheckOnClick = True
		Me.mnuNoArrange.CheckState = System.Windows.Forms.CheckState.Checked
		Me.mnuNoArrange.Name = "mnuNoArrange"
		Me.mnuNoArrange.Size = New System.Drawing.Size(157, 22)
		Me.mnuNoArrange.Text = "&Do Not Arrange"
		'
		'mnuTile
		'
		Me.mnuTile.CheckOnClick = True
		Me.mnuTile.Name = "mnuTile"
		Me.mnuTile.Size = New System.Drawing.Size(157, 22)
		Me.mnuTile.Text = "&Tile"
		'
		'mnuCascade
		'
		Me.mnuCascade.CheckOnClick = True
		Me.mnuCascade.Name = "mnuCascade"
		Me.mnuCascade.Size = New System.Drawing.Size(157, 22)
		Me.mnuCascade.Text = "&Cascade"
		'
		'mnuHelp
		'
		Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAbout, Me.mnuViewReadMe, Me.mnuViewLicense})
		Me.mnuHelp.Name = "mnuHelp"
		Me.mnuHelp.Size = New System.Drawing.Size(44, 20)
		Me.mnuHelp.Text = "&Help"
		'
		'mnuAbout
		'
		Me.mnuAbout.Name = "mnuAbout"
		Me.mnuAbout.Size = New System.Drawing.Size(232, 22)
		Me.mnuAbout.Text = "&About SQL Database Manager"
		'
		'StatusStrip1
		'
		Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatusLabel})
		Me.StatusStrip1.Location = New System.Drawing.Point(0, 451)
		Me.StatusStrip1.Name = "StatusStrip1"
		Me.StatusStrip1.Size = New System.Drawing.Size(777, 22)
		Me.StatusStrip1.TabIndex = 2
		Me.StatusStrip1.Tag = ""
		Me.StatusStrip1.Text = "StatusStrip1"
		'
		'StatusLabel
		'
		Me.StatusLabel.AutoSize = False
		Me.StatusLabel.Name = "StatusLabel"
		Me.StatusLabel.Size = New System.Drawing.Size(300, 17)
		Me.StatusLabel.Tag = "NoColor"
		Me.StatusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		'
		'OpenFileDialog1
		'
		Me.OpenFileDialog1.FileName = "OpenFileDialog1"
		'
		'mnuViewReadMe
		'
		Me.mnuViewReadMe.Name = "mnuViewReadMe"
		Me.mnuViewReadMe.Size = New System.Drawing.Size(232, 22)
		Me.mnuViewReadMe.Text = "View &ReadMe"
		'
		'mnuViewLicense
		'
		Me.mnuViewLicense.Name = "mnuViewLicense"
		Me.mnuViewLicense.Size = New System.Drawing.Size(232, 22)
		Me.mnuViewLicense.Text = "View &License"
		'
		'frmMain
		'
		Me.AllowDrop = True
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ClientSize = New System.Drawing.Size(777, 473)
		Me.Controls.Add(Me.StatusStrip1)
		Me.Controls.Add(Me.MenuStrip1)
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.IsMdiContainer = True
		Me.MainMenuStrip = Me.MenuStrip1
		Me.Name = "frmMain"
		Me.Text = "SQL Database Manager"
		Me.MenuStrip1.ResumeLayout(False)
		Me.MenuStrip1.PerformLayout()
		Me.StatusStrip1.ResumeLayout(False)
		Me.StatusStrip1.PerformLayout()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
	Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNewDatabase As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOpenDatabase As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCloseDatabase As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MRUSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents MRU1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MRU2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MRU3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MRU4 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAbout As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents mnuUtilities As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents mnuOrganizeDB As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuWindow As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCascade As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNoArrange As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents ToolStripMenuItem2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuTableDesigner As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuScriptEditor As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents MRU5 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MRU6 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MRU7 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MRU8 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuResetSettings As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTheme As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCompactDB As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuRelationships As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuGenerateCommands As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FontDialog1 As FontDialog
	Friend WithEvents mnuViewReadMe As ToolStripMenuItem
	Friend WithEvents mnuViewLicense As ToolStripMenuItem
End Class
