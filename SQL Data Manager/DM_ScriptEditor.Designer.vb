<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmScriptEditor
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
		Me.components = New System.ComponentModel.Container()
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmScriptEditor))
		Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
		Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
		Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuNewSQL = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuOpenSQL = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuSave = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuSaveAs = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuCloseSQL = New System.Windows.Forms.ToolStripMenuItem()
		Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripSeparator()
		Me.mnuExecute = New System.Windows.Forms.ToolStripMenuItem()
		Me.ToolStripMenuItem4 = New System.Windows.Forms.ToolStripSeparator()
		Me.mnuClose = New System.Windows.Forms.ToolStripMenuItem()
		Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
		Me.StatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
		Me.ToolStripDropDownButton1 = New System.Windows.Forms.ToolStripDropDownButton()
		Me.mnu200 = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnu175 = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnu150 = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnu125 = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnu100 = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnu75 = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnu50 = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnu25 = New System.Windows.Forms.ToolStripMenuItem()
		Me.Rtb1 = New System.Windows.Forms.RichTextBox()
		Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
		Me.MenuStrip1.SuspendLayout()
		Me.StatusStrip1.SuspendLayout()
		Me.SuspendLayout()
		'
		'MenuStrip1
		'
		Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1, Me.FileToolStripMenuItem})
		Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
		Me.MenuStrip1.Name = "MenuStrip1"
		Me.MenuStrip1.Size = New System.Drawing.Size(681, 24)
		Me.MenuStrip1.TabIndex = 0
		Me.MenuStrip1.Text = "MenuStrip1"
		Me.MenuStrip1.Visible = False
		'
		'ToolStripMenuItem1
		'
		Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
		Me.ToolStripMenuItem1.Size = New System.Drawing.Size(12, 20)
		'
		'FileToolStripMenuItem
		'
		Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewSQL, Me.mnuOpenSQL, Me.mnuSave, Me.mnuSaveAs, Me.mnuCloseSQL, Me.ToolStripMenuItem2, Me.mnuExecute, Me.ToolStripMenuItem4, Me.mnuClose})
		Me.FileToolStripMenuItem.MergeAction = System.Windows.Forms.MergeAction.Insert
		Me.FileToolStripMenuItem.MergeIndex = 1
		Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
		Me.FileToolStripMenuItem.Size = New System.Drawing.Size(49, 20)
		Me.FileToolStripMenuItem.Text = "&Script"
		'
		'mnuNewSQL
		'
		Me.mnuNewSQL.Name = "mnuNewSQL"
		Me.mnuNewSQL.Size = New System.Drawing.Size(180, 22)
		Me.mnuNewSQL.Text = "&New SQL Script"
		'
		'mnuOpenSQL
		'
		Me.mnuOpenSQL.Name = "mnuOpenSQL"
		Me.mnuOpenSQL.Size = New System.Drawing.Size(180, 22)
		Me.mnuOpenSQL.Text = "&Open SQL Script..."
		'
		'mnuSave
		'
		Me.mnuSave.Name = "mnuSave"
		Me.mnuSave.Size = New System.Drawing.Size(180, 22)
		Me.mnuSave.Text = "&Save"
		'
		'mnuSaveAs
		'
		Me.mnuSaveAs.Name = "mnuSaveAs"
		Me.mnuSaveAs.Size = New System.Drawing.Size(180, 22)
		Me.mnuSaveAs.Text = "Save &As..."
		'
		'mnuCloseSQL
		'
		Me.mnuCloseSQL.Enabled = False
		Me.mnuCloseSQL.Name = "mnuCloseSQL"
		Me.mnuCloseSQL.Size = New System.Drawing.Size(180, 22)
		Me.mnuCloseSQL.Text = "&Close SQL Script"
		'
		'ToolStripMenuItem2
		'
		Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
		Me.ToolStripMenuItem2.Size = New System.Drawing.Size(177, 6)
		'
		'mnuExecute
		'
		Me.mnuExecute.Name = "mnuExecute"
		Me.mnuExecute.Size = New System.Drawing.Size(180, 22)
		Me.mnuExecute.Text = "&Execute Script"
		'
		'ToolStripMenuItem4
		'
		Me.ToolStripMenuItem4.Name = "ToolStripMenuItem4"
		Me.ToolStripMenuItem4.Size = New System.Drawing.Size(177, 6)
		'
		'mnuClose
		'
		Me.mnuClose.Name = "mnuClose"
		Me.mnuClose.Size = New System.Drawing.Size(180, 22)
		Me.mnuClose.Text = "&Close"
		'
		'StatusStrip1
		'
		Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatusLabel1, Me.ToolStripDropDownButton1})
		Me.StatusStrip1.Location = New System.Drawing.Point(0, 466)
		Me.StatusStrip1.Name = "StatusStrip1"
		Me.StatusStrip1.ShowItemToolTips = True
		Me.StatusStrip1.Size = New System.Drawing.Size(681, 22)
		Me.StatusStrip1.TabIndex = 2
		Me.StatusStrip1.Text = "StatusStrip1"
		'
		'StatusLabel1
		'
		Me.StatusLabel1.Name = "StatusLabel1"
		Me.StatusLabel1.Size = New System.Drawing.Size(637, 17)
		Me.StatusLabel1.Spring = True
		Me.StatusLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		'
		'ToolStripDropDownButton1
		'
		Me.ToolStripDropDownButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
		Me.ToolStripDropDownButton1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnu200, Me.mnu175, Me.mnu150, Me.mnu125, Me.mnu100, Me.mnu75, Me.mnu50, Me.mnu25})
		Me.ToolStripDropDownButton1.Image = Global.SQL_Database_Manager.My.Resources.Resources.View
		Me.ToolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta
		Me.ToolStripDropDownButton1.Name = "ToolStripDropDownButton1"
		Me.ToolStripDropDownButton1.Size = New System.Drawing.Size(29, 20)
		Me.ToolStripDropDownButton1.ToolTipText = "Text Size"
		'
		'mnu200
		'
		Me.mnu200.Name = "mnu200"
		Me.mnu200.Size = New System.Drawing.Size(102, 22)
		Me.mnu200.Text = "200%"
		'
		'mnu175
		'
		Me.mnu175.Name = "mnu175"
		Me.mnu175.Size = New System.Drawing.Size(102, 22)
		Me.mnu175.Text = "175%"
		'
		'mnu150
		'
		Me.mnu150.Name = "mnu150"
		Me.mnu150.Size = New System.Drawing.Size(102, 22)
		Me.mnu150.Text = "150%"
		'
		'mnu125
		'
		Me.mnu125.Name = "mnu125"
		Me.mnu125.Size = New System.Drawing.Size(102, 22)
		Me.mnu125.Text = "125%"
		'
		'mnu100
		'
		Me.mnu100.Name = "mnu100"
		Me.mnu100.Size = New System.Drawing.Size(102, 22)
		Me.mnu100.Text = "100%"
		'
		'mnu75
		'
		Me.mnu75.Name = "mnu75"
		Me.mnu75.Size = New System.Drawing.Size(102, 22)
		Me.mnu75.Text = "75%"
		'
		'mnu50
		'
		Me.mnu50.Name = "mnu50"
		Me.mnu50.Size = New System.Drawing.Size(102, 22)
		Me.mnu50.Text = "50%"
		'
		'mnu25
		'
		Me.mnu25.Name = "mnu25"
		Me.mnu25.Size = New System.Drawing.Size(102, 22)
		Me.mnu25.Text = "25%"
		'
		'Rtb1
		'
		Me.Rtb1.AcceptsTab = True
		Me.Rtb1.EnableAutoDragDrop = True
		Me.Rtb1.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Rtb1.Location = New System.Drawing.Point(0, 27)
		Me.Rtb1.Name = "Rtb1"
		Me.Rtb1.Size = New System.Drawing.Size(681, 436)
		Me.Rtb1.TabIndex = 3
		Me.Rtb1.Text = ""
		'
		'Timer1
		'
		Me.Timer1.Interval = 5000
		'
		'frmScriptEditor
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(681, 488)
		Me.Controls.Add(Me.Rtb1)
		Me.Controls.Add(Me.StatusStrip1)
		Me.Controls.Add(Me.MenuStrip1)
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.KeyPreview = True
		Me.MainMenuStrip = Me.MenuStrip1
		Me.Name = "frmScriptEditor"
		Me.Text = "SQL Script Editor"
		Me.MenuStrip1.ResumeLayout(False)
		Me.MenuStrip1.PerformLayout()
		Me.StatusStrip1.ResumeLayout(False)
		Me.StatusStrip1.PerformLayout()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
	Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNewSQL As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOpenSQL As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSave As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSaveAs As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCloseSQL As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuClose As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Rtb1 As System.Windows.Forms.RichTextBox
	Friend WithEvents mnuExecute As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents ToolStripMenuItem4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripDropDownButton1 As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents mnu200 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnu175 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnu150 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnu125 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnu100 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnu75 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnu50 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnu25 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
End Class
