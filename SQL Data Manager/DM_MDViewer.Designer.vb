<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMDViewer
	Inherits System.Windows.Forms.Form

	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()>
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
	<System.Diagnostics.DebuggerStepThrough()>
	Private Sub InitializeComponent()
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMDViewer))
		Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
		Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnu = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuSave = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuSaveAs = New System.Windows.Forms.ToolStripMenuItem()
		Me.ViewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuViewRender = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuViewEditor = New System.Windows.Forms.ToolStripMenuItem()
		Me.TextBox1 = New System.Windows.Forms.TextBox()
		Me.MarkdownViewer1 = New SQL_Database_Manager.MarkdownViewer()
		Me.MenuStrip1.SuspendLayout()
		Me.SuspendLayout()
		'
		'MenuStrip1
		'
		Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ViewToolStripMenuItem})
		Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
		Me.MenuStrip1.Name = "MenuStrip1"
		Me.MenuStrip1.Size = New System.Drawing.Size(872, 24)
		Me.MenuStrip1.TabIndex = 1
		Me.MenuStrip1.Text = "MenuStrip1"
		Me.MenuStrip1.Visible = False
		'
		'FileToolStripMenuItem
		'
		Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnu, Me.mnuSave, Me.mnuSaveAs})
		Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
		Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
		Me.FileToolStripMenuItem.Text = "&File"
		'
		'mnu
		'
		Me.mnu.Name = "mnu"
		Me.mnu.Size = New System.Drawing.Size(200, 22)
		Me.mnu.Text = "&Open Markdown File"
		'
		'mnuSave
		'
		Me.mnuSave.Name = "mnuSave"
		Me.mnuSave.Size = New System.Drawing.Size(200, 22)
		Me.mnuSave.Text = "&Save Markdown File"
		'
		'mnuSaveAs
		'
		Me.mnuSaveAs.Name = "mnuSaveAs"
		Me.mnuSaveAs.Size = New System.Drawing.Size(200, 22)
		Me.mnuSaveAs.Text = "Save Markdown Files &As"
		'
		'ViewToolStripMenuItem
		'
		Me.ViewToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuViewRender, Me.mnuViewEditor})
		Me.ViewToolStripMenuItem.Name = "ViewToolStripMenuItem"
		Me.ViewToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
		Me.ViewToolStripMenuItem.Text = "&View"
		'
		'mnuViewRender
		'
		Me.mnuViewRender.Name = "mnuViewRender"
		Me.mnuViewRender.Size = New System.Drawing.Size(128, 22)
		Me.mnuViewRender.Text = "&Rendering"
		'
		'mnuViewEditor
		'
		Me.mnuViewEditor.Name = "mnuViewEditor"
		Me.mnuViewEditor.Size = New System.Drawing.Size(128, 22)
		Me.mnuViewEditor.Text = "&Editing"
		'
		'TextBox1
		'
		Me.TextBox1.Dock = System.Windows.Forms.DockStyle.Fill
		Me.TextBox1.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TextBox1.Location = New System.Drawing.Point(0, 0)
		Me.TextBox1.Multiline = True
		Me.TextBox1.Name = "TextBox1"
		Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.TextBox1.Size = New System.Drawing.Size(872, 533)
		Me.TextBox1.TabIndex = 0
		Me.TextBox1.Visible = False
		'
		'MarkdownViewer1
		'
		Me.MarkdownViewer1.AutoScroll = True
		Me.MarkdownViewer1.BackColor = System.Drawing.SystemColors.Window
		Me.MarkdownViewer1.Dock = System.Windows.Forms.DockStyle.Fill
		Me.MarkdownViewer1.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MarkdownViewer1.Location = New System.Drawing.Point(0, 0)
		Me.MarkdownViewer1.Name = "MarkdownViewer1"
		Me.MarkdownViewer1.RawText = Nothing
		Me.MarkdownViewer1.Size = New System.Drawing.Size(872, 533)
		Me.MarkdownViewer1.TabIndex = 2
		'
		'frmMDViewer
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(872, 533)
		Me.Controls.Add(Me.MarkdownViewer1)
		Me.Controls.Add(Me.TextBox1)
		Me.Controls.Add(Me.MenuStrip1)
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.MainMenuStrip = Me.MenuStrip1
		Me.Name = "frmMDViewer"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.MenuStrip1.ResumeLayout(False)
		Me.MenuStrip1.PerformLayout()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents MenuStrip1 As MenuStrip
	Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
	Friend WithEvents mnu As ToolStripMenuItem
	Friend WithEvents mnuSave As ToolStripMenuItem
	Friend WithEvents mnuSaveAs As ToolStripMenuItem
	Friend WithEvents TextBox1 As TextBox
	Friend WithEvents ViewToolStripMenuItem As ToolStripMenuItem
	Friend WithEvents mnuViewRender As ToolStripMenuItem
	Friend WithEvents mnuViewEditor As ToolStripMenuItem
	Friend WithEvents MarkdownViewer1 As MarkdownViewer
End Class
