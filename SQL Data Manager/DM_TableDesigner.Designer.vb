<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTableDesigner
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
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTableDesigner))
		Me.VScrollBar1 = New System.Windows.Forms.VScrollBar()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.TextBox1 = New System.Windows.Forms.TextBox()
		Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
		Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
		Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuNewSQL = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuOpenSQL = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuSave = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuSaveAs = New System.Windows.Forms.ToolStripMenuItem()
		Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripSeparator()
		Me.mnuView = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuUpload = New System.Windows.Forms.ToolStripMenuItem()
		Me.ToolStripMenuItem4 = New System.Windows.Forms.ToolStripSeparator()
		Me.mnuClose = New System.Windows.Forms.ToolStripMenuItem()
		Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
		Me.StatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
		Me.Panel1 = New System.Windows.Forms.Panel()
		Me.Panel2 = New System.Windows.Forms.Panel()
		Me.GroupBox1 = New System.Windows.Forms.GroupBox()
		Me.optIndex = New System.Windows.Forms.RadioButton()
		Me.optTable = New System.Windows.Forms.RadioButton()
		Me.Panel3 = New System.Windows.Forms.Panel()
		Me.Panel4 = New System.Windows.Forms.Panel()
		Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
		Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
		Me.mnuInsertRow = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuDeleteRow = New System.Windows.Forms.ToolStripMenuItem()
		Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
		Me.mnuDeleteField = New System.Windows.Forms.ToolStripMenuItem()
		Me.MenuStrip1.SuspendLayout()
		Me.StatusStrip1.SuspendLayout()
		Me.Panel1.SuspendLayout()
		Me.GroupBox1.SuspendLayout()
		Me.Panel3.SuspendLayout()
		Me.ContextMenuStrip2.SuspendLayout()
		Me.ContextMenuStrip1.SuspendLayout()
		Me.SuspendLayout()
		'
		'VScrollBar1
		'
		Me.VScrollBar1.Location = New System.Drawing.Point(547, 158)
		Me.VScrollBar1.Name = "VScrollBar1"
		Me.VScrollBar1.Size = New System.Drawing.Size(17, 246)
		Me.VScrollBar1.TabIndex = 18
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(197, 42)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(68, 13)
		Me.Label1.TabIndex = 19
		Me.Label1.Text = "Table &Name:"
		'
		'TextBox1
		'
		Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TextBox1.Location = New System.Drawing.Point(284, 40)
		Me.TextBox1.Name = "TextBox1"
		Me.TextBox1.Size = New System.Drawing.Size(142, 22)
		Me.TextBox1.TabIndex = 20
		Me.TextBox1.Text = "New Table Name"
		'
		'MenuStrip1
		'
		Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1, Me.FileToolStripMenuItem})
		Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
		Me.MenuStrip1.Name = "MenuStrip1"
		Me.MenuStrip1.Size = New System.Drawing.Size(578, 24)
		Me.MenuStrip1.TabIndex = 21
		Me.MenuStrip1.Text = "MenuStrip1"
		'
		'ToolStripMenuItem1
		'
		Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
		Me.ToolStripMenuItem1.Size = New System.Drawing.Size(12, 20)
		'
		'FileToolStripMenuItem
		'
		Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewSQL, Me.mnuOpenSQL, Me.mnuSave, Me.mnuSaveAs, Me.ToolStripMenuItem2, Me.mnuView, Me.mnuUpload, Me.ToolStripMenuItem4, Me.mnuClose})
		Me.FileToolStripMenuItem.MergeAction = System.Windows.Forms.MergeAction.Insert
		Me.FileToolStripMenuItem.MergeIndex = 1
		Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
		Me.FileToolStripMenuItem.Size = New System.Drawing.Size(49, 20)
		Me.FileToolStripMenuItem.Text = "&Script"
		'
		'mnuNewSQL
		'
		Me.mnuNewSQL.Name = "mnuNewSQL"
		Me.mnuNewSQL.Size = New System.Drawing.Size(179, 22)
		Me.mnuNewSQL.Text = "&New SQL Script"
		'
		'mnuOpenSQL
		'
		Me.mnuOpenSQL.Name = "mnuOpenSQL"
		Me.mnuOpenSQL.Size = New System.Drawing.Size(179, 22)
		Me.mnuOpenSQL.Text = "&Open SQL Script..."
		'
		'mnuSave
		'
		Me.mnuSave.Name = "mnuSave"
		Me.mnuSave.Size = New System.Drawing.Size(179, 22)
		Me.mnuSave.Text = "&Save"
		'
		'mnuSaveAs
		'
		Me.mnuSaveAs.Name = "mnuSaveAs"
		Me.mnuSaveAs.Size = New System.Drawing.Size(179, 22)
		Me.mnuSaveAs.Text = "Save &As..."
		'
		'ToolStripMenuItem2
		'
		Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
		Me.ToolStripMenuItem2.Size = New System.Drawing.Size(176, 6)
		'
		'mnuView
		'
		Me.mnuView.Name = "mnuView"
		Me.mnuView.Size = New System.Drawing.Size(179, 22)
		Me.mnuView.Text = "&View in Script Editor"
		'
		'mnuUpload
		'
		Me.mnuUpload.Name = "mnuUpload"
		Me.mnuUpload.Size = New System.Drawing.Size(179, 22)
		Me.mnuUpload.Text = "&Upload to Database"
		'
		'ToolStripMenuItem4
		'
		Me.ToolStripMenuItem4.Name = "ToolStripMenuItem4"
		Me.ToolStripMenuItem4.Size = New System.Drawing.Size(176, 6)
		'
		'mnuClose
		'
		Me.mnuClose.Name = "mnuClose"
		Me.mnuClose.Size = New System.Drawing.Size(179, 22)
		Me.mnuClose.Text = "&Close"
		'
		'StatusStrip1
		'
		Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatusLabel1})
		Me.StatusStrip1.Location = New System.Drawing.Point(0, 407)
		Me.StatusStrip1.Name = "StatusStrip1"
		Me.StatusStrip1.ShowItemToolTips = True
		Me.StatusStrip1.Size = New System.Drawing.Size(578, 22)
		Me.StatusStrip1.TabIndex = 22
		Me.StatusStrip1.Text = "StatusStrip1"
		'
		'StatusLabel1
		'
		Me.StatusLabel1.Name = "StatusLabel1"
		Me.StatusLabel1.Size = New System.Drawing.Size(563, 17)
		Me.StatusLabel1.Spring = True
		Me.StatusLabel1.Tag = ""
		Me.StatusLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		'
		'Panel1
		'
		Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Panel1.Controls.Add(Me.Panel2)
		Me.Panel1.Location = New System.Drawing.Point(12, 158)
		Me.Panel1.Name = "Panel1"
		Me.Panel1.Size = New System.Drawing.Size(532, 246)
		Me.Panel1.TabIndex = 9
		'
		'Panel2
		'
		Me.Panel2.Location = New System.Drawing.Point(3, 3)
		Me.Panel2.Name = "Panel2"
		Me.Panel2.Size = New System.Drawing.Size(524, 300)
		Me.Panel2.TabIndex = 0
		'
		'GroupBox1
		'
		Me.GroupBox1.Controls.Add(Me.optIndex)
		Me.GroupBox1.Controls.Add(Me.optTable)
		Me.GroupBox1.Location = New System.Drawing.Point(12, 40)
		Me.GroupBox1.Name = "GroupBox1"
		Me.GroupBox1.Size = New System.Drawing.Size(146, 100)
		Me.GroupBox1.TabIndex = 23
		Me.GroupBox1.TabStop = False
		Me.GroupBox1.Text = "&Object Type"
		'
		'optIndex
		'
		Me.optIndex.AutoSize = True
		Me.optIndex.Location = New System.Drawing.Point(16, 50)
		Me.optIndex.Name = "optIndex"
		Me.optIndex.Size = New System.Drawing.Size(51, 17)
		Me.optIndex.TabIndex = 1
		Me.optIndex.Text = "&Index"
		Me.optIndex.UseVisualStyleBackColor = True
		'
		'optTable
		'
		Me.optTable.AutoSize = True
		Me.optTable.Location = New System.Drawing.Point(16, 27)
		Me.optTable.Name = "optTable"
		Me.optTable.Size = New System.Drawing.Size(52, 17)
		Me.optTable.TabIndex = 0
		Me.optTable.Text = "&Table"
		Me.optTable.UseVisualStyleBackColor = True
		'
		'Panel3
		'
		Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Panel3.Controls.Add(Me.Panel4)
		Me.Panel3.Location = New System.Drawing.Point(579, 158)
		Me.Panel3.Name = "Panel3"
		Me.Panel3.Size = New System.Drawing.Size(515, 246)
		Me.Panel3.TabIndex = 25
		Me.Panel3.Visible = False
		'
		'Panel4
		'
		Me.Panel4.AutoSize = True
		Me.Panel4.Location = New System.Drawing.Point(3, 3)
		Me.Panel4.Name = "Panel4"
		Me.Panel4.Size = New System.Drawing.Size(513, 100)
		Me.Panel4.TabIndex = 0
		'
		'Timer1
		'
		Me.Timer1.Interval = 5000
		'
		'ContextMenuStrip2
		'
		Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuInsertRow, Me.mnuDeleteRow})
		Me.ContextMenuStrip2.Name = "ContextMenuStrip1"
		Me.ContextMenuStrip2.Size = New System.Drawing.Size(131, 48)
		'
		'mnuInsertRow
		'
		Me.mnuInsertRow.Name = "mnuInsertRow"
		Me.mnuInsertRow.Size = New System.Drawing.Size(130, 22)
		Me.mnuInsertRow.Text = "&Insert Row"
		'
		'mnuDeleteRow
		'
		Me.mnuDeleteRow.Name = "mnuDeleteRow"
		Me.mnuDeleteRow.Size = New System.Drawing.Size(130, 22)
		Me.mnuDeleteRow.Text = "&DeleteRow"
		'
		'ContextMenuStrip1
		'
		Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuDeleteField})
		Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
		Me.ContextMenuStrip1.Size = New System.Drawing.Size(136, 26)
		'
		'mnuDeleteField
		'
		Me.mnuDeleteField.Name = "mnuDeleteField"
		Me.mnuDeleteField.Size = New System.Drawing.Size(135, 22)
		Me.mnuDeleteField.Text = "&Delete Field"
		'
		'frmTableDesigner
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(578, 429)
		Me.Controls.Add(Me.Panel3)
		Me.Controls.Add(Me.GroupBox1)
		Me.Controls.Add(Me.Panel1)
		Me.Controls.Add(Me.StatusStrip1)
		Me.Controls.Add(Me.MenuStrip1)
		Me.Controls.Add(Me.TextBox1)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.VScrollBar1)
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.Name = "frmTableDesigner"
		Me.Text = "Table Designer"
		Me.MenuStrip1.ResumeLayout(False)
		Me.MenuStrip1.PerformLayout()
		Me.StatusStrip1.ResumeLayout(False)
		Me.StatusStrip1.PerformLayout()
		Me.Panel1.ResumeLayout(False)
		Me.GroupBox1.ResumeLayout(False)
		Me.GroupBox1.PerformLayout()
		Me.Panel3.ResumeLayout(False)
		Me.Panel3.PerformLayout()
		Me.ContextMenuStrip2.ResumeLayout(False)
		Me.ContextMenuStrip1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
	Friend WithEvents VScrollBar1 As System.Windows.Forms.VScrollBar
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNewSQL As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOpenSQL As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSave As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSaveAs As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuUpload As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuClose As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents optIndex As System.Windows.Forms.RadioButton
    Friend WithEvents optTable As System.Windows.Forms.RadioButton
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents mnuView As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ContextMenuStrip2 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents mnuInsertRow As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuDeleteRow As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents mnuDeleteField As System.Windows.Forms.ToolStripMenuItem
End Class
