<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmObjectList
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
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmObjectList))
		Me.TreeView1 = New System.Windows.Forms.TreeView()
		Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
		Me.ContextMenuStrip1a = New System.Windows.Forms.ContextMenuStrip(Me.components)
		Me.mnuNewTable = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuScriptBuilder = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuScriptEditor = New System.Windows.Forms.ToolStripMenuItem()
		Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
		Me.mnuRemoveTable = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuModifyTable = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuViewSchema = New System.Windows.Forms.ToolStripMenuItem()
		Me.ContextMenuStrip3 = New System.Windows.Forms.ContextMenuStrip(Me.components)
		Me.mnuRemoveQuery = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuModifyQuery = New System.Windows.Forms.ToolStripMenuItem()
		Me.ContextMenuStrip4 = New System.Windows.Forms.ContextMenuStrip(Me.components)
		Me.mnuRemoveView = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuModifyView = New System.Windows.Forms.ToolStripMenuItem()
		Me.ContextMenuStrip5 = New System.Windows.Forms.ContextMenuStrip(Me.components)
		Me.mnuRemoveIndex = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuModifyIndex = New System.Windows.Forms.ToolStripMenuItem()
		Me.ContextMenuStrip6 = New System.Windows.Forms.ContextMenuStrip(Me.components)
		Me.mnuModifyFK = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuRemoveFK = New System.Windows.Forms.ToolStripMenuItem()
		Me.ContextMenuStrip1b = New System.Windows.Forms.ContextMenuStrip(Me.components)
		Me.mnuAddNew = New System.Windows.Forms.ToolStripMenuItem()
		Me.ContextMenuStrip7 = New System.Windows.Forms.ContextMenuStrip(Me.components)
		Me.mnuRemoveSettings = New System.Windows.Forms.ToolStripMenuItem()
		Me.ContextMenuStrip1a.SuspendLayout()
		Me.ContextMenuStrip2.SuspendLayout()
		Me.ContextMenuStrip3.SuspendLayout()
		Me.ContextMenuStrip4.SuspendLayout()
		Me.ContextMenuStrip5.SuspendLayout()
		Me.ContextMenuStrip6.SuspendLayout()
		Me.ContextMenuStrip1b.SuspendLayout()
		Me.ContextMenuStrip7.SuspendLayout()
		Me.SuspendLayout()
		'
		'TreeView1
		'
		Me.TreeView1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TreeView1.ImageIndex = 0
		Me.TreeView1.ImageList = Me.ImageList1
		Me.TreeView1.ItemHeight = 24
		Me.TreeView1.LabelEdit = True
		Me.TreeView1.LineColor = System.Drawing.Color.ForestGreen
		Me.TreeView1.Location = New System.Drawing.Point(0, 0)
		Me.TreeView1.Name = "TreeView1"
		Me.TreeView1.SelectedImageKey = "Pointer"
		Me.TreeView1.ShowLines = False
		Me.TreeView1.Size = New System.Drawing.Size(295, 384)
		Me.TreeView1.TabIndex = 1
		Me.TreeView1.Tag = "NoColor"
		'
		'ImageList1
		'
		Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
		Me.ImageList1.Images.SetKeyName(0, "Database")
		Me.ImageList1.Images.SetKeyName(1, "DataTable")
		Me.ImageList1.Images.SetKeyName(2, "Index")
		Me.ImageList1.Images.SetKeyName(3, "Gears")
		Me.ImageList1.Images.SetKeyName(4, "Pointer")
		Me.ImageList1.Images.SetKeyName(5, "Key")
		Me.ImageList1.Images.SetKeyName(6, "SQLQuery")
		Me.ImageList1.Images.SetKeyName(7, "View")
		Me.ImageList1.Images.SetKeyName(8, "Settings")
		'
		'ContextMenuStrip1a
		'
		Me.ContextMenuStrip1a.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewTable})
		Me.ContextMenuStrip1a.Name = "ContextMenuStrip1"
		Me.ContextMenuStrip1a.Size = New System.Drawing.Size(129, 26)
		'
		'mnuNewTable
		'
		Me.mnuNewTable.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuScriptBuilder, Me.mnuScriptEditor})
		Me.mnuNewTable.Name = "mnuNewTable"
		Me.mnuNewTable.Size = New System.Drawing.Size(128, 22)
		Me.mnuNewTable.Text = "&New Table"
		'
		'mnuScriptBuilder
		'
		Me.mnuScriptBuilder.Name = "mnuScriptBuilder"
		Me.mnuScriptBuilder.Size = New System.Drawing.Size(220, 22)
		Me.mnuScriptBuilder.Text = "Create Using &Table Designer"
		'
		'mnuScriptEditor
		'
		Me.mnuScriptEditor.Name = "mnuScriptEditor"
		Me.mnuScriptEditor.Size = New System.Drawing.Size(220, 22)
		Me.mnuScriptEditor.Text = "Create Using Script &Editor"
		'
		'ContextMenuStrip2
		'
		Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuRemoveTable, Me.mnuModifyTable, Me.mnuViewSchema})
		Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
		Me.ContextMenuStrip2.Size = New System.Drawing.Size(148, 70)
		'
		'mnuRemoveTable
		'
		Me.mnuRemoveTable.Name = "mnuRemoveTable"
		Me.mnuRemoveTable.Size = New System.Drawing.Size(147, 22)
		Me.mnuRemoveTable.Text = "&Remove Table"
		'
		'mnuModifyTable
		'
		Me.mnuModifyTable.Name = "mnuModifyTable"
		Me.mnuModifyTable.Size = New System.Drawing.Size(147, 22)
		Me.mnuModifyTable.Text = "&Modify Table"
		'
		'mnuViewSchema
		'
		Me.mnuViewSchema.Name = "mnuViewSchema"
		Me.mnuViewSchema.Size = New System.Drawing.Size(147, 22)
		Me.mnuViewSchema.Text = "&View Schema"
		'
		'ContextMenuStrip3
		'
		Me.ContextMenuStrip3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuRemoveQuery, Me.mnuModifyQuery})
		Me.ContextMenuStrip3.Name = "ContextMenuStrip3"
		Me.ContextMenuStrip3.Size = New System.Drawing.Size(153, 48)
		'
		'mnuRemoveQuery
		'
		Me.mnuRemoveQuery.Name = "mnuRemoveQuery"
		Me.mnuRemoveQuery.Size = New System.Drawing.Size(152, 22)
		Me.mnuRemoveQuery.Text = "&Remove Query"
		'
		'mnuModifyQuery
		'
		Me.mnuModifyQuery.Name = "mnuModifyQuery"
		Me.mnuModifyQuery.Size = New System.Drawing.Size(152, 22)
		Me.mnuModifyQuery.Text = "&Modify Query"
		'
		'ContextMenuStrip4
		'
		Me.ContextMenuStrip4.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuRemoveView, Me.mnuModifyView})
		Me.ContextMenuStrip4.Name = "ContextMenuStrip4"
		Me.ContextMenuStrip4.Size = New System.Drawing.Size(146, 48)
		'
		'mnuRemoveView
		'
		Me.mnuRemoveView.Name = "mnuRemoveView"
		Me.mnuRemoveView.Size = New System.Drawing.Size(145, 22)
		Me.mnuRemoveView.Text = "&Remove View"
		'
		'mnuModifyView
		'
		Me.mnuModifyView.Name = "mnuModifyView"
		Me.mnuModifyView.Size = New System.Drawing.Size(145, 22)
		Me.mnuModifyView.Text = "&Modify View"
		'
		'ContextMenuStrip5
		'
		Me.ContextMenuStrip5.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuRemoveIndex, Me.mnuModifyIndex})
		Me.ContextMenuStrip5.Name = "ContextMenuStrip5"
		Me.ContextMenuStrip5.Size = New System.Drawing.Size(150, 48)
		'
		'mnuRemoveIndex
		'
		Me.mnuRemoveIndex.Name = "mnuRemoveIndex"
		Me.mnuRemoveIndex.Size = New System.Drawing.Size(149, 22)
		Me.mnuRemoveIndex.Text = "&Remove Index"
		'
		'mnuModifyIndex
		'
		Me.mnuModifyIndex.Name = "mnuModifyIndex"
		Me.mnuModifyIndex.Size = New System.Drawing.Size(149, 22)
		Me.mnuModifyIndex.Text = "&Modify Index"
		'
		'ContextMenuStrip6
		'
		Me.ContextMenuStrip6.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuModifyFK, Me.mnuRemoveFK})
		Me.ContextMenuStrip6.Name = "ContextMenuStrip6"
		Me.ContextMenuStrip6.Size = New System.Drawing.Size(186, 48)
		'
		'mnuModifyFK
		'
		Me.mnuModifyFK.Name = "mnuModifyFK"
		Me.mnuModifyFK.Size = New System.Drawing.Size(185, 22)
		Me.mnuModifyFK.Text = "&Modify Relationship"
		'
		'mnuRemoveFK
		'
		Me.mnuRemoveFK.Name = "mnuRemoveFK"
		Me.mnuRemoveFK.Size = New System.Drawing.Size(185, 22)
		Me.mnuRemoveFK.Text = "&Remove Relationship"
		'
		'ContextMenuStrip1b
		'
		Me.ContextMenuStrip1b.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAddNew})
		Me.ContextMenuStrip1b.Name = "ContextMenuStrip1b"
		Me.ContextMenuStrip1b.Size = New System.Drawing.Size(99, 26)
		'
		'mnuAddNew
		'
		Me.mnuAddNew.Name = "mnuAddNew"
		Me.mnuAddNew.Size = New System.Drawing.Size(98, 22)
		Me.mnuAddNew.Text = "New"
		'
		'ContextMenuStrip7
		'
		Me.ContextMenuStrip7.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuRemoveSettings})
		Me.ContextMenuStrip7.Name = "ContextMenuStrip8"
		Me.ContextMenuStrip7.Size = New System.Drawing.Size(163, 26)
		'
		'mnuRemoveSettings
		'
		Me.mnuRemoveSettings.Name = "mnuRemoveSettings"
		Me.mnuRemoveSettings.Size = New System.Drawing.Size(162, 22)
		Me.mnuRemoveSettings.Text = "Remove &Settings"
		'
		'frmObjectList
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(294, 383)
		Me.Controls.Add(Me.TreeView1)
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimumSize = New System.Drawing.Size(250, 300)
		Me.Name = "frmObjectList"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "Database Objects"
		Me.ContextMenuStrip1a.ResumeLayout(False)
		Me.ContextMenuStrip2.ResumeLayout(False)
		Me.ContextMenuStrip3.ResumeLayout(False)
		Me.ContextMenuStrip4.ResumeLayout(False)
		Me.ContextMenuStrip5.ResumeLayout(False)
		Me.ContextMenuStrip6.ResumeLayout(False)
		Me.ContextMenuStrip1b.ResumeLayout(False)
		Me.ContextMenuStrip7.ResumeLayout(False)
		Me.ResumeLayout(False)

	End Sub
	Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
	Friend WithEvents ContextMenuStrip1a As System.Windows.Forms.ContextMenuStrip
	Friend WithEvents mnuNewTable As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents ContextMenuStrip2 As System.Windows.Forms.ContextMenuStrip
	Friend WithEvents mnuRemoveTable As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents mnuModifyTable As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
	Friend WithEvents ContextMenuStrip3 As System.Windows.Forms.ContextMenuStrip
	Friend WithEvents mnuRemoveQuery As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents mnuModifyQuery As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents ContextMenuStrip4 As System.Windows.Forms.ContextMenuStrip
	Friend WithEvents mnuRemoveView As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents mnuModifyView As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents ContextMenuStrip5 As System.Windows.Forms.ContextMenuStrip
	Friend WithEvents mnuRemoveIndex As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents mnuModifyIndex As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents ContextMenuStrip6 As System.Windows.Forms.ContextMenuStrip
	Friend WithEvents mnuRemoveFK As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents mnuScriptBuilder As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents mnuScriptEditor As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents mnuViewSchema As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents ContextMenuStrip1b As System.Windows.Forms.ContextMenuStrip
	Friend WithEvents mnuAddNew As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents ContextMenuStrip7 As System.Windows.Forms.ContextMenuStrip
	Friend WithEvents mnuRemoveSettings As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents mnuModifyFK As System.Windows.Forms.ToolStripMenuItem
End Class
