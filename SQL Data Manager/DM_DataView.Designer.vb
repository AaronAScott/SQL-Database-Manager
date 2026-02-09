<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDataView
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
		Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDataView))
		Me.DataGridView1 = New System.Windows.Forms.DataGridView()
		Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
		Me.btnFirst = New System.Windows.Forms.ToolStripButton()
		Me.StatusLabel = New System.Windows.Forms.ToolStripTextBox()
		Me.btnLast = New System.Windows.Forms.ToolStripButton()
		Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
		Me.MessageLabel = New System.Windows.Forms.ToolStripLabel()
		Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
		Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
		Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuSearch = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuReplace = New System.Windows.Forms.ToolStripMenuItem()
		Me.FormatToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
		Me.mnuFontSize = New System.Windows.Forms.ToolStripMenuItem()
		CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ToolStrip1.SuspendLayout()
		Me.MenuStrip1.SuspendLayout()
		Me.SuspendLayout()
		'
		'DataGridView1
		'
		DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ActiveCaption
		DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
		DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
		Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
		Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.DataGridView1.Location = New System.Drawing.Point(0, 27)
		Me.DataGridView1.Name = "DataGridView1"
		Me.DataGridView1.Size = New System.Drawing.Size(837, 364)
		Me.DataGridView1.TabIndex = 0
		'
		'ToolStrip1
		'
		Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnFirst, Me.StatusLabel, Me.btnLast, Me.ToolStripSeparator1, Me.MessageLabel})
		Me.ToolStrip1.Location = New System.Drawing.Point(0, 394)
		Me.ToolStrip1.Name = "ToolStrip1"
		Me.ToolStrip1.Size = New System.Drawing.Size(836, 25)
		Me.ToolStrip1.TabIndex = 1
		Me.ToolStrip1.Text = "ToolStrip1"
		'
		'btnFirst
		'
		Me.btnFirst.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
		Me.btnFirst.Image = CType(resources.GetObject("btnFirst.Image"), System.Drawing.Image)
		Me.btnFirst.ImageTransparentColor = System.Drawing.Color.Magenta
		Me.btnFirst.Name = "btnFirst"
		Me.btnFirst.Size = New System.Drawing.Size(23, 22)
		Me.btnFirst.Text = " btnFirst"
		Me.btnFirst.ToolTipText = "Move to First Record"
		'
		'StatusLabel
		'
		Me.StatusLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.StatusLabel.Font = New System.Drawing.Font("Segoe UI", 9.0!)
		Me.StatusLabel.Name = "StatusLabel"
		Me.StatusLabel.ReadOnly = True
		Me.StatusLabel.Size = New System.Drawing.Size(150, 25)
		Me.StatusLabel.TextBoxTextAlign = System.Windows.Forms.HorizontalAlignment.Center
		'
		'btnLast
		'
		Me.btnLast.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
		Me.btnLast.Image = CType(resources.GetObject("btnLast.Image"), System.Drawing.Image)
		Me.btnLast.ImageTransparentColor = System.Drawing.Color.Magenta
		Me.btnLast.Name = "btnLast"
		Me.btnLast.Size = New System.Drawing.Size(23, 22)
		Me.btnLast.Text = "btnLast"
		Me.btnLast.ToolTipText = "Move to Last Record"
		'
		'ToolStripSeparator1
		'
		Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
		Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
		'
		'MessageLabel
		'
		Me.MessageLabel.AutoSize = False
		Me.MessageLabel.Name = "MessageLabel"
		Me.MessageLabel.Size = New System.Drawing.Size(550, 22)
		Me.MessageLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		'
		'Timer1
		'
		Me.Timer1.Interval = 5000
		'
		'ToolTip1
		'
		Me.ToolTip1.IsBalloon = True
		Me.ToolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Warning
		Me.ToolTip1.ToolTipTitle = "Invalid Column Value"
		'
		'MenuStrip1
		'
		Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EditToolStripMenuItem, Me.FormatToolStripMenuItem})
		Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
		Me.MenuStrip1.Name = "MenuStrip1"
		Me.MenuStrip1.Size = New System.Drawing.Size(836, 24)
		Me.MenuStrip1.TabIndex = 2
		Me.MenuStrip1.Text = "MenuStrip1"
		'
		'EditToolStripMenuItem
		'
		Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuSearch, Me.mnuReplace})
		Me.EditToolStripMenuItem.MergeAction = System.Windows.Forms.MergeAction.Insert
		Me.EditToolStripMenuItem.MergeIndex = 1
		Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
		Me.EditToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
		Me.EditToolStripMenuItem.Text = "&Edit"
		'
		'mnuSearch
		'
		Me.mnuSearch.Name = "mnuSearch"
		Me.mnuSearch.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
		Me.mnuSearch.Size = New System.Drawing.Size(219, 22)
		Me.mnuSearch.Text = "&Search Table"
		'
		'mnuReplace
		'
		Me.mnuReplace.Name = "mnuReplace"
		Me.mnuReplace.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.H), System.Windows.Forms.Keys)
		Me.mnuReplace.Size = New System.Drawing.Size(219, 22)
		Me.mnuReplace.Text = "Search and &Replace"
		'
		'FormatToolStripMenuItem
		'
		Me.FormatToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFontSize})
		Me.FormatToolStripMenuItem.Name = "FormatToolStripMenuItem"
		Me.FormatToolStripMenuItem.Size = New System.Drawing.Size(57, 20)
		Me.FormatToolStripMenuItem.Text = "&Format"
		'
		'mnuFontSize
		'
		Me.mnuFontSize.Name = "mnuFontSize"
		Me.mnuFontSize.Size = New System.Drawing.Size(180, 22)
		Me.mnuFontSize.Text = "&Cell Font..."
		'
		'frmDataView
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(836, 419)
		Me.Controls.Add(Me.ToolStrip1)
		Me.Controls.Add(Me.MenuStrip1)
		Me.Controls.Add(Me.DataGridView1)
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.KeyPreview = True
		Me.MainMenuStrip = Me.MenuStrip1
		Me.MaximizeBox = False
		Me.Name = "frmDataView"
		CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ToolStrip1.ResumeLayout(False)
		Me.ToolStrip1.PerformLayout()
		Me.MenuStrip1.ResumeLayout(False)
		Me.MenuStrip1.PerformLayout()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
	Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents btnFirst As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnLast As System.Windows.Forms.ToolStripButton
    Friend WithEvents MessageLabel As System.Windows.Forms.ToolStripLabel
    Friend WithEvents StatusLabel As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSearch As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuReplace As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents FormatToolStripMenuItem As ToolStripMenuItem
	Friend WithEvents mnuFontSize As ToolStripMenuItem
End Class
