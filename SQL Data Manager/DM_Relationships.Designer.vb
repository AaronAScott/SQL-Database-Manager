<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRelationships
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRelationships))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.RelationshipsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAddTable = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAddRelationship = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditRelationship = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuRemoveTable = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.StatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.MenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RelationshipsToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(693, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        Me.MenuStrip1.Visible = False
        '
        'RelationshipsToolStripMenuItem
        '
        Me.RelationshipsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAddTable, Me.mnuRemoveTable, Me.mnuAddRelationship, Me.mnuEditRelationship})
        Me.RelationshipsToolStripMenuItem.MergeAction = System.Windows.Forms.MergeAction.Insert
        Me.RelationshipsToolStripMenuItem.MergeIndex = 1
        Me.RelationshipsToolStripMenuItem.Name = "RelationshipsToolStripMenuItem"
        Me.RelationshipsToolStripMenuItem.Size = New System.Drawing.Size(89, 20)
        Me.RelationshipsToolStripMenuItem.Text = "&Relationships"
        '
        'mnuAddTable
        '
        Me.mnuAddTable.Name = "mnuAddTable"
        Me.mnuAddTable.Size = New System.Drawing.Size(164, 22)
        Me.mnuAddTable.Text = "Add &Table"
        '
        'mnuAddRelationship
        '
        Me.mnuAddRelationship.Name = "mnuAddRelationship"
        Me.mnuAddRelationship.Size = New System.Drawing.Size(164, 22)
        Me.mnuAddRelationship.Text = "&Add Relationship"
        '
        'mnuEditRelationship
        '
        Me.mnuEditRelationship.Name = "mnuEditRelationship"
        Me.mnuEditRelationship.Size = New System.Drawing.Size(164, 22)
        Me.mnuEditRelationship.Text = "&Edit Relationship"
        '
        'mnuRemoveTable
        '
        Me.mnuRemoveTable.Name = "mnuRemoveTable"
        Me.mnuRemoveTable.Size = New System.Drawing.Size(164, 22)
        Me.mnuRemoveTable.Text = "Re&move Table"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatusLabel})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 464)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(693, 22)
        Me.StatusStrip1.TabIndex = 1
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'StatusLabel
        '
        Me.StatusLabel.AutoSize = False
        Me.StatusLabel.Name = "StatusLabel"
        Me.StatusLabel.Size = New System.Drawing.Size(200, 17)
        Me.StatusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Timer1
        '
        Me.Timer1.Interval = 10000
        '
        'frmRelationships
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(693, 486)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmRelationships"
        Me.Text = "Table Relationships"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents RelationshipsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAddTable As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAddRelationship As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuEditRelationship As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRemoveTable As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
End Class
