<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSearch
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.optAny = New System.Windows.Forms.RadioButton()
        Me.optEntire = New System.Windows.Forms.RadioButton()
        Me.optStart = New System.Windows.Forms.RadioButton()
        Me.btnFind = New System.Windows.Forms.Button()
        Me.btnFindNext = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(13, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Search For:"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(81, 10)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(176, 20)
        Me.TextBox1.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.optAny)
        Me.GroupBox1.Controls.Add(Me.optEntire)
        Me.GroupBox1.Controls.Add(Me.optStart)
        Me.GroupBox1.Location = New System.Drawing.Point(81, 91)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(176, 100)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Match"
        '
        'optAny
        '
        Me.optAny.Checked = True
        Me.optAny.Location = New System.Drawing.Point(18, 65)
        Me.optAny.Name = "optAny"
        Me.optAny.Size = New System.Drawing.Size(106, 17)
        Me.optAny.TabIndex = 7
        Me.optAny.TabStop = True
        Me.optAny.Text = "Any Part of Field"
        Me.optAny.UseVisualStyleBackColor = True
        '
        'optEntire
        '
        Me.optEntire.Location = New System.Drawing.Point(18, 42)
        Me.optEntire.Name = "optEntire"
        Me.optEntire.Size = New System.Drawing.Size(106, 17)
        Me.optEntire.TabIndex = 6
        Me.optEntire.Text = "Entire Field"
        Me.optEntire.UseVisualStyleBackColor = True
        '
        'optStart
        '
        Me.optStart.Location = New System.Drawing.Point(18, 19)
        Me.optStart.Name = "optStart"
        Me.optStart.Size = New System.Drawing.Size(106, 17)
        Me.optStart.TabIndex = 5
        Me.optStart.Text = "Start of Field"
        Me.optStart.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Location = New System.Drawing.Point(304, 10)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(75, 26)
        Me.btnFind.TabIndex = 8
        Me.btnFind.Text = "Find &First"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnFindNext
        '
        Me.btnFindNext.Location = New System.Drawing.Point(304, 50)
        Me.btnFindNext.Name = "btnFindNext"
        Me.btnFindNext.Size = New System.Drawing.Size(75, 26)
        Me.btnFindNext.TabIndex = 9
        Me.btnFindNext.Text = "Find &Next"
        Me.btnFindNext.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(304, 90)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 26)
        Me.btnClose.TabIndex = 10
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(13, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Search In:"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(81, 49)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(176, 21)
        Me.ComboBox1.TabIndex = 3
        '
        'frmSearch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(402, 203)
        Me.ControlBox = False
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnFindNext)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSearch"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Search Table"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents optAny As System.Windows.Forms.RadioButton
    Friend WithEvents optEntire As System.Windows.Forms.RadioButton
    Friend WithEvents optStart As System.Windows.Forms.RadioButton
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnFindNext As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
End Class
