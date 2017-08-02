<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmViewSchema
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
        Me.cmbRestrictions = New System.Windows.Forms.ComboBox()
        Me.btnViewSchema = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.btnExit = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbRestrictions
        '
        Me.cmbRestrictions.FormattingEnabled = True
        Me.cmbRestrictions.Location = New System.Drawing.Point(104, 13)
        Me.cmbRestrictions.Name = "cmbRestrictions"
        Me.cmbRestrictions.Size = New System.Drawing.Size(532, 21)
        Me.cmbRestrictions.TabIndex = 41
        '
        'btnViewSchema
        '
        Me.btnViewSchema.Location = New System.Drawing.Point(12, 12)
        Me.btnViewSchema.Name = "btnViewSchema"
        Me.btnViewSchema.Size = New System.Drawing.Size(86, 22)
        Me.btnViewSchema.TabIndex = 40
        Me.btnViewSchema.Text = "View Schema:"
        Me.btnViewSchema.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 40)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(992, 553)
        Me.DataGridView1.TabIndex = 39
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(940, 11)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 38
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'frmViewSchema
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1016, 605)
        Me.Controls.Add(Me.cmbRestrictions)
        Me.Controls.Add(Me.btnViewSchema)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmViewSchema"
        Me.Text = "View Schema"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cmbRestrictions As ComboBox
    Friend WithEvents btnViewSchema As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents btnExit As Button
End Class
