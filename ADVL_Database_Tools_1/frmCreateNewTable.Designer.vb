<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCreateNewTable
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.txtTableName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnCreateTable = New System.Windows.Forms.Button()
        Me.btnReadTableDefFile = New System.Windows.Forms.Button()
        Me.txtTableDefFile = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(15, 66)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(677, 478)
        Me.DataGridView1.TabIndex = 21
        '
        'txtTableName
        '
        Me.txtTableName.Location = New System.Drawing.Point(86, 40)
        Me.txtTableName.Name = "txtTableName"
        Me.txtTableName.Size = New System.Drawing.Size(485, 20)
        Me.txtTableName.TabIndex = 20
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 13)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Table Name:"
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(628, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 18
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnCreateTable
        '
        Me.btnCreateTable.Location = New System.Drawing.Point(12, 12)
        Me.btnCreateTable.Name = "btnCreateTable"
        Me.btnCreateTable.Size = New System.Drawing.Size(94, 22)
        Me.btnCreateTable.TabIndex = 17
        Me.btnCreateTable.Text = "Create Table"
        Me.btnCreateTable.UseVisualStyleBackColor = True
        '
        'btnReadTableDefFile
        '
        Me.btnReadTableDefFile.Location = New System.Drawing.Point(112, 12)
        Me.btnReadTableDefFile.Name = "btnReadTableDefFile"
        Me.btnReadTableDefFile.Size = New System.Drawing.Size(155, 22)
        Me.btnReadTableDefFile.TabIndex = 16
        Me.btnReadTableDefFile.Text = "Read Table Definition File"
        Me.btnReadTableDefFile.UseVisualStyleBackColor = True
        '
        'txtTableDefFile
        '
        Me.txtTableDefFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTableDefFile.Location = New System.Drawing.Point(273, 13)
        Me.txtTableDefFile.Name = "txtTableDefFile"
        Me.txtTableDefFile.Size = New System.Drawing.Size(349, 20)
        Me.txtTableDefFile.TabIndex = 22
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'frmCreateNewTable
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(704, 556)
        Me.Controls.Add(Me.txtTableDefFile)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.txtTableName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnCreateTable)
        Me.Controls.Add(Me.btnReadTableDefFile)
        Me.Name = "frmCreateNewTable"
        Me.Text = "Create New Table"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents txtTableName As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnExit As Button
    Friend WithEvents btnCreateTable As Button
    Friend WithEvents btnReadTableDefFile As Button
    Friend WithEvents txtTableDefFile As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
End Class
