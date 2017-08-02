<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSaveDatabaseDefinition
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
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnSaveDatabaseDefintion = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtDatabaseDescr = New System.Windows.Forms.TextBox()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtFileName = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(305, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 10
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnSaveDatabaseDefintion
        '
        Me.btnSaveDatabaseDefintion.Location = New System.Drawing.Point(12, 12)
        Me.btnSaveDatabaseDefintion.Name = "btnSaveDatabaseDefintion"
        Me.btnSaveDatabaseDefintion.Size = New System.Drawing.Size(149, 22)
        Me.btnSaveDatabaseDefintion.TabIndex = 36
        Me.btnSaveDatabaseDefintion.Text = "Save Database Definition"
        Me.btnSaveDatabaseDefintion.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 82)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(110, 13)
        Me.Label1.TabIndex = 37
        Me.Label1.Text = "Database description:"
        '
        'txtDatabaseDescr
        '
        Me.txtDatabaseDescr.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDatabaseDescr.Location = New System.Drawing.Point(15, 98)
        Me.txtDatabaseDescr.Multiline = True
        Me.txtDatabaseDescr.Name = "txtDatabaseDescr"
        Me.txtDatabaseDescr.Size = New System.Drawing.Size(354, 76)
        Me.txtDatabaseDescr.TabIndex = 38
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 37)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(99, 13)
        Me.Label2.TabIndex = 39
        Me.Label2.Text = "Definition file name:"
        '
        'txtFileName
        '
        Me.txtFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFileName.Location = New System.Drawing.Point(15, 53)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(354, 20)
        Me.txtFileName.TabIndex = 40
        '
        'frmSaveDatabaseDefinition
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(381, 186)
        Me.Controls.Add(Me.txtFileName)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtDatabaseDescr)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnSaveDatabaseDefintion)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmSaveDatabaseDefinition"
        Me.Text = "Save Database Definition"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents btnSaveDatabaseDefintion As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents txtDatabaseDescr As TextBox
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents Label2 As Label
    Friend WithEvents txtFileName As TextBox
End Class
