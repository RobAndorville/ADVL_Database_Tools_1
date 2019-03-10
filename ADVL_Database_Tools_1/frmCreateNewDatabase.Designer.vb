<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCreateNewDatabase
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
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.txtDefaultDir = New System.Windows.Forms.TextBox()
        Me.txtDefaultName = New System.Windows.Forms.TextBox()
        Me.btnFind = New System.Windows.Forms.Button()
        Me.txtDefinitionFilePath = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtNewDatabaseDir = New System.Windows.Forms.TextBox()
        Me.txtNewDatabaseName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnCreateNewDatabase = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(571, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 9
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnSelect
        '
        Me.btnSelect.Location = New System.Drawing.Point(53, 85)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(58, 22)
        Me.btnSelect.TabIndex = 65
        Me.btnSelect.Text = "Select"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 109)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(63, 13)
        Me.Label7.TabIndex = 76
        Me.Label7.Text = "Description:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 51)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(52, 13)
        Me.Label6.TabIndex = 75
        Me.Label6.Text = "Directory:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(38, 13)
        Me.Label5.TabIndex = 74
        Me.Label5.Text = "Name:"
        '
        'txtDescription
        '
        Me.txtDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescription.Location = New System.Drawing.Point(75, 106)
        Me.txtDescription.Multiline = True
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.ReadOnly = True
        Me.txtDescription.Size = New System.Drawing.Size(542, 48)
        Me.txtDescription.TabIndex = 73
        '
        'txtDefaultDir
        '
        Me.txtDefaultDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDefaultDir.Location = New System.Drawing.Point(75, 48)
        Me.txtDefaultDir.Multiline = True
        Me.txtDefaultDir.Name = "txtDefaultDir"
        Me.txtDefaultDir.ReadOnly = True
        Me.txtDefaultDir.Size = New System.Drawing.Size(542, 48)
        Me.txtDefaultDir.TabIndex = 72
        '
        'txtDefaultName
        '
        Me.txtDefaultName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDefaultName.Location = New System.Drawing.Point(75, 19)
        Me.txtDefaultName.Name = "txtDefaultName"
        Me.txtDefaultName.ReadOnly = True
        Me.txtDefaultName.Size = New System.Drawing.Size(542, 20)
        Me.txtDefaultName.TabIndex = 71
        '
        'btnFind
        '
        Me.btnFind.Location = New System.Drawing.Point(53, 145)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(58, 22)
        Me.btnFind.TabIndex = 70
        Me.btnFind.Text = "Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'txtDefinitionFilePath
        '
        Me.txtDefinitionFilePath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDefinitionFilePath.Location = New System.Drawing.Point(117, 126)
        Me.txtDefinitionFilePath.Multiline = True
        Me.txtDefinitionFilePath.Name = "txtDefinitionFilePath"
        Me.txtDefinitionFilePath.Size = New System.Drawing.Size(518, 48)
        Me.txtDefinitionFilePath.TabIndex = 69
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(41, 129)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 13)
        Me.Label4.TabIndex = 68
        Me.Label4.Text = "Definition file:"
        '
        'txtNewDatabaseDir
        '
        Me.txtNewDatabaseDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNewDatabaseDir.Location = New System.Drawing.Point(117, 66)
        Me.txtNewDatabaseDir.Multiline = True
        Me.txtNewDatabaseDir.Name = "txtNewDatabaseDir"
        Me.txtNewDatabaseDir.Size = New System.Drawing.Size(518, 46)
        Me.txtNewDatabaseDir.TabIndex = 67
        '
        'txtNewDatabaseName
        '
        Me.txtNewDatabaseName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNewDatabaseName.Location = New System.Drawing.Point(117, 40)
        Me.txtNewDatabaseName.Name = "txtNewDatabaseName"
        Me.txtNewDatabaseName.Size = New System.Drawing.Size(518, 20)
        Me.txtNewDatabaseName.TabIndex = 66
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 69)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(99, 13)
        Me.Label3.TabIndex = 64
        Me.Label3.Text = "Database directory:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(26, 43)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(85, 13)
        Me.Label2.TabIndex = 63
        Me.Label2.Text = "Database name:"
        '
        'btnCreateNewDatabase
        '
        Me.btnCreateNewDatabase.Location = New System.Drawing.Point(12, 12)
        Me.btnCreateNewDatabase.Name = "btnCreateNewDatabase"
        Me.btnCreateNewDatabase.Size = New System.Drawing.Size(127, 22)
        Me.btnCreateNewDatabase.TabIndex = 62
        Me.btnCreateNewDatabase.Text = "Create New Database"
        Me.btnCreateNewDatabase.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.txtDefaultName)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtDefaultDir)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.txtDescription)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 180)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(623, 165)
        Me.GroupBox1.TabIndex = 77
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Information Stored in Definition File:"
        '
        'frmCreateNewDatabase
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(647, 356)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnSelect)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.txtDefinitionFilePath)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtNewDatabaseDir)
        Me.Controls.Add(Me.txtNewDatabaseName)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnCreateNewDatabase)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmCreateNewDatabase"
        Me.Text = "Create New Database"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents btnSelect As Button
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents txtDescription As TextBox
    Friend WithEvents txtDefaultDir As TextBox
    Friend WithEvents txtDefaultName As TextBox
    Friend WithEvents btnFind As Button
    Friend WithEvents txtDefinitionFilePath As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents txtNewDatabaseDir As TextBox
    Friend WithEvents txtNewDatabaseName As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents btnCreateNewDatabase As Button
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents GroupBox1 As GroupBox
End Class
