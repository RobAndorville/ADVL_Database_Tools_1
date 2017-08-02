<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main))
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnOnline = New System.Windows.Forms.Button()
        Me.btnMessages = New System.Windows.Forms.Button()
        Me.btnAppInfo = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.txtLastUsed = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtCreationDate = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtDataLocationPath = New System.Windows.Forms.TextBox()
        Me.txtDataLocationType = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtSettingsLocationPath = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtSettingsLocationType = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtProjectType = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtProjectDescription = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtProjectName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnProject = New System.Windows.Forms.Button()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.btnUpdateTables = New System.Windows.Forms.Button()
        Me.cmbDatabaseType = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.lstFields = New System.Windows.Forms.ListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lstTables = New System.Windows.Forms.ListBox()
        Me.btnDatabase = New System.Windows.Forms.Button()
        Me.txtDatabase = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnSaveDef = New System.Windows.Forms.Button()
        Me.btnCreateNew = New System.Windows.Forms.Button()
        Me.btnModify = New System.Windows.Forms.Button()
        Me.btnSqlCommand = New System.Windows.Forms.Button()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.btnCreateNewTable = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.btnApplyQuery = New System.Windows.Forms.Button()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtQuery = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.cmbSelectTable = New System.Windows.Forms.ComboBox()
        Me.btnShowTableProperties = New System.Windows.Forms.Button()
        Me.btnSaveTableDefinition = New System.Windows.Forms.Button()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.btnSaveChanges = New System.Windows.Forms.Button()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.cmbRestrictions = New System.Windows.Forms.ComboBox()
        Me.btnViewSchema = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.btnAndorville = New System.Windows.Forms.Button()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(652, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 8
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnOnline
        '
        Me.btnOnline.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOnline.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOnline.ForeColor = System.Drawing.Color.Red
        Me.btnOnline.Location = New System.Drawing.Point(590, 12)
        Me.btnOnline.Name = "btnOnline"
        Me.btnOnline.Size = New System.Drawing.Size(56, 22)
        Me.btnOnline.TabIndex = 37
        Me.btnOnline.Text = "Offline"
        Me.btnOnline.UseVisualStyleBackColor = True
        '
        'btnMessages
        '
        Me.btnMessages.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnMessages.Location = New System.Drawing.Point(512, 12)
        Me.btnMessages.Name = "btnMessages"
        Me.btnMessages.Size = New System.Drawing.Size(72, 22)
        Me.btnMessages.TabIndex = 38
        Me.btnMessages.Text = "Messages"
        Me.btnMessages.UseVisualStyleBackColor = True
        '
        'btnAppInfo
        '
        Me.btnAppInfo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAppInfo.Location = New System.Drawing.Point(411, 12)
        Me.btnAppInfo.Name = "btnAppInfo"
        Me.btnAppInfo.Size = New System.Drawing.Size(95, 22)
        Me.btnAppInfo.TabIndex = 46
        Me.btnAppInfo.Text = "Application Info"
        Me.btnAppInfo.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Location = New System.Drawing.Point(12, 40)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(704, 451)
        Me.TabControl1.TabIndex = 49
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.txtLastUsed)
        Me.TabPage1.Controls.Add(Me.Label11)
        Me.TabPage1.Controls.Add(Me.txtCreationDate)
        Me.TabPage1.Controls.Add(Me.Label10)
        Me.TabPage1.Controls.Add(Me.txtDataLocationPath)
        Me.TabPage1.Controls.Add(Me.txtDataLocationType)
        Me.TabPage1.Controls.Add(Me.Label9)
        Me.TabPage1.Controls.Add(Me.Label8)
        Me.TabPage1.Controls.Add(Me.txtSettingsLocationPath)
        Me.TabPage1.Controls.Add(Me.Label7)
        Me.TabPage1.Controls.Add(Me.txtSettingsLocationType)
        Me.TabPage1.Controls.Add(Me.Label6)
        Me.TabPage1.Controls.Add(Me.txtProjectType)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.txtProjectDescription)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.txtProjectName)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.btnProject)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(696, 425)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Project Information"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'txtLastUsed
        '
        Me.txtLastUsed.Location = New System.Drawing.Point(457, 109)
        Me.txtLastUsed.Name = "txtLastUsed"
        Me.txtLastUsed.Size = New System.Drawing.Size(150, 20)
        Me.txtLastUsed.TabIndex = 65
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(374, 112)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(78, 13)
        Me.Label11.TabIndex = 64
        Me.Label11.Text = "Date last used:"
        '
        'txtCreationDate
        '
        Me.txtCreationDate.Location = New System.Drawing.Point(217, 108)
        Me.txtCreationDate.Name = "txtCreationDate"
        Me.txtCreationDate.Size = New System.Drawing.Size(150, 20)
        Me.txtCreationDate.TabIndex = 63
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(138, 112)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(73, 13)
        Me.Label10.TabIndex = 62
        Me.Label10.Text = "Creation date:"
        '
        'txtDataLocationPath
        '
        Me.txtDataLocationPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDataLocationPath.Location = New System.Drawing.Point(215, 186)
        Me.txtDataLocationPath.Multiline = True
        Me.txtDataLocationPath.Name = "txtDataLocationPath"
        Me.txtDataLocationPath.Size = New System.Drawing.Size(462, 46)
        Me.txtDataLocationPath.TabIndex = 61
        '
        'txtDataLocationType
        '
        Me.txtDataLocationType.Location = New System.Drawing.Point(124, 186)
        Me.txtDataLocationType.Name = "txtDataLocationType"
        Me.txtDataLocationType.Size = New System.Drawing.Size(85, 20)
        Me.txtDataLocationType.TabIndex = 60
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(112, 215)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(97, 13)
        Me.Label9.TabIndex = 59
        Me.Label9.Text = "Data location path:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 189)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(96, 13)
        Me.Label8.TabIndex = 58
        Me.Label8.Text = "Data location type:"
        '
        'txtSettingsLocationPath
        '
        Me.txtSettingsLocationPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSettingsLocationPath.Location = New System.Drawing.Point(215, 134)
        Me.txtSettingsLocationPath.Multiline = True
        Me.txtSettingsLocationPath.Name = "txtSettingsLocationPath"
        Me.txtSettingsLocationPath.Size = New System.Drawing.Size(462, 46)
        Me.txtSettingsLocationPath.TabIndex = 57
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(97, 162)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 13)
        Me.Label7.TabIndex = 56
        Me.Label7.Text = "Settings location path:"
        '
        'txtSettingsLocationType
        '
        Me.txtSettingsLocationType.Location = New System.Drawing.Point(123, 134)
        Me.txtSettingsLocationType.Name = "txtSettingsLocationType"
        Me.txtSettingsLocationType.Size = New System.Drawing.Size(86, 20)
        Me.txtSettingsLocationType.TabIndex = 55
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 137)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(111, 13)
        Me.Label6.TabIndex = 54
        Me.Label6.Text = "Settings location type:"
        '
        'txtProjectType
        '
        Me.txtProjectType.Location = New System.Drawing.Point(46, 109)
        Me.txtProjectType.Name = "txtProjectType"
        Me.txtProjectType.Size = New System.Drawing.Size(86, 20)
        Me.txtProjectType.TabIndex = 53
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 111)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(34, 13)
        Me.Label5.TabIndex = 52
        Me.Label5.Text = "Type:"
        '
        'txtProjectDescription
        '
        Me.txtProjectDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtProjectDescription.Location = New System.Drawing.Point(123, 62)
        Me.txtProjectDescription.Multiline = True
        Me.txtProjectDescription.Name = "txtProjectDescription"
        Me.txtProjectDescription.Size = New System.Drawing.Size(554, 40)
        Me.txtProjectDescription.TabIndex = 51
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 65)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(63, 13)
        Me.Label4.TabIndex = 50
        Me.Label4.Text = "Description:"
        '
        'txtProjectName
        '
        Me.txtProjectName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtProjectName.Location = New System.Drawing.Point(123, 36)
        Me.txtProjectName.Name = "txtProjectName"
        Me.txtProjectName.Size = New System.Drawing.Size(554, 20)
        Me.txtProjectName.TabIndex = 49
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 39)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 13)
        Me.Label3.TabIndex = 48
        Me.Label3.Text = "Name:"
        '
        'btnProject
        '
        Me.btnProject.Location = New System.Drawing.Point(6, 6)
        Me.btnProject.Name = "btnProject"
        Me.btnProject.Size = New System.Drawing.Size(69, 22)
        Me.btnProject.TabIndex = 47
        Me.btnProject.Text = "Project"
        Me.btnProject.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.btnUpdateTables)
        Me.TabPage2.Controls.Add(Me.cmbDatabaseType)
        Me.TabPage2.Controls.Add(Me.Label13)
        Me.TabPage2.Controls.Add(Me.Label12)
        Me.TabPage2.Controls.Add(Me.lstFields)
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.lstTables)
        Me.TabPage2.Controls.Add(Me.btnDatabase)
        Me.TabPage2.Controls.Add(Me.txtDatabase)
        Me.TabPage2.Controls.Add(Me.Label1)
        Me.TabPage2.Controls.Add(Me.btnSaveDef)
        Me.TabPage2.Controls.Add(Me.btnCreateNew)
        Me.TabPage2.Controls.Add(Me.btnModify)
        Me.TabPage2.Controls.Add(Me.btnSqlCommand)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(696, 425)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Database"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'btnUpdateTables
        '
        Me.btnUpdateTables.Location = New System.Drawing.Point(6, 101)
        Me.btnUpdateTables.Name = "btnUpdateTables"
        Me.btnUpdateTables.Size = New System.Drawing.Size(53, 22)
        Me.btnUpdateTables.TabIndex = 51
        Me.btnUpdateTables.Text = "Update"
        Me.btnUpdateTables.UseVisualStyleBackColor = True
        '
        'cmbDatabaseType
        '
        Me.cmbDatabaseType.FormattingEnabled = True
        Me.cmbDatabaseType.Location = New System.Drawing.Point(46, 7)
        Me.cmbDatabaseType.Name = "cmbDatabaseType"
        Me.cmbDatabaseType.Size = New System.Drawing.Size(221, 21)
        Me.cmbDatabaseType.TabIndex = 26
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(6, 11)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(34, 13)
        Me.Label13.TabIndex = 25
        Me.Label13.Text = "Type:"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(6, 238)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(37, 13)
        Me.Label12.TabIndex = 24
        Me.Label12.Text = "Fields:"
        '
        'lstFields
        '
        Me.lstFields.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstFields.FormattingEnabled = True
        Me.lstFields.Location = New System.Drawing.Point(68, 238)
        Me.lstFields.Name = "lstFields"
        Me.lstFields.Size = New System.Drawing.Size(622, 160)
        Me.lstFields.TabIndex = 23
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 85)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(42, 13)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Tables:"
        '
        'lstTables
        '
        Me.lstTables.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstTables.FormattingEnabled = True
        Me.lstTables.Location = New System.Drawing.Point(68, 85)
        Me.lstTables.Name = "lstTables"
        Me.lstTables.Size = New System.Drawing.Size(622, 147)
        Me.lstTables.TabIndex = 21
        '
        'btnDatabase
        '
        Me.btnDatabase.Location = New System.Drawing.Point(6, 53)
        Me.btnDatabase.Name = "btnDatabase"
        Me.btnDatabase.Size = New System.Drawing.Size(53, 22)
        Me.btnDatabase.TabIndex = 20
        Me.btnDatabase.Text = "Find"
        Me.btnDatabase.UseVisualStyleBackColor = True
        '
        'txtDatabase
        '
        Me.txtDatabase.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDatabase.Location = New System.Drawing.Point(68, 34)
        Me.txtDatabase.Multiline = True
        Me.txtDatabase.Name = "txtDatabase"
        Me.txtDatabase.Size = New System.Drawing.Size(622, 45)
        Me.txtDatabase.TabIndex = 19
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 37)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 13)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "Database:"
        '
        'btnSaveDef
        '
        Me.btnSaveDef.Location = New System.Drawing.Point(550, 6)
        Me.btnSaveDef.Name = "btnSaveDef"
        Me.btnSaveDef.Size = New System.Drawing.Size(136, 22)
        Me.btnSaveDef.TabIndex = 17
        Me.btnSaveDef.Text = "Save Database Definition"
        Me.btnSaveDef.UseVisualStyleBackColor = True
        '
        'btnCreateNew
        '
        Me.btnCreateNew.Location = New System.Drawing.Point(423, 6)
        Me.btnCreateNew.Name = "btnCreateNew"
        Me.btnCreateNew.Size = New System.Drawing.Size(121, 22)
        Me.btnCreateNew.TabIndex = 16
        Me.btnCreateNew.Text = "Create New Database"
        Me.btnCreateNew.UseVisualStyleBackColor = True
        '
        'btnModify
        '
        Me.btnModify.Location = New System.Drawing.Point(369, 6)
        Me.btnModify.Name = "btnModify"
        Me.btnModify.Size = New System.Drawing.Size(48, 22)
        Me.btnModify.TabIndex = 15
        Me.btnModify.Text = "Modify"
        Me.btnModify.UseVisualStyleBackColor = True
        '
        'btnSqlCommand
        '
        Me.btnSqlCommand.Location = New System.Drawing.Point(273, 6)
        Me.btnSqlCommand.Name = "btnSqlCommand"
        Me.btnSqlCommand.Size = New System.Drawing.Size(90, 22)
        Me.btnSqlCommand.TabIndex = 14
        Me.btnSqlCommand.Text = "SQL Command"
        Me.btnSqlCommand.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.btnCreateNewTable)
        Me.TabPage3.Controls.Add(Me.DataGridView1)
        Me.TabPage3.Controls.Add(Me.btnApplyQuery)
        Me.TabPage3.Controls.Add(Me.Label15)
        Me.TabPage3.Controls.Add(Me.txtQuery)
        Me.TabPage3.Controls.Add(Me.Label14)
        Me.TabPage3.Controls.Add(Me.cmbSelectTable)
        Me.TabPage3.Controls.Add(Me.btnShowTableProperties)
        Me.TabPage3.Controls.Add(Me.btnSaveTableDefinition)
        Me.TabPage3.Controls.Add(Me.btnImport)
        Me.TabPage3.Controls.Add(Me.btnExport)
        Me.TabPage3.Controls.Add(Me.btnSaveChanges)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(696, 425)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Table"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'btnCreateNewTable
        '
        Me.btnCreateNewTable.Location = New System.Drawing.Point(564, 6)
        Me.btnCreateNewTable.Name = "btnCreateNewTable"
        Me.btnCreateNewTable.Size = New System.Drawing.Size(106, 22)
        Me.btnCreateNewTable.TabIndex = 39
        Me.btnCreateNewTable.Text = "Create NewTable"
        Me.btnCreateNewTable.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(6, 113)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(657, 281)
        Me.DataGridView1.TabIndex = 38
        '
        'btnApplyQuery
        '
        Me.btnApplyQuery.Location = New System.Drawing.Point(5, 85)
        Me.btnApplyQuery.Name = "btnApplyQuery"
        Me.btnApplyQuery.Size = New System.Drawing.Size(64, 22)
        Me.btnApplyQuery.TabIndex = 37
        Me.btnApplyQuery.Text = "Apply"
        Me.btnApplyQuery.UseVisualStyleBackColor = True
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(3, 64)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(38, 13)
        Me.Label15.TabIndex = 36
        Me.Label15.Text = "Query:"
        '
        'txtQuery
        '
        Me.txtQuery.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtQuery.Location = New System.Drawing.Point(75, 61)
        Me.txtQuery.Multiline = True
        Me.txtQuery.Name = "txtQuery"
        Me.txtQuery.Size = New System.Drawing.Size(588, 46)
        Me.txtQuery.TabIndex = 35
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(3, 37)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(66, 13)
        Me.Label14.TabIndex = 34
        Me.Label14.Text = "Select table:"
        '
        'cmbSelectTable
        '
        Me.cmbSelectTable.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbSelectTable.FormattingEnabled = True
        Me.cmbSelectTable.Location = New System.Drawing.Point(75, 34)
        Me.cmbSelectTable.Name = "cmbSelectTable"
        Me.cmbSelectTable.Size = New System.Drawing.Size(588, 21)
        Me.cmbSelectTable.TabIndex = 33
        '
        'btnShowTableProperties
        '
        Me.btnShowTableProperties.Location = New System.Drawing.Point(422, 6)
        Me.btnShowTableProperties.Name = "btnShowTableProperties"
        Me.btnShowTableProperties.Size = New System.Drawing.Size(136, 22)
        Me.btnShowTableProperties.TabIndex = 32
        Me.btnShowTableProperties.Text = "Show Table Properties"
        Me.btnShowTableProperties.UseVisualStyleBackColor = True
        '
        'btnSaveTableDefinition
        '
        Me.btnSaveTableDefinition.Location = New System.Drawing.Point(291, 6)
        Me.btnSaveTableDefinition.Name = "btnSaveTableDefinition"
        Me.btnSaveTableDefinition.Size = New System.Drawing.Size(125, 22)
        Me.btnSaveTableDefinition.TabIndex = 31
        Me.btnSaveTableDefinition.Text = "Save Table Definition"
        Me.btnSaveTableDefinition.UseVisualStyleBackColor = True
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(202, 6)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(83, 22)
        Me.btnImport.TabIndex = 30
        Me.btnImport.Text = "Import Data"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'btnExport
        '
        Me.btnExport.Location = New System.Drawing.Point(110, 6)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(86, 22)
        Me.btnExport.TabIndex = 29
        Me.btnExport.Text = "Export Data"
        Me.btnExport.UseVisualStyleBackColor = True
        '
        'btnSaveChanges
        '
        Me.btnSaveChanges.Location = New System.Drawing.Point(6, 6)
        Me.btnSaveChanges.Name = "btnSaveChanges"
        Me.btnSaveChanges.Size = New System.Drawing.Size(98, 22)
        Me.btnSaveChanges.TabIndex = 28
        Me.btnSaveChanges.Text = "Save Changes"
        Me.btnSaveChanges.UseVisualStyleBackColor = True
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.DataGridView2)
        Me.TabPage4.Controls.Add(Me.cmbRestrictions)
        Me.TabPage4.Controls.Add(Me.btnViewSchema)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(696, 425)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Schema"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'DataGridView2
        '
        Me.DataGridView2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Location = New System.Drawing.Point(6, 34)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.Size = New System.Drawing.Size(667, 393)
        Me.DataGridView2.TabIndex = 39
        '
        'cmbRestrictions
        '
        Me.cmbRestrictions.FormattingEnabled = True
        Me.cmbRestrictions.Location = New System.Drawing.Point(98, 7)
        Me.cmbRestrictions.Name = "cmbRestrictions"
        Me.cmbRestrictions.Size = New System.Drawing.Size(280, 21)
        Me.cmbRestrictions.TabIndex = 38
        '
        'btnViewSchema
        '
        Me.btnViewSchema.Location = New System.Drawing.Point(6, 6)
        Me.btnViewSchema.Name = "btnViewSchema"
        Me.btnViewSchema.Size = New System.Drawing.Size(86, 22)
        Me.btnViewSchema.TabIndex = 37
        Me.btnViewSchema.Text = "View Schema:"
        Me.btnViewSchema.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'btnAndorville
        '
        Me.btnAndorville.BackgroundImage = Global.ADVL_Database_Tools_1.My.Resources.Resources.Andorville_16May16_TM_Crop_Grey
        Me.btnAndorville.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnAndorville.Font = New System.Drawing.Font("Harlow Solid Italic", 14.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAndorville.Location = New System.Drawing.Point(5, 5)
        Me.btnAndorville.Name = "btnAndorville"
        Me.btnAndorville.Size = New System.Drawing.Size(118, 29)
        Me.btnAndorville.TabIndex = 50
        Me.btnAndorville.UseVisualStyleBackColor = True
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(728, 503)
        Me.Controls.Add(Me.btnAndorville)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnAppInfo)
        Me.Controls.Add(Me.btnMessages)
        Me.Controls.Add(Me.btnOnline)
        Me.Controls.Add(Me.btnExit)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Main"
        Me.Text = "Database Tools V1-0"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents btnOnline As Button
    Friend WithEvents btnMessages As Button
    Friend WithEvents btnAppInfo As Button
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents txtLastUsed As TextBox
    Friend WithEvents Label11 As Label
    Friend WithEvents txtCreationDate As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents txtDataLocationPath As TextBox
    Friend WithEvents txtDataLocationType As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents txtSettingsLocationPath As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents txtSettingsLocationType As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents txtProjectType As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txtProjectDescription As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents txtProjectName As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents btnProject As Button
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents Timer1 As Timer
    Friend WithEvents btnAndorville As Button
    Friend WithEvents cmbDatabaseType As ComboBox
    Friend WithEvents Label13 As Label
    Friend WithEvents Label12 As Label
    Friend WithEvents lstFields As ListBox
    Friend WithEvents Label2 As Label
    Friend WithEvents lstTables As ListBox
    Friend WithEvents btnDatabase As Button
    Friend WithEvents txtDatabase As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnSaveDef As Button
    Friend WithEvents btnCreateNew As Button
    Friend WithEvents btnModify As Button
    Friend WithEvents btnSqlCommand As Button
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents TabPage4 As TabPage
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents btnApplyQuery As Button
    Friend WithEvents Label15 As Label
    Friend WithEvents txtQuery As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents cmbSelectTable As ComboBox
    Friend WithEvents btnShowTableProperties As Button
    Friend WithEvents btnSaveTableDefinition As Button
    Friend WithEvents btnImport As Button
    Friend WithEvents btnExport As Button
    Friend WithEvents btnSaveChanges As Button
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents cmbRestrictions As ComboBox
    Friend WithEvents btnViewSchema As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents btnCreateNewTable As Button
    Friend WithEvents btnUpdateTables As Button
End Class
