<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmModifyDatabase
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
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.btnFindTableDef = New System.Windows.Forms.Button()
        Me.txtTableDefFileName = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.btnCreateTable = New System.Windows.Forms.Button()
        Me.txtNewTableName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.btnAutoNameFK = New System.Windows.Forms.Button()
        Me.btnDeleteRelationship = New System.Windows.Forms.Button()
        Me.btnCreateRelationship = New System.Windows.Forms.Button()
        Me.cmbRelatedColumn = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmbRelatedTable = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cmbFKColumn = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtNewForeignKeyName = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.cmbFKTable = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.chkUnique = New System.Windows.Forms.CheckBox()
        Me.btnAutoNameIndex = New System.Windows.Forms.Button()
        Me.cmbIndexOptions = New System.Windows.Forms.ComboBox()
        Me.lstColumns = New System.Windows.Forms.ListBox()
        Me.btnDeleteIndex = New System.Windows.Forms.Button()
        Me.btnCreateIndex = New System.Windows.Forms.Button()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.cmbNewIndexTable = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtIndexName = New System.Windows.Forms.TextBox()
        Me.DataGridView4 = New System.Windows.Forms.DataGridView()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.btnUpdateDescriptions = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbDescrSelectTable = New System.Windows.Forms.ComboBox()
        Me.DataGridView3 = New System.Windows.Forms.DataGridView()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.btnRenameTable = New System.Windows.Forms.Button()
        Me.txtNewTableName2 = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.btnAddColumn = New System.Windows.Forms.Button()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.cmbPrimaryKey = New System.Windows.Forms.ComboBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.cmbAutoIncrement = New System.Windows.Forms.ComboBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.cmbNull = New System.Windows.Forms.ComboBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.txtScale = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.txtPrecision = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txtSize = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.cmbColumnType = New System.Windows.Forms.ComboBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtColumnName = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.btnDeleteTable = New System.Windows.Forms.Button()
        Me.btnRenameColumn = New System.Windows.Forms.Button()
        Me.txtNewColumnName = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmbUtilitiesSelectColumn = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmbUtilitiesSelectTable = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.btnMoveUp = New System.Windows.Forms.Button()
        Me.btnMoveDown = New System.Windows.Forms.Button()
        Me.btnInsertAbove = New System.Windows.Forms.Button()
        Me.btnInsertBelow = New System.Windows.Forms.Button()
        Me.btnDeleteRow = New System.Windows.Forms.Button()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage5.SuspendLayout()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(935, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 10
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage5)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Location = New System.Drawing.Point(12, 40)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(987, 579)
        Me.TabControl1.TabIndex = 11
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.btnDeleteRow)
        Me.TabPage1.Controls.Add(Me.btnInsertBelow)
        Me.TabPage1.Controls.Add(Me.btnInsertAbove)
        Me.TabPage1.Controls.Add(Me.btnMoveDown)
        Me.TabPage1.Controls.Add(Me.btnMoveUp)
        Me.TabPage1.Controls.Add(Me.btnFindTableDef)
        Me.TabPage1.Controls.Add(Me.txtTableDefFileName)
        Me.TabPage1.Controls.Add(Me.Label23)
        Me.TabPage1.Controls.Add(Me.btnCreateTable)
        Me.TabPage1.Controls.Add(Me.txtNewTableName)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.DataGridView1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(979, 553)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Create New Table"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'btnFindTableDef
        '
        Me.btnFindTableDef.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindTableDef.Location = New System.Drawing.Point(918, 5)
        Me.btnFindTableDef.Name = "btnFindTableDef"
        Me.btnFindTableDef.Size = New System.Drawing.Size(55, 22)
        Me.btnFindTableDef.TabIndex = 12
        Me.btnFindTableDef.Text = "Find"
        Me.btnFindTableDef.UseVisualStyleBackColor = True
        '
        'txtTableDefFileName
        '
        Me.txtTableDefFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTableDefFileName.Location = New System.Drawing.Point(629, 6)
        Me.txtTableDefFileName.Name = "txtTableDefFileName"
        Me.txtTableDefFileName.Size = New System.Drawing.Size(283, 20)
        Me.txtTableDefFileName.TabIndex = 5
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(541, 9)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(82, 13)
        Me.Label23.TabIndex = 4
        Me.Label23.Text = "Table definition:"
        '
        'btnCreateTable
        '
        Me.btnCreateTable.Location = New System.Drawing.Point(453, 5)
        Me.btnCreateTable.Name = "btnCreateTable"
        Me.btnCreateTable.Size = New System.Drawing.Size(64, 22)
        Me.btnCreateTable.TabIndex = 3
        Me.btnCreateTable.Text = "Create"
        Me.btnCreateTable.UseVisualStyleBackColor = True
        '
        'txtNewTableName
        '
        Me.txtNewTableName.Location = New System.Drawing.Point(80, 6)
        Me.txtNewTableName.Name = "txtNewTableName"
        Me.txtNewTableName.Size = New System.Drawing.Size(367, 20)
        Me.txtNewTableName.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Table Name:"
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(6, 61)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(967, 486)
        Me.DataGridView1.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.btnAutoNameFK)
        Me.TabPage2.Controls.Add(Me.btnDeleteRelationship)
        Me.TabPage2.Controls.Add(Me.btnCreateRelationship)
        Me.TabPage2.Controls.Add(Me.cmbRelatedColumn)
        Me.TabPage2.Controls.Add(Me.Label6)
        Me.TabPage2.Controls.Add(Me.cmbRelatedTable)
        Me.TabPage2.Controls.Add(Me.Label10)
        Me.TabPage2.Controls.Add(Me.cmbFKColumn)
        Me.TabPage2.Controls.Add(Me.Label9)
        Me.TabPage2.Controls.Add(Me.txtNewForeignKeyName)
        Me.TabPage2.Controls.Add(Me.Label8)
        Me.TabPage2.Controls.Add(Me.cmbFKTable)
        Me.TabPage2.Controls.Add(Me.Label7)
        Me.TabPage2.Controls.Add(Me.DataGridView2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1091, 553)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Relationships"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'btnAutoNameFK
        '
        Me.btnAutoNameFK.Location = New System.Drawing.Point(301, 5)
        Me.btnAutoNameFK.Name = "btnAutoNameFK"
        Me.btnAutoNameFK.Size = New System.Drawing.Size(16, 22)
        Me.btnAutoNameFK.TabIndex = 5
        Me.btnAutoNameFK.Text = "<"
        Me.btnAutoNameFK.UseVisualStyleBackColor = True
        '
        'btnDeleteRelationship
        '
        Me.btnDeleteRelationship.Location = New System.Drawing.Point(6, 31)
        Me.btnDeleteRelationship.Name = "btnDeleteRelationship"
        Me.btnDeleteRelationship.Size = New System.Drawing.Size(124, 22)
        Me.btnDeleteRelationship.TabIndex = 5
        Me.btnDeleteRelationship.Text = "Delete Relationship"
        Me.btnDeleteRelationship.UseVisualStyleBackColor = True
        '
        'btnCreateRelationship
        '
        Me.btnCreateRelationship.Location = New System.Drawing.Point(136, 31)
        Me.btnCreateRelationship.Name = "btnCreateRelationship"
        Me.btnCreateRelationship.Size = New System.Drawing.Size(121, 22)
        Me.btnCreateRelationship.TabIndex = 5
        Me.btnCreateRelationship.Text = "Create Relationship"
        Me.btnCreateRelationship.UseVisualStyleBackColor = True
        '
        'cmbRelatedColumn
        '
        Me.cmbRelatedColumn.FormattingEnabled = True
        Me.cmbRelatedColumn.Location = New System.Drawing.Point(656, 33)
        Me.cmbRelatedColumn.Name = "cmbRelatedColumn"
        Me.cmbRelatedColumn.Size = New System.Drawing.Size(161, 21)
        Me.cmbRelatedColumn.TabIndex = 11
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(562, 36)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(85, 13)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Related Column:"
        '
        'cmbRelatedTable
        '
        Me.cmbRelatedTable.FormattingEnabled = True
        Me.cmbRelatedTable.Location = New System.Drawing.Point(375, 33)
        Me.cmbRelatedTable.Name = "cmbRelatedTable"
        Me.cmbRelatedTable.Size = New System.Drawing.Size(161, 21)
        Me.cmbRelatedTable.TabIndex = 9
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(292, 36)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(77, 13)
        Me.Label10.TabIndex = 8
        Me.Label10.Text = "Related Table:"
        '
        'cmbFKColumn
        '
        Me.cmbFKColumn.FormattingEnabled = True
        Me.cmbFKColumn.Location = New System.Drawing.Point(656, 6)
        Me.cmbFKColumn.Name = "cmbFKColumn"
        Me.cmbFKColumn.Size = New System.Drawing.Size(161, 21)
        Me.cmbFKColumn.TabIndex = 7
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(602, 9)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(45, 13)
        Me.Label9.TabIndex = 6
        Me.Label9.Text = "Column:"
        '
        'txtNewForeignKeyName
        '
        Me.txtNewForeignKeyName.Location = New System.Drawing.Point(136, 6)
        Me.txtNewForeignKeyName.Name = "txtNewForeignKeyName"
        Me.txtNewForeignKeyName.Size = New System.Drawing.Size(161, 20)
        Me.txtNewForeignKeyName.TabIndex = 5
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 9)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(122, 13)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "New Foreign Key Name:"
        '
        'cmbFKTable
        '
        Me.cmbFKTable.FormattingEnabled = True
        Me.cmbFKTable.Location = New System.Drawing.Point(375, 6)
        Me.cmbFKTable.Name = "cmbFKTable"
        Me.cmbFKTable.Size = New System.Drawing.Size(161, 21)
        Me.cmbFKTable.TabIndex = 3
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(332, 9)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(37, 13)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "Table:"
        '
        'DataGridView2
        '
        Me.DataGridView2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Location = New System.Drawing.Point(0, 60)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.Size = New System.Drawing.Size(960, 490)
        Me.DataGridView2.TabIndex = 0
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.chkUnique)
        Me.TabPage5.Controls.Add(Me.btnAutoNameIndex)
        Me.TabPage5.Controls.Add(Me.cmbIndexOptions)
        Me.TabPage5.Controls.Add(Me.lstColumns)
        Me.TabPage5.Controls.Add(Me.btnDeleteIndex)
        Me.TabPage5.Controls.Add(Me.btnCreateIndex)
        Me.TabPage5.Controls.Add(Me.Label13)
        Me.TabPage5.Controls.Add(Me.cmbNewIndexTable)
        Me.TabPage5.Controls.Add(Me.Label12)
        Me.TabPage5.Controls.Add(Me.txtIndexName)
        Me.TabPage5.Controls.Add(Me.DataGridView4)
        Me.TabPage5.Controls.Add(Me.Label11)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Size = New System.Drawing.Size(1091, 553)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Indexes"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'chkUnique
        '
        Me.chkUnique.AutoSize = True
        Me.chkUnique.Location = New System.Drawing.Point(101, 60)
        Me.chkUnique.Name = "chkUnique"
        Me.chkUnique.Size = New System.Drawing.Size(60, 17)
        Me.chkUnique.TabIndex = 11
        Me.chkUnique.Text = "Unique"
        Me.chkUnique.UseVisualStyleBackColor = True
        '
        'btnAutoNameIndex
        '
        Me.btnAutoNameIndex.Location = New System.Drawing.Point(290, 4)
        Me.btnAutoNameIndex.Name = "btnAutoNameIndex"
        Me.btnAutoNameIndex.Size = New System.Drawing.Size(16, 22)
        Me.btnAutoNameIndex.TabIndex = 10
        Me.btnAutoNameIndex.Text = "<"
        Me.btnAutoNameIndex.UseVisualStyleBackColor = True
        '
        'cmbIndexOptions
        '
        Me.cmbIndexOptions.FormattingEnabled = True
        Me.cmbIndexOptions.Location = New System.Drawing.Point(167, 60)
        Me.cmbIndexOptions.Name = "cmbIndexOptions"
        Me.cmbIndexOptions.Size = New System.Drawing.Size(166, 21)
        Me.cmbIndexOptions.TabIndex = 9
        '
        'lstColumns
        '
        Me.lstColumns.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstColumns.FormattingEnabled = True
        Me.lstColumns.Location = New System.Drawing.Point(381, 8)
        Me.lstColumns.Name = "lstColumns"
        Me.lstColumns.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.lstColumns.Size = New System.Drawing.Size(579, 108)
        Me.lstColumns.TabIndex = 8
        '
        'btnDeleteIndex
        '
        Me.btnDeleteIndex.Location = New System.Drawing.Point(6, 102)
        Me.btnDeleteIndex.Name = "btnDeleteIndex"
        Me.btnDeleteIndex.Size = New System.Drawing.Size(97, 22)
        Me.btnDeleteIndex.TabIndex = 7
        Me.btnDeleteIndex.Text = "Delete Index"
        Me.btnDeleteIndex.UseVisualStyleBackColor = True
        '
        'btnCreateIndex
        '
        Me.btnCreateIndex.Location = New System.Drawing.Point(6, 58)
        Me.btnCreateIndex.Name = "btnCreateIndex"
        Me.btnCreateIndex.Size = New System.Drawing.Size(89, 22)
        Me.btnCreateIndex.TabIndex = 5
        Me.btnCreateIndex.Text = "Create Index"
        Me.btnCreateIndex.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(319, 8)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(56, 13)
        Me.Label13.TabIndex = 5
        Me.Label13.Text = "Column(s):"
        '
        'cmbNewIndexTable
        '
        Me.cmbNewIndexTable.FormattingEnabled = True
        Me.cmbNewIndexTable.Location = New System.Drawing.Point(101, 31)
        Me.cmbNewIndexTable.Name = "cmbNewIndexTable"
        Me.cmbNewIndexTable.Size = New System.Drawing.Size(183, 21)
        Me.cmbNewIndexTable.TabIndex = 4
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(58, 34)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(37, 13)
        Me.Label12.TabIndex = 3
        Me.Label12.Text = "Table:"
        '
        'txtIndexName
        '
        Me.txtIndexName.Location = New System.Drawing.Point(101, 5)
        Me.txtIndexName.Name = "txtIndexName"
        Me.txtIndexName.Size = New System.Drawing.Size(183, 20)
        Me.txtIndexName.TabIndex = 2
        '
        'DataGridView4
        '
        Me.DataGridView4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView4.Location = New System.Drawing.Point(6, 130)
        Me.DataGridView4.Name = "DataGridView4"
        Me.DataGridView4.Size = New System.Drawing.Size(954, 420)
        Me.DataGridView4.TabIndex = 1
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 8)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(92, 13)
        Me.Label11.TabIndex = 0
        Me.Label11.Text = "New Index Name:"
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.btnUpdateDescriptions)
        Me.TabPage3.Controls.Add(Me.Label2)
        Me.TabPage3.Controls.Add(Me.cmbDescrSelectTable)
        Me.TabPage3.Controls.Add(Me.DataGridView3)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(1091, 553)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Column Descriptions"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'btnUpdateDescriptions
        '
        Me.btnUpdateDescriptions.Location = New System.Drawing.Point(547, 3)
        Me.btnUpdateDescriptions.Name = "btnUpdateDescriptions"
        Me.btnUpdateDescriptions.Size = New System.Drawing.Size(128, 22)
        Me.btnUpdateDescriptions.TabIndex = 5
        Me.btnUpdateDescriptions.Text = "Update Descriptions"
        Me.btnUpdateDescriptions.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Select Table:"
        '
        'cmbDescrSelectTable
        '
        Me.cmbDescrSelectTable.FormattingEnabled = True
        Me.cmbDescrSelectTable.Location = New System.Drawing.Point(79, 3)
        Me.cmbDescrSelectTable.Name = "cmbDescrSelectTable"
        Me.cmbDescrSelectTable.Size = New System.Drawing.Size(424, 21)
        Me.cmbDescrSelectTable.TabIndex = 1
        '
        'DataGridView3
        '
        Me.DataGridView3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.Location = New System.Drawing.Point(3, 30)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.Size = New System.Drawing.Size(955, 602)
        Me.DataGridView3.TabIndex = 0
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.btnRenameTable)
        Me.TabPage4.Controls.Add(Me.txtNewTableName2)
        Me.TabPage4.Controls.Add(Me.Label24)
        Me.TabPage4.Controls.Add(Me.btnAddColumn)
        Me.TabPage4.Controls.Add(Me.txtDescription)
        Me.TabPage4.Controls.Add(Me.Label22)
        Me.TabPage4.Controls.Add(Me.cmbPrimaryKey)
        Me.TabPage4.Controls.Add(Me.Label21)
        Me.TabPage4.Controls.Add(Me.cmbAutoIncrement)
        Me.TabPage4.Controls.Add(Me.Label20)
        Me.TabPage4.Controls.Add(Me.cmbNull)
        Me.TabPage4.Controls.Add(Me.Label19)
        Me.TabPage4.Controls.Add(Me.txtScale)
        Me.TabPage4.Controls.Add(Me.Label18)
        Me.TabPage4.Controls.Add(Me.txtPrecision)
        Me.TabPage4.Controls.Add(Me.Label17)
        Me.TabPage4.Controls.Add(Me.txtSize)
        Me.TabPage4.Controls.Add(Me.Label16)
        Me.TabPage4.Controls.Add(Me.cmbColumnType)
        Me.TabPage4.Controls.Add(Me.Label15)
        Me.TabPage4.Controls.Add(Me.txtColumnName)
        Me.TabPage4.Controls.Add(Me.Label14)
        Me.TabPage4.Controls.Add(Me.btnDeleteTable)
        Me.TabPage4.Controls.Add(Me.btnRenameColumn)
        Me.TabPage4.Controls.Add(Me.txtNewColumnName)
        Me.TabPage4.Controls.Add(Me.Label5)
        Me.TabPage4.Controls.Add(Me.cmbUtilitiesSelectColumn)
        Me.TabPage4.Controls.Add(Me.Label4)
        Me.TabPage4.Controls.Add(Me.cmbUtilitiesSelectTable)
        Me.TabPage4.Controls.Add(Me.Label3)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(1091, 553)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Miscellaneous"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'btnRenameTable
        '
        Me.btnRenameTable.Location = New System.Drawing.Point(574, 37)
        Me.btnRenameTable.Name = "btnRenameTable"
        Me.btnRenameTable.Size = New System.Drawing.Size(105, 22)
        Me.btnRenameTable.TabIndex = 29
        Me.btnRenameTable.Text = "Rename Table"
        Me.btnRenameTable.UseVisualStyleBackColor = True
        '
        'txtNewTableName2
        '
        Me.txtNewTableName2.Location = New System.Drawing.Point(116, 37)
        Me.txtNewTableName2.Name = "txtNewTableName2"
        Me.txtNewTableName2.Size = New System.Drawing.Size(442, 20)
        Me.txtNewTableName2.TabIndex = 28
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(9, 41)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(93, 13)
        Me.Label24.TabIndex = 27
        Me.Label24.Text = "New Table Name:"
        '
        'btnAddColumn
        '
        Me.btnAddColumn.Location = New System.Drawing.Point(607, 182)
        Me.btnAddColumn.Name = "btnAddColumn"
        Me.btnAddColumn.Size = New System.Drawing.Size(79, 22)
        Me.btnAddColumn.TabIndex = 26
        Me.btnAddColumn.Text = "Add Column"
        Me.btnAddColumn.UseVisualStyleBackColor = True
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(78, 184)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(523, 20)
        Me.txtDescription.TabIndex = 25
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(9, 187)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(63, 13)
        Me.Label22.TabIndex = 24
        Me.Label22.Text = "Description:"
        '
        'cmbPrimaryKey
        '
        Me.cmbPrimaryKey.FormattingEnabled = True
        Me.cmbPrimaryKey.Location = New System.Drawing.Point(607, 156)
        Me.cmbPrimaryKey.Name = "cmbPrimaryKey"
        Me.cmbPrimaryKey.Size = New System.Drawing.Size(79, 21)
        Me.cmbPrimaryKey.TabIndex = 23
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(607, 141)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(65, 13)
        Me.Label21.TabIndex = 22
        Me.Label21.Text = "Primary Key:"
        '
        'cmbAutoIncrement
        '
        Me.cmbAutoIncrement.FormattingEnabled = True
        Me.cmbAutoIncrement.Location = New System.Drawing.Point(522, 156)
        Me.cmbAutoIncrement.Name = "cmbAutoIncrement"
        Me.cmbAutoIncrement.Size = New System.Drawing.Size(79, 21)
        Me.cmbAutoIncrement.TabIndex = 21
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(519, 141)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(82, 13)
        Me.Label20.TabIndex = 20
        Me.Label20.Text = "Auto Increment:"
        '
        'cmbNull
        '
        Me.cmbNull.FormattingEnabled = True
        Me.cmbNull.Location = New System.Drawing.Point(445, 157)
        Me.cmbNull.Name = "cmbNull"
        Me.cmbNull.Size = New System.Drawing.Size(70, 21)
        Me.cmbNull.TabIndex = 19
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(442, 142)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(71, 13)
        Me.Label19.TabIndex = 18
        Me.Label19.Text = "Null/Not Null:"
        '
        'txtScale
        '
        Me.txtScale.Location = New System.Drawing.Point(388, 158)
        Me.txtScale.Name = "txtScale"
        Me.txtScale.Size = New System.Drawing.Size(51, 20)
        Me.txtScale.TabIndex = 17
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(387, 142)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(37, 13)
        Me.Label18.TabIndex = 16
        Me.Label18.Text = "Scale:"
        '
        'txtPrecision
        '
        Me.txtPrecision.Location = New System.Drawing.Point(331, 158)
        Me.txtPrecision.Name = "txtPrecision"
        Me.txtPrecision.Size = New System.Drawing.Size(51, 20)
        Me.txtPrecision.TabIndex = 15
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(328, 141)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(53, 13)
        Me.Label17.TabIndex = 14
        Me.Label17.Text = "Precision:"
        '
        'txtSize
        '
        Me.txtSize.Location = New System.Drawing.Point(272, 157)
        Me.txtSize.Name = "txtSize"
        Me.txtSize.Size = New System.Drawing.Size(51, 20)
        Me.txtSize.TabIndex = 13
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(269, 141)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(30, 13)
        Me.Label16.TabIndex = 12
        Me.Label16.Text = "Size:"
        '
        'cmbColumnType
        '
        Me.cmbColumnType.FormattingEnabled = True
        Me.cmbColumnType.Location = New System.Drawing.Point(156, 157)
        Me.cmbColumnType.Name = "cmbColumnType"
        Me.cmbColumnType.Size = New System.Drawing.Size(107, 21)
        Me.cmbColumnType.TabIndex = 11
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(153, 141)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(34, 13)
        Me.Label15.TabIndex = 10
        Me.Label15.Text = "Type:"
        '
        'txtColumnName
        '
        Me.txtColumnName.Location = New System.Drawing.Point(12, 157)
        Me.txtColumnName.Name = "txtColumnName"
        Me.txtColumnName.Size = New System.Drawing.Size(135, 20)
        Me.txtColumnName.TabIndex = 9
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(9, 141)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(76, 13)
        Me.Label14.TabIndex = 8
        Me.Label14.Text = "Column Name:"
        '
        'btnDeleteTable
        '
        Me.btnDeleteTable.Location = New System.Drawing.Point(574, 10)
        Me.btnDeleteTable.Name = "btnDeleteTable"
        Me.btnDeleteTable.Size = New System.Drawing.Size(105, 22)
        Me.btnDeleteTable.TabIndex = 7
        Me.btnDeleteTable.Text = "Delete Table"
        Me.btnDeleteTable.UseVisualStyleBackColor = True
        '
        'btnRenameColumn
        '
        Me.btnRenameColumn.Location = New System.Drawing.Point(574, 89)
        Me.btnRenameColumn.Name = "btnRenameColumn"
        Me.btnRenameColumn.Size = New System.Drawing.Size(105, 22)
        Me.btnRenameColumn.TabIndex = 6
        Me.btnRenameColumn.Text = "Rename Column"
        Me.btnRenameColumn.UseVisualStyleBackColor = True
        '
        'txtNewColumnName
        '
        Me.txtNewColumnName.Location = New System.Drawing.Point(116, 90)
        Me.txtNewColumnName.Name = "txtNewColumnName"
        Me.txtNewColumnName.Size = New System.Drawing.Size(442, 20)
        Me.txtNewColumnName.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 93)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(101, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "New Column Name:"
        '
        'cmbUtilitiesSelectColumn
        '
        Me.cmbUtilitiesSelectColumn.FormattingEnabled = True
        Me.cmbUtilitiesSelectColumn.Location = New System.Drawing.Point(116, 63)
        Me.cmbUtilitiesSelectColumn.Name = "cmbUtilitiesSelectColumn"
        Me.cmbUtilitiesSelectColumn.Size = New System.Drawing.Size(442, 21)
        Me.cmbUtilitiesSelectColumn.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 66)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(78, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Select Column:"
        '
        'cmbUtilitiesSelectTable
        '
        Me.cmbUtilitiesSelectTable.FormattingEnabled = True
        Me.cmbUtilitiesSelectTable.Location = New System.Drawing.Point(116, 10)
        Me.cmbUtilitiesSelectTable.Name = "cmbUtilitiesSelectTable"
        Me.cmbUtilitiesSelectTable.Size = New System.Drawing.Size(442, 21)
        Me.cmbUtilitiesSelectTable.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 13)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(70, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Select Table:"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'btnMoveUp
        '
        Me.btnMoveUp.Location = New System.Drawing.Point(9, 33)
        Me.btnMoveUp.Name = "btnMoveUp"
        Me.btnMoveUp.Size = New System.Drawing.Size(64, 22)
        Me.btnMoveUp.TabIndex = 13
        Me.btnMoveUp.Text = "Move Up"
        Me.btnMoveUp.UseVisualStyleBackColor = True
        '
        'btnMoveDown
        '
        Me.btnMoveDown.Location = New System.Drawing.Point(79, 33)
        Me.btnMoveDown.Name = "btnMoveDown"
        Me.btnMoveDown.Size = New System.Drawing.Size(80, 22)
        Me.btnMoveDown.TabIndex = 14
        Me.btnMoveDown.Text = "Move Down"
        Me.btnMoveDown.UseVisualStyleBackColor = True
        '
        'btnInsertAbove
        '
        Me.btnInsertAbove.Location = New System.Drawing.Point(165, 33)
        Me.btnInsertAbove.Name = "btnInsertAbove"
        Me.btnInsertAbove.Size = New System.Drawing.Size(80, 22)
        Me.btnInsertAbove.TabIndex = 15
        Me.btnInsertAbove.Text = "Insert Above"
        Me.btnInsertAbove.UseVisualStyleBackColor = True
        '
        'btnInsertBelow
        '
        Me.btnInsertBelow.Location = New System.Drawing.Point(251, 33)
        Me.btnInsertBelow.Name = "btnInsertBelow"
        Me.btnInsertBelow.Size = New System.Drawing.Size(80, 22)
        Me.btnInsertBelow.TabIndex = 16
        Me.btnInsertBelow.Text = "Insert Below"
        Me.btnInsertBelow.UseVisualStyleBackColor = True
        '
        'btnDeleteRow
        '
        Me.btnDeleteRow.Location = New System.Drawing.Point(337, 33)
        Me.btnDeleteRow.Name = "btnDeleteRow"
        Me.btnDeleteRow.Size = New System.Drawing.Size(80, 22)
        Me.btnDeleteRow.TabIndex = 17
        Me.btnDeleteRow.Text = "Delete Row"
        Me.btnDeleteRow.UseVisualStyleBackColor = True
        '
        'frmModifyDatabase
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1011, 631)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmModifyDatabase"
        Me.Text = "Modify Database"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage5.ResumeLayout(False)
        Me.TabPage5.PerformLayout()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents btnCreateTable As Button
    Friend WithEvents txtNewTableName As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents btnAutoNameFK As Button
    Friend WithEvents btnDeleteRelationship As Button
    Friend WithEvents btnCreateRelationship As Button
    Friend WithEvents cmbRelatedColumn As ComboBox
    Friend WithEvents Label6 As Label
    Friend WithEvents cmbRelatedTable As ComboBox
    Friend WithEvents Label10 As Label
    Friend WithEvents cmbFKColumn As ComboBox
    Friend WithEvents Label9 As Label
    Friend WithEvents txtNewForeignKeyName As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents cmbFKTable As ComboBox
    Friend WithEvents Label7 As Label
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents TabPage5 As TabPage
    Friend WithEvents chkUnique As CheckBox
    Friend WithEvents btnAutoNameIndex As Button
    Friend WithEvents cmbIndexOptions As ComboBox
    Friend WithEvents lstColumns As ListBox
    Friend WithEvents btnDeleteIndex As Button
    Friend WithEvents btnCreateIndex As Button
    Friend WithEvents Label13 As Label
    Friend WithEvents cmbNewIndexTable As ComboBox
    Friend WithEvents Label12 As Label
    Friend WithEvents txtIndexName As TextBox
    Friend WithEvents DataGridView4 As DataGridView
    Friend WithEvents Label11 As Label
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents btnUpdateDescriptions As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents cmbDescrSelectTable As ComboBox
    Friend WithEvents DataGridView3 As DataGridView
    Friend WithEvents TabPage4 As TabPage
    Friend WithEvents btnAddColumn As Button
    Friend WithEvents txtDescription As TextBox
    Friend WithEvents Label22 As Label
    Friend WithEvents cmbPrimaryKey As ComboBox
    Friend WithEvents Label21 As Label
    Friend WithEvents cmbAutoIncrement As ComboBox
    Friend WithEvents Label20 As Label
    Friend WithEvents cmbNull As ComboBox
    Friend WithEvents Label19 As Label
    Friend WithEvents txtScale As TextBox
    Friend WithEvents Label18 As Label
    Friend WithEvents txtPrecision As TextBox
    Friend WithEvents Label17 As Label
    Friend WithEvents txtSize As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents cmbColumnType As ComboBox
    Friend WithEvents Label15 As Label
    Friend WithEvents txtColumnName As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents btnDeleteTable As Button
    Friend WithEvents btnRenameColumn As Button
    Friend WithEvents txtNewColumnName As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents cmbUtilitiesSelectColumn As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents cmbUtilitiesSelectTable As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents btnFindTableDef As Button
    Friend WithEvents txtTableDefFileName As TextBox
    Friend WithEvents Label23 As Label
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents btnRenameTable As Button
    Friend WithEvents txtNewTableName2 As TextBox
    Friend WithEvents Label24 As Label
    Friend WithEvents btnInsertBelow As Button
    Friend WithEvents btnInsertAbove As Button
    Friend WithEvents btnMoveDown As Button
    Friend WithEvents btnMoveUp As Button
    Friend WithEvents btnDeleteRow As Button
End Class
