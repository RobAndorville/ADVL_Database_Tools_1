Public Class frmModifyDatabase
    'Form used to modify the database.

#Region " Coding Notes - Notes on the code used in this class." '------------------------------------------------------------------------------------------------------------------------------

    'Add References:
    'Project \ Add Reference...

    'Assemblies \ Extensions
    'Microsoft.Office.Interop.Access.Dao 15.0.0.0

    'COM \ Type Libraries
    'Microsoft ActiveX Data Objects 6.1 Library
    'Microsoft ADO Ext 6.0 for DDL and Security

    'List of Access reserved words:
    'https://support.microsoft.com/en-au/kb/286335
    'https://support.office.com/en-us/article/Access-2007-reserved-words-and-symbols-e33eb3a9-8baa-4335-9f57-da237c63eabe


#End Region 'Coding Notes ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Variable Declarations - All the variables and class objects used in this form and this application." '-------------------------------------------------------------------------------------------------

    Dim tableDefFileName As String 'The Database deinition file name.
    Dim tableDefXDoc As System.Xml.Linq.XDocument 'The database definition XDocument.
    Dim WithEvents Zip As ADVL_Utilities_Library_1.ZipComp

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form and this application" '------------------------------------------------------------------------------------------------------------

#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Process XML files - Read and write XML files." '-------------------------------------------------------------------------------------------------------------------------------------

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <!---->
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Main.Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"

        If Main.Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Main.Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore form position and size:
            If Settings.<FormSettings>.<Left>.Value = Nothing Then
                'Form setting not saved.
            Else
                Me.Left = Settings.<FormSettings>.<Left>.Value
            End If

            If Settings.<FormSettings>.<Top>.Value = Nothing Then
                'Form setting not saved.
            Else
                Me.Top = Settings.<FormSettings>.<Top>.Value
            End If

            If Settings.<FormSettings>.<Height>.Value = Nothing Then
                'Form setting not saved.
            Else
                Me.Height = Settings.<FormSettings>.<Height>.Value
            End If

            If Settings.<FormSettings>.<Width>.Value = Nothing Then
                'Form setting not saved.
            Else
                Me.Width = Settings.<FormSettings>.<Width>.Value
            End If

            'Add code to read other saved setting here:

        End If
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message) 'Save the form settings before the form is minimised:
        If m.Msg = &H112 Then 'SysCommand
            If m.WParam.ToInt32 = &HF020 Then 'Form is being minimised
                SaveFormSettings()
            End If
        End If
        MyBase.WndProc(m)
    End Sub



#End Region 'Process XML Files ----------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Display Methods - Code used to display this form." '----------------------------------------------------------------------------------------------------------------------------

    Private Sub frmModifyDatabase_Load(sender As Object, e As EventArgs) Handles Me.Load
        RestoreFormSettings()   'Restore the form settings

        'Set up DataGridView1:
        Dim TextBoxCol0 As New DataGridViewTextBoxColumn
        DataGridView1.Columns.Add(TextBoxCol0)
        DataGridView1.Columns(0).HeaderText = "Column Name"
        DataGridView1.Columns(0).Width = 160

        'Dim TextBoxCol1 As New DataGridViewTextBoxColumn
        'DataGridView1.Columns.Add(TextBoxCol1)
        'DataGridView1.Columns(1).HeaderText = "Type Code"
        'DataGridView1.Columns(1).Width = 160



        Dim ComboBoxCol1 As New DataGridViewComboBoxColumn
        DataGridView1.Columns.Add(ComboBoxCol1)
        DataGridView1.Columns(1).HeaderText = "Type"
        DataGridView1.Columns(1).Width = 120
        'See Data Type schema for the list of data types.
        ComboBoxCol1.Items.Add("Short (Integer)")
        ComboBoxCol1.Items.Add("Long (Integer)")
        ComboBoxCol1.Items.Add("Single")
        ComboBoxCol1.Items.Add("Double")
        ComboBoxCol1.Items.Add("Currency")
        ComboBoxCol1.Items.Add("DateTime")
        ComboBoxCol1.Items.Add("Bit (Boolean)")
        ComboBoxCol1.Items.Add("Byte")
        ComboBoxCol1.Items.Add("GUID")
        ComboBoxCol1.Items.Add("BigBinary")
        ComboBoxCol1.Items.Add("LongBinary")
        ComboBoxCol1.Items.Add("VarBinary")
        ComboBoxCol1.Items.Add("LongText")
        ComboBoxCol1.Items.Add("VarChar")
        ComboBoxCol1.Items.Add("Decimal")

        'Dim ComboBoxCol2 As New DataGridViewComboBoxColumn
        'DataGridView1.Columns.Add(ComboBoxCol2)
        'DataGridView1.Columns(2).HeaderText = "Type"
        'DataGridView1.Columns(2).Width = 120
        ''See Data Type schema for the list of data types.
        'ComboBoxCol2.Items.Add("Short (Integer)")
        'ComboBoxCol2.Items.Add("Long (Integer)")
        'ComboBoxCol2.Items.Add("Single")
        'ComboBoxCol2.Items.Add("Double")
        'ComboBoxCol2.Items.Add("Currency")
        'ComboBoxCol2.Items.Add("DateTime")
        'ComboBoxCol2.Items.Add("Bit (Boolean)")
        'ComboBoxCol2.Items.Add("Byte")
        'ComboBoxCol2.Items.Add("GUID")
        'ComboBoxCol2.Items.Add("BigBinary")
        'ComboBoxCol2.Items.Add("LongBinary")
        'ComboBoxCol2.Items.Add("VarBinary")
        'ComboBoxCol2.Items.Add("LongText")
        'ComboBoxCol2.Items.Add("VarChar")
        'ComboBoxCol2.Items.Add("Decimal")

        'Short      2       5                               System.Int16
        'Long       3       10                              System.Int32
        'Single     4       7                               System.Single
        'Double     5       15                              System.Double
        'Currency   6       19                              System.Decimal
        'DateTime   7       8                               System.DateTime
        'Bit        11      2                               System.Boolean
        'Byte       17      3                               System.Byte
        'GUID       72      16                              System.Guid
        'BigBinary  204     4000                            System.Byte
        'LongBinary 205     1073741823                      System.Byte
        'VarBinary  204     510         max length          System.Byte
        'LongText   203     536870910                       System.String
        'VarChar    202     255         max length          System.String
        'Decimal    131     28          precision, scale    System.Decimal

        'http://www.vb6.us/tutorials/access-sql
        'BINARY BIT BYTE COUNTER CURRENCY DATETIME SINGLE DOUBLE SHORT LONG LONGTEXT LONGBINARY TEXT 

        'Create New Table data:
        Dim TextBoxCol2 As New DataGridViewTextBoxColumn
        DataGridView1.Columns.Add(TextBoxCol2)
        DataGridView1.Columns(2).HeaderText = "Size"
        Dim TextBoxCol3 As New DataGridViewTextBoxColumn
        DataGridView1.Columns.Add(TextBoxCol3)
        DataGridView1.Columns(3).HeaderText = "Precision"
        Dim TextBoxCol4 As New DataGridViewTextBoxColumn
        DataGridView1.Columns.Add(TextBoxCol4)
        DataGridView1.Columns(4).HeaderText = "Scale"
        Dim ComboBoxCol5 As New DataGridViewComboBoxColumn
        DataGridView1.Columns.Add(ComboBoxCol5)
        DataGridView1.Columns(5).HeaderText = "Null/Not Null"
        ComboBoxCol5.Items.Add("")
        ComboBoxCol5.Items.Add("Null")
        ComboBoxCol5.Items.Add("Not Null")
        Dim ComboBoxCol6 As New DataGridViewComboBoxColumn
        DataGridView1.Columns.Add(ComboBoxCol6)
        DataGridView1.Columns(6).HeaderText = "Auto Increment"
        ComboBoxCol6.Items.Add("")
        ComboBoxCol6.Items.Add("Auto Increment")
        Dim ComboBoxCol7 As New DataGridViewComboBoxColumn
        DataGridView1.Columns.Add(ComboBoxCol7)
        DataGridView1.Columns(7).HeaderText = "Primary Key"
        ComboBoxCol7.Items.Add("")
        ComboBoxCol7.Items.Add("Primary Key")
        Dim TextBoxCol8 As New DataGridViewTextBoxColumn
        DataGridView1.Columns.Add(TextBoxCol8)
        DataGridView1.Columns(8).HeaderText = "Description"

        ''Create New Table data:
        'Dim TextBoxCol3 As New DataGridViewTextBoxColumn
        'DataGridView1.Columns.Add(TextBoxCol3)
        'DataGridView1.Columns(3).HeaderText = "Size"
        'Dim TextBoxCol4 As New DataGridViewTextBoxColumn
        'DataGridView1.Columns.Add(TextBoxCol4)
        'DataGridView1.Columns(4).HeaderText = "Precision"
        'Dim TextBoxCol5 As New DataGridViewTextBoxColumn
        'DataGridView1.Columns.Add(TextBoxCol5)
        'DataGridView1.Columns(5).HeaderText = "Scale"
        'Dim ComboBoxCol6 As New DataGridViewComboBoxColumn
        'DataGridView1.Columns.Add(ComboBoxCol6)
        'DataGridView1.Columns(6).HeaderText = "Null/Not Null"
        'ComboBoxCol6.Items.Add("")
        'ComboBoxCol6.Items.Add("Null")
        'ComboBoxCol6.Items.Add("Not Null")
        'Dim ComboBoxCol7 As New DataGridViewComboBoxColumn
        'DataGridView1.Columns.Add(ComboBoxCol7)
        'DataGridView1.Columns(7).HeaderText = "Auto Increment"
        'ComboBoxCol7.Items.Add("")
        'ComboBoxCol7.Items.Add("Auto Increment")
        'Dim ComboBoxCol8 As New DataGridViewComboBoxColumn
        'DataGridView1.Columns.Add(ComboBoxCol8)
        'DataGridView1.Columns(8).HeaderText = "Primary Key"
        'ComboBoxCol8.Items.Add("")
        'ComboBoxCol8.Items.Add("Primary Key")
        'Dim TextBoxCol9 As New DataGridViewTextBoxColumn
        'DataGridView1.Columns.Add(TextBoxCol9)
        'DataGridView1.Columns(9).HeaderText = "Description"

        'Column descriptions:
        Dim TextBoxCol3_0 As New DataGridViewTextBoxColumn
        DataGridView3.Columns.Add(TextBoxCol3_0)
        DataGridView3.Columns(0).HeaderText = "Column Name"
        DataGridView3.Columns(0).Width = 160
        Dim TextBoxCol3_1 As New DataGridViewTextBoxColumn
        DataGridView3.Columns.Add(TextBoxCol3_1)
        DataGridView3.Columns(1).HeaderText = "Description"

        'Relationships data:
        Dim TextBoxCol2_0 As New DataGridViewTextBoxColumn
        DataGridView2.Columns.Add(TextBoxCol2_0)
        DataGridView2.Columns(0).HeaderText = "PK_TABLE_NAME"
        Dim TextBoxCol2_1 As New DataGridViewTextBoxColumn
        DataGridView2.Columns.Add(TextBoxCol2_1)
        DataGridView2.Columns(1).HeaderText = "PK_COLUMN_NAME"
        Dim TextBoxCol2_2 As New DataGridViewTextBoxColumn
        DataGridView2.Columns.Add(TextBoxCol2_2)
        DataGridView2.Columns(2).HeaderText = "FK_TABLE_NAME"
        Dim TextBoxCol2_3 As New DataGridViewTextBoxColumn
        DataGridView2.Columns.Add(TextBoxCol2_3)
        DataGridView2.Columns(3).HeaderText = "FK_COLUMN_NAME"
        Dim TextBoxCol2_4 As New DataGridViewTextBoxColumn
        DataGridView2.Columns.Add(TextBoxCol2_4)
        DataGridView2.Columns(4).HeaderText = "ORDINAL"
        Dim TextBoxCol2_5 As New DataGridViewTextBoxColumn
        DataGridView2.Columns.Add(TextBoxCol2_5)
        DataGridView2.Columns(5).HeaderText = "UPDATE_RULE"
        Dim TextBoxCol2_6 As New DataGridViewTextBoxColumn
        DataGridView2.Columns.Add(TextBoxCol2_6)
        DataGridView2.Columns(6).HeaderText = "DELETE_RULE"
        Dim TextBoxCol2_7 As New DataGridViewTextBoxColumn
        DataGridView2.Columns.Add(TextBoxCol2_7)
        DataGridView2.Columns(7).HeaderText = "PK_NAME"
        DataGridView2.Columns(7).Width = 160
        Dim TextBoxCol2_8 As New DataGridViewTextBoxColumn
        DataGridView2.Columns.Add(TextBoxCol2_8)
        DataGridView2.Columns(8).HeaderText = "FK_NAME"
        DataGridView2.Columns(8).Width = 160
        DataGridView2.AutoResizeColumns()

        FillRelationshipsGrid()

        FillCmbTableLists()

        'Index data:
        'Show Disallow Null and Ignore Null options???
        Dim TextBoxCol4_0 As New DataGridViewTextBoxColumn
        DataGridView4.Columns.Add(TextBoxCol4_0)
        DataGridView4.Columns(0).HeaderText = "Table Name"
        DataGridView4.Columns(0).Width = 120
        Dim TextBoxCol4_1 As New DataGridViewTextBoxColumn
        DataGridView4.Columns.Add(TextBoxCol4_1)
        DataGridView4.Columns(1).HeaderText = "Index Name"
        DataGridView4.Columns(1).Width = 120
        Dim CheckBoxCol4_2 As New DataGridViewCheckBoxColumn
        DataGridView4.Columns.Add(CheckBoxCol4_2)
        DataGridView4.Columns(2).HeaderText = "Primary Key"
        DataGridView4.Columns(2).Width = 120
        Dim CheckBoxCol4_3 As New DataGridViewCheckBoxColumn
        DataGridView4.Columns.Add(CheckBoxCol4_3)
        DataGridView4.Columns(3).HeaderText = "Unique"
        DataGridView4.Columns(3).Width = 120
        Dim TextBoxCol4_4 As New DataGridViewTextBoxColumn
        DataGridView4.Columns.Add(TextBoxCol4_4)
        DataGridView4.Columns(4).HeaderText = "Column Name"
        DataGridView4.Columns(4).Width = 120

        FillIndexesGrid()

        'Add Index options:
        'http://msdn.microsoft.com/en-us/library/aa140011(v=office.10).aspx
        cmbIndexOptions.Items.Clear()
        cmbIndexOptions.Items.Add("")
        cmbIndexOptions.Items.Add("Primary")
        cmbIndexOptions.Items.Add("Disallow Null")
        cmbIndexOptions.Items.Add("Ignore Null")
        'cmbIndexOptions.Items.Add("Unique")

        'Add data types to the Add Columns section of the Miscellaneous tab:
        cmbColumnType.Items.Add("Short (Integer)")
        cmbColumnType.Items.Add("Long (Integer)")
        cmbColumnType.Items.Add("Single")
        cmbColumnType.Items.Add("Double")
        cmbColumnType.Items.Add("Currency")
        cmbColumnType.Items.Add("DateTime")
        cmbColumnType.Items.Add("Bit (Boolean)")
        cmbColumnType.Items.Add("Byte")
        cmbColumnType.Items.Add("GUID")
        cmbColumnType.Items.Add("BigBinary")
        cmbColumnType.Items.Add("LongBinary")
        cmbColumnType.Items.Add("VarBinary")
        cmbColumnType.Items.Add("LongText")
        cmbColumnType.Items.Add("VarChar")
        cmbColumnType.Items.Add("Decimal")

        'Add Null/Not Null options to the Add Columns section of the Mescellaneous tab:
        cmbNull.Items.Add("")
        cmbNull.Items.Add("Null")
        cmbNull.Items.Add("Not Null")

        'Add AutoIncrement option to the Add Columns section of the Mescellaneous tab:
        cmbAutoIncrement.Items.Add("")
        cmbAutoIncrement.Items.Add("Auto Increment")

        'Add PrimaryKey option to the Add Columns section of the Mescellaneous tab:
        cmbPrimaryKey.Items.Add("")
        cmbPrimaryKey.Items.Add("Primary Key")

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Me.Close() 'Close the form
    End Sub

    Private Sub frmModifyDatabase_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if form is minimised.
        End If
    End Sub





#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Open and Close Forms - Code used to open and close other forms." '-------------------------------------------------------------------------------------------------------------------

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Methods - The main actions performed by this form." '---------------------------------------------------------------------------------------------------------------------------

    Private Sub FillCmbTableLists()
        'Fill the cmbSelectTable combobox with the availalble tables in the selected database.

        If Main.DatabasePath = "" Then
            'No database selected
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        cmbDescrSelectTable.Items.Clear()
        cmbUtilitiesSelectTable.Items.Clear()
        cmbNewIndexTable.Items.Clear()
        cmbFKTable.Items.Clear()
        cmbRelatedTable.Items.Clear()

        'Specify the connection string:
        'Access 2003
        'connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + _
        '"data source = " + txtDatabase.Text

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + Main.DatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
        dt = conn.GetSchema("Tables", restrictions)

        'Fill lstSelectTable
        'Dim dr As DataRow
        Dim I As Integer 'Loop index
        Dim MaxI As Integer

        MaxI = dt.Rows.Count
        For I = 0 To MaxI - 1
            'dr = dt.Rows(0)
            cmbDescrSelectTable.Items.Add(dt.Rows(I).Item(2).ToString)
            cmbUtilitiesSelectTable.Items.Add(dt.Rows(I).Item(2).ToString)
            cmbNewIndexTable.Items.Add(dt.Rows(I).Item(2).ToString)
            cmbFKTable.Items.Add(dt.Rows(I).Item(2).ToString)
            cmbRelatedTable.Items.Add(dt.Rows(I).Item(2).ToString)
        Next I

        conn.Close()

    End Sub

#Region " Create new table code" 'Code used to create a new table -----------------------------------------------------------------------------------------------------------------------------

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        Dim Row As Integer = e.RowIndex
        Dim Col As Integer = e.ColumnIndex

        If DataGridView1.Rows(Row).Cells(Col).ReadOnly = False Then
            If DataGridView1.Columns(Col).CellType = GetType(DataGridViewComboBoxCell) Then
                'OR: If TypeOf(DataGridView1.EditingControl) Is ComboBox Then
                DataGridView1.BeginEdit(True)
                'Casting the editing control and fire DropDown event:
                CType(DataGridView1.EditingControl, ComboBox).DroppedDown = True
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit

        Dim Row As Integer
        Dim Col As Integer

        Row = e.RowIndex
        Col = e.ColumnIndex

        'List of columns in DataGridView1:
        '0: Column Name   1: Type   2: Size   3: Precision   4: Scale   5: Null/Not Null   6: Auto Increment   7: Primary Key 8: Description

        If Col = 1 Then 'Column Type selected:
            Select Case DataGridView1.Rows(Row).Cells(Col).Value
                Case "Short (Integer)"
                    Main.Message.Add("Column type is Short (Integer)" & vbCrLf & vbCrLf)
                    DataGridView1.Rows(Row).Cells(2).Value = "n/a" 'Disable Size
                    DataGridView1.Rows(Row).Cells(2).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(3).Value = "n/a" 'Disable Precision
                    DataGridView1.Rows(Row).Cells(3).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(4).Value = "n/a" 'Disable Scale
                    DataGridView1.Rows(Row).Cells(4).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(6).Value = "" 'Disable Auto Increment
                    DataGridView1.Rows(Row).Cells(6).ReadOnly = True

                Case "Long (Integer)"
                    Main.Message.Add("Column type is Long (Integer)" & vbCrLf & vbCrLf)
                    DataGridView1.Rows(Row).Cells(2).Value = "n/a" 'Disable Size
                    DataGridView1.Rows(Row).Cells(2).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(3).Value = "n/a" 'Disable Precision
                    DataGridView1.Rows(Row).Cells(3).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(4).Value = "n/a" 'Disable Scale
                    DataGridView1.Rows(Row).Cells(4).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(6).ReadOnly = False 'Enable Auto Increment

                Case "Single"
                    Main.Message.Add("Column type is Single" & vbCrLf & vbCrLf)
                    DataGridView1.Rows(Row).Cells(2).Value = "n/a" 'Disable Size
                    DataGridView1.Rows(Row).Cells(2).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(3).Value = "n/a" 'Disable Precision
                    DataGridView1.Rows(Row).Cells(3).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(4).Value = "n/a" 'Disable Scale
                    DataGridView1.Rows(Row).Cells(4).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(6).Value = "" 'Disable Auto Increment
                    DataGridView1.Rows(Row).Cells(6).ReadOnly = True

                Case "Double"
                    Main.Message.Add("Column type is Double" & vbCrLf & vbCrLf)
                    DataGridView1.Rows(Row).Cells(2).Value = "n/a" 'Disable Size
                    DataGridView1.Rows(Row).Cells(2).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(3).Value = "n/a" 'Disable Precision
                    DataGridView1.Rows(Row).Cells(3).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(4).Value = "n/a" 'Disable Scale
                    DataGridView1.Rows(Row).Cells(4).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(6).Value = "" 'Disable Auto Increment
                    DataGridView1.Rows(Row).Cells(6).ReadOnly = True

                Case "Currency"
                    Main.Message.Add("Column type is Currency" & vbCrLf & vbCrLf)
                    DataGridView1.Rows(Row).Cells(2).Value = "n/a" 'Disable Size
                    DataGridView1.Rows(Row).Cells(2).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(3).Value = "n/a" 'Disable Precision
                    DataGridView1.Rows(Row).Cells(3).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(4).Value = "n/a" 'Disable Scale
                    DataGridView1.Rows(Row).Cells(4).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(6).Value = "" 'Disable Auto Increment
                    DataGridView1.Rows(Row).Cells(6).ReadOnly = True

                Case "DateTime"
                    Main.Message.Add("Column type is DateTime" & vbCrLf & vbCrLf)
                    DataGridView1.Rows(Row).Cells(2).Value = "n/a" 'Disable Size
                    DataGridView1.Rows(Row).Cells(2).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(3).Value = "n/a" 'Disable Precision
                    DataGridView1.Rows(Row).Cells(3).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(4).Value = "n/a" 'Disable Scale
                    DataGridView1.Rows(Row).Cells(4).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(6).Value = "" 'Disable Auto Increment
                    DataGridView1.Rows(Row).Cells(6).ReadOnly = True

                Case "Bit (Boolean)"
                    Main.Message.Add("Column type is Bit (Boolean)" & vbCrLf & vbCrLf)
                    DataGridView1.Rows(Row).Cells(2).Value = "n/a" 'Disable Size
                    DataGridView1.Rows(Row).Cells(2).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(3).Value = "n/a" 'Disable Precision
                    DataGridView1.Rows(Row).Cells(3).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(4).Value = "n/a" 'Disable Scale
                    DataGridView1.Rows(Row).Cells(4).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(6).Value = "" 'Disable Auto Increment
                    DataGridView1.Rows(Row).Cells(6).ReadOnly = True

                Case "Byte"
                    Main.Message.Add("Column type is Byte" & vbCrLf & vbCrLf)
                    DataGridView1.Rows(Row).Cells(2).Value = "n/a" 'Disable Size
                    DataGridView1.Rows(Row).Cells(2).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(3).Value = "n/a" 'Disable Precision
                    DataGridView1.Rows(Row).Cells(3).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(4).Value = "n/a" 'Disable Scale
                    DataGridView1.Rows(Row).Cells(4).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(6).Value = "" 'Disable Auto Increment
                    DataGridView1.Rows(Row).Cells(6).ReadOnly = True

                Case "GUID"
                    Main.Message.Add("Column type is GUID" & vbCrLf & vbCrLf)
                    DataGridView1.Rows(Row).Cells(2).Value = "n/a" 'Disable Size
                    DataGridView1.Rows(Row).Cells(2).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(3).Value = "n/a" 'Disable Precision
                    DataGridView1.Rows(Row).Cells(3).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(4).Value = "n/a" 'Disable Scale
                    DataGridView1.Rows(Row).Cells(4).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(6).Value = "" 'Disable Auto Increment
                    DataGridView1.Rows(Row).Cells(6).ReadOnly = True

                Case "BigBinary"
                    Main.Message.Add("Column type is BigBinary" & vbCrLf & vbCrLf)
                    DataGridView1.Rows(Row).Cells(2).Value = "n/a" 'Disable Size
                    DataGridView1.Rows(Row).Cells(2).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(3).Value = "n/a" 'Disable Precision
                    DataGridView1.Rows(Row).Cells(3).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(4).Value = "n/a" 'Disable Scale
                    DataGridView1.Rows(Row).Cells(4).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(6).Value = "" 'Disable Auto Increment
                    DataGridView1.Rows(Row).Cells(6).ReadOnly = True
                    Main.Message.Add("BigBinary: Maximum size: 4000 " & vbCrLf)

                Case "LongBinary"
                    Main.Message.Add("Column type is LongBinary" & vbCrLf & vbCrLf)
                    DataGridView1.Rows(Row).Cells(2).Value = "n/a" 'Disable Size
                    DataGridView1.Rows(Row).Cells(2).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(3).Value = "n/a" 'Disable Precision
                    DataGridView1.Rows(Row).Cells(3).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(4).Value = "n/a" 'Disable Scale
                    DataGridView1.Rows(Row).Cells(4).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(6).Value = "" 'Disable Auto Increment
                    DataGridView1.Rows(Row).Cells(6).ReadOnly = True
                    Main.Message.Add("LongBinary: Maximum size: 1073741823" & vbCrLf)

                Case "VarBinary"
                    Main.Message.Add("Column type is VarBinary" & vbCrLf & vbCrLf)
                    DataGridView1.Rows(Row).Cells(2).Value = "" 'Enable Size
                    DataGridView1.Rows(Row).Cells(2).ReadOnly = False
                    DataGridView1.Rows(Row).Cells(3).Value = "n/a" 'Disable Precision
                    DataGridView1.Rows(Row).Cells(3).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(4).Value = "n/a" 'Disable Scale
                    DataGridView1.Rows(Row).Cells(4).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(6).Value = "" 'Disable Auto Increment
                    DataGridView1.Rows(Row).Cells(6).ReadOnly = True
                    Main.Message.Add("VarBinary: Maximum size: 510" & vbCrLf)

                Case "LongText"
                    Main.Message.Add("Column type is LongText" & vbCrLf & vbCrLf)
                    DataGridView1.Rows(Row).Cells(2).Value = "n/a" 'Disable Size
                    DataGridView1.Rows(Row).Cells(2).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(3).Value = "n/a" 'Disable Precision
                    DataGridView1.Rows(Row).Cells(3).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(4).Value = "n/a" 'Disable Scale
                    DataGridView1.Rows(Row).Cells(4).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(6).Value = "" 'Disable Auto Increment
                    DataGridView1.Rows(Row).Cells(6).ReadOnly = True
                    Main.Message.Add("LongText: Maximum size: 536870910" & vbCrLf)

                Case "VarChar"
                    Main.Message.Add("Column type is VarChar" & vbCrLf & vbCrLf)
                    DataGridView1.Rows(Row).Cells(2).Value = "" 'Enable Size
                    DataGridView1.Rows(Row).Cells(2).ReadOnly = False
                    DataGridView1.Rows(Row).Cells(3).Value = "n/a" 'Disable Precision
                    DataGridView1.Rows(Row).Cells(3).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(4).Value = "n/a" 'Disable Scale
                    DataGridView1.Rows(Row).Cells(4).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(6).Value = "" 'Disable Auto Increment
                    DataGridView1.Rows(Row).Cells(6).ReadOnly = True
                    Main.Message.Add("VarChar: Maximum size: 255" & vbCrLf)

                Case "Decimal"
                    Main.Message.Add("Column type is Decimal" & vbCrLf & vbCrLf)
                    DataGridView1.Rows(Row).Cells(2).Value = "n/a" 'Disable Size
                    DataGridView1.Rows(Row).Cells(2).ReadOnly = True
                    DataGridView1.Rows(Row).Cells(3).Value = "" 'Enable Precision
                    DataGridView1.Rows(Row).Cells(3).ReadOnly = False
                    DataGridView1.Rows(Row).Cells(4).Value = "" 'Enable Scale
                    DataGridView1.Rows(Row).Cells(4).ReadOnly = False
                    DataGridView1.Rows(Row).Cells(6).Value = "" 'Disable Auto Increment
                    DataGridView1.Rows(Row).Cells(6).ReadOnly = True
                    Main.Message.Add("Decimal: Specify Precision and Scale." & vbCrLf)
                    Main.Message.Add("Precision is the number of digits." & vbCrLf)
                    Main.Message.Add("Scale is the number of digits to the right of the decimal point." & vbCrLf)

            End Select
        End If

    End Sub

    Private Sub btnCreateTable_Click(sender As System.Object, e As System.EventArgs) Handles btnCreateTable.Click
        'Create the new table:
        'Use SQL DDL to create the new table.
        'Use ADOX to add the column descriptions. (This cannot be done using DDL.)

        If Trim(txtNewTableName.Text) = "" Then
            Main.Message.SetWarningStyle()
            Main.Message.Add("Table name not specified." & vbCrLf)
            Main.Message.SetNormalStyle()
            Exit Sub
        End If


        Dim bldCmd As New System.Text.StringBuilder
        bldCmd.Clear()
        bldCmd.Append("CREATE TABLE " & Trim(txtNewTableName.Text) & vbCrLf)
        Dim I As Integer
        Dim LastRow As Integer
        Dim DataTypeString As String
        LastRow = DataGridView1.RowCount - 1 'This is the number of columns specified in the DataGridView

        'Check for invalid field specifications:
        Dim InvalidSpec As Boolean = False
        Dim NPrimaryKeyCols As Integer = 0 'Keeps a count of the number of columns used in the primary key.
        Dim PrimaryKeyCols() As String 'Array of primary key column names.
        Dim PrimaryKeyColNo As Integer = 0 'The current primary key column number.
        For I = 1 To LastRow
            If Trim(DataGridView1.Rows(I - 1).Cells(1).Value) = "VarChar" Then
                If Trim(DataGridView1.Rows(I - 1).Cells(2).Value) = "" Then
                    Main.Message.SetWarningStyle()
                    Main.Message.Add("The maximum size of the character string in Column " & DataGridView1.Rows(I - 1).Cells(0).Value & " is not specified." & vbCrLf)
                    Main.Message.SetNormalStyle()
                    InvalidSpec = True
                End If
            End If
            If Trim(DataGridView1.Rows(I - 1).Cells(1).Value) = "Decimal" Then
                If Trim(DataGridView1.Rows(I - 1).Cells(3).Value) = "" Then
                    Main.Message.SetWarningStyle()
                    Main.Message.Add("The Precision of the Decimal type in Column " & DataGridView1.Rows(I - 1).Cells(0).Value & " is not specified." & vbCrLf)
                    Main.Message.SetNormalStyle()
                    InvalidSpec = True
                End If
                If Trim(DataGridView1.Rows(I - 1).Cells(4).Value) = "" Then
                    Main.Message.SetWarningStyle()
                    Main.Message.Add("The Scale of the Decimal type in Column " & DataGridView1.Rows(I - 1).Cells(0).Value & " is not specified." & vbCrLf)
                    Main.Message.SetNormalStyle()
                    InvalidSpec = True
                End If
            End If
            If DataGridView1.Rows(I - 1).Cells(7).Value = "Primary Key" Then
                NPrimaryKeyCols = NPrimaryKeyCols + 1
            End If
        Next
        Main.Message.Add("The number of columns used in the primary key: " & NPrimaryKeyCols & vbCrLf)

        If NPrimaryKeyCols > 1 Then
            ReDim PrimaryKeyCols(0 To NPrimaryKeyCols - 1)
        End If

        If InvalidSpec = True Then
            Exit Sub
        End If

        For I = 1 To LastRow
            If I = 1 Then
                bldCmd.Append("    (")
            Else
                bldCmd.Append("    ")
            End If
            GetDataTypeString(I - 1, DataTypeString)
            bldCmd.Append("[" & Trim(DataGridView1.Rows(I - 1).Cells(0).Value) & "] " & DataTypeString) 'Add the new column name

            'Add Null / Not Null:
            If Trim(DataGridView1.Rows(I - 1).Cells(5).Value) <> "" Then
                bldCmd.Append(" " & UCase(DataGridView1.Rows(I - 1).Cells(5).Value))
            End If

            'Add Auto Increment:
            If Trim(DataGridView1.Rows(I - 1).Cells(6).Value) = "Auto Increment" Then
                'bldCmd.Append(" AUTO_INCREMENT")
                bldCmd.Append(" IDENTITY(1,1)")
            End If

            'Add Primary Key:
            If Trim(DataGridView1.Rows(I - 1).Cells(7).Value) = "Primary Key" Then
                If NPrimaryKeyCols = 1 Then
                    bldCmd.Append(" PRIMARY KEY")
                ElseIf NPrimaryKeyCols > 1 Then
                    PrimaryKeyCols(PrimaryKeyColNo) = Trim(DataGridView1.Rows(I - 1).Cells(0).Value)
                    PrimaryKeyColNo = PrimaryKeyColNo + 1
                End If

            End If

            If I = LastRow Then
                'bldCmd.Append(");" & vbCrLf) 'close the brackets containing the column specifications
                If NPrimaryKeyCols > 1 Then
                    bldCmd.Append("," & vbCrLf)
                End If

            Else
                bldCmd.Append("," & vbCrLf) 'end the line
            End If
        Next

        'Add Primary Key statement if required:
        If NPrimaryKeyCols > 1 Then
            bldCmd.Append("CONSTRAINT pk_" & PrimaryKeyCols(0) & " PRIMARY KEY (")
            For I = 0 To NPrimaryKeyCols - 1
                bldCmd.Append(PrimaryKeyCols(I))
                If I = NPrimaryKeyCols - 1 Then
                    bldCmd.Append(")")
                Else
                    bldCmd.Append(",")
                End If
            Next
        End If

        'Close the Create Table bracket:
        bldCmd.Append(");" & vbCrLf)

        Main.Message.Add("Create Table SQL command: " & bldCmd.ToString & vbCrLf)


        'Apply the SQL DDL command:
        Main.SqlCommandText = bldCmd.ToString 'Set the SqlCommandText property on the Database form.
        Main.ApplySqlCommand()


        Main.Message.Add("Create Table SQL command has been applied. " & vbCrLf)


        'Add the Column Descriptions:
        'Project \ Add Reference \ COM \ Microsoft ActiveX Data Objects 6.1 Library
        Dim aConn As New ADODB.Connection
        'Project \ Add Reference \ COM \ Microsoft ADO Ext 6.0 for DDL and Security
        Dim aDB As New ADOX.Catalog
        Dim aTable As ADOX.Table
        Dim aField As ADOX.Column

        aConn.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" + "data source = " + Main.DatabasePath
        aConn.Open()
        aDB.ActiveConnection = aConn

        'Check that the new table has been created:
        Dim NTables As Integer
        Dim TableFound As Boolean = False
        NTables = aDB.Tables.Count
        For I = 0 To NTables - 1
            If aDB.Tables(I).Name = txtNewTableName.Text Then
                TableFound = True
                Exit For
            End If
        Next

        If TableFound = False Then
            Main.Message.SetWarningStyle()
            Main.Message.Add("Error. New table was not created! " & vbCrLf)
            Main.Message.SetNormalStyle()
        Else
            aTable = aDB.Tables(txtNewTableName.Text)

            Dim NFields As Integer
            NFields = aTable.Columns.Count

            For I = 0 To NFields - 1
                aTable.Columns(Trim(DataGridView1.Rows(I).Cells(0).Value)).Properties("Description").Value = Trim(DataGridView1.Rows(I).Cells(8).Value)
            Next

            'Update the list of tables on the Database tab:
            Main.FillLstTables()
            'Update the list of tables on the Tables tab:
            Main.FillCmbSelectTable()

        End If

        aConn.Close()

    End Sub

    Private Sub GetDataTypeString(ByVal RowNo As Integer, ByRef DataTypeString As String)
        'Return the string used to define the data type in the SQL command:

        Dim DataType As String
        DataType = DataGridView1.Rows(RowNo).Cells(1).Value
        Dim Size As Integer

        Select Case DataType
            Case "Short (Integer)"
                DataTypeString = "SHORT"
            Case "Long (Integer)"
                DataTypeString = "LONG"
            Case "Single"
                DataTypeString = "SINGLE"
            Case "Double"
                DataTypeString = "DOUBLE"
            Case "Currency"
                DataTypeString = "CURRENCY"
            Case "DateTime"
                DataTypeString = "DATETIME"
            Case "Bit (Boolean)"
                DataTypeString = "BIT"
            Case "Byte"
                DataTypeString = "BYTE"
            Case "GUID"
                DataTypeString = "GUID"
            Case "BigBinary"
                DataTypeString = "BIGBINARY"
            Case "LongBinary"
                DataTypeString = "LONGBINARY"
            Case "VarBinary"
                Size = DataGridView1.Rows(RowNo).Cells(2).Value
                DataTypeString = "VARBINARY (" & Size & ")"
            Case "LongText"
                DataTypeString = "LONGTEXT"
            Case "VarChar"
                Size = DataGridView1.Rows(RowNo).Cells(2).Value
                DataTypeString = "VARCHAR (" & Size & ")"
            Case "Decimal"
                Dim Precision As Integer
                Precision = DataGridView1.Rows(RowNo).Cells(3).Value
                Dim Scale As Integer
                Scale = DataGridView1.Rows(RowNo).Cells(4).Value
                DataTypeString = "DECIMAL (" & Precision & "," & Scale & ")"
        End Select
    End Sub

#End Region 'Create new table code ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Column Descriptions Code" 'Code used to view and edit column dexcriptions -----------------------------------------------------------------------------------------------------------

    Private Sub cmbDescrSelectTable_TextChanged(sender As Object, e As System.EventArgs) Handles cmbDescrSelectTable.TextChanged
        'Fill DataGridView3 with a list of Column names and corresponding descriptions.
        RefreshDescriptions()
    End Sub

    Private Sub RefreshDescriptions()
        'Refresh the Columns Names and Descriptions shown on DataGridView3:

        If Main.DatabasePath = "" Then
            'No database selected.
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim commandString As String 'Declare a command string - contains the query to be passed to the database.
        Dim ds As DataSet 'Declate a Dataset.
        Dim dt As DataTable

        'Project \ Add Reference \ COM \ Microsoft ActiveX Data Objects 6.1 Library
        Dim aConn As New ADODB.Connection
        'Project \ Add Reference \ COM \ Microsoft ADO Ext 6.0 for DDL and Security
        Dim aDB As New ADOX.Catalog
        Dim aTable As ADOX.Table
        Dim aField As ADOX.Column

        If cmbDescrSelectTable.SelectedIndex = -1 Then 'No item is selected

        Else 'A table has been selected. List its fields:
            aConn.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" + "data source = " + Main.DatabasePath
            aConn.Open()
            aDB.ActiveConnection = aConn
            aTable = aDB.Tables(cmbDescrSelectTable.Text)
            Dim NFields As Integer
            NFields = aTable.Columns.Count
            DataGridView3.Rows.Clear()
            Dim I As Integer
            For I = 0 To NFields - 1
                DataGridView3.Rows.Add()
                DataGridView3.Rows(I).Cells(0).Value = aTable.Columns(I).Name
                If IsNothing(aTable.Columns(I).Properties("Description")) Then
                    aTable.Columns(I).Properties("Description").Value = ""
                    DataGridView3.Rows(I).Cells(1).Value = aTable.Columns(I).Properties("Description").Value.ToString
                Else
                    DataGridView3.Rows(I).Cells(1).Value = aTable.Columns(I).Properties("Description").Value.ToString
                End If
            Next
            aDB = Nothing
            aTable = Nothing
            aConn.Close()

            DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DataGridView3.AutoResizeColumns()
        End If
    End Sub

    Private Sub btnUpdateDescriptions_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdateDescriptions.Click
        'Update the database with the descriptions shown on the DataGridView3:

        If Main.DatabasePath = "" Then
            'No database selected.
            Exit Sub
        End If

        'Project \ Add Reference \ COM \ Microsoft ActiveX Data Objects 6.1 Library
        Dim aConn As New ADODB.Connection
        'Project \ Add Reference \ COM \ Microsoft ADO Ext 6.0 for DDL and Security
        Dim aDB As New ADOX.Catalog
        Dim aTable As ADOX.Table
        Dim aField As ADOX.Column

        If cmbDescrSelectTable.SelectedIndex = -1 Then 'No item is selected

        Else 'Update the selected table:
            aConn.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" + "data source = " + Main.DatabasePath
            aConn.Open()
            aDB.ActiveConnection = aConn
            aTable = aDB.Tables(cmbDescrSelectTable.Text)

            Dim NFields As Integer
            NFields = aTable.Columns.Count

            Main.Message.Add("Udating column descriptions table: " & vbCrLf)
            Dim I As Integer
            For I = 0 To NFields - 1
                'aTable.Columns(I).Properties("Description").Value = DataGridView3.Rows(I).Cells(1).Value
                'Main.Message.Add("Column number: " & I & "   Name: " & aTable.Columns(I).Name & "   Description: " & aTable.Columns(I).Properties("Description").Value.ToString & vbCrLf)
                If DataGridView3.Rows(I).Cells(1).Value = Nothing Then
                    Main.Message.Add("No description specified for column: " & DataGridView3.Rows(I).Cells(0).Value & vbCrLf)
                Else
                    aTable.Columns(DataGridView3.Rows(I).Cells(0).Value).Properties("Description").Value = DataGridView3.Rows(I).Cells(1).Value
                    Main.Message.Add("Column Name: " & DataGridView3.Rows(I).Cells(0).Value & "   Description: " & aTable.Columns(DataGridView3.Rows(I).Cells(0).Value).Properties("Description").Value.ToString & vbCrLf)
                End If

            Next
            aConn.Close()
        End If
    End Sub

#End Region 'Column Descriptions Code ---------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Utilities Code" 'Code used in the Utilities tab -------------------------------------------------------------------------------------------------------------------------------------

    Private Sub cmbUtilitiesSelectTable_TextChanged(sender As Object, e As System.EventArgs) Handles cmbUtilitiesSelectTable.TextChanged
        'A new table has been selected. Fill the Columns combo box:
        UpdateCmbUtilitiesSelectColumn()
    End Sub

    Private Sub btnRenameColumn_Click(sender As System.Object, e As System.EventArgs) Handles btnRenameColumn.Click
        'Change the name of a Column:
        'The table is selected in cmbUtilitiesSelectTable.
        'The column is selected in cmbUtilitiesSelectColumn.
        'The new name is specified in txtNewColumnName.

        If Main.DatabasePath = "" Then
            'No database selected.
            Exit Sub
        End If

        'Project \ Add Reference \ COM \ Microsoft ActiveX Data Objects 6.1 Library
        Dim aConn As New ADODB.Connection
        'Project \ Add Reference \ COM \ Microsoft ADO Ext 6.0 for DDL and Security
        Dim aDB As New ADOX.Catalog
        Dim aTable As ADOX.Table
        Dim aField As ADOX.Column


        aConn.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" + "data source = " + Main.DatabasePath
        aConn.Open()
        aDB.ActiveConnection = aConn
        aTable = aDB.Tables(cmbUtilitiesSelectTable.Text)

        aTable.Columns(cmbUtilitiesSelectColumn.Text).Name = txtNewColumnName.Text

        aDB = Nothing
        aTable = Nothing
        aConn.Close()

        Dim SelectedColumnIndex As Integer
        SelectedColumnIndex = cmbUtilitiesSelectColumn.SelectedIndex
        UpdateCmbUtilitiesSelectColumn()
        cmbUtilitiesSelectColumn.SelectedIndex = SelectedColumnIndex

        If cmbUtilitiesSelectTable.Text = cmbDescrSelectTable.Text Then
            'The Descriptions list is from the same table that has just been changed.
            'Redisplay the discriptions list to show the changed column name:
            RefreshDescriptions()
        End If

    End Sub

    Private Sub UpdateCmbUtilitiesSelectColumn()
        'Update the cmbUtilitiesSelectColumn combo box:

        If Main.DatabasePath = "" Then
            'No database selected.
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim commandString As String 'Declare a command string - contains the query to be passed to the database.
        Dim ds As DataSet 'Declate a Dataset.
        Dim dt As DataTable

        If cmbUtilitiesSelectTable.SelectedIndex = -1 Then 'No item is selected

        Else
            'Access 2007:
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
            "data source = " + Main.DatabasePath

            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            'Specify the commandString to query the database:
            'commandString = "SELECT * FROM " + cmbUtilitiesSelectTable.SelectedItem.ToString
            commandString = "SELECT Top 500 * FROM " + cmbUtilitiesSelectTable.SelectedItem.ToString
            Dim dataAdapter As New System.Data.OleDb.OleDbDataAdapter(commandString, conn)
            ds = New DataSet
            dataAdapter.Fill(ds, "SelectedTable") 'ds was defined earlier as a DataSet
            dt = ds.Tables("SelectedTable")

            Dim NFields As Integer
            NFields = dt.Columns.Count

            cmbUtilitiesSelectColumn.Items.Clear()
            Dim I As Integer
            For I = 0 To NFields - 1
                cmbUtilitiesSelectColumn.Items.Add(dt.Columns(I).ColumnName)
            Next

            conn.Close()

        End If
    End Sub

#End Region 'Utilities Code -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Relationships Code" '----------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Database Schema Relationships Table columns:
    'PK_TABLE_CATALOG PK_TABLE_SCHEMA PK_TABLE_NAME PK_COLUMN_NAME PK_COLUMN_PROPID FK_TABLE_CATALOG FK_TABLE_SCHEMA FK_TABLE_NAME FK_COLUMN_NAME FK_COLUMN_GUID FK_COLUMN_PROPID ORDINAL UPDATE_RULE DELETE_RULE PK_NAME_FK_NAME DEFERRABILITY
    'See http://msdn.microsoft.com/en-us/library/windows/desktop/ms711276(v=vs.85).aspx for

    'Columns displayed on DataGridView2:
    'PK_TABLE_NAME PK_COLUMN_NAME FK_TABLE_NAME FK_COLUMN_NAME ORDINAL UPDATE_RULE DELETE_RULE PK_NAME FK_NAME

    Private Sub FillRelationshipsGrid()

        If Main.DatabasePath = "" Then
            'No database selected.
            Exit Sub
        End If

        Dim connString As String
        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim schemaTable As DataTable = New DataTable

        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & Main.DatabasePath
        myConnection.ConnectionString = connString
        myConnection.Open()

        schemaTable = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Foreign_Keys, New String() {Nothing, Nothing, Nothing, Nothing, Nothing, Nothing})
        Dim NRelationships As Integer
        NRelationships = schemaTable.Rows.Count

        DataGridView2.Rows.Clear()

        Dim I As Integer
        For I = 1 To NRelationships
            DataGridView2.Rows.Add()
            DataGridView2.Rows(I - 1).Cells(0).Value = schemaTable.Rows(I - 1).Item("PK_TABLE_NAME")
            DataGridView2.Rows(I - 1).Cells(1).Value = schemaTable.Rows(I - 1).Item("PK_COLUMN_NAME")
            DataGridView2.Rows(I - 1).Cells(2).Value = schemaTable.Rows(I - 1).Item("FK_TABLE_NAME")
            DataGridView2.Rows(I - 1).Cells(3).Value = schemaTable.Rows(I - 1).Item("FK_COLUMN_NAME")
            DataGridView2.Rows(I - 1).Cells(4).Value = schemaTable.Rows(I - 1).Item("ORDINAL")
            DataGridView2.Rows(I - 1).Cells(5).Value = schemaTable.Rows(I - 1).Item("UPDATE_RULE")
            DataGridView2.Rows(I - 1).Cells(6).Value = schemaTable.Rows(I - 1).Item("DELETE_RULE")
            DataGridView2.Rows(I - 1).Cells(7).Value = schemaTable.Rows(I - 1).Item("PK_NAME")
            DataGridView2.Rows(I - 1).Cells(8).Value = schemaTable.Rows(I - 1).Item("FK_NAME")
        Next
        myConnection.Close()
        DataGridView2.AutoResizeColumns()
    End Sub

#End Region 'Relationships Code ---------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub cmbRelTable_TextChanged(sender As Object, e As System.EventArgs) Handles cmbFKTable.TextChanged
        'If a new table has been selected then update list of avaialble columns:

        If Main.DatabasePath = "" Then
            'No database selected.
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        If cmbFKTable.SelectedIndex = -1 Then 'No table has been selected
            cmbFKColumn.Items.Clear()
        Else 'A table has been selected. List its columns:
            cmbFKColumn.Items.Clear()

            'Access 2007:
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
            "data source = " + Main.DatabasePath

            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            'Restrictions: TABLE_CATALOG TABLE_SCHEMA TABLE_NAME COLUMN_NAME
            Dim restrictions As String() = New String() {Nothing, Nothing, cmbFKTable.Text, Nothing} 'This restriction removes system tables
            dt = conn.GetSchema("Columns", restrictions)

            Dim I As Integer
            Dim MaxI As Integer

            MaxI = dt.Rows.Count
            For I = 0 To MaxI - 1
                cmbFKColumn.Items.Add(dt.Rows(I).Item(3).ToString)
            Next

            conn.Close()

        End If

    End Sub

    Private Sub cmbRelatedTable_TextChanged(sender As Object, e As System.EventArgs) Handles cmbRelatedTable.TextChanged
        'If a new table has been selected then update list of avaialble columns:

        If Main.DatabasePath = "" Then
            'No database selected.
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        If cmbRelatedTable.SelectedIndex = -1 Then 'No table has been selected
            cmbRelatedColumn.Items.Clear()
        Else 'A table has been selected. List its columns:
            cmbRelatedColumn.Items.Clear()

            'Access 2007:
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
            "data source = " + Main.DatabasePath

            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            'Restrictions: TABLE_CATALOG TABLE_SCHEMA TABLE_NAME COLUMN_NAME
            Dim restrictions As String() = New String() {Nothing, Nothing, cmbRelatedTable.Text, Nothing} 'This restriction removes system tables
            dt = conn.GetSchema("Columns", restrictions)

            Dim I As Integer
            Dim MaxI As Integer

            MaxI = dt.Rows.Count
            For I = 0 To MaxI - 1
                cmbRelatedColumn.Items.Add(dt.Rows(I).Item(3).ToString)
            Next

            conn.Close()

        End If
    End Sub

    Private Sub FillIndexesGrid()

        If Main.DatabasePath = "" Then
            'No database selected.
            Exit Sub
        End If

        Dim connString As String
        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim schemaTable As DataTable = New DataTable

        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & Main.DatabasePath
        myConnection.ConnectionString = connString
        myConnection.Open()

        schemaTable = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Indexes, New String() {Nothing, Nothing, Nothing, Nothing, Nothing})
        Dim NIndexes As Integer
        NIndexes = schemaTable.Rows.Count

        DataGridView4.Rows.Clear()

        Dim I As Integer
        Dim LastRow As Integer = 0
        For I = 1 To NIndexes
            If schemaTable.Rows(I - 1).Item("TABLE_NAME") = "MSysAccessStorage" Then
                Continue For
            End If
            If schemaTable.Rows(I - 1).Item("TABLE_NAME") = "MSysNavPaneGroupCategories" Then
                Continue For
            End If
            If schemaTable.Rows(I - 1).Item("TABLE_NAME") = "MSysNavPaneGroups" Then
                Continue For
            End If
            If schemaTable.Rows(I - 1).Item("TABLE_NAME") = "MSysNavPaneGroupToObjects" Then
                Continue For
            End If
            DataGridView4.Rows.Add()
            DataGridView4.Rows(LastRow).Cells(0).Value = schemaTable.Rows(I - 1).Item("TABLE_NAME") 'Table Name
            DataGridView4.Rows(LastRow).Cells(1).Value = schemaTable.Rows(I - 1).Item("INDEX_NAME") 'Index Name
            DataGridView4.Rows(LastRow).Cells(2).Value = schemaTable.Rows(I - 1).Item("PRIMARY_KEY") 'Primary Key
            DataGridView4.Rows(LastRow).Cells(3).Value = schemaTable.Rows(I - 1).Item("UNIQUE") 'Unique
            DataGridView4.Rows(LastRow).Cells(4).Value = schemaTable.Rows(I - 1).Item("COLUMN_NAME") 'Column Name
            LastRow = LastRow + 1
        Next

        myConnection.Close()

        DataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        DataGridView4.AutoResizeColumns()

    End Sub

    Private Sub cmbNewIndexTable_TextChanged(sender As Object, e As System.EventArgs) Handles cmbNewIndexTable.TextChanged
        'Update lstColumns with the list of columns in the selected table

        If Main.DatabasePath = "" Then
            'No database selected.
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        If cmbNewIndexTable.SelectedIndex = -1 Then 'No table has been selected
            lstColumns.Items.Clear()
        Else
            'Access 2007:
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
            "data source = " + Main.DatabasePath

            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            'Restrictions: TABLE_CATALOG TABLE_SCHEMA TABLE_NAME COLUMN_NAME
            Dim restrictions As String() = New String() {Nothing, Nothing, cmbNewIndexTable.Text, Nothing} 'This restriction removes system tables
            dt = conn.GetSchema("Columns", restrictions)

            Dim I As Integer
            Dim MaxI As Integer

            MaxI = dt.Rows.Count
            For I = 0 To MaxI - 1
                lstColumns.Items.Add(dt.Rows(I).Item(3).ToString)
            Next

            conn.Close()

        End If

    End Sub

    Private Sub btnCreateRelationship_Click(sender As System.Object, e As System.EventArgs) Handles btnCreateRelationship.Click
        'Create the specified relationship
        'http://msdn.microsoft.com/en-us/library/aa140015(v=office.10).aspx
        '
        'ALTER TABLE tblShipping
        '   ADD CONSTRAINT FK_tblShipping
        '   FOREIGN KEY (CustomerID) REFERENCES
        '      tblCustomers (CustomerID)
        '   ON UPDATE CASCADE
        '   ON DELETE CASCADE

        Dim bldCmd As New System.Text.StringBuilder
        bldCmd.Clear()
        bldCmd.Append("ALTER TABLE " & Trim(cmbFKTable.Text) & vbCrLf)
        bldCmd.Append("ADD CONSTRAINT " & Trim(txtNewForeignKeyName.Text) & vbCrLf)
        bldCmd.Append("FOREIGN KEY (" & Trim(cmbFKColumn.Text) & ")" & vbCrLf & "REFERENCES ")
        bldCmd.Append(Trim(cmbRelatedTable.Text) & " (" & Trim(cmbRelatedColumn.Text) & ")" & vbCrLf)
        bldCmd.Append("ON UPDATE CASCADE " & vbCrLf)
        bldCmd.Append("ON DELETE CASCADE " & vbCrLf)

        Main.Message.Add("Create Relationship SQL command: " & vbCrLf & bldCmd.ToString & vbCrLf)

        'Apply the SQL DDL command:
        Main.SqlCommandText = bldCmd.ToString 'Set the SqlCommandText property on the Database form.
        Main.ApplySqlCommand()

    End Sub

    Private Sub btnDeleteTable_Click(sender As System.Object, e As System.EventArgs) Handles btnDeleteTable.Click
        'Delete the selected table:

        Dim bldCmd As New System.Text.StringBuilder
        bldCmd.Clear()
        bldCmd.Append("DROP TABLE " & Trim(cmbUtilitiesSelectTable.Text & ";") & vbCrLf)

        Main.SqlCommandText = bldCmd.ToString  'Set the SqlCommandText property on the Database form.
        Main.ApplySqlCommand()

        'Update relationships form:
        Dim OrigRelTable As String = cmbFKTable.Text
        Dim OrigRelColumn As String = cmbFKColumn.Text
        If OrigRelTable = cmbUtilitiesSelectTable.Text Then 'The selected table has been deleted.
            OrigRelTable = ""
            OrigRelColumn = ""
        End If

        Dim OrigRelatedTable As String = cmbRelatedTable.Text
        Dim OrigRelatedColumn As String = cmbRelatedColumn.Text
        If OrigRelatedTable = cmbUtilitiesSelectTable.Text Then
            OrigRelatedTable = ""
            OrigRelatedColumn = ""
        End If

        FillRelationshipsGrid()

        'Update Indexes form:
        Dim OrigNewIndexTable As String = cmbNewIndexTable.Text
        If OrigNewIndexTable = cmbUtilitiesSelectTable.Text Then
            OrigNewIndexTable = ""
        End If

        FillIndexesGrid()

        'Update Column Descriptions form:
        Dim OrigDescrSelectTable As String = cmbDescrSelectTable.Text
        If OrigDescrSelectTable = cmbUtilitiesSelectTable.Text Then
            OrigDescrSelectTable = ""
            DataGridView3.Rows.Clear()
        End If

        FillCmbTableLists()

        'Restore original table selections:
        cmbFKTable.Text = OrigRelTable
        cmbRelatedTable.Text = OrigRelatedTable
        cmbNewIndexTable.Text = OrigNewIndexTable
        cmbDescrSelectTable.Text = OrigDescrSelectTable

        'Update the list of tables on the Database form:
        Main.FillLstTables()

    End Sub

    Private Sub cmbColumnType_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbColumnType.SelectedIndexChanged

        Main.Message.Add("Column Type changed." & vbCrLf)

        'List of columns in DataGridView1:
        '0: Column Name   1: Type   2: Size   3: Precision   4: Scale   5: Null/Not Null   6: Auto Increment   7: Primary Key 8: Description

        'If Col = 1 Then 'Column Type selected:
        Select Case cmbColumnType.Text
            Case "Short (Integer)"
                Main.Message.Add("Column type is Short (Integer)" & vbCrLf & vbCrLf)
                txtSize.Text = "n/a" 'Disable Size
                txtSize.ReadOnly = True
                txtPrecision.Text = "n/a" 'Disable Precision
                txtPrecision.ReadOnly = True
                txtScale.Text = "n/a" 'Disable Scale
                txtScale.ReadOnly = True
                cmbAutoIncrement.Text = "" 'Disable AutoIncrement
                cmbAutoIncrement.Enabled = False

            Case "Long (Integer)"
                Main.Message.Add("Column type is Long (Integer)" & vbCrLf & vbCrLf)
                txtSize.Text = "n/a" 'Disable Size
                txtSize.ReadOnly = True
                txtPrecision.Text = "n/a" 'Disable Precision
                txtPrecision.ReadOnly = True
                txtScale.Text = "n/a" 'Disable Scale
                txtScale.ReadOnly = True
                cmbAutoIncrement.Enabled = True 'Enable Auto Increment

            Case "Single"
                Main.Message.Add("Column type is Single" & vbCrLf & vbCrLf)
                txtSize.Text = "n/a" 'Disable Size
                txtSize.ReadOnly = True
                txtPrecision.Text = "n/a" 'Disable Precision
                txtPrecision.ReadOnly = True
                txtScale.Text = "n/a" 'Disable Scale
                txtScale.ReadOnly = True
                cmbAutoIncrement.Text = "" 'Disable AutoIncrement
                cmbAutoIncrement.Enabled = False

            Case "Double"
                Main.Message.Add("Column type is Double" & vbCrLf & vbCrLf)
                txtSize.Text = "n/a" 'Disable Size
                txtSize.ReadOnly = True
                txtPrecision.Text = "n/a" 'Disable Precision
                txtPrecision.ReadOnly = True
                txtScale.Text = "n/a" 'Disable Scale
                txtScale.ReadOnly = True
                cmbAutoIncrement.Text = "" 'Disable AutoIncrement
                cmbAutoIncrement.Enabled = False

            Case "Currency"
                Main.Message.Add("Column type is Currency" & vbCrLf & vbCrLf)
                txtSize.Text = "n/a" 'Disable Size
                txtSize.ReadOnly = True
                txtPrecision.Text = "n/a" 'Disable Precision
                txtPrecision.ReadOnly = True
                txtScale.Text = "n/a" 'Disable Scale
                txtScale.ReadOnly = True
                cmbAutoIncrement.Text = "" 'Disable AutoIncrement
                cmbAutoIncrement.Enabled = False

            Case "DateTime"
                Main.Message.Add("Column type is DateTime" & vbCrLf & vbCrLf)
                txtSize.Text = "n/a" 'Disable Size
                txtSize.ReadOnly = True
                txtPrecision.Text = "n/a" 'Disable Precision
                txtPrecision.ReadOnly = True
                txtScale.Text = "n/a" 'Disable Scale
                txtScale.ReadOnly = True
                cmbAutoIncrement.Text = "" 'Disable AutoIncrement
                cmbAutoIncrement.Enabled = False

            Case "Bit (Boolean)"
                Main.Message.Add("Column type is Bit (Boolean)" & vbCrLf & vbCrLf)
                txtSize.Text = "n/a" 'Disable Size
                txtSize.ReadOnly = True
                txtPrecision.Text = "n/a" 'Disable Precision
                txtPrecision.ReadOnly = True
                txtScale.Text = "n/a" 'Disable Scale
                txtScale.ReadOnly = True
                cmbAutoIncrement.Text = "" 'Disable AutoIncrement
                cmbAutoIncrement.Enabled = False

            Case "Byte"
                Main.Message.Add("Column type is Byte" & vbCrLf & vbCrLf)
                txtSize.Text = "n/a" 'Disable Size
                txtSize.ReadOnly = True
                txtPrecision.Text = "n/a" 'Disable Precision
                txtPrecision.ReadOnly = True
                txtScale.Text = "n/a" 'Disable Scale
                txtScale.ReadOnly = True
                cmbAutoIncrement.Text = "" 'Disable AutoIncrement
                cmbAutoIncrement.Enabled = False

            Case "GUID"
                Main.Message.Add("Column type is GUID" & vbCrLf & vbCrLf)
                txtSize.Text = "n/a" 'Disable Size
                txtSize.ReadOnly = True
                txtPrecision.Text = "n/a" 'Disable Precision
                txtPrecision.ReadOnly = True
                txtScale.Text = "n/a" 'Disable Scale
                txtScale.ReadOnly = True
                cmbAutoIncrement.Text = "" 'Disable AutoIncrement
                cmbAutoIncrement.Enabled = False

            Case "BigBinary"
                Main.Message.Add("Column type is BigBinary" & vbCrLf & vbCrLf)
                txtSize.Text = "n/a" 'Disable Size
                txtSize.ReadOnly = True
                txtPrecision.Text = "n/a" 'Disable Precision
                txtPrecision.ReadOnly = True
                txtScale.Text = "n/a" 'Disable Scale
                txtScale.ReadOnly = True
                cmbAutoIncrement.Text = "" 'Disable AutoIncrement
                cmbAutoIncrement.Enabled = False
                Main.Message.Add("BigBinary: Maximum size: 4000 " & vbCrLf)

            Case "LongBinary"
                Main.Message.Add("Column type is LongBinary" & vbCrLf & vbCrLf)
                txtSize.Text = "n/a" 'Disable Size
                txtSize.ReadOnly = True
                txtPrecision.Text = "n/a" 'Disable Precision
                txtPrecision.ReadOnly = True
                txtScale.Text = "n/a" 'Disable Scale
                txtScale.ReadOnly = True
                cmbAutoIncrement.Text = "" 'Disable AutoIncrement
                cmbAutoIncrement.Enabled = False
                Main.Message.Add("LongBinary: Maximum size: 1073741823" & vbCrLf)

            Case "VarBinary"
                Main.Message.Add("Column type is VarBinary" & vbCrLf & vbCrLf)
                txtSize.Text = "" 'Enable Size
                txtSize.ReadOnly = False
                txtPrecision.Text = "n/a" 'Disable Precision
                txtPrecision.ReadOnly = True
                txtScale.Text = "n/a" 'Disable Scale
                txtScale.ReadOnly = True
                cmbAutoIncrement.Text = "" 'Disable AutoIncrement
                cmbAutoIncrement.Enabled = False
                Main.Message.Add("VarBinary: Maximum size: 510" & vbCrLf)

            Case "LongText"
                Main.Message.Add("Column type is LongText" & vbCrLf & vbCrLf)
                txtSize.Text = "n/a" 'Disable Size
                txtSize.ReadOnly = True
                txtPrecision.Text = "n/a" 'Disable Precision
                txtPrecision.ReadOnly = True
                txtScale.Text = "n/a" 'Disable Scale
                txtScale.ReadOnly = True
                cmbAutoIncrement.Text = "" 'Disable AutoIncrement
                cmbAutoIncrement.Enabled = False
                Main.Message.Add("LongText: Maximum size: 536870910" & vbCrLf)

            Case "VarChar"
                Main.Message.Add("Column type is VarChar" & vbCrLf & vbCrLf)
                txtSize.Text = "" 'Enable Size
                txtSize.ReadOnly = False
                txtPrecision.Text = "n/a" 'Disable Precision
                txtPrecision.ReadOnly = True
                txtScale.Text = "n/a" 'Disable Scale
                txtScale.ReadOnly = True
                cmbAutoIncrement.Text = "" 'Disable AutoIncrement
                cmbAutoIncrement.Enabled = False
                Main.Message.Add("VarChar: Maximum size: 255" & vbCrLf)

            Case "Decimal"
                Main.Message.Add("Column type is Decimal" & vbCrLf & vbCrLf)
                txtSize.Text = "n/a" 'Disable Size
                txtSize.ReadOnly = True
                txtPrecision.Text = "" 'Enable Precision
                txtPrecision.ReadOnly = False
                txtScale.Text = "" 'Enable Scale
                txtScale.ReadOnly = False
                cmbAutoIncrement.Text = "" 'Disable AutoIncrement
                cmbAutoIncrement.Enabled = False
                Main.Message.Add("Decimal: Specify Precision and Scale." & vbCrLf)
                Main.Message.Add("Precision is the number of digits." & vbCrLf)
                Main.Message.Add("Scale is the number of digits to the right of the decimal point." & vbCrLf)
        End Select
    End Sub

    Private Sub btnAddColumn_Click(sender As System.Object, e As System.EventArgs) Handles btnAddColumn.Click
        'Add the specified Column to the selected Table:

        'Create the new table:
        'Use SQL DDL to create the new table.
        'Use ADOX to add the column descriptions. (This cannot be done using DDL.)

        If Trim(cmbUtilitiesSelectTable.Text) = "" Then
            Main.Message.SetWarningStyle()
            Main.Message.Add("Table name not specified." & vbCrLf)
            Main.Message.SetNormalStyle()
            Exit Sub
        End If

        'Check for invalid field specifications:
        Dim InvalidSpec As Boolean = False
        If Trim(cmbColumnType.Text) = "VarChar" Then
            If Trim(txtSize.Text) = "" Then
                Main.Message.SetWarningStyle()
                Main.Message.Add("The maximum size of the character string is not specified." & vbCrLf)
                Main.Message.SetNormalStyle()
                InvalidSpec = True
            End If
        End If
        If Trim(cmbColumnType.Text) = "Decimal" Then
            If Trim(txtPrecision.Text) = "" Then
                Main.Message.SetWarningStyle()
                Main.Message.Add("The Precision of the Decimal type is not specified." & vbCrLf)
                Main.Message.SetNormalStyle()
                InvalidSpec = True
            End If
            If Trim(txtScale.Text) = "" Then
                Main.Message.SetWarningStyle()
                Main.Message.Add("The Scale of the Decimal type is not specified." & vbCrLf)
                Main.Message.SetNormalStyle()
                InvalidSpec = True
            End If
        End If

        If InvalidSpec = True Then
            Exit Sub
        End If

        Dim bldCmd As New System.Text.StringBuilder
        bldCmd.Clear()

        bldCmd.Append("ALTER TABLE " & Trim(cmbUtilitiesSelectTable.Text) & vbCrLf)
        bldCmd.Append("    ADD COLUMN [" & Trim(txtColumnName.Text) & "]") 'Add the column name

        Select Case cmbColumnType.Text
            Case "Short (Integer)"
                bldCmd.Append(" SHORT")
            Case "Long (Integer)"
                bldCmd.Append(" LONG")
            Case "Single"
                bldCmd.Append(" SINGLE")
            Case "Double"
                bldCmd.Append(" DOUBLE")
            Case "Currency"
                bldCmd.Append(" CURRENCY")
            Case "DateTime"
                bldCmd.Append(" DATETIME")
            Case "Bit (Boolean)"
                bldCmd.Append(" BIT")
            Case "Byte"
                bldCmd.Append(" BYTE")
            Case "GUID"
                bldCmd.Append(" GUID")
            Case "BigBinary"
                bldCmd.Append(" BIGBINARY")
            Case "LongBinary"
                bldCmd.Append(" LONGBINARY")
            Case "VarBinary"
                bldCmd.Append(" VARBINARY (" & txtSize.Text & ")")
            Case "LongText"
                bldCmd.Append(" LONGTEXT")
            Case "VarChar"
                bldCmd.Append(" VARCHAR (" & txtSize.Text & ")")
            Case "Decimal"
                bldCmd.Append(" DECIMAL (" & txtPrecision.Text & ", " & txtScale.Text & ")")
        End Select

        'Add Null / Not Null:
        If Trim(cmbNull.Text) <> "" Then
            bldCmd.Append(" " & UCase(cmbNull.Text))
        End If

        'Add Auto Increment:
        If Trim(cmbAutoIncrement.Text) = "Auto Increment" Then
            bldCmd.Append(" IDENTITY(1,1)")
        End If

        'Add Primary Key:
        If Trim(cmbPrimaryKey.Text) = "Primary Key" Then
            bldCmd.Append(" PRIMARY KEY")
        End If

        bldCmd.Append(");" & vbCrLf)

        'Apply the SQL DDL command:
        Main.SqlCommandText = bldCmd.ToString 'Set the SqlCommandText property on the Database form.
        Main.ApplySqlCommand()

        If Main.SqlCommandResult = "OK" Then
            'Add the Column Description:
            'Project \ Add Reference \ COM \ Microsoft ActiveX Data Objects 6.1 Library
            Dim aConn As New ADODB.Connection
            'Project \ Add Reference \ COM \ Microsoft ADO Ext 6.0 for DDL and Security
            Dim aDB As New ADOX.Catalog
            Dim aTable As ADOX.Table
            Dim aField As ADOX.Column

            aConn.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" + "data source = " + Main.DatabasePath
            aConn.Open()
            aDB.ActiveConnection = aConn
            aTable = aDB.Tables(cmbUtilitiesSelectTable.Text)

            Dim NFields As Integer
            NFields = aTable.Columns.Count

            'If aTable.Columns(Trim(txtColumnName.Text)).Properties("Description").Value = Nothing Then
            'If IsNothing(aTable.Columns(Trim(txtColumnName.Text))) Then

            'Else
            aTable.Columns(Trim(txtColumnName.Text)).Properties("Description").Value = Trim(txtDescription.Text)
            'End If


            aConn.Close()

            'Clear the column parameters:
            txtColumnName.Text = ""
            txtSize.Text = ""
            txtPrecision.Text = ""
            txtScale.Text = ""
            cmbNull.SelectedIndex = 0
            cmbAutoIncrement.SelectedIndex = 0
            cmbPrimaryKey.SelectedIndex = 0
            txtDescription.Text = ""

            'Update the list of tables on the Database form:
            Main.FillLstTables()

            'Update the list of columns:
            UpdateCmbUtilitiesSelectColumn()
        End If



    End Sub

    Private Sub btnAutoNameFK_Click(sender As System.Object, e As System.EventArgs) Handles btnAutoNameFK.Click
        'Auto fill the Foreign Key name:
        txtNewForeignKeyName.Text = "fk_" & cmbFKTable.Text & "_" & cmbFKColumn.Text
    End Sub

    Private Sub btnDeleteRelationship_Click(sender As System.Object, e As System.EventArgs) Handles btnDeleteRelationship.Click
        'Delete the selected table relationship:

        If DataGridView2.SelectedRows.Count = 0 Then
            Main.Message.SetWarningStyle()
            Main.Message.Add("No table relationship has been selected." & vbCrLf)
            Main.Message.SetNormalStyle()
            Exit Sub
        End If

        If DataGridView2.SelectedRows.Count > 1 Then
            Main.Message.SetWarningStyle()
            Main.Message.Add("Select only one table relationship to be deleted." & vbCrLf)
            Main.Message.SetNormalStyle()
            Exit Sub
        End If

        Dim RowNo As Integer
        RowNo = DataGridView2.SelectedRows(0).Index

        Dim SelectedForeignKey As String
        SelectedForeignKey = DataGridView2.Rows(RowNo).Cells(8).Value.ToString

        Dim bldCmd As New System.Text.StringBuilder
        bldCmd.Clear()

        bldCmd.Append("ALTER TABLE " & Trim(DataGridView2.Rows(RowNo).Cells(2).Value & vbCrLf))
        bldCmd.Append("DROP CONSTRAINT " & Trim(DataGridView2.Rows(RowNo).Cells(8).Value & vbCrLf))

        Main.Message.Add("Drop Constraint SQL command: " & vbCrLf & bldCmd.ToString & vbCrLf)

        'Apply the SQL DDL command:
        Main.SqlCommandText = bldCmd.ToString 'Set the SqlCommandText property on the Database form.
        Main.ApplySqlCommand()

        FillRelationshipsGrid()

    End Sub

    Private Sub btnAutoNameIndex_Click(sender As System.Object, e As System.EventArgs) Handles btnAutoNameIndex.Click
        'Auto name the index:
        txtIndexName.Text = "idx_" & cmbNewIndexTable.Text & "_" & lstColumns.Text
    End Sub

    Private Sub btnCreateIndex_Click(sender As System.Object, e As System.EventArgs) Handles btnCreateIndex.Click
        'Create the specified index

        'http://www.w3schools.com/sql/sql_create_index.asp

        'http://msdn.microsoft.com/en-us/library/aa140015(v=office.10).aspx
        '
        Dim bldCmd As New System.Text.StringBuilder
        bldCmd.Clear()
        If chkUnique.Checked = True Then
            bldCmd.Append("CREATE UNIQUE INDEX " & Trim(txtIndexName.Text) & vbCrLf)
        Else
            bldCmd.Append("CREATE INDEX " & Trim(txtIndexName.Text) & vbCrLf)
        End If

        bldCmd.Append("ON " & Trim(cmbNewIndexTable.Text) & " (")

        Dim NColumns As Integer = lstColumns.SelectedItems.Count

        If NColumns = 0 Then
            Main.Message.SetWarningStyle()
            Main.Message.Add("No columns have been selected for an index." & vbCrLf)
            Main.Message.SetNormalStyle()
            Exit Sub
        ElseIf NColumns = 1 Then
            bldCmd.Append("[" & Trim(lstColumns.Text) & "])" & vbCrLf)
        ElseIf NColumns > 1 Then
            bldCmd.Append("[" & Trim(lstColumns.SelectedItems(0).ToString) & "]")
            Dim I As Integer
            For I = 1 To NColumns
                bldCmd.Append(", " & "[" & Trim(lstColumns.SelectedItems(I).ToString) & "]")
            Next
            bldCmd.Append(")")
        End If

        If cmbIndexOptions.Text = "" Then
            'No option selected
            bldCmd.Append(";" & vbCrLf)
        Else
            bldCmd.Append("WITH " & Trim(cmbIndexOptions.Text) & ";" & vbCrLf)
        End If

        Main.Message.Add("Create Index SQL command: " & vbCrLf & bldCmd.ToString & vbCrLf)

        'Apply the SQL DDL command:
        Main.SqlCommandText = bldCmd.ToString 'Set the SqlCommandText property on the Database form.
        Main.ApplySqlCommand()
    End Sub

    Private Sub btnDeleteIndex_Click(sender As System.Object, e As System.EventArgs) Handles btnDeleteIndex.Click
        'Delete the selected index:

        If DataGridView4.SelectedRows.Count = 0 Then
            Main.Message.SetWarningStyle()
            Main.Message.Add("No index has been selected." & vbCrLf)
            Main.Message.SetNormalStyle()
            Exit Sub
        End If

        If DataGridView4.SelectedRows.Count > 1 Then
            Main.Message.SetWarningStyle()
            Main.Message.Add("Select only one index to be deleted." & vbCrLf)
            Main.Message.SetNormalStyle()
            Exit Sub
        End If

        Dim RowNo As Integer
        RowNo = DataGridView4.SelectedRows(0).Index

        Dim SelectedIndex As String
        SelectedIndex = Trim(DataGridView4.Rows(RowNo).Cells(1).Value.ToString)
        Dim TableName As String
        TableName = Trim(DataGridView4.Rows(RowNo).Cells(0).Value.ToString)

        Dim bldCmd As New System.Text.StringBuilder
        bldCmd.Clear()

        bldCmd.Append("DROP INDEX " & SelectedIndex & vbCrLf)
        bldCmd.Append("ON " & TableName & vbCrLf)

        Main.Message.Add("Drop Index SQL command: " & vbCrLf & bldCmd.ToString & vbCrLf)

        'Apply the SQL DDL command:
        Main.SqlCommandText = bldCmd.ToString 'Set the SqlCommandText property on the Database form.
        Main.ApplySqlCommand()

        FillIndexesGrid()
    End Sub

    Private Sub btnFindTableDef_Click(sender As Object, e As EventArgs) Handles btnFindTableDef.Click
        Select Case Main.Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                OpenFileDialog1.InitialDirectory = Main.Project.DataLocn.Path
                OpenFileDialog1.Filter = "Table Definition |*.TableDef"
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    tableDefFileName = OpenFileDialog1.FileName
                    txtTableDefFileName.Text = tableDefFileName
                    tableDefXDoc = XDocument.Load(tableDefFileName)
                    ReadTableDefXDoc()
                    'Read database name, directory and description:
                    'txtDefaultName.Text = databaseDefXDoc.<DatabaseDefinition>.<Summary>.<DatabaseName>.Value
                    'txtDefaultDir.Text = databaseDefXDoc.<DatabaseDefinition>.<Summary>.<DatabaseDirectory>.Value
                    'txtDescription.Text = databaseDefXDoc.<DatabaseDefinition>.<Description>.Value
                End If
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                'Select a Database Definition file from the project archive:
                'Show the zip archive file selection form:
                Zip = New ADVL_Utilities_Library_1.ZipComp
                Zip.ArchivePath = Main.Project.DataLocn.Path
                Zip.SelectFile()
                Zip.SelectFileForm.ApplicationName = Main.Project.ApplicationName
                Zip.SelectFileForm.SettingsLocn = Main.Project.SettingsLocn
                Zip.SelectFileForm.Show()
                Zip.SelectFileForm.RestoreFormSettings()
                Zip.SelectFileForm.FileExtension = ".TableDef"
                Zip.SelectFileForm.GetFileList()
                'Process file selection in the Zip.FileSelected event.
        End Select
    End Sub

    Private Sub Zip_FileSelected(FileName As String) Handles Zip.FileSelected
        tableDefFileName = FileName
        txtTableDefFileName.Text = tableDefFileName
        Main.Project.DataLocn.ReadXmlData(FileName, tableDefXDoc)
        ReadTableDefXDoc()
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        'This code stops an unnecessary error dialog appearing.
        If e.Context = DataGridViewDataErrorContexts.Formatting Or e.Context = DataGridViewDataErrorContexts.PreferredSize Then
            e.ThrowException = False
        End If
    End Sub

    Private Sub ReadTableDefXDoc()

        DataGridView1.AllowUserToAddRows = False ''This removes the last edit row from the DataGridView.

        DataGridView1.Rows.Clear()
        Dim Database As String = tableDefXDoc.<TableDefinition>.<Summary>.<Database>.Value
        Dim TableName As String = tableDefXDoc.<TableDefinition>.<Summary>.<TableName>.Value
        txtNewTableName.Text = Trim(TableName)
        Dim NumberOfColumns As Integer = tableDefXDoc.<TableDefinition>.<Summary>.<NumberOfColumns>.Value
        Dim NumberOfPrimaryKeys As Integer = tableDefXDoc.<TableDefinition>.<Summary>.<NumberOfPrimaryKeys>.Value
        Dim PrimaryKeyColName As String
        Dim I As Integer

        Dim NRows As Integer = tableDefXDoc.<TableDefinition>.<Summary>.<NumberOfColumns>.Value
        DataGridView1.RowCount = NRows
        Dim RowNo As Integer

        For Each item In tableDefXDoc.<TableDefinition>.<Columns>.<Column>
            RowNo = item.<OrdinalPosition>.Value
            DataGridView1.Rows(RowNo - 1).Cells(0).Value = item.<ColumnName>.Value 'Write the Column Name.

            'Select Case item.<DataType>.Value
            '        'List of database data types:
            '        'http://support.microsoft.com/kb/320435

            '        'Visual Basic data types:
            '        'System.Boolean
            '        'System.Byte
            '        'System.Char
            '        'System.DateTime
            '        'System.Decimal
            '        'System.Double
            '        'System.Int16
            '        'System.Int32
            '        'System.Int64
            '        'System.Object
            '        'System.SByte
            '        'System.Single
            '        'System.String
            '        'System.UInt16
            '        'System.UInt32
            '        'System.UInt64

            Select Case item.<DataType>.Value
                Case 2 'SmallInt (Short)
                    DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Short (Integer)"
                Case 3 'Integer (Long)
                    DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Long (Integer)"
                Case 4 'Single
                    DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Single"
                Case 5 'Double
                    DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Double"
                Case 6 'Currency
                    DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Currency"
                Case 7 'Date (DateTime)
                    DataGridView1.Rows(RowNo - 1).Cells(1).Value = "DateTime"
                Case 11 'Boolean (Bit)
                    DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Bit (Boolean)"
                Case 17 'UnsignedTinyInt (Byte)
                    DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Byte"
                Case 72 'Guid (GUID)
                    DataGridView1.Rows(RowNo - 1).Cells(1).Value = "GUID"
                      'View Schema: Data Types: 
                            'Type Name  Provider Db Type    Native Data Type
                            'BigBinary  204                 128 (Column size: 4000)
                            'LongBinary 205                 128 (Column size: 1073741823)
                            'VarBinary  204                 128 (Column size: 510) (Max length parameter required)
                Case 128 'Binary
                    If item.<CharMaxLength>.Value = 4000 Then
                        DataGridView1.Rows(RowNo - 1).Cells(1).Value = "BigBinary"
                    ElseIf item.<CharMaxLength>.Value = 1073741823 Then
                        DataGridView1.Rows(RowNo - 1).Cells(1).Value = "LongBinary"
                    Else
                        DataGridView1.Rows(RowNo - 1).Cells(1).Value = "VarBinary"
                        DataGridView1.Rows(RowNo - 1).Cells(2).Value = item.<CharMaxLength>.Value
                    End If

                Case 130 'WChar
                    DataGridView1.Rows(RowNo - 1).Cells(1).Value = "VarChar"
                    DataGridView1.Rows(RowNo - 1).Cells(2).Value = item.<CharMaxLength>.Value
                Case 131 'Numeric (Decimal)
                    DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Decimal"
                    DataGridView1.Rows(RowNo - 1).Cells(3).Value = item.<Precision>.Value
                    DataGridView1.Rows(RowNo - 1).Cells(4).Value = item.<Scale>.Value
                Case Else
                    Main.Message.SetWarningStyle()
                    Main.Message.Add("Unrecognized data type: " & item.<DataType>.Value & vbCrLf)
                    Main.Message.SetNormalStyle()
            End Select

            If item.<IsNullable>.Value = "True" Then
                DataGridView1.Rows(RowNo - 1).Cells(5).Value = "Null"
            Else
                DataGridView1.Rows(RowNo - 1).Cells(5).Value = "Not Null"
            End If

            If item.<AutoIncrement>.Value = "true" Then
                DataGridView1.Rows(RowNo - 1).Cells(6).Value = "Auto Increment"
            Else
                DataGridView1.Rows(RowNo - 1).Cells(6).Value = ""
            End If

            If item.<CharMaxLength>.Value = "" Then
                DataGridView1.Rows(RowNo - 1).Cells(2).Value = ""
            Else
                DataGridView1.Rows(RowNo - 1).Cells(2).Value = item.<CharMaxLength>.Value
            End If
            DataGridView1.Rows(RowNo - 1).Cells(8).Value = item.<Description>.Value

        Next

        For Each item In tableDefXDoc.<TableDefinition>.<PrimaryKeys>.<Key>
            PrimaryKeyColName = item.Value
            For I = 1 To DataGridView1.Rows.Count
                If DataGridView1.Rows(I - 1).Cells(0).Value = PrimaryKeyColName Then
                    DataGridView1.Rows(I - 1).Cells(7).Value = "Primary Key"
                Else
                    'DataGridView1.Rows(I - 1).Cells(8).Value = "" 'Dont do this. If there are multiple keys, it will change earlier Primary Key flags.
                End If
            Next
        Next

        DataGridView1.AutoResizeColumns()
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None 'Allow used the resize columns.
        DataGridView1.AllowUserToAddRows = True 'Allow user to add rows again.

    End Sub

    Private Sub btnRenameTable_Click(sender As Object, e As EventArgs) Handles btnRenameTable.Click
        'Change the name of a Table:
        'The table is selected in cmbUtilitiesSelectTable.
        'The new name is specified in txtNewTableName2.

        If Main.DatabasePath = "" Then
            'No database selected.
            Exit Sub
        End If

        'Project \ Add Reference \ COM \ Microsoft ActiveX Data Objects 6.1 Library
        Dim aConn As New ADODB.Connection
        'Project \ Add Reference \ COM \ Microsoft ADO Ext 6.0 for DDL and Security
        Dim aDB As New ADOX.Catalog
        Dim aTable As ADOX.Table
        Dim aField As ADOX.Column


        aConn.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" + "data source = " + Main.DatabasePath
        aConn.Open()
        aDB.ActiveConnection = aConn
        aTable = aDB.Tables(cmbUtilitiesSelectTable.Text)

        'aTable.Columns(cmbUtilitiesSelectColumn.Text).Name = txtNewColumnName.Text
        aTable.Name = txtNewTableName2.Text

        aDB = Nothing
        aTable = Nothing
        aConn.Close()

        'Dim SelectedColumnIndex As Integer
        'SelectedColumnIndex = cmbUtilitiesSelectColumn.SelectedIndex
        'UpdateCmbUtilitiesSelectColumn()
        'cmbUtilitiesSelectColumn.SelectedIndex = SelectedColumnIndex

        'If cmbUtilitiesSelectTable.Text = cmbDescrSelectTable.Text Then
        '    'The Descriptions list is from the same table that has just been changed.
        '    'Redisplay the discriptions list to show the changed column name:
        '    RefreshDescriptions()
        'End If

        'Update the list of tables on the Database form:
        Main.FillLstTables()

        'Update the list of tables:
        FillCmbTableLists()
        Main.FillCmbSelectTable()

        'Update the list of columns:
        UpdateCmbUtilitiesSelectColumn()

    End Sub

    Private Sub btnMoveUp_Click(sender As Object, e As EventArgs) Handles btnMoveUp.Click
        'Move the Column Definition up one line.

        Dim RowIndex As Integer = DataGridView1.SelectedCells(0).RowIndex

        Dim TempColName As String = ""
        Dim TempType As String = ""
        Dim TempSize As String = ""
        Dim TempPrecision As String = ""
        Dim TempScale As String = ""
        Dim TempNullOrNot As String = ""
        Dim TempAutoInc As String = ""
        Dim TempPrimKey As String = ""
        Dim TempDescr As String = ""

        If RowIndex = -1 Then
            'No row selected
        ElseIf RowIndex = 0 Then
            'Already at top row.
        Else
            'Save the current row settings:
            TempColName = DataGridView1.Rows(RowIndex).Cells(0).Value
            TempType = DataGridView1.Rows(RowIndex).Cells(1).Value
            TempSize = DataGridView1.Rows(RowIndex).Cells(2).Value
            TempPrecision = DataGridView1.Rows(RowIndex).Cells(3).Value
            TempScale = DataGridView1.Rows(RowIndex).Cells(4).Value
            TempNullOrNot = DataGridView1.Rows(RowIndex).Cells(5).Value
            TempAutoInc = DataGridView1.Rows(RowIndex).Cells(6).Value
            TempPrimKey = DataGridView1.Rows(RowIndex).Cells(7).Value
            TempDescr = DataGridView1.Rows(RowIndex).Cells(8).Value

            'Move the row above down:
            DataGridView1.Rows(RowIndex).Cells(0).Value = DataGridView1.Rows(RowIndex - 1).Cells(0).Value
            DataGridView1.Rows(RowIndex).Cells(1).Value = DataGridView1.Rows(RowIndex - 1).Cells(1).Value
            DataGridView1.Rows(RowIndex).Cells(2).Value = DataGridView1.Rows(RowIndex - 1).Cells(2).Value
            DataGridView1.Rows(RowIndex).Cells(3).Value = DataGridView1.Rows(RowIndex - 1).Cells(3).Value
            DataGridView1.Rows(RowIndex).Cells(4).Value = DataGridView1.Rows(RowIndex - 1).Cells(4).Value
            DataGridView1.Rows(RowIndex).Cells(5).Value = DataGridView1.Rows(RowIndex - 1).Cells(5).Value
            DataGridView1.Rows(RowIndex).Cells(6).Value = DataGridView1.Rows(RowIndex - 1).Cells(6).Value
            DataGridView1.Rows(RowIndex).Cells(7).Value = DataGridView1.Rows(RowIndex - 1).Cells(7).Value
            DataGridView1.Rows(RowIndex).Cells(8).Value = DataGridView1.Rows(RowIndex - 1).Cells(8).Value

            'Replace the row above with the saved row:
            DataGridView1.Rows(RowIndex - 1).Cells(0).Value = TempColName
            DataGridView1.Rows(RowIndex - 1).Cells(1).Value = TempType
            DataGridView1.Rows(RowIndex - 1).Cells(2).Value = TempSize
            DataGridView1.Rows(RowIndex - 1).Cells(3).Value = TempPrecision
            DataGridView1.Rows(RowIndex - 1).Cells(4).Value = TempScale
            DataGridView1.Rows(RowIndex - 1).Cells(5).Value = TempNullOrNot
            DataGridView1.Rows(RowIndex - 1).Cells(6).Value = TempAutoInc
            DataGridView1.Rows(RowIndex - 1).Cells(7).Value = TempPrimKey
            DataGridView1.Rows(RowIndex - 1).Cells(8).Value = TempDescr

            'Move the row selection up
            DataGridView1.ClearSelection()
            DataGridView1.Rows(RowIndex - 1).Selected = True
        End If

    End Sub

    Private Sub btnMoveDown_Click(sender As Object, e As EventArgs) Handles btnMoveDown.Click
        'Move the Column Definition down one line.

        Dim RowIndex As Integer = DataGridView1.SelectedCells(0).RowIndex

        Dim TempColName As String = ""
        Dim TempType As String = ""
        Dim TempSize As String = ""
        Dim TempPrecision As String = ""
        Dim TempScale As String = ""
        Dim TempNullOrNot As String = ""
        Dim TempAutoInc As String = ""
        Dim TempPrimKey As String = ""
        Dim TempDescr As String = ""

        If RowIndex = -1 Then
            'No row selected
        ElseIf RowIndex = DataGridView1.RowCount - 1 Then
            'Already at last row
        Else
            'Save the current row settings:
            TempColName = DataGridView1.Rows(RowIndex).Cells(0).Value
            TempType = DataGridView1.Rows(RowIndex).Cells(1).Value
            TempSize = DataGridView1.Rows(RowIndex).Cells(2).Value
            TempPrecision = DataGridView1.Rows(RowIndex).Cells(3).Value
            TempScale = DataGridView1.Rows(RowIndex).Cells(4).Value
            TempNullOrNot = DataGridView1.Rows(RowIndex).Cells(5).Value
            TempAutoInc = DataGridView1.Rows(RowIndex).Cells(6).Value
            TempPrimKey = DataGridView1.Rows(RowIndex).Cells(7).Value
            TempDescr = DataGridView1.Rows(RowIndex).Cells(8).Value

            'Move the row below up:
            DataGridView1.Rows(RowIndex).Cells(0).Value = DataGridView1.Rows(RowIndex + 1).Cells(0).Value
            DataGridView1.Rows(RowIndex).Cells(1).Value = DataGridView1.Rows(RowIndex + 1).Cells(1).Value
            DataGridView1.Rows(RowIndex).Cells(2).Value = DataGridView1.Rows(RowIndex + 1).Cells(2).Value
            DataGridView1.Rows(RowIndex).Cells(3).Value = DataGridView1.Rows(RowIndex + 1).Cells(3).Value
            DataGridView1.Rows(RowIndex).Cells(4).Value = DataGridView1.Rows(RowIndex + 1).Cells(4).Value
            DataGridView1.Rows(RowIndex).Cells(5).Value = DataGridView1.Rows(RowIndex + 1).Cells(5).Value
            DataGridView1.Rows(RowIndex).Cells(6).Value = DataGridView1.Rows(RowIndex + 1).Cells(6).Value
            DataGridView1.Rows(RowIndex).Cells(7).Value = DataGridView1.Rows(RowIndex + 1).Cells(7).Value
            DataGridView1.Rows(RowIndex).Cells(8).Value = DataGridView1.Rows(RowIndex + 1).Cells(8).Value

            'Replace the row above with the saved row:
            DataGridView1.Rows(RowIndex + 1).Cells(0).Value = TempColName
            DataGridView1.Rows(RowIndex + 1).Cells(1).Value = TempType
            DataGridView1.Rows(RowIndex + 1).Cells(2).Value = TempSize
            DataGridView1.Rows(RowIndex + 1).Cells(3).Value = TempPrecision
            DataGridView1.Rows(RowIndex + 1).Cells(4).Value = TempScale
            DataGridView1.Rows(RowIndex + 1).Cells(5).Value = TempNullOrNot
            DataGridView1.Rows(RowIndex + 1).Cells(6).Value = TempAutoInc
            DataGridView1.Rows(RowIndex + 1).Cells(7).Value = TempPrimKey
            DataGridView1.Rows(RowIndex + 1).Cells(8).Value = TempDescr

            'Move the row selection down
            DataGridView1.ClearSelection()
            DataGridView1.Rows(RowIndex + 1).Selected = True

        End If


    End Sub

    Private Sub btnInsertAbove_Click(sender As Object, e As EventArgs) Handles btnInsertAbove.Click
        'Inset a new Column Definition above the current line.

        Dim RowIndex As Integer = DataGridView1.SelectedCells(0).RowIndex

        DataGridView1.Rows.Insert(RowIndex, 1)

        'Select the new row
        DataGridView1.ClearSelection()
        DataGridView1.Rows(RowIndex).Selected = True


    End Sub

    Private Sub btnInsertBelow_Click(sender As Object, e As EventArgs) Handles btnInsertBelow.Click
        'Insert a new Column Definition below the current line.

        Dim RowIndex As Integer = DataGridView1.SelectedCells(0).RowIndex

        DataGridView1.Rows.Insert(RowIndex + 1, 1)

        'Select the new row
        DataGridView1.ClearSelection()
        DataGridView1.Rows(RowIndex + 1).Selected = True

    End Sub

    Private Sub btnDeleteRow_Click(sender As Object, e As EventArgs) Handles btnDeleteRow.Click
        'Delete the selected row.

        Dim RowIndex As Integer = DataGridView1.SelectedCells(0).RowIndex

        'Select the entire row
        DataGridView1.ClearSelection()
        DataGridView1.Rows(RowIndex).Selected = True

        If MessageBox.Show("Confirm row deletion") = DialogResult.OK Then
            DataGridView1.Rows.RemoveAt(RowIndex)
        End If

    End Sub


#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Events - Events raised by this form." '-----------------------------------------------------------------------------------------------------------------------------------------

#End Region 'Form Events ----------------------------------------------------------------------------------------------------------------------------------------------------------------------




End Class