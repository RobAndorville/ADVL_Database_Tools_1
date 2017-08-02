Public Class frmCreateNewTable
    'This form is used to create a new table in the selected database.

#Region " Variable Declarations - All the variables used in this form and this application." '-------------------------------------------------------------------------------------------------

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

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        RestoreFormSettings()   'Restore the form settings

        'Set up DataGridView1:
        'DataGridView1.ColumnCount = 5
        'DataGridView1.Columns(0).HeaderText = "Name"
        'DataGridView1.Columns(1).HeaderText = "Primary Key"
        'DataGridView1.Columns(2).HeaderText = "Data Type"
        'DataGridView1.Columns(3).HeaderText = "String Length"
        'DataGridView1.Columns(4).HeaderText = "Allow DB Null"
        'DataGridView1.Columns(5).HeaderText = "Auto Increment"

        DataGridView1.ColumnCount = 1
        DataGridView1.Columns(0).HeaderText = "Name"
        Dim dgvCheckPrimKey As New DataGridViewCheckBoxColumn
        DataGridView1.Columns.Add(dgvCheckPrimKey)
        DataGridView1.Columns(1).HeaderText = "Primary Key"
        Dim dgvCombo As New DataGridViewComboBoxColumn
        dgvCombo.HeaderText = "Data Type"
        DataGridView1.Columns.Add(dgvCombo)
        'DataGridView1.Columns(2).HeaderText = "Primary Key"
        DataGridView1.Columns.Add("StringLength", "String Length")
        Dim dgvCheckAllowNull As New DataGridViewCheckBoxColumn
        DataGridView1.Columns.Add(dgvCheckAllowNull)
        DataGridView1.Columns(4).HeaderText = "Allow DB Null"
        Dim dgvCheckAutoInc As New DataGridViewCheckBoxColumn
        DataGridView1.Columns.Add(dgvCheckAutoInc)
        DataGridView1.Columns(5).HeaderText = "AutoIncrement"

        'dgvCombo.Items.Add("VarChar")
        'dgvCombo.Items.Add("Bit")
        'dgvCombo.Items.Add("Date")
        'dgvCombo.Items.Add("Int")
        dgvCombo.Items.Add("String")
        'dgvCombo.Items.Add("Integer")
        dgvCombo.Items.Add("Short (16 bit Integer)")
        dgvCombo.Items.Add("Long (32 bit Integer)")
        dgvCombo.Items.Add("Single")
        dgvCombo.Items.Add("Double")
        dgvCombo.Items.Add("Numeric")
        dgvCombo.Items.Add("Currency")
        'dgvCombo.Items.Add("Date")
        dgvCombo.Items.Add("DateTime")
        dgvCombo.Items.Add("Boolean")
        dgvCombo.Items.Add("Bit")
        dgvCombo.Items.Add("Byte")
        dgvCombo.Items.Add("GUID")


    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Me.Close() 'Close the form
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub btnReadTableDefFile_Click(sender As Object, e As EventArgs) Handles btnReadTableDefFile.Click
        Select Case Main.Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                OpenFileDialog1.InitialDirectory = Main.Project.DataLocn.Path
                OpenFileDialog1.Filter = "Table Definition |*.TableDef"
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    tableDefFileName = OpenFileDialog1.FileName
                    txtTableDefFile.Text = tableDefFileName
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
        txtTableDefFile.Text = tableDefFileName
        Main.Project.DataLocn.ReadXmlData(FileName, tableDefXDoc)
        ReadTableDefXDoc()
    End Sub

    Private Sub ReadTableDefXDoc()

        DataGridView1.AllowUserToAddRows = False ''This removes the last edit row from the DataGridView.

        DataGridView1.Rows.Clear()
        Dim Database As String = tableDefXDoc.<TableDefinition>.<Summary>.<Database>.Value
        Dim TableName As String = tableDefXDoc.<TableDefinition>.<Summary>.<TableName>.Value
        'txtNewTableName.Text = Trim(TableName)
        txtTableName.Text = Trim(TableName)
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
                    'DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Short (Integer)"
                    DataGridView1.Rows(RowNo - 1).Cells(2).Value = "Short (Integer)"
                Case 3 'Integer (Long)
                    'DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Long (Integer)"
                    DataGridView1.Rows(RowNo - 1).Cells(2).Value = "Long (Integer)"
                Case 4 'Single
                    'DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Single"
                    DataGridView1.Rows(RowNo - 1).Cells(2).Value = "Single"
                Case 5 'Double
                    'DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Double"
                    DataGridView1.Rows(RowNo - 1).Cells(2).Value = "Double"
                Case 6 'Currency
                    'DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Currency"
                    DataGridView1.Rows(RowNo - 1).Cells(2).Value = "Currency"
                Case 7 'Date (DateTime)
                    'DataGridView1.Rows(RowNo - 1).Cells(1).Value = "DateTime"
                    DataGridView1.Rows(RowNo - 1).Cells(2).Value = "DateTime"
                Case 11 'Boolean (Bit)
                    'DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Bit (Boolean)"
                    DataGridView1.Rows(RowNo - 1).Cells(2).Value = "Bit (Boolean)"
                Case 17 'UnsignedTinyInt (Byte)
                    'DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Byte"
                    DataGridView1.Rows(RowNo - 1).Cells(2).Value = "Byte"
                Case 72 'Guid (GUID)
                    'DataGridView1.Rows(RowNo - 1).Cells(1).Value = "GUID"
                    DataGridView1.Rows(RowNo - 1).Cells(2).Value = "GUID"
                      'View Schema: Data Types: 
                            'Type Name  Provider Db Type    Native Data Type
                            'BigBinary  204                 128 (Column size: 4000)
                            'LongBinary 205                 128 (Column size: 1073741823)
                            'VarBinary  204                 128 (Column size: 510) (Max length parameter required)
                Case 128 'Binary
                    If item.<CharMaxLength>.Value = 4000 Then
                        'DataGridView1.Rows(RowNo - 1).Cells(1).Value = "BigBinary"
                        DataGridView1.Rows(RowNo - 1).Cells(2).Value = "BigBinary"
                    ElseIf item.<CharMaxLength>.Value = 1073741823 Then
                        'DataGridView1.Rows(RowNo - 1).Cells(1).Value = "LongBinary"
                        DataGridView1.Rows(RowNo - 1).Cells(2).Value = "LongBinary"
                    Else
                        'DataGridView1.Rows(RowNo - 1).Cells(1).Value = "VarBinary"
                        DataGridView1.Rows(RowNo - 1).Cells(2).Value = "VarBinary"
                        'DataGridView1.Rows(RowNo - 1).Cells(2).Value = item.<CharMaxLength>.Value
                        DataGridView1.Rows(RowNo - 1).Cells(3).Value = item.<CharMaxLength>.Value
                    End If

                Case 130 'WChar
                    'DataGridView1.Rows(RowNo - 1).Cells(1).Value = "VarChar"
                    DataGridView1.Rows(RowNo - 1).Cells(2).Value = "VarChar"
                    'DataGridView1.Rows(RowNo - 1).Cells(2).Value = item.<CharMaxLength>.Value
                    DataGridView1.Rows(RowNo - 1).Cells(3).Value = item.<CharMaxLength>.Value
                Case 131 'Numeric (Decimal)
                    'DataGridView1.Rows(RowNo - 1).Cells(1).Value = "Decimal"
                    DataGridView1.Rows(RowNo - 1).Cells(2).Value = "Decimal"
                    'DataGridView1.Rows(RowNo - 1).Cells(3).Value = item.<Precision>.Value
                    '   DataGridView1.Rows(RowNo - 1).Cells(3).Value = item.<Precision>.Value
                    'DataGridView1.Rows(RowNo - 1).Cells(4).Value = item.<Scale>.Value
                    '   DataGridView1.Rows(RowNo - 1).Cells(4).Value = item.<Scale>.Value
                Case Else
                    Main.Message.SetWarningStyle()
                    Main.Message.Add("Unrecognized data type: " & item.<DataType>.Value & vbCrLf)
                    Main.Message.SetNormalStyle()
            End Select

            If item.<IsNullable>.Value = "True" Then
                'DataGridView1.Rows(RowNo - 1).Cells(5).Value = "Null"
                DataGridView1.Rows(RowNo - 1).Cells(4).Value = True
            Else
                'DataGridView1.Rows(RowNo - 1).Cells(5).Value = "Not Null"
                DataGridView1.Rows(RowNo - 1).Cells(4).Value = False
            End If

            If item.<AutoIncrement>.Value = "true" Then
                'DataGridView1.Rows(RowNo - 1).Cells(6).Value = "Auto Increment"
                DataGridView1.Rows(RowNo - 1).Cells(5).Value = True
            Else
                'DataGridView1.Rows(RowNo - 1).Cells(6).Value = ""
                DataGridView1.Rows(RowNo - 1).Cells(5).Value = False
            End If

            If item.<CharMaxLength>.Value = "" Then
                'DataGridView1.Rows(RowNo - 1).Cells(2).Value = ""
                DataGridView1.Rows(RowNo - 1).Cells(3).Value = ""
            Else
                'DataGridView1.Rows(RowNo - 1).Cells(2).Value = item.<CharMaxLength>.Value
                DataGridView1.Rows(RowNo - 1).Cells(3).Value = item.<CharMaxLength>.Value
            End If
            'DataGridView1.Rows(RowNo - 1).Cells(8).Value = item.<Description>.Value
            'DataGridView1.Rows(RowNo - 1).Cells(8).Value = item.<Description>.Value

        Next

        For Each item In tableDefXDoc.<TableDefinition>.<PrimaryKeys>.<Key>
            PrimaryKeyColName = item.Value
            For I = 1 To DataGridView1.Rows.Count
                If DataGridView1.Rows(I - 1).Cells(0).Value = PrimaryKeyColName Then
                    'DataGridView1.Rows(I - 1).Cells(7).Value = "Primary Key"
                    DataGridView1.Rows(I - 1).Cells(1).Value = True
                Else
                    'DataGridView1.Rows(I - 1).Cells(8).Value = "" 'Dont do this. If there are multiple keys, it will change earlier Primary Key flags.
                End If
            Next
        Next

        DataGridView1.AutoResizeColumns()
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None 'Allow used the resize columns.
        DataGridView1.AllowUserToAddRows = True 'Allow user to add rows again.

    End Sub

    Private Sub ReadTableDefXDoc_Old()
        DataGridView1.Rows.Clear()

        Dim Database As String = tableDefXDoc.<TableDefinition>.<Summary>.<Database>.Value
        Dim TableName As String = tableDefXDoc.<TableDefinition>.<Summary>.<TableName>.Value
        txtTableName.Text = Trim(TableName)
        Dim NumberOfColumns As Integer = tableDefXDoc.<TableDefinition>.<Summary>.<NumberOfColumns>.Value
        Dim NumberOfPrimaryKeys As Integer = tableDefXDoc.<TableDefinition>.<Summary>.<NumberOfPrimaryKeys>.Value

        Dim PrimaryKey As String
        Dim I As Integer

        Dim RowCount As Integer
        For Each item In tableDefXDoc.<TableDefinition>.<Columns>.<Column>
            DataGridView1.Rows.Add()
            RowCount = DataGridView1.Rows.Count
            DataGridView1.Rows(RowCount - 2).Cells(0).Value = item.<ColumnName>.Value

            Select Case item.<DataType>.Value
                    'List of database data types:
                    'http://support.microsoft.com/kb/320435

                    'Visual Basic data types:
                    'System.Boolean
                    'System.Byte
                    'System.Char
                    'System.DateTime
                    'System.Decimal
                    'System.Double
                    'System.Int16
                    'System.Int32
                    'System.Int64
                    'System.Object
                    'System.SByte
                    'System.Single
                    'System.String
                    'System.UInt16
                    'System.UInt32
                    'System.UInt64

                Case "System.String"
                    DataGridView1.Rows(RowCount - 2).Cells(2).Value = "String"
                Case "System.Int16"
                    DataGridView1.Rows(RowCount - 2).Cells(2).Value = "Short (16 bit Integer)"
                'Case "System.Int32"
                '    DataGridView1.Rows(RowCount - 2).Cells(2).Value = "Integer"
                Case "System.Int32"
                    DataGridView1.Rows(RowCount - 2).Cells(2).Value = "Long (32 bit Integer)"
                Case "System.Single"
                    DataGridView1.Rows(RowCount - 2).Cells(2).Value = "Single"
                Case "System.Double"
                    DataGridView1.Rows(RowCount - 2).Cells(2).Value = "Double"
                'Case "System.Decimal"
                '    DataGridView1.Rows(RowCount - 2).Cells(2).Value = "Numeric"
                Case "System.Decimal"
                    DataGridView1.Rows(RowCount - 2).Cells(2).Value = "Currency"
                'Case "System.DateTime"
                '    DataGridView1.Rows(RowCount - 2).Cells(2).Value = "Date"
                Case "System.DateTime"
                    DataGridView1.Rows(RowCount - 2).Cells(2).Value = "DateTime"
                Case "System.Boolean"
                    DataGridView1.Rows(RowCount - 2).Cells(2).Value = "Boolean"
                Case "System.Boolean"
                    DataGridView1.Rows(RowCount - 2).Cells(2).Value = "Bit"
                Case "System.Byte"
                    DataGridView1.Rows(RowCount - 2).Cells(2).Value = "Byte"
                Case "System.Guid"
                    DataGridView1.Rows(RowCount - 2).Cells(2).Value = "GUID"





                Case Else
                    Main.Message.SetWarningStyle()
                    Main.Message.Add("Unrecognized data type: " & item.<DataType>.Value & vbCrLf)
                    Main.Message.SetNormalStyle()
            End Select

            If item.<AllowDBNull>.Value = "true" Then
                DataGridView1.Rows(RowCount - 2).Cells(4).Value = vbTrue
            End If

            If item.<AutoIncrement>.Value = "true" Then
                DataGridView1.Rows(RowCount - 2).Cells(5).Value = True
            End If

            'If item.<StringFieldLength>.Value = "-1" Then
            If item.<MaxLength>.Value = "-1" Then
                DataGridView1.Rows(RowCount - 2).Cells(3).Value = ""
            Else
                'DataGridView1.Rows(RowCount - 2).Cells(3).Value = item.<StringFieldLength>.Value
                DataGridView1.Rows(RowCount - 2).Cells(3).Value = item.<MaxLength>.Value
            End If
        Next

        For Each item In tableDefXDoc.<TableDefinition>.<PrimaryKeys>.<Key>
            'PrimaryKeys(I) = item.Value
            'I = I + 1
            PrimaryKey = item.Value
            For I = 1 To DataGridView1.Rows.Count
                If DataGridView1.Rows(I - 1).Cells(0).Value = PrimaryKey Then
                    DataGridView1.Rows(I - 1).Cells(1).Value = True
                End If
            Next
        Next
    End Sub

    Private Sub btnCreateTable_Click(sender As Object, e As EventArgs) Handles btnCreateTable.Click
        'Create a new table with the specified fields:

        'Data types:
        ' http://msdn.microsoft.com/en-us/library/aa157100(office.10).aspx 
        ' http://support.microsoft.com/kb/320435 

        ' http://www.vb6.us/tutorials/access-sql 
        ' http://www.techotopia.com/index.php/Creating_Databases_and_Tables_Using_SQL_Commands 


        Dim connString As String
        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        'Dim ds As DataSet = New DataSet
        'Dim da As OleDb.OleDbDataAdapter
        'Dim tables As DataTableCollection = ds.Tables

        'Dim dt As DataTable


        'TableName = cmbSelectTable.SelectedItem.ToString

        'txtQuery.Text = "Select * From " & TableName

        'NOTE: Check that a database has been created before adding a table!!!

        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & Main.DatabasePath
        myConnection.ConnectionString = connString
        myConnection.Open()
        'da = New OleDb.OleDbDataAdapter("Select * from " & TableName, myConnection)


        Dim bldCmd As New System.Text.StringBuilder
        Dim RowNo As Integer
        Dim RowCount As Integer = DataGridView1.Rows.Count - 1

        'transactionid INT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
        'salesman CHAR(20) NOT NULL,
        'amount FLOAT

        'customer_id int NOT NULL AUTO_INCREMENT,
        'customer_name char(20) NOT NULL,
        'PRIMARY KEY (customer_id)

        'Img_ID INT NOT NULL AUTO_INCREMENT PRIMARY KEY,

        'Img_Id INT NOT NULL AUTO_INCREMENT FOREIGN KEY,


        bldCmd.Append("Create Table " & txtTableName.Text & " (")
        For RowNo = 1 To RowCount
            Select Case DataGridView1.Rows(RowNo - 1).Cells(2).Value
                Case "String"
                    bldCmd.Append(DataGridView1.Rows(RowNo - 1).Cells(0).Value & " VarChar(" & DataGridView1.Rows(RowNo - 1).Cells(3).Value & ")")
                    If DataGridView1.Rows(RowNo - 1).Cells(4).Value = False Then
                        bldCmd.Append(" Not Null")
                    Else
                        bldCmd.Append(" Null")
                    End If
                    If DataGridView1.Rows(RowNo - 1).Cells(1).Value = True Then
                        bldCmd.Append(" Primary Key")
                    End If
                'Case "Date"
                Case "DateTime"
                    bldCmd.Append(DataGridView1.Rows(RowNo - 1).Cells(0).Value & " DateTime")
                    If DataGridView1.Rows(RowNo - 1).Cells(4).Value = False Then
                        bldCmd.Append(" Not Null")
                    Else
                        bldCmd.Append(" Null")
                    End If
                    If DataGridView1.Rows(RowNo - 1).Cells(1).Value = True Then
                        bldCmd.Append(" Primary Key")
                    End If
                Case "Integer"
                    bldCmd.Append(DataGridView1.Rows(RowNo - 1).Cells(0).Value & " Integer")
                    If DataGridView1.Rows(RowNo - 1).Cells(4).Value = False Then
                        bldCmd.Append(" Not Null")
                    Else
                        bldCmd.Append(" Null")
                    End If
                    If DataGridView1.Rows(RowNo - 1).Cells(5).Value = True Then
                        bldCmd.Append(" Auto_Increment")
                    End If
                    If DataGridView1.Rows(RowNo - 1).Cells(1).Value = True Then
                        bldCmd.Append(" Primary Key")
                    End If
                Case "Numeric"
                    bldCmd.Append(DataGridView1.Rows(RowNo - 1).Cells(0).Value & " Float")
                    If DataGridView1.Rows(RowNo - 1).Cells(4).Value = False Then
                        bldCmd.Append(" Not Null")
                    Else
                        bldCmd.Append(" Null")
                    End If
                    If DataGridView1.Rows(RowNo - 1).Cells(1).Value = True Then
                        bldCmd.Append(" Primary Key")
                    End If
                Case "Boolean"
                    bldCmd.Append(DataGridView1.Rows(RowNo - 1).Cells(0).Value & " Boolean")
                    If DataGridView1.Rows(RowNo - 1).Cells(4).Value = False Then
                        bldCmd.Append(" Not Null")
                    Else
                        bldCmd.Append(" Null")
                    End If
                    If DataGridView1.Rows(RowNo - 1).Cells(1).Value = True Then
                        bldCmd.Append(" Primary Key")
                    End If

            End Select

            If RowNo < RowCount Then
                'Add a comma after the field definition:
                bldCmd.Append(", ")
            Else
                'At last field definition
                'Add closing bracket:
                bldCmd.Append(")")
            End If
        Next

        Main.Message.Add("Create Table command: " & bldCmd.ToString & vbCrLf)

        Dim cmd As New OleDb.OleDbCommand
        cmd.CommandText = bldCmd.ToString
        cmd.Connection = myConnection
        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Main.Message.SetWarningStyle()
            Main.Message.Add("Error creating new table: " & ex.Message & vbCrLf)
            Main.Message.SetNormalStyle()
        End Try

        myConnection.Close()
    End Sub


#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class