Public Class frmSaveTableDefinition
    'This form is used to save a table definition.

#Region " Variable Declarations - All the variables used in this form and this application." '-------------------------------------------------------------------------------------------------

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

    Private Sub frmSaveTableDefinition_Load(sender As Object, e As EventArgs) Handles Me.Load
        RestoreFormSettings()   'Restore the form settings
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Me.Close() 'Close the form
    End Sub

    Private Sub frmSaveTableDefinition_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub btnSaveTableDefintion_Click(sender As Object, e As EventArgs) Handles btnSaveTableDefintion.Click

        Dim FileName As String = Trim(txtFileName.Text)
        If FileName.EndsWith(".TableDef") Then
            'FileName has the correct extension.
        Else
            'Add the file extension to FileName
            FileName = FileName & ".TableDef"
            txtFileName.Text = FileName
        End If

        Dim doc = New XDocument 'Create the XDocument to hold the XML data.

        Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & Main.DatabasePath
        Main.myConnection.ConnectionString = connString
        Main.myConnection.Open()

        'Dim myTable As DataTable = Nothing
        'myTable = Main.myConnection.GetSchema("Tables", New String() {Nothing, Nothing, Nothing, Nothing, "TABLE"})
        Dim myColumns As DataTable = Nothing
        myColumns = Main.myConnection.GetSchema("Columns", New String() {Nothing, Nothing, Main.ds.Tables(0).TableName})
        Dim ColumnName As String

        'Add the XML declaration:
        Dim decl = New XDeclaration("1.0", "utf-8", "yes")
        doc.Declaration = decl

        doc.Add(New XComment(""))
        doc.Add(New XComment("Exported table definition."))

        Dim tableData As New XElement("TableDefinition")

        tableData.Add(New XComment(""))
        tableData.Add(New XComment("Table summary."))

        'Add table summary
        Dim summary = New XElement("Summary")
        summary.Add(New XElement("Database", Main.DatabasePath))
        'summary.Add(New XElement("TableName", ds.Tables(0).TableName))
        summary.Add(New XElement("TableName", Main.ds.Tables(0).TableName))
        'summary.Add(New XElement("NumberOfColumns", ds.Tables(0).Columns.Count))
        summary.Add(New XElement("NumberOfColumns", Main.ds.Tables(0).Columns.Count))
        summary.Add(New XElement("NumberOfPrimaryKeys", Main.ds.Tables(0).PrimaryKey.Count))
        'summary.Add(New XElement("PrimaryKey", ds.Tables(0).PrimaryKey))

        tableData.Add(summary)

        tableData.Add(New XComment(""))
        tableData.Add(New XComment("Primary keys."))
        Dim NPrimaryKeys As Integer = Main.ds.Tables(0).PrimaryKey.Count
        Dim I As Integer
        Dim primaryKeys = New XElement("PrimaryKeys")
        For I = 1 To NPrimaryKeys
            primaryKeys.Add(New XElement("Key", Main.ds.Tables(0).PrimaryKey(I - 1)))
        Next

        tableData.Add(primaryKeys)

        'Add column definitions:
        'Dim ColNo As Integer
        'Dim NCols As Integer
        'NCols = Main.ds.Tables(0).Columns.Count

        Dim columns As New XElement("Columns")

        For Each myRow In myColumns.DefaultView
            Dim column As New XElement("Column")
            ColumnName = myRow("COLUMN_NAME").ToString
            Dim setting1 As New XElement("ColumnName", ColumnName)
            Dim setting2 As New XElement("OrdinalPosition", myRow("ORDINAL_POSITION").ToString)
            Dim setting3 As New XElement("IsNullable", myRow("IS_NULLABLE").ToString)
            Dim setting4 As New XElement("DataType", myRow("DATA_TYPE"))
            Dim setting4b As New XElement("DataTypeName", CType(myRow("DATA_TYPE"), OleDb.OleDbType))
            'If DATA_TYPE is Numeric, we need to obtain the Precision and Scale parameters
            'NUMERIC_PRECISION
            'NUMERIC_SCALE
            Dim setting4c As XElement
            Dim setting4d As XElement
            If myRow("DATA_TYPE") = 131 Then 'Data type is Numeric.
                setting4c = New XElement("Precision", myRow("NUMERIC_PRECISION"))
                setting4d = New XElement("Scale", myRow("NUMERIC_SCALE"))
            End If
            Dim setting5 As New XElement("CharMaxLength", myRow("CHARACTER_MAXIMUM_LENGTH").ToString)
            'Dim setting6 As New XElement("AutoIncrement", ds.Tables(TableName).Columns(ColumnName).AutoIncrement)
            Dim setting6 As New XElement("AutoIncrement", Main.ds.Tables(0).Columns(ColumnName).AutoIncrement)

            ''Find any Indexed fields:
            'Dim foundRows() As DataRow
            'foundRows = schemaTable.Select("COLUMN_NAME = '" & ColumnName & "'") 'Search for Indexes associated with with the column having the name ColumnName.
            'Dim setting7 As XElement
            'If foundRows.Count = 0 Then
            '    setting7 = New XElement("Indexed", "No")
            'Else
            '    If foundRows(0).Item("PRIMARY_KEY") = True Then
            '        'This field is a primary key: these are indexed by default.
            '        setting7 = New XElement("Indexed", "PrimaryKey")
            '    Else
            '        If foundRows(0).Item("UNIQUE") = True Then
            '            setting7 = New XElement("Indexed", "Yes_NoDupl")
            '        Else
            '            setting7 = New XElement("Indexed", "Yes_DuplOk")
            '        End If
            '    End If
            'End If

            Dim setting8 As New XElement("Description", myRow("DESCRIPTION").ToString)

            column.Add(setting1)
            column.Add(setting2)
            column.Add(setting3)
            column.Add(setting4)
            column.Add(setting4b)
            If myRow("DATA_TYPE") = 131 Then 'Data type is Numeric.
                column.Add(setting4c)
                column.Add(setting4d)
            End If
            'column.Add(setting4e)
            column.Add(setting5)
            column.Add(setting6)
            'column.Add(setting7)
            column.Add(setting8)

            columns.Add(column) 'Add the next column to the set of columns.

        Next

        'doc.Add(records)
        tableData.Add(New XComment(""))
        tableData.Add(New XComment("List of column definitions."))
        tableData.Add(columns)

        'Add Relations:
        Dim relations = New XElement("Relations")
        Dim RelCount As Integer
        'RelCount = ds.Tables(0).ChildRelations.Count
        RelCount = Main.ds.Tables(0).DataSet.Relations.Count

        relations.Add(New XElement("NumberOfChildRelations", RelCount))
        For I = 1 To RelCount
            Dim relation = New XElement("Relation")
            Dim relName = New XElement("RelationName", Main.ds.Tables(0).ChildRelations(I - 1).RelationName)
            relation.Add(relName)
            Dim childTable = New XElement("ChildTable", Main.ds.Tables(0).ChildRelations(I - 1).ChildTable)
            relation.Add(childTable)
            Dim childColumn = New XElement("ChildColumn", Main.ds.Tables(0).ChildRelations(I - 1).ChildColumns(0).ColumnName)
            relation.Add(childColumn)
            relations.Add(New XComment(""))
            relations.Add(relation)
        Next

        tableData.Add(New XComment(""))
        tableData.Add(New XComment("List of table relations."))
        tableData.Add(relations)


        doc.Add(tableData)

        'Dim SaveFilePath As String
        'SaveFilePath = SaveFileDialog1.FileName

        'If IO.Path.GetExtension = "
        'SaveFilePath = IO.Path.GetFileNameWithoutExtension(SaveFilePath) & ".TableDef"
        'SaveFilePath = IO.Path.GetFullPath(SaveFilePath) & IO.Path.GetFileNameWithoutExtension(SaveFilePath) & ".TableDef"

        'SaveFilePath = IO.Path.GetDirectoryName(SaveFilePath) & "\" & IO.Path.GetFileNameWithoutExtension(SaveFilePath) & ".TableDef"

        'doc.Save(SaveFileDialog1.FileName)
        'doc.Save(SaveFilePath)
        Main.Project.SaveXmlData(FileName, doc)

        Main.myConnection.Close()

    End Sub

    Sub SaveTableDefinition_Old()

        'Save the table defintion in an XML file:
        'TableDefinition
        '   Summary
        '   Primary Keys
        '   Columns
        '       Column1
        '       ...


        'SaveFileDialog1.Filter = "Table Definition |*.TableDef"

        'Dim FilePath As String
        'If Trim(Main.ProjectPath) <> "" Then 'Write the Form Settings file in the Project Directory
        '    FilePath = Main.ProjectPath
        'Else 'Write the Form Settings file in the Application Directory
        '    FilePath = Main.ApplicationDir
        'End If

        'SaveFileDialog1.InitialDirectory = FilePath

        Dim FileName As String = Trim(txtFileName.Text)
        If FileName.EndsWith(".TableDef") Then
            'FileName has the correct extension.
        Else
            'Add the file extension to FileName
            FileName = FileName & ".TableDef"
            txtFileName.Text = FileName
        End If


        'If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then

        'ds.Tables(0).WriteXmlSchema(SaveFileDialog1.FileName)
        'ds.Tables(0).WriteXml(SaveFileDialog1.FileName)
        'ds.Tables(0).WriteXmlSchema(SaveFileDialog1.FileName, True)

        Dim doc = New XDocument 'Create the XDocument to hold the XML data.

        'Add the XML declaration:
        Dim decl = New XDeclaration("1.0", "utf-8", "yes")
        doc.Declaration = decl

        doc.Add(New XComment(""))
        doc.Add(New XComment("Exported table definition."))

        Dim tableData As New XElement("TableDefinition")

        tableData.Add(New XComment(""))
        tableData.Add(New XComment("Table summary."))

        'Add table summary
        Dim summary = New XElement("Summary")
        summary.Add(New XElement("Database", Main.DatabasePath))
        'summary.Add(New XElement("TableName", ds.Tables(0).TableName))
        summary.Add(New XElement("TableName", Main.ds.Tables(0).TableName))
        'summary.Add(New XElement("NumberOfColumns", ds.Tables(0).Columns.Count))
        summary.Add(New XElement("NumberOfColumns", Main.ds.Tables(0).Columns.Count))
        summary.Add(New XElement("NumberOfPrimaryKeys", Main.ds.Tables(0).PrimaryKey.Count))
        'summary.Add(New XElement("PrimaryKey", ds.Tables(0).PrimaryKey))

        tableData.Add(summary)

        tableData.Add(New XComment(""))
        tableData.Add(New XComment("Primary keys."))
        Dim NPrimaryKeys As Integer = Main.ds.Tables(0).PrimaryKey.Count
        Dim I As Integer
        Dim primaryKeys = New XElement("PrimaryKeys")
        For I = 1 To NPrimaryKeys
            primaryKeys.Add(New XElement("Key", Main.ds.Tables(0).PrimaryKey(I - 1)))
        Next

        tableData.Add(primaryKeys)

        'Add column definitions:
        Dim ColNo As Integer
        Dim NCols As Integer
        NCols = Main.ds.Tables(0).Columns.Count
        Dim columns = New XElement("Columns")
        For ColNo = 1 To NCols
            Dim column = New XElement("Column")
            'For ColNo = 1 To NCols
            'Dim field = New XElement(FieldName(ColNo - 1), TDS_Finances.ViewTables.DataGridView1.Rows(RowNo - 1).Cells(ColNo - 1).Value.ToString)
            'record.Add(field)
            'Next
            Dim setting1 = New XElement("ColumnName", Main.ds.Tables(0).Columns(ColNo - 1).ColumnName)
            column.Add(setting1)
            Dim setting2 = New XElement("DataType", Main.ds.Tables(0).Columns(ColNo - 1).DataType)
            column.Add(setting2)
            'Dim setting3 = New XElement("MaxLength", ds.Tables(0).Columns(ColNo - 1).MaxLength) 'This returns -1
            'column.Add(setting3)
            'Dim setting3 = New XElement("Capton", ds.Tables(0).Columns(ColNo - 1).Caption)
            'column.Add(setting3)
            'Dim setting8 = New XElement("Attributes", ds.Tables(0).Columns(ColNo - 1).DataType.Attributes)
            'column.Add(setting8)
            Dim setting4 = New XElement("AllowDBNull", Main.ds.Tables(0).Columns(ColNo - 1).AllowDBNull)
            column.Add(setting4)
            Dim setting5 = New XElement("AutoIncrement", Main.ds.Tables(0).Columns(ColNo - 1).AutoIncrement)
            column.Add(setting5)
            'Dim setting6 = New XElement("StringFieldLength", Main.ds.Tables(0).Columns(ColNo - 1).MaxLength)
            Dim setting6 = New XElement("MaxLength", Main.ds.Tables(0).Columns(ColNo - 1).MaxLength)
            column.Add(setting6)
            'Dim settings7 = New XElement("MaxLen", ds.Tables(0).Columns(ColNo - 1).MaxLength)
            columns.Add(column)

            'Dim setting7 = New XElement("Test", Main.ds.Tables(0).Columns(ColNo - 1).ExtendedProperties())
        Next

        'doc.Add(records)
        tableData.Add(New XComment(""))
        tableData.Add(New XComment("List of column definitions."))
        tableData.Add(columns)

        'Add Relations:
        Dim relations = New XElement("Relations")
        Dim RelCount As Integer
        'RelCount = ds.Tables(0).ChildRelations.Count
        RelCount = Main.ds.Tables(0).DataSet.Relations.Count

        relations.Add(New XElement("NumberOfChildRelations", RelCount))
        For I = 1 To RelCount
            Dim relation = New XElement("Relation")
            Dim relName = New XElement("RelationName", Main.ds.Tables(0).ChildRelations(I - 1).RelationName)
            relation.Add(relName)
            Dim childTable = New XElement("ChildTable", Main.ds.Tables(0).ChildRelations(I - 1).ChildTable)
            relation.Add(childTable)
            Dim childColumn = New XElement("ChildColumn", Main.ds.Tables(0).ChildRelations(I - 1).ChildColumns(0).ColumnName)
            relation.Add(childColumn)
            relations.Add(New XComment(""))
            relations.Add(relation)
        Next

        tableData.Add(New XComment(""))
        tableData.Add(New XComment("List of table relations."))
        tableData.Add(relations)


        doc.Add(tableData)

        'Dim SaveFilePath As String
        'SaveFilePath = SaveFileDialog1.FileName

        'If IO.Path.GetExtension = "
        'SaveFilePath = IO.Path.GetFileNameWithoutExtension(SaveFilePath) & ".TableDef"
        'SaveFilePath = IO.Path.GetFullPath(SaveFilePath) & IO.Path.GetFileNameWithoutExtension(SaveFilePath) & ".TableDef"

        'SaveFilePath = IO.Path.GetDirectoryName(SaveFilePath) & "\" & IO.Path.GetFileNameWithoutExtension(SaveFilePath) & ".TableDef"

        'doc.Save(SaveFileDialog1.FileName)
        'doc.Save(SaveFilePath)
        Main.Project.SaveXmlData(FileName, doc)

        'Else

        'End If


    End Sub



#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class