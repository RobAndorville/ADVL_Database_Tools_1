Public Class frmSaveDatabaseDefinition
    'This form is used to save a database definition.

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

            CheckFormPos()
        End If
    End Sub

    Private Sub CheckFormPos()
        'Check that the form can be seen on a screen.

        Dim MinWidthVisible As Integer = 192 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        Dim MinHeightVisible As Integer = 64 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.

        Dim FormRect As New Rectangle(Me.Left, Me.Top, Me.Width, Me.Height)
        Dim WARect As Rectangle = Screen.GetWorkingArea(FormRect) 'The Working Area rectangle - the usable area of the screen containing the form.

        ''Check if the top of the form is less than zero:
        'If Me.Top < 0 Then Me.Top = 0

        'Check if the top of the form is above the top of the Working Area:
        If Me.Top < WARect.Top Then
            Me.Top = WARect.Top
        End If

        'Check if the top of the form is too close to the bottom of the Working Area:
        If (Me.Top + MinHeightVisible) > (WARect.Top + WARect.Height) Then
            Me.Top = WARect.Top + WARect.Height - MinHeightVisible
        End If

        'Check if the left edge of the form is too close to the right edge of the Working Area:
        If (Me.Left + MinWidthVisible) > (WARect.Left + WARect.Width) Then
            Me.Left = WARect.Left + WARect.Width - MinWidthVisible
        End If

        'Check if the right edge of the form is too close to the left edge of the Working Area:
        If (Me.Left + Me.Width - MinWidthVisible) < WARect.Left Then
            Me.Left = WARect.Left - Me.Width + MinWidthVisible
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

    Private Sub frmSaveDatabaseDefinition_Load(sender As Object, e As EventArgs) Handles Me.Load
        RestoreFormSettings()   'Restore the form settings
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Me.Close() 'Close the form
    End Sub

    Private Sub frmSaveDatabaseDefinition_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub btnSaveDatabaseDefintion_Click(sender As Object, e As EventArgs) Handles btnSaveDatabaseDefintion.Click
        'Save the database defintion in an XML file:
        'DatabaseDefinition
        '   Summary (Database, Number of tables)
        '   Tables
        '       Table
        '           Summary (Table Name, Number of Columns, Number of Primary Keys)
        '           Primary Keys (Key1, ...)
        '           Columns
        '               Column (Column name, Data type, Allow db null, Auto increment, String filed length)
        '               ...
        '           /Columns
        '       /Table
        '       ...
        '   /Tables
        '   Relationships


        Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & Main.DatabasePath
        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        myConnection.ConnectionString = connString
        myConnection.Open()

        Dim TableName As String
        Dim ColumnName As String
        Dim NPrimaryKeys As Integer
        Dim NIndexFields As Integer
        Dim I As Integer

        Dim FileName As String = Trim(txtFileName.Text)
        If FileName.EndsWith(".DbDef") Then
            'FileName has the correct extension.
        Else
            'Add the file extension to FileName
            FileName = FileName & ".DbDef"
            txtFileName.Text = FileName
        End If

        Dim myTables As DataTable = Nothing
        Dim doc = New XDocument 'Create the XDocument to hold the XML data.

        'Add the XML declaration:
        Dim decl = New XDeclaration("1.0", "utf-8", "yes")
        doc.Declaration = decl

        doc.Add(New XComment(""))
        doc.Add(New XComment("Exported database definition."))

        Dim databaseData As New XElement("DatabaseDefinition")

        databaseData.Add(New XComment(""))
        databaseData.Add(New XComment("Database summary."))
        databaseData.Add(New XComment("Generated by " & Main.ApplicationInfo.Name & " application."))
        databaseData.Add(New XComment(""))

            Dim descr = New XElement("Description", txtDatabaseDescr.Text)
        databaseData.Add(descr)
        databaseData.Add(New XComment(""))

        'Get a list of the tables:
        'Restrictions:
        'Catalog    TABLE_CATALOG
        'Owner      TABLE_SCHEMA
        'Table      TABLE_NAME      eg: Customers, Items, MSysAccessStorage
        'TableType  TABLE_TYPE      eg: TABE, ACCESS TABLE, SYSTEM TABLE
        'To get the user tables, set the TableType restriction to TABLE
        myTables = myConnection.GetSchema("Tables", New String() {Nothing, Nothing, Nothing, "TABLE"})

        'Add database summary:
        Dim summary = New XElement("Summary")
        summary.Add(New XElement("DatabaseName", System.IO.Path.GetFileName(Main.DatabasePath)))
            summary.Add(New XElement("DatabaseDirectory", System.IO.Path.GetDirectoryName(Main.DatabasePath)))
            summary.Add(New XElement("NumberOfTables", myTables.Rows.Count))

            databaseData.Add(summary)
            databaseData.Add(New XComment(""))
            databaseData.Add(New XComment("List of tables."))

            'Add table data:
            Dim Row As DataRow
            Dim tables As New XElement("Tables")

            Dim Query As String
            Dim da As OleDb.OleDbDataAdapter
            Dim ds As DataSet = New DataSet
            Dim schemaTable As DataTable = New DataTable 'Used to find the Indexes in a table

            For Each Row In myTables.Rows 'Process each row (each table in the database)
                Dim table As New XElement("Table")
                'Debug.Print(Row.Item("TABLE_NAME").ToString)
                Dim myColumns As DataTable
                Dim columns As New XElement("Columns")
                myColumns = myConnection.GetSchema("Columns", New String() {Nothing, Nothing, Row.Item("TABLE_NAME").ToString})

                Dim myCol As DataColumn

            'This section of code opens the table in a DataSet. 
            TableName = Row.Item("TABLE_NAME").ToString
                table.Add(New XElement("TableName", TableName))
                table.Add(New XElement("NumberOfFields", myColumns.Rows.Count.ToString))


            'Query = "Select Top 500 * From " & TableName
            Query = "Select Top 500 * From [" & TableName & "]"
            da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey 'This statement is required to obtain the correct result from the statement: ds.Tables(0).Columns(0).MaxLength (This fixes a Microsoft bug: http://support.microsoft.com/kb/317175 )
                ds.Clear()
                ds.Reset()
                da.FillSchema(ds, SchemaType.Source, TableName)
                da.Fill(ds, TableName)
                'Dataset ds now contains the first 500 records in table with name TableName. 

                schemaTable.Clear()
                schemaTable.Reset()
                schemaTable = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Indexes, New String() {Nothing, Nothing, Nothing, Nothing, TableName})
                'schemaTable now contains the Indexes schema for the table with name TableName.

                NPrimaryKeys = ds.Tables(TableName).PrimaryKey.Count 'The number of primary keys is obtained from ds.
                table.Add(New XElement("ColumnsInPrimaryKey", ds.Tables(TableName).PrimaryKey.Count))
                Dim primaryKey = New XElement("PrimaryKey")
                For I = 1 To NPrimaryKeys
                    primaryKey.Add(New XElement("ColumnName", ds.Tables(TableName).PrimaryKey(I - 1))) 'the name of each primary key is obtained from ds.
                Next

                table.Add(New XComment(""))
                table.Add(primaryKey)

            Dim myRow As DataRowView
                Debug.Print("Table: " & Row.Item("TABLE_NAME").ToString)
                myColumns.DefaultView.Sort = "ORDINAL_POSITION ASC" 'Sort by ordinal position. By default, columns are sorted by column name.
                'DataView.Sort Property:
                'http://msdn.microsoft.com/en-au/library/system.data.dataview.sort.aspx

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
                    Dim setting6 As New XElement("AutoIncrement", ds.Tables(TableName).Columns(ColumnName).AutoIncrement)

                    'http://cislab.moorparkcollege.edu/gcampbell/advVB3-2010.htm Try this code...

                    'Find any Indexed fields:
                    Dim foundRows() As DataRow
                    foundRows = schemaTable.Select("COLUMN_NAME = '" & ColumnName & "'") 'Search for Indexes associated with with the column having the name ColumnName.
                    Dim setting7 As XElement
                    If foundRows.Count = 0 Then
                        setting7 = New XElement("Indexed", "No")
                    Else
                        If foundRows(0).Item("PRIMARY_KEY") = True Then
                            'This field is a primary key: these are indexed by default.
                            setting7 = New XElement("Indexed", "PrimaryKey")
                        Else
                            If foundRows(0).Item("UNIQUE") = True Then
                                setting7 = New XElement("Indexed", "Yes_NoDupl")
                            Else
                                setting7 = New XElement("Indexed", "Yes_DuplOk")
                            End If
                        End If
                    End If

                    Dim setting8 As New XElement("Description", myRow("DESCRIPTION").ToString)

                'Debug.Print("COLUMN_NAME " & myRow("COLUMN_NAME").ToString)
                'Debug.Print("ORDINAL_POSITION " & myRow("ORDINAL_POSITION").ToString)
                'Debug.Print("IS_NULLABLE " & myRow("IS_NULLABLE").ToString)
                'Debug.Print("DATA_TYPE " & myRow("DATA_TYPE").ToString)
                'Debug.Print("CHARACTER_MAXIMUM_LENGTH " & myRow("CHARACTER_MAXIMUM_LENGTH").ToString)
                'Debug.Print("COLUMN_FLAGS " & myRow("COLUMN_FLAGS").ToString)
                'Note: If Data_Type is 3 and Column_Flags is 90 then AutoIncrement is True???
                'http://www.vbforums.com/archive/index.php/t-621834.html
                'BUT: Long Integer with AllowDbNull is False will have DATA_TYPE = 3 and COLUMN_FLAGS = 90!!!
                'http://www.daniweb.com/software-development/vbnet/threads/373942/getoledbschematable-data_types#

                Debug.Print("")
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
                    column.Add(setting7)
                    column.Add(setting8)

                    columns.Add(column) 'Add the next column to the set of columns.
                Next 'Process the next column (the next field in the current table)

                Debug.Print("")
                table.Add(New XComment(""))
                table.Add(columns) 'Add the set of columns to the table.
                tables.Add(table) 'Add the table to the set of tables.
                tables.Add(New XComment(""))
            Next 'Process the next row (the next table in the database.

            'Fields returned by myConnection.GetSchema("Columns", New String() {Nothing, Nothing, Row.Item("TABLE_NAME").ToString})
            'TABLE_CATALOG
            'TABLE_SCHEMA
            'TABLE_NAME         The name of the table
            'COLUMN_NAME        The name of the column
            'COLUMN_GUID
            'COLUMN_PROPID
            'ORDINAL_POSITION   The order number of the column
            'COLUMN_HASDEFAULT
            'COLUMN_DEFAULT
            'COLUMN_FLAGS
            'IS_NULLABLE
            'DATA_TYPE
            'TYPE_GUID
            'CHARACTER_MAXIMUM_LENGTH
            'CHARACTER_OCTET_LENGTH
            'NUMERIC_PRECISION
            'NUMERIC_SCALE
            'DATETIME_PRECISION
            'CHARACTER_SET_CATALOG
            'CHARACTER_SET_SCHEMA
            'CHARACTER_SET_NAME
            'COLLATION_CATALOG
            'COLLATION_SCHEMA
            'COLLATION_NAME
            'DOMAIN_CATALOG
            'DOMAIN_SCHEMA
            'DOMAIN_NAME
            'DESCRIPTION

            'List of DataTypes:
            'TypeName   TypeNo  ColSize     Params              Data Type
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
            'Long is AutoIncrementable (IsAutoIncrementable = True)
            'All except BigBinary, LongBinary, VarBinary, LongText and VarChar are FixedLength (IsFixedLength = True)
            'All except Single, Double and BigBinary are Fixed Precision
            'All except Bit are nullable (IsNullable = True)

            databaseData.Add(tables) 'Add the set og tables to the databaseData

            'Get the table of table relations:
            schemaTable.Clear()
            schemaTable.Reset()
            schemaTable = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Foreign_Keys, New String() {Nothing, Nothing, Nothing, Nothing, Nothing, Nothing})

            databaseData.Add(New XComment(""))
            databaseData.Add(New XComment("List of table relationships."))

            Dim NRelationships As Integer
            NRelationships = schemaTable.Rows.Count
            databaseData.Add(New XElement("NumberOfRelationships", NRelationships))
            Dim relationships = New XElement("Relationships")

            For I = 1 To NRelationships
                Dim relationship = New XElement("relationship")
                relationship.Add(New XElement("PK_TABLE_NAME", schemaTable.Rows(I - 1).Item("PK_TABLE_NAME")))
                relationship.Add(New XElement("PK_COLUMN_NAME", schemaTable.Rows(I - 1).Item("PK_COLUMN_NAME")))
                relationship.Add(New XElement("FK_TABLE_NAME", schemaTable.Rows(I - 1).Item("FK_TABLE_NAME")))
                relationship.Add(New XElement("FK_COLUMN_NAME", schemaTable.Rows(I - 1).Item("FK_COLUMN_NAME")))
                relationships.Add(relationship)
            Next

            databaseData.Add(relationships)

            databaseData.Add(New XComment(""))

            doc.Add(databaseData) 'Add the databaseData to the XML document.

        Main.Project.SaveXmlData(FileName, doc)

        myConnection.Close()

    End Sub

    Private Sub txtFileName_TextChanged(sender As Object, e As EventArgs) Handles txtFileName.TextChanged

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub


#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class