Public Class frmCreateNewDatabase
    'Form used to create a new database.

#Region " Variable Declarations - All the variables used in this form and this application." '-------------------------------------------------------------------------------------------------

    Dim databaseDefFileName As String 'The Database deinition file name.
    Dim databaseDefXDoc As System.Xml.Linq.XDocument 'The database definition XDocument.
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
                               <NewDatabaseName><%= txtNewDatabaseName.Text %></NewDatabaseName>
                               <NewDatabaseDirectory><%= txtNewDatabaseDir.Text %></NewDatabaseDirectory>
                               <DBDefFilePath><%= txtDefinitionFilePath.Text %></DBDefFilePath>
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
            If Settings.<FormSettings>.<NewDatabaseName>.Value <> Nothing Then txtNewDatabaseName.Text = Settings.<FormSettings>.<NewDatabaseName>.Value
            If Settings.<FormSettings>.<NewDatabaseDirectory>.Value <> Nothing Then txtNewDatabaseDir.Text = Settings.<FormSettings>.<NewDatabaseDirectory>.Value
            If Settings.<FormSettings>.<DBDefFilePath>.Value <> Nothing Then
                databaseDefFileName = Settings.<FormSettings>.<DBDefFilePath>.Value
                txtDefinitionFilePath.Text = databaseDefFileName
                databaseDefXDoc = XDocument.Load(databaseDefFileName)
                'Read database name, directory and description:
                txtDefaultName.Text = databaseDefXDoc.<DatabaseDefinition>.<Summary>.<DatabaseName>.Value
                txtDefaultDir.Text = databaseDefXDoc.<DatabaseDefinition>.<Summary>.<DatabaseDirectory>.Value
                txtDescription.Text = databaseDefXDoc.<DatabaseDefinition>.<Description>.Value
            End If

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

    Private Sub frmCreateNewDatabase_Load(sender As Object, e As EventArgs) Handles Me.Load
        RestoreFormSettings()   'Restore the form settings
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form

        Me.Close() 'Close the form
    End Sub

    Private Sub frmCreateNewDatabase_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        If txtNewDatabaseDir.Text = "" Then
            FolderBrowserDialog1.SelectedPath = My.Computer.FileSystem.SpecialDirectories.MyDocuments.ToString()
        Else
            FolderBrowserDialog1.SelectedPath = txtNewDatabaseDir.Text
        End If

        FolderBrowserDialog1.ShowDialog()
        txtNewDatabaseDir.Text = FolderBrowserDialog1.SelectedPath
    End Sub

    Private Sub btnFind_Click(sender As Object, e As EventArgs) Handles btnFind.Click
        'Find a databse definition file in the current project.

        Select Case Main.Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                OpenFileDialog1.InitialDirectory = Main.Project.DataLocn.Path
                OpenFileDialog1.Filter = "Database Definition |*.DbDef"
                OpenFileDialog1.FileName = ""
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    databaseDefFileName = OpenFileDialog1.FileName
                    txtDefinitionFilePath.Text = databaseDefFileName
                    databaseDefXDoc = XDocument.Load(databaseDefFileName)
                    'Read database name, directory and description:
                    txtDefaultName.Text = databaseDefXDoc.<DatabaseDefinition>.<Summary>.<DatabaseName>.Value
                    txtDefaultDir.Text = databaseDefXDoc.<DatabaseDefinition>.<Summary>.<DatabaseDirectory>.Value
                    txtDescription.Text = databaseDefXDoc.<DatabaseDefinition>.<Description>.Value
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
                Zip.SelectFileForm.FileExtension = ".DbDef"
                Zip.SelectFileForm.GetFileList()
                'Process file selection in the Zip.FileSelected event.
        End Select

    End Sub

    Private Sub Zip_FileSelected(FileName As String) Handles Zip.FileSelected
        databaseDefFileName = FileName
        txtDefinitionFilePath.Text = databaseDefFileName
        Main.Project.DataLocn.ReadXmlData(FileName, databaseDefXDoc)

        'Read database name, directory and description:
        txtDefaultName.Text = databaseDefXDoc.<DatabaseDefinition>.<Summary>.<DatabaseName>.Value
        txtDefaultDir.Text = databaseDefXDoc.<DatabaseDefinition>.<Summary>.<DatabaseDirectory>.Value
        txtDescription.Text = databaseDefXDoc.<DatabaseDefinition>.<Description>.Value
    End Sub

    Private Sub CreateNewDatabase()
        'Create a new database using a .DbDef database definition file.

        Dim bldCmd As New System.Text.StringBuilder
        Dim NColsInPrimaryKey As Integer = 0 'The number of column in the primary key.

        Try
            'First add the reference: Project \ Add Reference \ Assemblies \ Extensions \ Microsoft.Office.Interop.Access.Dao Version 15.0.0.0
            Dim AccessDbEngine As New Microsoft.Office.Interop.Access.Dao.DBEngine
            Dim AccessDb As Microsoft.Office.Interop.Access.Dao.Database
            txtNewDatabaseName.Text = Trim(txtNewDatabaseName.Text)
            If txtNewDatabaseName.Text.EndsWith(".accdb") Then
            Else
                txtNewDatabaseName.Text = txtNewDatabaseName.Text & ".accdb"
            End If
            AccessDb = AccessDbEngine.CreateDatabase(txtNewDatabaseDir.Text & "\" & txtNewDatabaseName.Text, Microsoft.Office.Interop.Access.Dao.LanguageConstants.dbLangGeneral, Microsoft.Office.Interop.Access.Dao.DatabaseTypeEnum.dbVersion120)

            If Trim(txtDefinitionFilePath.Text) = "" Then
                'No database definition file specified
                'Leave the database blank.
                AccessDb.Close()
                AccessDb = Nothing
                AccessDbEngine = Nothing
                Main.Message.Add("No database definition file specified. " & bldCmd.ToString & vbCrLf)
                Main.Message.Add("A blank database file has been created. " & bldCmd.ToString & vbCrLf & vbCrLf)
                Exit Sub
            End If

            'Add the tables to the database:
            Dim AtFirstColumn As Boolean
            For Each item In databaseDefXDoc.<DatabaseDefinition>.<Tables>.<Table>
                bldCmd.Clear()
                bldCmd.Append("Create Table " & item.<TableName>.Value & " (")
                NColsInPrimaryKey = item.<ColumnsInPrimaryKey>.Value

                If NColsInPrimaryKey > 1 Then
                    Dim NFields As Integer = item.<NumberOfFields>.Value
                    'Dim PKeys(0 To NColsInPrimaryKey - 1) As Integer 'Array to hold the number of each primary key Field.
                    Dim PKeys(0 To NColsInPrimaryKey - 1) As String 'Array to hold the name of each primary key Field.
                    Dim PKeyNo As Integer = 0
                    Dim ColNumber As Integer = 1 'Start at the first column
                    Dim J As Integer

                    For Each colItem In item.<Columns>.<Column>
                        If ColNumber > 1 Then
                            'Add a comma to end previous column statement
                            bldCmd.Append(", ")
                        End If
                        Select Case colItem.<DataType>.Value
                        'Datatype 2: SmallInt
                        'Datatype 3: Integer
                        'Datatype 4: Single
                        'Datatype 5: Double
                        'Datatype 6: Currency
                        'Datatype 7: Date
                        'Datatype 11: Boolean (Yes No)
                        'Datatype 17: UnsignedTinyInt
                        'Datatype 72: Guid
                        'Datatype 128: Binary (OLE Object)
                        'Datatype 130: WChar
                        'Datatype 131: Numeric (Decimal)

                            Case 2 'SmallInt (Short)
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " SHORT")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    PKeys(PKeyNo) = colItem.<ColumnName>.Value 'Save the primary key column name in PKeys()
                                    PKeyNo += 1 'Increment the Primary Key number
                                End If

                            Case 3 'Integer (Long)
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " LONG")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    PKeys(PKeyNo) = colItem.<ColumnName>.Value 'Save the primary key column name in PKeys()
                                    PKeyNo += 1 'Increment the Primary Key number
                                End If

                            Case 4 'Single
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " SINGLE")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    PKeys(PKeyNo) = colItem.<ColumnName>.Value 'Save the primary key column name in PKeys()
                                    PKeyNo += 1 'Increment the Primary Key number
                                End If

                            Case 5 'Double
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " DOUBLE")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    PKeys(PKeyNo) = colItem.<ColumnName>.Value 'Save the primary key column name in PKeys()
                                    PKeyNo += 1 'Increment the Primary Key number
                                End If

                            Case 6 'Currency
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " CURRENCY")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    PKeys(PKeyNo) = colItem.<ColumnName>.Value 'Save the primary key column name in PKeys()
                                    PKeyNo += 1 'Increment the Primary Key number
                                End If

                            Case 7 'Date (DateTime)
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " DATETIME")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    PKeys(PKeyNo) = colItem.<ColumnName>.Value 'Save the primary key column name in PKeys()
                                    PKeyNo += 1 'Increment the Primary Key number
                                End If

                            Case 11 'Boolean (Bit)
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " BIT")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    PKeys(PKeyNo) = colItem.<ColumnName>.Value 'Save the primary key column name in PKeys()
                                    PKeyNo += 1 'Increment the Primary Key number
                                End If

                            Case 17 'UnsignedTinyInt (Byte)
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " BYTE")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    PKeys(PKeyNo) = colItem.<ColumnName>.Value 'Save the primary key column name in PKeys()
                                    PKeyNo += 1 'Increment the Primary Key number
                                End If

                            Case 72 'Guid (GUID)
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " GUID")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    PKeys(PKeyNo) = colItem.<ColumnName>.Value 'Save the primary key column name in PKeys()
                                    PKeyNo += 1 'Increment the Primary Key number
                                End If

                            'View Schema: Data Types: 
                            'Type Name  Provider Db Type    Native Data Type
                            'BigBinary  204                 128 (Column size: 4000)
                            'LongBinary 205                 128 (Column size: 1073741823)
                            'VarBinary  204                 128 (Column size: 510) (Max length parameter required)
                            Case 128 'Binary
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " BINARY")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    PKeys(PKeyNo) = colItem.<ColumnName>.Value 'Save the primary key column name in PKeys()
                                    PKeyNo += 1 'Increment the Primary Key number
                                End If

                            'View Schema: Data Types: 
                            'Type Name  Provider Db Type    Native Data Type
                            'LongText   203                 130 (Column size: 536870910)
                            'VarChar    202                 130 (Column size: 255) (Max length parameter required)
                            Case 130 'WChar
                                If colItem.<CharMaxLength>.Value = 0 Then
                                    'bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " VARCHAR(1)")
                                    bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " LONGTEXT")
                                Else
                                    bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " VARCHAR(" & colItem.<CharMaxLength>.Value & ")")
                                End If

                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    PKeys(PKeyNo) = colItem.<ColumnName>.Value 'Save the primary key column name in PKeys()
                                    PKeyNo += 1 'Increment the Primary Key number
                                End If

                            Case 131 'Numeric (Decimal)
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " DECIMAL")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    PKeys(PKeyNo) = colItem.<ColumnName>.Value 'Save the primary key column name in PKeys()
                                    PKeyNo += 1 'Increment the Primary Key number
                                End If

                        End Select

                        If ColNumber = NFields Then 'At the last column.
                            bldCmd.Append("," & vbCrLf) 'End the line.
                            bldCmd.Append("CONSTRAINT MultiColPrimKey PRIMARY KEY (")
                            For J = 0 To NColsInPrimaryKey - 2
                                bldCmd.Append(PKeys(J) & ", ")
                            Next
                            bldCmd.Append(PKeys(NColsInPrimaryKey - 1) & ")")
                        End If
                        ColNumber += 1 'Increment the column number.
                    Next

                    bldCmd.Append(")")

                    Main.Message.Add(" " & vbCrLf)
                    Main.Message.Add("Create Table command: " & bldCmd.ToString & vbCrLf)

                    Try
                        AccessDb.Execute(bldCmd.ToString)
                    Catch ex As Exception
                        Main.Message.SetWarningStyle()
                        Main.Message.Add("Error creating new table: " & ex.Message & vbCrLf)
                        Main.Message.SetNormalStyle()
                    End Try


                Else 'No primary key or primary key with one column.
                    AtFirstColumn = True
                    For Each colItem In item.<Columns>.<Column>
                        If AtFirstColumn = False Then
                            'Add a comma to end previous column statement
                            bldCmd.Append(", ")
                        End If
                        Select Case colItem.<DataType>.Value
                        'Datatype 2: SmallInt
                        'Datatype 3: Integer
                        'Datatype 4: Single
                        'Datatype 5: Double
                        'Datatype 6: Currency
                        'Datatype 7: Date
                        'Datatype 11: Boolean (Yes No)
                        'Datatype 17: UnsignedTinyInt
                        'Datatype 72: Guid
                        'Datatype 128: Binary (OLE Object)
                        'Datatype 130: WChar
                        'Datatype 131: Numeric (Decimal)

                            Case 2 'SmallInt (Short)
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " SHORT")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    bldCmd.Append(" Primary Key")
                                End If

                            Case 3 'Integer (Long)
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " LONG")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    bldCmd.Append(" Primary Key")
                                End If

                            Case 4 'Single
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " SINGLE")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    bldCmd.Append(" Primary Key")
                                End If

                            Case 5 'Double
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " DOUBLE")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    bldCmd.Append(" Primary Key")
                                End If

                            Case 6 'Currency
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " CURRENCY")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    bldCmd.Append(" Primary Key")
                                End If

                            Case 7 'Date (DateTime)
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " DATETIME")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    bldCmd.Append(" Primary Key")
                                End If

                            Case 11 'Boolean (Bit)
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " BIT")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    bldCmd.Append(" Primary Key")
                                End If

                            Case 17 'UnsignedTinyInt (Byte)
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " BYTE")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    bldCmd.Append(" Primary Key")
                                End If

                            Case 72 'Guid (GUID)
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " GUID")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    bldCmd.Append(" Primary Key")
                                End If

                            'View Schema: Data Types: 
                            'Type Name  Provider Db Type    Native Data Type
                            'BigBinary  204                 128 (Column size: 4000)
                            'LongBinary 205                 128 (Column size: 1073741823)
                            'VarBinary  204                 128 (Column size: 510) (Max length parameter required)
                            Case 128 'Binary
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " BINARY")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    bldCmd.Append(" Primary Key")
                                End If

                            'View Schema: Data Types: 
                            'Type Name  Provider Db Type    Native Data Type
                            'LongText   203                 130 (Column size: 536870910)
                            'VarChar    202                 130 (Column size: 255) (Max length parameter required)
                            Case 130 'WChar
                                If colItem.<CharMaxLength>.Value = 0 Then
                                    bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " VARCHAR(1)")
                                Else
                                    bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " VARCHAR(" & colItem.<CharMaxLength>.Value & ")")
                                End If

                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    bldCmd.Append(" Primary Key")
                                End If

                            Case 131 'Numeric (Decimal)
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " DECIMAL")
                                If colItem.<IsNullable>.Value = "false" Then
                                    bldCmd.Append(" Not Null")
                                Else

                                End If
                                If colItem.<Indexed>.Value = "PrimaryKey" Then
                                    bldCmd.Append(" Primary Key")
                                End If

                        End Select
                        AtFirstColumn = False
                    Next

                    bldCmd.Append(")")

                    Main.Message.Add(" " & vbCrLf)
                    Main.Message.Add("Create Table command: " & bldCmd.ToString & vbCrLf)

                    Try
                        AccessDb.Execute(bldCmd.ToString)
                    Catch ex As Exception
                        Main.Message.SetWarningStyle()
                        Main.Message.Add("Error creating new table: " & ex.Message & vbCrLf)
                        Main.Message.SetNormalStyle()
                    End Try
                End If 'NColsInPrimaryKey > 1 
            Next

            AccessDb.Close()
            AccessDb = Nothing
            AccessDbEngine = Nothing

        Catch ex As Exception
            Main.Message.AddWarning("Error creating new database. " & ex.Message & vbCrLf)
        End Try

    End Sub

    Private Sub btnCreateNewDatabase_Click(sender As Object, e As EventArgs) Handles btnCreateNewDatabase.Click


        CreateNewDatabase()

        Exit Sub

        'OLD CODE VERSION:

        Dim bldCmd As New System.Text.StringBuilder

        'Creat a new database:
        ' http://social.msdn.microsoft.com/Forums/en/vbgeneral/thread/5711484a-1c5b-4550-aada-d3c849d08d58 


        Try
            'NOTE: This propject was created using Visual Basic 2010.
            ' When it was opened in Visual Basic 2013, Microsoft.Office.Interop.Access.Dao was not recognised.
            ' In Solution Explorer \ References the Microsoft.Office.interop.access.dao reference shows a warning symbol.
            ' The Microsoft Office 2010 Primary Interop Assemblies Redistributable was downloaded and installed.
            ' This did not fix the problem.
            ' The Microsoft.Office.interop.access.dao reference was deleted then added again.
            ' The problem was fixed. (I am not sure if the redistributable needed to be installed.)

            'First add the reference: Project \ Add Reference \ .NET \ Microsoft.Office.interop.access.dao
            'UPDATE: First add the reference: Project \ Add Reference \ Assemblies \ Extensions \ Microsoft.Office.Interop.Access.Dao Version 15.0.0.0
            Dim AccessDbEngine As New Microsoft.Office.Interop.Access.Dao.DBEngine
            Dim AccessDb As Microsoft.Office.Interop.Access.Dao.Database
            txtNewDatabaseName.Text = Trim(txtNewDatabaseName.Text)
            If txtNewDatabaseName.Text.EndsWith(".accdb") Then
            Else
                txtNewDatabaseName.Text = txtNewDatabaseName.Text & ".accdb"
            End If
            AccessDb = AccessDbEngine.CreateDatabase(txtNewDatabaseDir.Text & "\" & txtNewDatabaseName.Text, Microsoft.Office.Interop.Access.Dao.LanguageConstants.dbLangGeneral, Microsoft.Office.Interop.Access.Dao.DatabaseTypeEnum.dbVersion120)

            If Trim(txtDefinitionFilePath.Text) = "" Then
                'No database definition file specified
                'Leave the database blank.
                AccessDb.Close()
                AccessDb = Nothing
                AccessDbEngine = Nothing
                Main.Message.Add("No database definition file specified. " & bldCmd.ToString & vbCrLf)
                Main.Message.Add("A blank database file has been created. " & bldCmd.ToString & vbCrLf & vbCrLf)
                Exit Sub
            End If

            'Add the tables to the database:
            'Use: databaseDefXDoc
            Dim AtFirstColumn As Boolean
            'For Each item In databaseDefinitionData.<DatabaseDefinition>.<Tables>.<Table>
            For Each item In databaseDefXDoc.<DatabaseDefinition>.<Tables>.<Table>
                bldCmd.Clear()
                bldCmd.Append("Create Table " & item.<TableName>.Value & " (")

                AtFirstColumn = True
                For Each colItem In item.<Columns>.<Column>
                    If AtFirstColumn = False Then
                        'Add a comma to end previous column statement
                        bldCmd.Append(", ")
                    End If
                    Select Case colItem.<DataType>.Value
                        'Datatype 2: SmallInt
                        'Datatype 3: Integer
                        'Datatype 4: Single
                        'Datatype 5: Double
                        'Datatype 6: Currency
                        'Datatype 7: Date
                        'Datatype 11: Boolean (Yes No)
                        'Datatype 17: UnsignedTinyInt
                        'Datatype 72: Guid
                        'Datatype 128: Binary (OLE Object)
                        'Datatype 130: WChar
                        'Datatype 131: Numeric (Decimal)

                        Case 2 'SmallInt (Short)
                            bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " SHORT")
                            If colItem.<IsNullable>.Value = "false" Then
                                bldCmd.Append(" Not Null")
                            Else

                            End If
                            If colItem.<Indexed>.Value = "PrimaryKey" Then
                                bldCmd.Append(" Primary Key")
                            End If

                        Case 3 'Integer (Long)
                            bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " LONG")
                            If colItem.<IsNullable>.Value = "false" Then
                                bldCmd.Append(" Not Null")
                            Else

                            End If
                            If colItem.<Indexed>.Value = "PrimaryKey" Then
                                bldCmd.Append(" Primary Key")
                            End If

                        Case 4 'Single
                            bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " SINGLE")
                            If colItem.<IsNullable>.Value = "false" Then
                                bldCmd.Append(" Not Null")
                            Else

                            End If
                            If colItem.<Indexed>.Value = "PrimaryKey" Then
                                bldCmd.Append(" Primary Key")
                            End If

                        Case 5 'Double
                            bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " DOUBLE")
                            If colItem.<IsNullable>.Value = "false" Then
                                bldCmd.Append(" Not Null")
                            Else

                            End If
                            If colItem.<Indexed>.Value = "PrimaryKey" Then
                                bldCmd.Append(" Primary Key")
                            End If

                        Case 6 'Currency
                            bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " CURRENCY")
                            If colItem.<IsNullable>.Value = "false" Then
                                bldCmd.Append(" Not Null")
                            Else

                            End If
                            If colItem.<Indexed>.Value = "PrimaryKey" Then
                                bldCmd.Append(" Primary Key")
                            End If

                        Case 7 'Date (DateTime)
                            bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " DATETIME")
                            If colItem.<IsNullable>.Value = "false" Then
                                bldCmd.Append(" Not Null")
                            Else

                            End If
                            If colItem.<Indexed>.Value = "PrimaryKey" Then
                                bldCmd.Append(" Primary Key")
                            End If

                        Case 11 'Boolean (Bit)
                            bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " BIT")
                            If colItem.<IsNullable>.Value = "false" Then
                                bldCmd.Append(" Not Null")
                            Else

                            End If
                            If colItem.<Indexed>.Value = "PrimaryKey" Then
                                bldCmd.Append(" Primary Key")
                            End If

                        Case 17 'UnsignedTinyInt (Byte)
                            bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " BYTE")
                            If colItem.<IsNullable>.Value = "false" Then
                                bldCmd.Append(" Not Null")
                            Else

                            End If
                            If colItem.<Indexed>.Value = "PrimaryKey" Then
                                bldCmd.Append(" Primary Key")
                            End If

                        Case 72 'Guid (GUID)
                            bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " GUID")
                            If colItem.<IsNullable>.Value = "false" Then
                                bldCmd.Append(" Not Null")
                            Else

                            End If
                            If colItem.<Indexed>.Value = "PrimaryKey" Then
                                bldCmd.Append(" Primary Key")
                            End If

                            'View Schema: Data Types: 
                            'Type Name  Provider Db Type    Native Data Type
                            'BigBinary  204                 128 (Column size: 4000)
                            'LongBinary 205                 128 (Column size: 1073741823)
                            'VarBinary  204                 128 (Column size: 510) (Max length parameter required)
                        Case 128 'Binary
                            bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " BINARY")
                            If colItem.<IsNullable>.Value = "false" Then
                                bldCmd.Append(" Not Null")
                            Else

                            End If
                            If colItem.<Indexed>.Value = "PrimaryKey" Then
                                bldCmd.Append(" Primary Key")
                            End If

                            'View Schema: Data Types: 
                            'Type Name  Provider Db Type    Native Data Type
                            'LongText   203                 130 (Column size: 536870910)
                            'VarChar    202                 130 (Column size: 255) (Max length parameter required)
                        Case 130 'WChar
                            If colItem.<CharMaxLength>.Value = 0 Then
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " VARCHAR(1)")
                            Else
                                bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " VARCHAR(" & colItem.<CharMaxLength>.Value & ")")
                            End If

                            If colItem.<IsNullable>.Value = "false" Then
                                bldCmd.Append(" Not Null")
                            Else

                            End If
                            If colItem.<Indexed>.Value = "PrimaryKey" Then
                                bldCmd.Append(" Primary Key")
                            End If

                        Case 131 'Numeric (Decimal)
                            bldCmd.Append("[" & colItem.<ColumnName>.Value & "]" & " DECIMAL")
                            If colItem.<IsNullable>.Value = "false" Then
                                bldCmd.Append(" Not Null")
                            Else

                            End If
                            If colItem.<Indexed>.Value = "PrimaryKey" Then
                                bldCmd.Append(" Primary Key")
                            End If

                    End Select
                    AtFirstColumn = False
                Next

                bldCmd.Append(")")

                Main.Message.Add(" " & vbCrLf)
                Main.Message.Add("Create Table command: " & bldCmd.ToString & vbCrLf)

                Try
                    AccessDb.Execute(bldCmd.ToString)
                Catch ex As Exception
                    Main.Message.SetWarningStyle()
                    Main.Message.Add("Error creating new table: " & ex.Message & vbCrLf)
                    Main.Message.SetNormalStyle()
                End Try
            Next

            'Read database name, directory and description:
            txtDefaultName.Text = databaseDefXDoc.<DatabaseDefinition>.<Summary>.<DatabaseName>.Value
            AccessDb.Close()
            AccessDb = Nothing
            AccessDbEngine = Nothing
        Catch ex As Exception
            Main.Message.SetWarningStyle()
            Main.Message.Add("Error creating new database. " & ex.Message & vbCrLf)
            Main.Message.SetNormalStyle()
        End Try
    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class