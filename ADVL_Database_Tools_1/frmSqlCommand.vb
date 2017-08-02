Public Class frmSqlCommand
    'Form used to modify the database using SQL commands.

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

    Private Sub frmSqlCommand_Load(sender As Object, e As EventArgs) Handles Me.Load

        RestoreFormSettings()   'Restore the form settings

        'Set up DataGridView1:
        Dim TextBoxCol0 As New DataGridViewTextBoxColumn
        DataGridView1.Columns.Add(TextBoxCol0)
        DataGridView1.Columns(0).HeaderText = "Column Name"
        DataGridView1.Columns(0).Width = 160
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

        FillCmbSelectTable()

        cmbAddColumnType.Items.Add("Short (Integer)")
        cmbAddColumnType.Items.Add("Long (Integer)")
        cmbAddColumnType.Items.Add("Single")
        cmbAddColumnType.Items.Add("Double")
        cmbAddColumnType.Items.Add("Currency")
        cmbAddColumnType.Items.Add("DateTime")
        cmbAddColumnType.Items.Add("Bit (Boolean)")
        cmbAddColumnType.Items.Add("Byte")
        cmbAddColumnType.Items.Add("GUID")
        cmbAddColumnType.Items.Add("BigBinary")
        cmbAddColumnType.Items.Add("LongBinary")
        cmbAddColumnType.Items.Add("VarBinary")
        cmbAddColumnType.Items.Add("LongText")
        cmbAddColumnType.Items.Add("VarChar")
        cmbAddColumnType.Items.Add("Decimal")

        cmbAddColumnNull.Items.Add("")
        cmbAddColumnNull.Items.Add("Null")
        cmbAddColumnNull.Items.Add("Not Null")

        cmbAddColumnAutoInc.Items.Add("")
        cmbAddColumnAutoInc.Items.Add("Auto Increment")

        cmbAddColumnPrimaryKey.Items.Add("")
        cmbAddColumnPrimaryKey.Items.Add("Primary Key")

        cmbAlterColumnType.Items.Add("Short (Integer)")
        cmbAlterColumnType.Items.Add("Long (Integer)")
        cmbAlterColumnType.Items.Add("Single")
        cmbAlterColumnType.Items.Add("Double")
        cmbAlterColumnType.Items.Add("Currency")
        cmbAlterColumnType.Items.Add("DateTime")
        cmbAlterColumnType.Items.Add("Bit (Boolean)")
        cmbAlterColumnType.Items.Add("Byte")
        cmbAlterColumnType.Items.Add("GUID")
        cmbAlterColumnType.Items.Add("BigBinary")
        cmbAlterColumnType.Items.Add("LongBinary")
        cmbAlterColumnType.Items.Add("VarBinary")
        cmbAlterColumnType.Items.Add("LongText")
        cmbAlterColumnType.Items.Add("VarChar")
        cmbAlterColumnType.Items.Add("Decimal")

        cmbAlterColumnNull.Items.Add("")
        cmbAlterColumnNull.Items.Add("Null")
        cmbAlterColumnNull.Items.Add("Not Null")

        cmbAlterColumnAutoInc.Items.Add("")
        cmbAlterColumnAutoInc.Items.Add("Auto Increment")

        cmbAlterColumnPrimaryKey.Items.Add("")
        cmbAlterColumnPrimaryKey.Items.Add("Primary Key")

        cmbCreateIndexUnique.Items.Add("")
        cmbCreateIndexUnique.Items.Add("Unique")

        cmbCreateIndexHandleNull.Items.Add("")
        cmbCreateIndexHandleNull.Items.Add("Primary")
        cmbCreateIndexHandleNull.Items.Add("Disallow Null")
        cmbCreateIndexHandleNull.Items.Add("Ignore Null")

        cmbConstraintType.Items.Add("")
        cmbConstraintType.Items.Add("Not Null")
        cmbConstraintType.Items.Add("Default")
        cmbConstraintType.Items.Add("Unique")
        cmbConstraintType.Items.Add("Check")
        cmbConstraintType.Items.Add("Primary Key")
        cmbConstraintType.Items.Add("Foreign Key")


    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form

        Me.Close() 'Close the form
    End Sub

    Private Sub frmSqlCommand_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        Dim Row As Integer = e.RowIndex
        Dim Col As Integer = e.ColumnIndex

        If Col = -1 Then Exit Sub

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

        If Col = 1 Then 'Type

            'Column Name   Type   Size   Precision   Scale   Null/Not Null   Auto Increment   Primary Key

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
                    'DataGridView1.Rows(Row).Cells(6).Value = "" 'Enable Auto Increment
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
                    'DataGridView1.Rows(Row).Cells(2).Value = "" 'Enable Size
                    'DataGridView1.Rows(Row).Cells(2).ReadOnly = False
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
                    'DataGridView1.Rows(Row).Cells(2).Value = "" 'Enable Size
                    'DataGridView1.Rows(Row).Cells(2).ReadOnly = False
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
                    'DataGridView1.Rows(Row).Cells(2).Value = "" 'Enable Size
                    'DataGridView1.Rows(Row).Cells(2).ReadOnly = False
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

    Private Sub btnCreateCommand_Click(sender As System.Object, e As System.EventArgs) Handles btnCreateCommand.Click

        If TabControl1.SelectedTab.Text = "Create Table" Then
            GenerateCreateTableCommand()
        ElseIf TabControl1.SelectedTab.Text = "Alter Table" Then
            GenerateAlterTableCommand()
        End If
    End Sub

    Private Sub GenerateCreateTableCommand()
        'Generate the Create Table command:

        If Trim(txtTableName.Text) = "" Then
            Main.Message.SetWarningStyle()
            Main.Message.Add("Table name not specified." & vbCrLf)
            Main.Message.SetNormalStyle()
            Exit Sub
        End If

        Dim NPrimaryKeys As Integer = 0 'The number of primary key columns
        Dim LastRow As Integer = DataGridView1.RowCount - 1
        Dim I As Integer

        'Count the number of primary keys:
        For I = 1 To LastRow
            If Trim(DataGridView1.Rows(I - 1).Cells(7).Value) = "Primary Key" Then
                NPrimaryKeys += 1
            End If
        Next

        Dim bldCmd As New System.Text.StringBuilder
        bldCmd.Clear()
        bldCmd.Append("CREATE TABLE " & Trim(txtTableName.Text) & vbCrLf)
        'Dim I As Integer
        'Dim LastRow As Integer
        Dim DataTypeString As String
        'LastRow = DataGridView1.RowCount - 1

        If NPrimaryKeys > 1 Then
            Dim J As Integer
            'Dim PKeys(0 To NPrimaryKeys) As Integer
            Dim PKeys(0 To NPrimaryKeys - 1) As Integer
            'Dim PKeyNo As Integer = 1
            Dim PKeyNo As Integer = 0
            For I = 1 To LastRow
                If I = 1 Then
                    bldCmd.Append("    (")
                Else
                    bldCmd.Append("    ")
                End If
                GetDataTypeString(I - 1, DataTypeString)
                bldCmd.Append(Trim(DataGridView1.Rows(I - 1).Cells(0).Value) & " " & DataTypeString) 'Add the new column name

                'Add Null / Not Null:
                If Trim(DataGridView1.Rows(I - 1).Cells(5).Value) <> "" Then
                    bldCmd.Append(" " & UCase(DataGridView1.Rows(I - 1).Cells(5).Value))
                End If

                'Add Auto Increment:
                If Trim(DataGridView1.Rows(I - 1).Cells(6).Value) = "Auto Increment" Then
                    bldCmd.Append(" IDENTITY(1,1)")
                End If

                'Add any Primary Keys to the PKeys() list:
                If Trim(DataGridView1.Rows(I - 1).Cells(7).Value) = "Primary Key" Then
                    'bldCmd.Append(" PRIMARY KEY")
                    'PKeys(PKeyNo) = I 'Add the position of the primary key to the PKeys() list
                    PKeys(PKeyNo) = I - 1 'Add the position of the primary key to the PKeys() list
                    PKeyNo += 1 'Increment PKeyNo
                End If

                If I = LastRow Then
                    'Add the multi-column primary key contraint:
                    bldCmd.Append("," & vbCrLf) 'end the line
                    bldCmd.Append("CONSTRAINT MultiColPrimKey PRIMARY KEY (")
                    'For J = 1 To NPrimaryKeys - 1
                    'For J = 0 To NPrimaryKeys - 1
                    For J = 0 To NPrimaryKeys - 2
                        bldCmd.Append(Trim(DataGridView1.Rows(PKeys(J)).Cells(0).Value) & ", ")
                    Next
                    'bldCmd.Append(Trim(DataGridView1.Rows(PKeys(NPrimaryKeys)).Cells(0).Value) & ")")
                    bldCmd.Append(Trim(DataGridView1.Rows(PKeys(NPrimaryKeys - 1)).Cells(0).Value) & ")")


                    bldCmd.Append(");" & vbCrLf) 'close the brackets containing the column specifications
                Else
                    bldCmd.Append("," & vbCrLf) 'end the line
                End If



            Next
        Else
            For I = 1 To LastRow
                If I = 1 Then
                    bldCmd.Append("    (")
                Else
                    bldCmd.Append("    ")
                End If
                GetDataTypeString(I - 1, DataTypeString)
                bldCmd.Append(Trim(DataGridView1.Rows(I - 1).Cells(0).Value) & " " & DataTypeString) 'Add the new column name

                'Add Null / Not Null:
                If Trim(DataGridView1.Rows(I - 1).Cells(5).Value) <> "" Then
                    bldCmd.Append(" " & UCase(DataGridView1.Rows(I - 1).Cells(5).Value))
                End If

                'Add Auto Increment:
                If Trim(DataGridView1.Rows(I - 1).Cells(6).Value) = "Auto Increment" Then
                    bldCmd.Append(" IDENTITY(1,1)")
                End If

                'Add Primary Key:
                If Trim(DataGridView1.Rows(I - 1).Cells(7).Value) = "Primary Key" Then
                    bldCmd.Append(" PRIMARY KEY")
                End If

                If I = LastRow Then
                    bldCmd.Append(");" & vbCrLf) 'close the brackets containing the column specifications
                Else
                    bldCmd.Append("," & vbCrLf) 'end the line
                End If

            Next
        End If

        txtCommand.AppendText(bldCmd.ToString)

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

    Private Sub GenerateAlterTableCommand()
        'Generates the Alter Table command:

        Dim bldCmd As New System.Text.StringBuilder
        bldCmd.Clear()

        If Trim(cmbSelectTable.Text) = "" Then
            Main.Message.SetWarningStyle()
            Main.Message.Add("A table has not been selected!." & vbCrLf)
            Main.Message.SetNormalStyle()
            Exit Sub
        End If

        If rbDropTable.Checked = True Then
            bldCmd.Append("DROP TABLE " & Trim(cmbSelectTable.Text & ";") & vbCrLf)

        ElseIf rbAddColumn.Checked = True Then
            Main.Message.Add("Add column selected." & vbCrLf)
            bldCmd.Append("ALTER TABLE " & Trim(cmbSelectTable.Text) & vbCrLf)
            bldCmd.Append("    ADD COLUMN [" & Trim(txtAddColumnName.Text) & "]") 'Add the column name

            Select Case cmbAddColumnType.Text
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
                    bldCmd.Append(" VARBINARY (" & txtAddColumnSize.Text & ")")
                Case "LongText"
                    bldCmd.Append(" LONGTEXT")
                Case "VarChar"
                    bldCmd.Append(" VARCHAR (" & txtAddColumnSize.Text & ")")
                Case "Decimal"
                    bldCmd.Append(" DECIMAL (" & txtAddColumnPrecision.Text & ", " & txtAddColumnScale.Text & ")")

            End Select

            'Add Null / Not Null:
            If Trim(cmbAddColumnNull.Text) <> "" Then
                bldCmd.Append(" " & UCase(cmbAddColumnNull.Text))
            End If

            'Add Auto Increment:
            If Trim(cmbAddColumnAutoInc.Text) = "Auto Increment" Then
                bldCmd.Append(" IDENTITY(1,1)")
            End If

            'Add Primary Key:
            If Trim(cmbAddColumnPrimaryKey.Text) = "Primary Key" Then
                bldCmd.Append(" PRIMARY KEY")
            End If

            bldCmd.Append(");" & vbCrLf)


        ElseIf rbAlterColumn.Checked = True Then
            bldCmd.Append("ALTER TABLE " & Trim(cmbSelectTable.Text) & vbCrLf)
            bldCmd.Append("    ALTER COLUMN " & Trim(cmbAlterColumnName.Text)) 'Add the column name

            'Add Null / Not Null:
            If Trim(cmbAlterColumnNull.Text) <> "" Then
                bldCmd.Append(" " & UCase(cmbAlterColumnNull.Text))
            End If

            'Add Auto Increment:
            If Trim(cmbAlterColumnAutoInc.Text) = "Auto Increment" Then
                bldCmd.Append(" IDENTITY(1,1)")
            End If

            'Add Primary Key:
            If Trim(cmbAlterColumnPrimaryKey.Text) = "Primary Key" Then
                bldCmd.Append(" PRIMARY KEY")
            End If

            bldCmd.Append(");" & vbCrLf)

        ElseIf rbDropColumn.Checked = True Then
            bldCmd.Append("ALTER TABLE " & Trim(cmbSelectTable.Text) & vbCrLf)
            bldCmd.Append("    DROP COLUMN " & Trim(cmbDropColumnName.Text)) 'Add the column name
            bldCmd.Append(");" & vbCrLf)

        ElseIf rbCreateIndex.Checked = True Then
            If Trim(cmbCreateIndexUnique.Text) = "Unique" Then
                bldCmd.Append("CREATE UNIQUE INDEX ")
            Else
                bldCmd.Append("CREATE INDEX ")
            End If

            bldCmd.Append(txtIndexName.Text & vbCrLf) 'Add index name

            bldCmd.Append("    ON " & cmbSelectTable.Text) 'Add table name

            'Add column name(s):
            Dim NCols As Integer
            NCols = lbCreateIndexColumnName.SelectedItems.Count
            If NCols = 0 Then

            ElseIf NCols = 1 Then
                bldCmd.Append(" ([" & lbCreateIndexColumnName.SelectedItem.ToString & "])" & vbCrLf)
            ElseIf NCols > 1 Then

            End If

            If Trim(cmbCreateIndexHandleNull.Text) = "" Then
                bldCmd.Append(";" & vbCrLf)
            ElseIf cmbCreateIndexHandleNull.Text = "Disallow Null" Then
                bldCmd.Append("    WITH DISALLOW NULL;" & vbCrLf)
            ElseIf cmbCreateIndexHandleNull.Text = "Ignore Null" Then
                bldCmd.Append("    WITH IGNORE NULL;" & vbCrLf)
            ElseIf cmbCreateIndexHandleNull.Text = "Primary" Then
                bldCmd.Append("    WITH PRIMARY;" & vbCrLf)

            End If


        ElseIf rbDropIndex.Checked = True Then
            Dim IndexName As String
            Dim SelRow As Integer
            SelRow = DataGridView2.SelectedCells(0).RowIndex
            IndexName = DataGridView2.Rows(SelRow).Cells(0).Value
            Dim ColumnName As String
            ColumnName = DataGridView2.Rows(SelRow).Cells(1).Value

            bldCmd.Append("DROP INDEX " & IndexName & " ON " & ColumnName & ";" & vbCrLf)

        ElseIf rbAddForeignKey.Checked = True Then
            'http://www.w3schools.com/sql/sql_foreignkey.asp
            bldCmd.Append("ALTER TABLE " & Trim(cmbSelectTable.Text) & vbCrLf)
            bldCmd.Append("    ADD FOREIGN KEY (" & cmbForeignKeyColumnName.Text & ")" & vbCrLf)
            bldCmd.Append("    REFERENCES " & cmbRelatedTable.Text & " (" & cmbPrimaryKey.Text & ");" & vbCrLf)

        ElseIf rbAddConstraint.Checked = True Then
            Main.Message.Add("Add Constraint selected." & vbCrLf)


            'Examples: http://www.w3schools.com/sql/sql_unique.asp
            'To create a UNIQUE constraint on the "P_Id" column when the table is already created, use the following SQL:
            'ALTER TABLE Persons
            'ADD UNIQUE (P_Id)
            '
            'To allow naming of a UNIQUE constraint, and for defining a UNIQUE constraint on multiple columns, use the following SQL syntax:
            'ALTER TABLE Persons
            'ADD CONSTRAINT uc_PersonID UNIQUE (P_Id,LastName)
            '
            '
            '
            '
            '
            '
            '
            '
            '
            '
            '

        End If

        txtCommand.AppendText(bldCmd.ToString)

    End Sub

    Private Sub FillCmbSelectTable()
        'Fill the cmbSelectTable listbox with the availalble tables in the selected database.

        If Main.DatabasePath = "" Then
            'No database selected.
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        cmbSelectTable.Items.Clear()

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
        Dim I As Integer 'Loop index
        Dim MaxI As Integer

        MaxI = dt.Rows.Count
        For I = 0 To MaxI - 1
            'dr = dt.Rows(0)
            cmbSelectTable.Items.Add(dt.Rows(I).Item(2).ToString)
        Next I

        conn.Close()

    End Sub

    Private Sub btnClearCommand_Click(sender As System.Object, e As System.EventArgs) Handles btnClearCommand.Click
        txtCommand.Text = ""
    End Sub

    Private Sub btnApplyCommand_Click(sender As System.Object, e As System.EventArgs) Handles btnApplyCommand.Click
        'Apply the SQL Command:
        Main.SqlCommandText = txtCommand.Text 'Set the SqlCommandText property on the Database form.
        Main.ApplySqlCommand()
    End Sub

    Private Sub FillCmbFields()
        'Fill the Column Name ComboBoxes with the availalble fields in the selected table.

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

        If cmbSelectTable.SelectedIndex = -1 Then 'No item is selected
            'lstFields.Items.Clear()
            cmbAlterColumnName.Items.Clear()
            cmbDropColumnName.Items.Clear()
        Else 'A table has been selected. List its fields:
            cmbAlterColumnName.Items.Clear()
            cmbDropColumnName.Items.Clear()
            lbCreateIndexColumnName.Items.Clear()
            cmbForeignKeyColumnName.Items.Clear()

            'Access 2007:
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
            "data source = " + Main.DatabasePath

            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            'Specify the commandString to query the database:
            'commandString = "SELECT * FROM " + cmbSelectTable.SelectedItem.ToString
            commandString = "SELECT Top 500 * FROM " + cmbSelectTable.SelectedItem.ToString
            Dim dataAdapter As New System.Data.OleDb.OleDbDataAdapter(commandString, conn)
            ds = New DataSet
            dataAdapter.Fill(ds, "SelTable") 'ds was defined earlier as a DataSet
            dt = ds.Tables("SelTable")

            Dim NFields As Integer
            NFields = dt.Columns.Count

            Dim I As Integer
            For I = 0 To NFields - 1
                cmbAlterColumnName.Items.Add(dt.Columns(I).ColumnName.ToString)
                cmbDropColumnName.Items.Add(dt.Columns(I).ColumnName.ToString)
                lbCreateIndexColumnName.Items.Add(dt.Columns(I).ColumnName.ToString)
                cmbForeignKeyColumnName.Items.Add(dt.Columns(I).ColumnName.ToString)
            Next

            'Fill Related Table combo box:
            cmbRelatedTable.Items.Clear()
            dt.Clear()
            Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
            dt = conn.GetSchema("Tables", restrictions)

            Dim MaxI As Integer
            MaxI = dt.Rows.Count
            For I = 0 To MaxI - 1
                If dt.Rows(I).Item("TABLE_NAME").ToString = cmbSelectTable.Text Then

                Else
                    cmbRelatedTable.Items.Add(dt.Rows(I).Item("TABLE_NAME").ToString)
                End If
            Next

            conn.Close()

        End If
    End Sub

    Private Sub cmbSelectTable_TextChanged(sender As Object, e As System.EventArgs) Handles cmbSelectTable.TextChanged
        FillCmbFields()
        FillIndexList()
    End Sub

    Private Sub cmbAddColumnType_TextChanged(sender As Object, e As System.EventArgs) Handles cmbAddColumnType.TextChanged
        Select Case cmbAddColumnType.Text
            Case "Short (Integer)"
                txtAddColumnSize.Text = "n/a" 'Disable Size
                txtAddColumnSize.ReadOnly = True
                txtAddColumnPrecision.Text = "n/a" 'Disable Precision
                txtAddColumnPrecision.ReadOnly = True
                txtAddColumnScale.Text = "n/a" 'Disable Scale
                txtAddColumnScale.ReadOnly = True
                cmbAddColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAddColumnAutoInc.Enabled = False

            Case "Long (Integer)"
                txtAddColumnSize.Text = "n/a" 'Disable Size
                txtAddColumnSize.ReadOnly = True
                txtAddColumnPrecision.Text = "n/a" 'Disable Precision
                txtAddColumnPrecision.ReadOnly = True
                txtAddColumnScale.Text = "n/a" 'Disable Scale
                txtAddColumnScale.ReadOnly = True
                cmbAddColumnAutoInc.Text = "" 'Enable AutoInc
                cmbAddColumnAutoInc.Enabled = True

            Case "Single"
                txtAddColumnSize.Text = "n/a" 'Disable Size
                txtAddColumnSize.ReadOnly = True
                txtAddColumnPrecision.Text = "n/a" 'Disable Precision
                txtAddColumnPrecision.ReadOnly = True
                txtAddColumnScale.Text = "n/a" 'Disable Scale
                txtAddColumnScale.ReadOnly = True
                cmbAddColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAddColumnAutoInc.Enabled = False

            Case "Double"
                txtAddColumnSize.Text = "n/a" 'Disable Size
                txtAddColumnSize.ReadOnly = True
                txtAddColumnPrecision.Text = "n/a" 'Disable Precision
                txtAddColumnPrecision.ReadOnly = True
                txtAddColumnScale.Text = "n/a" 'Disable Scale
                txtAddColumnScale.ReadOnly = True
                cmbAddColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAddColumnAutoInc.Enabled = False

            Case "Currency"
                txtAddColumnSize.Text = "n/a" 'Disable Size
                txtAddColumnSize.ReadOnly = True
                txtAddColumnPrecision.Text = "n/a" 'Disable Precision
                txtAddColumnPrecision.ReadOnly = True
                txtAddColumnScale.Text = "n/a" 'Disable Scale
                txtAddColumnScale.ReadOnly = True
                cmbAddColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAddColumnAutoInc.Enabled = False

            Case "DateTime"
                txtAddColumnSize.Text = "n/a" 'Disable Size
                txtAddColumnSize.ReadOnly = True
                txtAddColumnPrecision.Text = "n/a" 'Disable Precision
                txtAddColumnPrecision.ReadOnly = True
                txtAddColumnScale.Text = "n/a" 'Disable Scale
                txtAddColumnScale.ReadOnly = True
                cmbAddColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAddColumnAutoInc.Enabled = False

            Case "Bit (Boolean)"
                txtAddColumnSize.Text = "n/a" 'Disable Size
                txtAddColumnSize.ReadOnly = True
                txtAddColumnPrecision.Text = "n/a" 'Disable Precision
                txtAddColumnPrecision.ReadOnly = True
                txtAddColumnScale.Text = "n/a" 'Disable Scale
                txtAddColumnScale.ReadOnly = True
                cmbAddColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAddColumnAutoInc.Enabled = False

            Case "Byte"
                txtAddColumnSize.Text = "n/a" 'Disable Size
                txtAddColumnSize.ReadOnly = True
                txtAddColumnPrecision.Text = "n/a" 'Disable Precision
                txtAddColumnPrecision.ReadOnly = True
                txtAddColumnScale.Text = "n/a" 'Disable Scale
                txtAddColumnScale.ReadOnly = True
                cmbAddColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAddColumnAutoInc.Enabled = False

            Case "GUID"
                txtAddColumnSize.Text = "n/a" 'Disable Size
                txtAddColumnSize.ReadOnly = True
                txtAddColumnPrecision.Text = "n/a" 'Disable Precision
                txtAddColumnPrecision.ReadOnly = True
                txtAddColumnScale.Text = "n/a" 'Disable Scale
                txtAddColumnScale.ReadOnly = True
                cmbAddColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAddColumnAutoInc.Enabled = False


            Case "BigBinary"
                txtAddColumnSize.Text = "n/a" 'Disable Size
                txtAddColumnSize.ReadOnly = True
                txtAddColumnPrecision.Text = "n/a" 'Disable Precision
                txtAddColumnPrecision.ReadOnly = True
                txtAddColumnScale.Text = "n/a" 'Disable Scale
                txtAddColumnScale.ReadOnly = True
                cmbAddColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAddColumnAutoInc.Enabled = False
                Main.Message.Add("BigBinary: Maximum size: 4000 " & vbCrLf)

            Case "LongBinary"
                txtAddColumnSize.Text = "n/a" 'Disable Size
                txtAddColumnSize.ReadOnly = True
                txtAddColumnPrecision.Text = "n/a" 'Disable Precision
                txtAddColumnPrecision.ReadOnly = True
                txtAddColumnScale.Text = "n/a" 'Disable Scale
                txtAddColumnScale.ReadOnly = True
                cmbAddColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAddColumnAutoInc.Enabled = False
                Main.Message.Add("LongBinary: Maximum size: 1073741823 " & vbCrLf)

            Case "VarBinary"
                txtAddColumnSize.Text = "" 'Enable Size
                txtAddColumnSize.ReadOnly = False
                txtAddColumnPrecision.Text = "n/a" 'Disable Precision
                txtAddColumnPrecision.ReadOnly = True
                txtAddColumnScale.Text = "n/a" 'Disable Scale
                txtAddColumnScale.ReadOnly = True
                cmbAddColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAddColumnAutoInc.Enabled = False
                Main.Message.Add("VarBinary: Maximum size: 510 " & vbCrLf)

            Case "LongText"
                txtAddColumnSize.Text = "n/a" 'Disable Size
                txtAddColumnSize.ReadOnly = True
                txtAddColumnPrecision.Text = "n/a" 'Disable Precision
                txtAddColumnPrecision.ReadOnly = True
                txtAddColumnScale.Text = "n/a" 'Disable Scale
                txtAddColumnScale.ReadOnly = True
                cmbAddColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAddColumnAutoInc.Enabled = False
                Main.Message.Add("LongText: Maximum size: 536870910 " & vbCrLf)

            Case "VarChar"
                txtAddColumnSize.Text = "" 'Enable Size
                txtAddColumnSize.ReadOnly = False
                txtAddColumnPrecision.Text = "n/a" 'Disable Precision
                txtAddColumnPrecision.ReadOnly = True
                txtAddColumnScale.Text = "n/a" 'Disable Scale
                txtAddColumnScale.ReadOnly = True
                cmbAddColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAddColumnAutoInc.Enabled = False
                Main.Message.Add("VarChar: Maximum size: 255 " & vbCrLf)

            Case "Decimal"
                txtAddColumnSize.Text = "n/a" 'Disable Size
                txtAddColumnSize.ReadOnly = False
                txtAddColumnPrecision.Text = "" 'Enable Precision
                txtAddColumnPrecision.ReadOnly = False
                txtAddColumnScale.Text = "" 'Enable Scale
                txtAddColumnScale.ReadOnly = False
                cmbAddColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAddColumnAutoInc.Enabled = False
                Main.Message.Add("Decimal: Specify Precision and Scale." & vbCrLf)
                Main.Message.Add("Precision is the number of digits." & vbCrLf)
                Main.Message.Add("Scale is the number of digits to the right of the decimal point." & vbCrLf)

        End Select
    End Sub

    Private Sub cmbAlterColumnType_TextChanged(sender As Object, e As System.EventArgs) Handles cmbAlterColumnType.TextChanged

        Select Case cmbAlterColumnType.Text
            Case "Short (Integer)"
                txtAlterColumnSize.Text = "n/a" 'Disable Size
                txtAlterColumnSize.ReadOnly = True
                txtAlterColumnPrecision.Text = "n/a" 'Disable Precision
                txtAlterColumnPrecision.ReadOnly = True
                txtAlterColumnScale.Text = "n/a" 'Disable Scale
                txtAlterColumnScale.ReadOnly = True
                cmbAlterColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAlterColumnAutoInc.Enabled = False

            Case "Long (Integer)"
                txtAlterColumnSize.Text = "n/a" 'Disable Size
                txtAlterColumnSize.ReadOnly = True
                txtAlterColumnPrecision.Text = "n/a" 'Disable Precision
                txtAlterColumnPrecision.ReadOnly = True
                txtAlterColumnScale.Text = "n/a" 'Disable Scale
                txtAlterColumnScale.ReadOnly = True
                cmbAlterColumnAutoInc.Text = "" 'Enable AutoInc
                cmbAlterColumnAutoInc.Enabled = True

            Case "Single"
                txtAlterColumnSize.Text = "n/a" 'Disable Size
                txtAlterColumnSize.ReadOnly = True
                txtAlterColumnPrecision.Text = "n/a" 'Disable Precision
                txtAlterColumnPrecision.ReadOnly = True
                txtAlterColumnScale.Text = "n/a" 'Disable Scale
                txtAlterColumnScale.ReadOnly = True
                cmbAlterColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAlterColumnAutoInc.Enabled = False

            Case "Double"
                txtAlterColumnSize.Text = "n/a" 'Disable Size
                txtAlterColumnSize.ReadOnly = True
                txtAlterColumnPrecision.Text = "n/a" 'Disable Precision
                txtAlterColumnPrecision.ReadOnly = True
                txtAlterColumnScale.Text = "n/a" 'Disable Scale
                txtAlterColumnScale.ReadOnly = True
                cmbAlterColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAlterColumnAutoInc.Enabled = False

            Case "Currency"
                txtAlterColumnSize.Text = "n/a" 'Disable Size
                txtAlterColumnSize.ReadOnly = True
                txtAlterColumnPrecision.Text = "n/a" 'Disable Precision
                txtAlterColumnPrecision.ReadOnly = True
                txtAlterColumnScale.Text = "n/a" 'Disable Scale
                txtAlterColumnScale.ReadOnly = True
                cmbAlterColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAlterColumnAutoInc.Enabled = False

            Case "DateTime"
                txtAlterColumnSize.Text = "n/a" 'Disable Size
                txtAlterColumnSize.ReadOnly = True
                txtAlterColumnPrecision.Text = "n/a" 'Disable Precision
                txtAlterColumnPrecision.ReadOnly = True
                txtAlterColumnScale.Text = "n/a" 'Disable Scale
                txtAlterColumnScale.ReadOnly = True
                cmbAlterColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAlterColumnAutoInc.Enabled = False

            Case "Bit (Boolean)"
                txtAlterColumnSize.Text = "n/a" 'Disable Size
                txtAlterColumnSize.ReadOnly = True
                txtAlterColumnPrecision.Text = "n/a" 'Disable Precision
                txtAlterColumnPrecision.ReadOnly = True
                txtAlterColumnScale.Text = "n/a" 'Disable Scale
                txtAlterColumnScale.ReadOnly = True
                cmbAlterColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAlterColumnAutoInc.Enabled = False

            Case "Byte"
                txtAlterColumnSize.Text = "n/a" 'Disable Size
                txtAlterColumnSize.ReadOnly = True
                txtAlterColumnPrecision.Text = "n/a" 'Disable Precision
                txtAlterColumnPrecision.ReadOnly = True
                txtAlterColumnScale.Text = "n/a" 'Disable Scale
                txtAlterColumnScale.ReadOnly = True
                cmbAlterColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAlterColumnAutoInc.Enabled = False

            Case "GUID"
                txtAlterColumnSize.Text = "n/a" 'Disable Size
                txtAlterColumnSize.ReadOnly = True
                txtAlterColumnPrecision.Text = "n/a" 'Disable Precision
                txtAlterColumnPrecision.ReadOnly = True
                txtAlterColumnScale.Text = "n/a" 'Disable Scale
                txtAlterColumnScale.ReadOnly = True
                cmbAlterColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAlterColumnAutoInc.Enabled = False

            Case "BigBinary"
                txtAlterColumnSize.Text = "" 'Enable Size
                txtAlterColumnSize.ReadOnly = False
                txtAlterColumnPrecision.Text = "n/a" 'Disable Precision
                txtAlterColumnPrecision.ReadOnly = True
                txtAlterColumnScale.Text = "n/a" 'Disable Scale
                txtAlterColumnScale.ReadOnly = True
                cmbAlterColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAlterColumnAutoInc.Enabled = False
                Main.Message.Add("BigBinary: Maximum size: 4000 " & vbCrLf)

            Case "LongBinary"
                txtAlterColumnSize.Text = "" 'Enable Size
                txtAlterColumnSize.ReadOnly = False
                txtAlterColumnPrecision.Text = "n/a" 'Disable Precision
                txtAlterColumnPrecision.ReadOnly = True
                txtAlterColumnScale.Text = "n/a" 'Disable Scale
                txtAlterColumnScale.ReadOnly = True
                cmbAlterColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAlterColumnAutoInc.Enabled = False
                Main.Message.Add("LongBinary: Maximum size: 1073741823 " & vbCrLf)

            Case "VarBinary"
                txtAlterColumnSize.Text = "" 'Enable Size
                txtAlterColumnSize.ReadOnly = False
                txtAlterColumnPrecision.Text = "n/a" 'Disable Precision
                txtAlterColumnPrecision.ReadOnly = True
                txtAlterColumnScale.Text = "n/a" 'Disable Scale
                txtAlterColumnScale.ReadOnly = True
                cmbAlterColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAlterColumnAutoInc.Enabled = False
                Main.Message.Add("VarBinary: Maximum size: 510 " & vbCrLf)

            Case "LongText"
                txtAlterColumnSize.Text = "" 'Enable Size
                txtAlterColumnSize.ReadOnly = False
                txtAlterColumnPrecision.Text = "n/a" 'Disable Precision
                txtAlterColumnPrecision.ReadOnly = True
                txtAlterColumnScale.Text = "n/a" 'Disable Scale
                txtAlterColumnScale.ReadOnly = True
                cmbAlterColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAlterColumnAutoInc.Enabled = False
                Main.Message.Add("LongText: Maximum size: 536870910 " & vbCrLf)

            Case "VarChar"
                txtAlterColumnSize.Text = "" 'Enable Size
                txtAlterColumnSize.ReadOnly = False
                txtAlterColumnPrecision.Text = "n/a" 'Disable Precision
                txtAlterColumnPrecision.ReadOnly = True
                txtAlterColumnScale.Text = "n/a" 'Disable Scale
                txtAlterColumnScale.ReadOnly = True
                cmbAlterColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAlterColumnAutoInc.Enabled = False
                Main.Message.Add("VarChar: Maximum size: 255 " & vbCrLf)

            Case "Decimal"
                txtAlterColumnSize.Text = "n/a" 'Disable Size
                txtAlterColumnSize.ReadOnly = True
                txtAlterColumnPrecision.Text = "" 'Enable Precision
                txtAlterColumnPrecision.ReadOnly = False
                txtAlterColumnScale.Text = "" 'Enable Scale
                txtAlterColumnScale.ReadOnly = False
                cmbAlterColumnAutoInc.Text = "" 'Disable AutoInc
                cmbAlterColumnAutoInc.Enabled = False
                Main.Message.Add("Decimal: Specify Precision and Scale." & vbCrLf)
                Main.Message.Add("Precision is the number of digits." & vbCrLf)
                Main.Message.Add("Scale is the number of digits to the right of the decimal point." & vbCrLf)

        End Select
    End Sub

    Private Sub FillIndexList()
        'Fill the index list in the listbox: lbIndexList

        If Main.DatabasePath = "" Then
            'No database selected.
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

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

        'Index restrictions: TABLE_CATALOG TABLE_SCHEMA INDEX_NAME TYPE TABLE_NAME
        Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, Nothing, cmbSelectTable.Text} 'This restriction limits indexes to the selected table
        dt = conn.GetSchema("Indexes", restrictions)
        Dim I As Integer

        DataGridView2.Rows.Clear()
        DataGridView2.ColumnCount = 2
        DataGridView2.Columns(0).HeaderText = "Index Name"
        DataGridView2.Columns(0).Width = 130
        DataGridView2.Columns(1).HeaderText = "Column Name"
        DataGridView2.Columns(1).Width = 130

        For I = 1 To dt.Rows.Count
            DataGridView2.Rows.Add()
            DataGridView2.Rows(I - 1).Cells(0).Value = dt.Rows(I - 1).Item("INDEX_NAME").ToString
            DataGridView2.Rows(I - 1).Cells(1).Value = dt.Rows(I - 1).Item("COLUMN_NAME").ToString
        Next

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Dim SelRow As Integer

        'Highlight the selected row:
        SelRow = DataGridView2.SelectedCells(0).RowIndex
        DataGridView2.Rows(SelRow).Selected = True
    End Sub

    Private Sub cmbRelatedTable_TextChanged(sender As Object, e As System.EventArgs) Handles cmbRelatedTable.TextChanged
        'Update cmbPrimaryKey:

        If Main.DatabasePath = "" Then
            'No database selected.
            Exit Sub
        End If

        Dim RelatedTableName As String
        RelatedTableName = cmbRelatedTable.Text

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

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

        'Index restrictions: TABLE_CATALOG TABLE_SCHEMA INDEX_NAME TYPE TABLE_NAME
        Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, Nothing, cmbRelatedTable.Text} 'This restriction limits indexes to the selected table
        dt = conn.GetSchema("Indexes", restrictions)

        Dim I As Integer
        cmbPrimaryKey.Items.Clear()
        For I = 1 To dt.Rows.Count
            If dt.Rows(I - 1).Item("PRIMARY_KEY").ToString = "True" Then
                cmbPrimaryKey.Items.Add(dt.Rows(I - 1).Item("COLUMN_NAME").ToString)
            End If
        Next
    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------



End Class