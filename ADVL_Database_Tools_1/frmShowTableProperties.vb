Public Class frmShowTableProperties
    'Form displays the properties of the table.

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

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        RestoreFormSettings()   'Restore the form settings
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

    Public Sub UpdateProperties(ByRef ds As DataSet)
        'Update DataGridView1 with the properties of the table in dataset ds
        'The name of each property is shown in the header of each column

        If ds.Tables.Count = 0 Then
            Main.Message.Add("No table selected")
            Exit Sub
        End If

        Dim NCols As Integer
        NCols = 5 'The number of properties shown
        DataGridView1.ColumnCount = NCols
        DataGridView1.Columns(0).HeaderText = "Column Name"
        DataGridView1.Columns(1).HeaderText = "Data Type"
        DataGridView1.Columns(2).HeaderText = "Allow DB Null"
        DataGridView1.Columns(3).HeaderText = "Auto Increment"
        DataGridView1.Columns(4).HeaderText = "Max Length"

        Dim NRows As Integer
        NRows = ds.Tables(0).Columns.Count 'The number of fields in the table

        DataGridView1.Rows.Clear()

        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        DataGridView1.AutoResizeColumns()

        Dim I As Integer
        'For I = 1 To NCols
        For I = 1 To NRows
            DataGridView1.Rows.Add()
            DataGridView1.Rows(I - 1).Cells(0).Value = ds.Tables(0).Columns(I - 1).ColumnName
            DataGridView1.Rows(I - 1).Cells(1).Value = ds.Tables(0).Columns(I - 1).DataType
            DataGridView1.Rows(I - 1).Cells(2).Value = ds.Tables(0).Columns(I - 1).AllowDBNull
            DataGridView1.Rows(I - 1).Cells(3).Value = ds.Tables(0).Columns(I - 1).AutoIncrement
            DataGridView1.Rows(I - 1).Cells(4).Value = ds.Tables(0).Columns(I - 1).MaxLength

        Next
        'List of column properties:
        'AllowDBNull, AutoIncrement, AutoIncrementSeed, AutoIncrementStep, Caption, ColumnMapping, ColumnName, Container, DataType, DateTimeMode
        'DefaultValue, DesignMode, MaxLength, Ordinal, ReadOnly, Site, Table, Unique

    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class