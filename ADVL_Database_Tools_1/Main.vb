
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598

'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
''WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
'
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Class Main
    'The ADVL_Database_Tools_1 application is used to design and create databases.
    'The database design is stored as an xml file than can be edited.
    'A new database can be created using the design.
    'An existing database can be analysed and its design stored as an xml file.
    'Currently, only Access databases are supported.

#Region " Coding Notes - Notes on the code used in this class." '------------------------------------------------------------------------------------------------------------------------------

    'ADD THE SYSTEM UTILITIES REFERENCE: ==========================================================================================
    'The following references are required by this software: 
    'Project \ Add Reference... \ ADVL_Utilities_Library_1.dll
    'The Utilities Library is used for Project Management, Archive file management, running XSequence files and running XMessage files.
    'If there are problems with a reference, try deleting it from the references list and adding it again.

    'ADD THE SERVICE REFERENCE: ===================================================================================================
    'A service reference to the Message Service must be added to the source code before this service can be used.
    'This is used to connect to the Application Network.

    'Adding the service reference to a project that includes the WcfMsgServiceLib project: -----------------------------------------
    'Project \ Add Service Reference
    'Press the Discover button.
    'Expand the items in the Services window and select IMsgService.
    'Press OK.
    '------------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------------
    'Adding the service reference to other projects that dont include the WcfMsgServiceLib project: -------------------------------
    'Run the ADVL_Application_Network_1 application to start the Application Network message service.
    'In Microsoft Visual Studio select: Project \ Add Service Reference
    'Enter the address: http://localhost:8733/ADVLService
    'Press the Go button.
    'MsgService is found.
    'Press OK to add ServiceReference1 to the project.
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'ADD THE MsgServiceCallback CODE: =============================================================================================
    'This is used to connect to the Application Network.
    'In Microsoft Visual Studio select: Project \ Add Class
    'MsgServiceCallback.vb
    'Add the following code to the class:
    'Imports System.ServiceModel
    'Public Class MsgServiceCallback
    '    Implements ServiceReference1.IMsgServiceCallback
    '    Public Sub OnSendMessage(message As String) Implements ServiceReference1.IMsgServiceCallback.OnSendMessage
    '        'A message has been received.
    '        'Set the InstrReceived property value to the message (usually in XMessage format). This will also apply the instructions in the XMessage.
    '        Main.InstrReceived = message
    '    End Sub
    'End Class
    '------------------------------------------------------------------------------------------------------------------------------

#End Region 'Coding Notes ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Variable Declarations - All the variables used in this form and this application." '-------------------------------------------------------------------------------------------------

    Public WithEvents ApplicationInfo As New ADVL_Utilities_Library_1.ApplicationInfo 'This object is used to store application information.
    Public WithEvents Project As New ADVL_Utilities_Library_1.Project 'This object is used to store Project information.
    Public WithEvents Message As New ADVL_Utilities_Library_1.Message 'This object is used to display messages in the Messages window.
    Public WithEvents ApplicationUsage As New ADVL_Utilities_Library_1.Usage 'This object stores application usage information.

    'Declare Forms used by the application:
    'Public WithEvents TemplateForm As frmTemplate
    Public WithEvents SqlCommand As frmSqlCommand
    Public WithEvents ModifyDatabase As frmModifyDatabase
    Public WithEvents CreateNewDatabase As frmCreateNewDatabase
    Public WithEvents SaveDatabaseDefinition As frmSaveDatabaseDefinition
    Public WithEvents SaveTableDefinition As frmSaveTableDefinition
    Public WithEvents ShowTableProperties As frmShowTableProperties
    Public WithEvents CreateNewTable As frmCreateNewTable


    'Declare objects used to connect to the Application Network:
    Public client As ServiceReference1.MsgServiceClient
    Public WithEvents XMsg As New ADVL_Utilities_Library_1.XMessage
    Dim XDoc As New System.Xml.XmlDocument
    Public Status As New System.Collections.Specialized.StringCollection
    Dim ClientName As String 'The name of the client requesting service
    Dim MessageText As String 'The text of a message sent through the Application Network
    Dim MessageDest As String 'The destination of a message sent through the Application Network

    'Variables used to connect to a database and open a table - Displayed in the Table tab:
    Dim connString As String
    'Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
    Public myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
    'Dim ds As DataSet = New DataSet
    Public ds As DataSet = New DataSet
    Dim da As OleDb.OleDbDataAdapter
    Dim tables As DataTableCollection = ds.Tables
    Dim UpdateNeeded As Boolean 'If False, data displayed on the Table tab has not been changed and the database table does not need to be updated.


#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form and this application" '------------------------------------------------------------------------------------------------------------

    Private _connectionHashcode As Integer 'The Application Network connection hashcode. This is used to identify a connection in the Application Netowrk when reconnecting.
    Property ConnectionHashcode As Integer
        Get
            Return _connectionHashcode
        End Get
        Set(value As Integer)
            _connectionHashcode = value
        End Set
    End Property


    Private _connectedToAppNet As Boolean = False  'True if the application is connected to the Application Network.
    Property ConnectedToAppnet As Boolean
        Get
            Return _connectedToAppNet
        End Get
        Set(value As Boolean)
            _connectedToAppNet = value
        End Set
    End Property

    Private _instrReceived As String = "" 'Contains Instructions received from the Application Network message service.
    Property InstrReceived As String
        Get
            Return _instrReceived
        End Get
        Set(value As String)
            If value = Nothing Then
                Message.Add("Empty message received!")
            Else
                _instrReceived = value

                'Add the message to the XMessages window:
                Message.Color = Color.CadetBlue
                Message.FontStyle = FontStyle.Bold
                Message.XAdd("Message received: " & vbCrLf)
                Message.SetNormalStyle()
                Message.XAdd(_instrReceived & vbCrLf & vbCrLf)

                If _instrReceived.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
                    Try
                        Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
                        XDoc.LoadXml(XmlHeader & vbCrLf & _instrReceived)
                        XMsg.Run(XDoc, Status)
                    Catch ex As Exception
                        Message.Add("Error running XMsg: " & ex.Message & vbCrLf)
                    End Try

                    'XMessage has been run.
                    'Reply to this message:
                    'Add the message reply to the XMessages window:
                    If ClientName = "" Then
                        'No client to send a message to!
                    Else
                        If MessageText = "" Then
                            'No message to send!
                        Else
                            Message.Color = Color.Red
                            Message.FontStyle = FontStyle.Bold
                            Message.XAdd("Message sent to " & ClientName & ":" & vbCrLf)
                            Message.SetNormalStyle()
                            Message.XAdd(MessageText & vbCrLf & vbCrLf)
                            MessageDest = ClientName
                            'SendMessage sends the contents of MessageText to MessageDest.
                            SendMessage() 'This subroutine triggers the timer to send the message after a short delay.
                        End If
                    End If
                Else

                End If
            End If

        End Set
    End Property

    Public Enum DatabaseTypes
        Access
        Unknown
    End Enum

    Private _databaseType As DatabaseTypes = DatabaseTypes.Access 'The type of database. Currently only Access is supported.
    Property DatabaseType As DatabaseTypes
        Get
            Return _databaseType
        End Get
        Set(value As DatabaseTypes)
            _databaseType = value
        End Set
    End Property

    Private _databasePath As String = "" 'The path to the selected database.
    Property DatabasePath As String
        Get
            Return _databasePath
        End Get
        Set(value As String)
            _databasePath = value
        End Set
    End Property

    Private _tableName As String = ""  'The TableName property stores the name of the table selected for viewing.
    Property TableName As String
        Get
            Return _tableName
        End Get
        Set(value As String)
            _tableName = value
        End Set
    End Property

    Private _query As String = "" 'The text of the query used to display table values.
    Property Query As String
        Get
            Return _query
        End Get
        Set(value As String)
            _query = value
        End Set
    End Property

    Private _sqlCommandText As String
    Property SqlCommandText As String
        Get
            Return _sqlCommandText
        End Get
        Set(value As String)
            _sqlCommandText = value
        End Set
    End Property

    Private _sqlCommandResult As String = "Error" 'The result of applying the SQL Command in SqlCommandText. Either "OK" or "Error"
    Property SqlCommandResult As String
        Get
            Return _sqlCommandResult
        End Get
        Set(value As String)
            _sqlCommandResult = value
        End Set
    End Property



#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Process XML files - Read and write XML files." '-------------------------------------------------------------------------------------------------------------------------------------

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Form settings for Main form.-->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <SelectedMainTabIndex><%= TabControl1.SelectedIndex %></SelectedMainTabIndex>
                               <!---->
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"

        If Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Project.ReadXmlSettings(SettingsFileName, Settings)

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
            If Settings.<FormSettings>.<SelectedMainTabIndex>.Value <> Nothing Then TabControl1.SelectedIndex = Settings.<FormSettings>.<SelectedMainTabIndex>.Value

        End If
    End Sub

    Private Sub ReadApplicationInfo()
        'Read the Application Information.

        If ApplicationInfo.FileExists Then
            ApplicationInfo.ReadFile()
        Else
            'There is no Application_Info.xml file.
            DefaultAppProperties() 'Create a new Application Info file with default application properties:
        End If
    End Sub

    Private Sub DefaultAppProperties()
        'These properties will be saved in the Application_Info.xml file in the application directory.
        'If this file is deleted, it will be re-created using these default application properties.

        'Change this to show your application Name, Description and Creation Date.
        ApplicationInfo.Name = "ADVL_Database_Tools_1"

        'ApplicationInfo.ApplicationDir is set when the application is started.
        ApplicationInfo.ExecutablePath = Application.ExecutablePath

        ApplicationInfo.Description = "The Database Tools application is used to design, modify and create databases. The database designs are saved in XML files."
        ApplicationInfo.CreationDate = "14-Aug-2016 12:00:00"

        'Author -----------------------------------------------------------------------------------------------------------
        'Change this to show your Name, Description and Contact information.
        ApplicationInfo.Author.Name = "Signalworks Pty Ltd"
        ApplicationInfo.Author.Description = "Signalworks Pty Ltd" & vbCrLf &
            "Australian Proprietary Company" & vbCrLf &
            "ABN 26 066 681 598" & vbCrLf &
            "Registration Date 05/10/1994"

        ApplicationInfo.Author.Contact = "http://www.andorville.com.au/"

        'File Associations: -----------------------------------------------------------------------------------------------
        'Add any file associations here.
        'The file extension and a description of files that can be opened by this application are specified.
        'The example below specifies a coordinate system parameter file type with the file extension .ADVLCoord.
        'Dim Assn1 As New ADVL_System_Utilities.FileAssociation
        'Assn1.Extension = "ADVLCoord"
        'Assn1.Description = "Andorville (TM) software coordinate system parameter file"
        'ApplicationInfo.FileAssociations.Add(Assn1)

        'Version ----------------------------------------------------------------------------------------------------------
        ApplicationInfo.Version.Major = My.Application.Info.Version.Major
        ApplicationInfo.Version.Minor = My.Application.Info.Version.Minor
        ApplicationInfo.Version.Build = My.Application.Info.Version.Build
        ApplicationInfo.Version.Revision = My.Application.Info.Version.Revision

        'Copyright --------------------------------------------------------------------------------------------------------
        'Add your copyright information here.
        ApplicationInfo.Copyright.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        ApplicationInfo.Copyright.PublicationYear = "2016"

        'Trademarks -------------------------------------------------------------------------------------------------------
        'Add your trademark information here.
        Dim Trademark1 As New ADVL_Utilities_Library_1.Trademark
        Trademark1.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark1.Text = "Andorville"
        Trademark1.Registered = False
        Trademark1.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark1)
        Dim Trademark2 As New ADVL_Utilities_Library_1.Trademark
        Trademark2.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark2.Text = "AL-H7"
        Trademark2.Registered = False
        Trademark2.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark2)

        'License -------------------------------------------------------------------------------------------------------
        'Add your license information here.
        ApplicationInfo.License.CopyrightOwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        ApplicationInfo.License.PublicationYear = "2016"

        'License Links:
        'http://choosealicense.com/
        'http://www.apache.org/licenses/
        'http://opensource.org/

        'Apache License 2.0 ---------------------------------------------
        ApplicationInfo.License.Code = ADVL_Utilities_Library_1.License.Codes.Apache_License_2_0
        ApplicationInfo.License.Notice = ApplicationInfo.License.ApacheLicenseNotice 'Get the pre-defined Aapche license notice.
        ApplicationInfo.License.Text = ApplicationInfo.License.ApacheLicenseText     'Get the pre-defined Apache license text.

        'Code to use other pre-defined license types is shown below:

        'GNU General Public License, version 3 --------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.GNU_GPL_V3_0
        'ApplicationInfo.License.Notice = 'Add the License Notice to ADVL_Utilities_Library_1 License class.
        'ApplicationInfo.License.Text = 'Add the License Text to ADVL_Utilities_Library_1 License class.

        'The MIT License ------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.MIT_License
        'ApplicationInfo.License.Notice = ApplicationInfo.License.MITLicenseNotice
        'ApplicationInfo.License.Text = ApplicationInfo.License.MITLicenseText

        'No License Specified -------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.None
        'ApplicationInfo.License.Notice = ""
        'ApplicationInfo.License.Text = ""

        'The Unlicense --------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.The_Unlicense
        'ApplicationInfo.License.Notice = ApplicationInfo.License.UnLicenseNotice
        'ApplicationInfo.License.Text = ApplicationInfo.License.UnLicenseText

        'Unknown License ------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.Unknown
        'ApplicationInfo.License.Notice = ""
        'ApplicationInfo.License.Text = ""

        'Source Code: --------------------------------------------------------------------------------------------------
        'Add your source code information here if required.
        'THIS SECTION WILL BE UPDATED TO ALLOW A GITHUB LINK.
        ApplicationInfo.SourceCode.Language = "Visual Basic 2015"
        ApplicationInfo.SourceCode.FileName = ""
        ApplicationInfo.SourceCode.FileSize = 0
        ApplicationInfo.SourceCode.FileHash = ""
        ApplicationInfo.SourceCode.WebLink = ""
        ApplicationInfo.SourceCode.Contact = ""
        ApplicationInfo.SourceCode.Comments = ""

        'ModificationSummary: -----------------------------------------------------------------------------------------
        'Add any source code modification here is required.
        ApplicationInfo.ModificationSummary.BaseCodeName = ""
        ApplicationInfo.ModificationSummary.BaseCodeDescription = ""
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Major = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Minor = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Build = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Revision = 0
        ApplicationInfo.ModificationSummary.Description = "This is the first released version of the application. No earlier base code used."

        'Library List: ------------------------------------------------------------------------------------------------
        'Add the ADVL_Utilties_Library_1 library:
        Dim NewLib As New ADVL_Utilities_Library_1.LibrarySummary
        NewLib.Name = "ADVL_System_Utilities"
        NewLib.Description = "System Utility classes used in Andorville (TM) software development system applications"
        NewLib.CreationDate = "7-Jan-2016 12:00:00"
        NewLib.LicenseNotice = "Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598" & vbCrLf &
                               vbCrLf &
                               "Licensed under the Apache License, Version 2.0 (the ""License"");" & vbCrLf &
                               "you may not use this file except in compliance with the License." & vbCrLf &
                               "You may obtain a copy of the License at" & vbCrLf &
                               vbCrLf &
                               "http://www.apache.org/licenses/LICENSE-2.0" & vbCrLf &
                               vbCrLf &
                               "Unless required by applicable law or agreed to in writing, software" & vbCrLf &
                               "distributed under the License is distributed on an ""AS IS"" BASIS," & vbCrLf &
                               "WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied." & vbCrLf &
                               "See the License for the specific language governing permissions and" & vbCrLf &
                               "limitations under the License." & vbCrLf

        NewLib.CopyrightNotice = "Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598"

        NewLib.Version.Major = 1
        NewLib.Version.Minor = 0
        NewLib.Version.Build = 1
        NewLib.Version.Revision = 0

        NewLib.Author.Name = "Signalworks Pty Ltd"
        NewLib.Author.Description = "Signalworks Pty Ltd" & vbCrLf &
            "Australian Proprietary Company" & vbCrLf &
            "ABN 26 066 681 598" & vbCrLf &
            "Registration Date 05/10/1994"

        NewLib.Author.Contact = "http://www.andorville.com.au/"

        Dim NewClass1 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass1.Name = "ZipComp"
        NewClass1.Description = "The ZipComp class is used to compress files into and extract files from a zip file."
        NewLib.Classes.Add(NewClass1)
        Dim NewClass2 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass2.Name = "XSequence"
        NewClass2.Description = "The XSequence class is used to run an XML property sequence (XSequence) file. XSequence files are used to record and replay processing sequences in Andorville (TM) software applications."
        NewLib.Classes.Add(NewClass2)
        Dim NewClass3 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass3.Name = "XMessage"
        NewClass3.Description = "The XMessage class is used to read an XML Message (XMessage). An XMessage is a simplified XSequence used to exchange information between Andorville (TM) software applications."
        NewLib.Classes.Add(NewClass3)
        Dim NewClass4 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass4.Name = "Location"
        NewClass4.Description = "The Location class consists of properties and methods to store data in a location, which is either a directory or archive file."
        NewLib.Classes.Add(NewClass4)
        Dim NewClass5 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass5.Name = "Project"
        NewClass5.Description = "An Andorville (TM) software application can store data within one or more projects. Each project stores a set of related data files. The Project class contains properties and methods used to manage a project."
        NewLib.Classes.Add(NewClass5)
        Dim NewClass6 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass6.Name = "ProjectSummary"
        NewClass6.Description = "ProjectSummary stores a summary of a project."
        NewLib.Classes.Add(NewClass6)
        Dim NewClass7 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass7.Name = "DataFileInfo"
        NewClass7.Description = "The DataFileInfo class stores information about a data file."
        NewLib.Classes.Add(NewClass7)
        Dim NewClass8 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass8.Name = "Message"
        NewClass8.Description = "The Message class contains text properties and methods used to display messages in an Andorville (TM) software application."
        NewLib.Classes.Add(NewClass8)
        Dim NewClass9 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass9.Name = "ApplicationSummary"
        NewClass9.Description = "The ApplicationSummary class stores a summary of an Andorville (TM) software application."
        NewLib.Classes.Add(NewClass9)
        Dim NewClass10 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass10.Name = "LibrarySummary"
        NewClass10.Description = "The LibrarySummary class stores a summary of a software library used by an application."
        NewLib.Classes.Add(NewClass10)
        Dim NewClass11 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass11.Name = "ClassSummary"
        NewClass11.Description = "The ClassSummary class stores a summary of a class contained in a software library."
        NewLib.Classes.Add(NewClass11)
        Dim NewClass12 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass12.Name = "ModificationSummary"
        NewClass12.Description = "The ModificationSummary class stores a summary of any modifications made to an application or library."
        NewLib.Classes.Add(NewClass12)
        Dim NewClass13 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass13.Name = "ApplicationInfo"
        NewClass13.Description = "The ApplicationInfo class stores information about an Andorville (TM) software application."
        NewLib.Classes.Add(NewClass13)
        Dim NewClass14 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass14.Name = "Version"
        NewClass14.Description = "The Version class stores application, library or project version information."
        NewLib.Classes.Add(NewClass14)
        Dim NewClass15 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass15.Name = "Author"
        NewClass15.Description = "The Author class stores information about an Author."
        NewLib.Classes.Add(NewClass15)
        Dim NewClass16 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass16.Name = "FileAssociation"
        NewClass16.Description = "The FileAssociation class stores the file association extension and description. An application can open files on its file association list."
        NewLib.Classes.Add(NewClass16)
        Dim NewClass17 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass17.Name = "Copyright"
        NewClass17.Description = "The Copyright class stores copyright information."
        NewLib.Classes.Add(NewClass17)
        Dim NewClass18 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass18.Name = "License"
        NewClass18.Description = "The License class stores license information."
        NewLib.Classes.Add(NewClass18)
        Dim NewClass19 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass19.Name = "SourceCode"
        NewClass19.Description = "The SourceCode class stores information about the source code for the application."
        NewLib.Classes.Add(NewClass19)
        Dim NewClass20 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass20.Name = "Usage"
        NewClass20.Description = "The Usage class stores information about application or project usage."
        NewLib.Classes.Add(NewClass20)
        Dim NewClass21 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass21.Name = "Trademark"
        NewClass21.Description = "The Trademark class stored information about a trademark used by the author of an application or data."
        NewLib.Classes.Add(NewClass21)

        ApplicationInfo.Libraries.Add(NewLib)

        'Add other library information here: --------------------------------------------------------------------------

    End Sub


    'Save the form settings if the form is being minimised:
    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = &H112 Then 'SysCommand
            If m.WParam.ToInt32 = &HF020 Then 'Form is being minimised
                SaveFormSettings()
            End If
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub SaveProjectSettings()
        'Save project settings in an xml file.

        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Project settings for ADVL_Coordinates_1 application.-->
                           <ProjectSettings>
                               <DatabaseType><%= DatabaseType.ToString %></DatabaseType>
                               <DatabasePath><%= DatabasePath %></DatabasePath>
                               <TableName><%= TableName %></TableName>
                           </ProjectSettings>

        Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Project.SaveXmlSettings(SettingsFileName, settingsData)

    End Sub

    Private Sub RestoreProjectSettings()
        'Read the project settings from an XML document.

        Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"

        If Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore the database type:
            If Settings.<ProjectSettings>.<DatabaseType>.Value = Nothing Then
                'Project setting not saved.
                DatabaseType = DatabaseTypes.Access
            Else
                Select Case Settings.<ProjectSettings>.<DatabaseType>.Value
                    Case "Access"
                        DatabaseType = DatabaseTypes.Access
                    Case Else
                        DatabaseType = DatabaseTypes.Unknown
                End Select
            End If

            'Restore the database path:
            If Settings.<ProjectSettings>.<DatabasePath>.Value = Nothing Then
                'Project setting not saved.
                DatabasePath = ""
            Else
                DatabasePath = Settings.<ProjectSettings>.<DatabasePath>.Value
            End If

            'Restore the selected table name:
            If Settings.<ProjectSettings>.<TableName>.Value = Nothing Then
                'Table name not saved.
                TableName = ""
            Else
                TableName = Settings.<ProjectSettings>.<TableName>.Value
            End If

        End If

    End Sub

#End Region 'Process XML Files ----------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Display Methods - Code used to display this form." '----------------------------------------------------------------------------------------------------------------------------

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles Me.Load

        'Write the startup messages in a stringbuilder object.
        'Messages cannot be written using Message.Add until this is set up later in the startup sequence.
        Dim sb As New System.Text.StringBuilder
        sb.Append("------------------- Starting Application: ADVL Database Tools ------------------------------------------------------------------------ " & vbCrLf)

        'Set the Application Directory path: ------------------------------------------------
        Project.ApplicationDir = My.Application.Info.DirectoryPath.ToString

        'Read the Application Information file: ---------------------------------------------
        ApplicationInfo.ApplicationDir = My.Application.Info.DirectoryPath.ToString 'Set the Application Directory property

        If ApplicationInfo.ApplicationLocked Then
            MessageBox.Show("The application is locked. If the application is not already in use, remove the 'Application_Info.lock file from the application directory: " & ApplicationInfo.ApplicationDir, "Notice", MessageBoxButtons.OK)
            Dim dr As System.Windows.Forms.DialogResult
            dr = MessageBox.Show("Press 'Yes' to unlock the application", "Notice", MessageBoxButtons.YesNo)
            If dr = System.Windows.Forms.DialogResult.Yes Then
                ApplicationInfo.UnlockApplication()
            Else
                Application.Exit()
                'System.Windows.Forms.Application.Exit()
            End If
        End If

        ReadApplicationInfo()
        ApplicationInfo.LockApplication()

        'Read the Application Usage information: --------------------------------------------
        ApplicationUsage.StartTime = Now
        ApplicationUsage.SaveLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        ApplicationUsage.SaveLocn.Path = Project.ApplicationDir
        ApplicationUsage.RestoreUsageInfo()
        sb.Append("Application usage: Total duration = " & Format(ApplicationUsage.TotalDuration.TotalHours, "#0.##") & " hours" & vbCrLf)

        'Restore Project information: -------------------------------------------------------
        Project.ApplicationName = ApplicationInfo.Name
        Project.ReadLastProjectInfo()
        Project.ReadProjectInfoFile()
        Project.Usage.StartTime = Now

        Project.ReadProjectInfoFile()

        ApplicationInfo.SettingsLocn = Project.SettingsLocn

        'Set up the Message object:
        Message.ApplicationName = ApplicationInfo.Name
        Message.SettingsLocn = Project.SettingsLocn

        'Set up the Database Type combobox:
        cmbDatabaseType.Items.Clear()
        cmbDatabaseType.Items.Add("Access")
        cmbDatabaseType.Items.Add("Unknown")
        'cmbDatabaseType.SelectedIndex = 0

        'Restore the form settings: ---------------------------------------------------------
        RestoreFormSettings()

        RestoreProjectSettings()
        cmbDatabaseType.SelectedIndex = cmbDatabaseType.FindStringExact(DatabaseType.ToString)
        txtDatabase.Text = DatabasePath
        FillLstTables()
        FillCmbSelectTable()

        'Show the project information: ------------------------------------------------------
        txtProjectName.Text = Project.Name
        txtProjectDescription.Text = Project.Description
        Select Case Project.Type
            Case ADVL_Utilities_Library_1.Project.Types.Directory
                txtProjectType.Text = "Directory"
            Case ADVL_Utilities_Library_1.Project.Types.Archive
                txtProjectType.Text = "Archive"
            Case ADVL_Utilities_Library_1.Project.Types.Hybrid
                txtProjectType.Text = "Hybrid"
            Case ADVL_Utilities_Library_1.Project.Types.None
                txtProjectType.Text = "None"
        End Select
        txtCreationDate.Text = Format(Project.Usage.FirstUsed, "d-MMM-yyyy H:mm:ss")
        txtLastUsed.Text = Format(Project.Usage.LastUsed, "d-MMM-yyyy H:mm:ss")
        Select Case Project.SettingsLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSettingsLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSettingsLocationType.Text = "Archive"
        End Select
        txtSettingsLocationPath.Text = Project.SettingsLocn.Path
        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtDataLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtDataLocationType.Text = "Archive"
        End Select
        txtDataLocationPath.Text = Project.DataLocn.Path

        'Set up the Table tab: ------------------------------------------------------------------------------------
        FillCmbSelectTable()
        UpdateNeeded = False

        'Set the Restrictions options in the Scheme tab:
        cmbRestrictions.Items.Add("<No restrictions>")
        cmbRestrictions.Items.Add("Tables")
        cmbRestrictions.Items.Add("Tables (excluding System and Access tables)")
        cmbRestrictions.Items.Add("Restrictions")
        'cmbRestrictions.Items.Add("Tables, <Nothing>, <Nothing>, <Nothing>, TABLE")
        cmbRestrictions.Items.Add("Columns")
        cmbRestrictions.Items.Add("Indexes")
        cmbRestrictions.Items.Add("DataSourceInformation")
        cmbRestrictions.Items.Add("DataTypes")
        cmbRestrictions.Items.Add("Table Relationships")
        cmbRestrictions.Items.Add("Index List")
        'cmbRestrictions.Items.Add("Test")


        sb.Append("------------------- Started OK ------------------------------------------------------------------------------------------------------------------------ " & vbCrLf & vbCrLf)
        Me.Show() 'Show this form before showing the Message form
        Message.Add(sb.ToString)

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Application

        DisconnectFromAppNet() 'Disconnect from the Application Network.

        SaveFormSettings() 'Save the settings of this form.

        SaveProjectSettings()  'Save the project settings.

        ApplicationInfo.WriteFile() 'Update the Application Information file.
        ApplicationInfo.UnlockApplication()

        Project.SaveLastProjectInfo() 'Save information about the last project used.

        'Project.SaveProjectInfoFile() 'Update the Project Information file. This is not required unless there is a change made to the project.

        Project.Usage.SaveUsageInfo() 'Save Project usage information.

        ApplicationUsage.SaveUsageInfo() 'Save Application usage information.

        Application.Exit()

    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'Save the form settings if the form state is normal. (A minimised form will have the incorrect size and location.)
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        End If
    End Sub

#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Open and Close Forms - Code used to open and close other forms." '-------------------------------------------------------------------------------------------------------------------

    Private Sub btnSqlCommand_Click(sender As Object, e As EventArgs) Handles btnSqlCommand.Click
        'Show the SQL Command form:

        If IsNothing(SqlCommand) Then
            SqlCommand = New frmSqlCommand
            SqlCommand.Show()
        Else
            SqlCommand.Show()
        End If
    End Sub

    Private Sub SqlCommand_FormClosed(sender As Object, e As FormClosedEventArgs) Handles SqlCommand.FormClosed
        SqlCommand = Nothing
    End Sub

    Private Sub btnModify_Click(sender As Object, e As EventArgs) Handles btnModify.Click
        'Show the Modify Database form:

        If IsNothing(ModifyDatabase) Then
            ModifyDatabase = New frmModifyDatabase
            ModifyDatabase.Show()
        Else
            ModifyDatabase.Show()
        End If
    End Sub

    Private Sub ModifyDatabase_FormClosed(sender As Object, e As FormClosedEventArgs) Handles ModifyDatabase.FormClosed
        ModifyDatabase = Nothing
    End Sub

    Private Sub btnCreateNew_Click(sender As Object, e As EventArgs) Handles btnCreateNew.Click
        'Show the Create New Database form:

        If IsNothing(CreateNewDatabase) Then
            CreateNewDatabase = New frmCreateNewDatabase
            CreateNewDatabase.Show()
        Else
            CreateNewDatabase.Show()
        End If
    End Sub

    Private Sub CreateNewDatabase_FormClosed(sender As Object, e As FormClosedEventArgs) Handles CreateNewDatabase.FormClosed
        CreateNewDatabase = Nothing
    End Sub

    Private Sub btnSaveDef_Click(sender As Object, e As EventArgs) Handles btnSaveDef.Click
        'Show the Save Databse Definition form:

        If IsNothing(SaveDatabaseDefinition) Then
            SaveDatabaseDefinition = New frmSaveDatabaseDefinition
            SaveDatabaseDefinition.Show()
        Else
            SaveDatabaseDefinition.Show()
        End If
    End Sub

    Private Sub SaveDatabaseDefinition_FormClosed(sender As Object, e As FormClosedEventArgs) Handles SaveDatabaseDefinition.FormClosed
        SaveDatabaseDefinition = Nothing
    End Sub

    Private Sub btnSaveTableDefinition_Click(sender As Object, e As EventArgs) Handles btnSaveTableDefinition.Click
        'Show the Save Table Definition form:

        If IsNothing(SaveTableDefinition) Then
            SaveTableDefinition = New frmSaveTableDefinition
            SaveTableDefinition.Show()
        Else
            SaveTableDefinition.Show()
        End If
    End Sub

    Private Sub SaveTableDefinition_FormClosed(sender As Object, e As FormClosedEventArgs) Handles SaveTableDefinition.FormClosed
        SaveTableDefinition = Nothing
    End Sub

    Private Sub btnShowTableProperties_Click(sender As Object, e As EventArgs) Handles btnShowTableProperties.Click
        'Show the Show Table Properties form:

        If IsNothing(ShowTableProperties) Then
            ShowTableProperties = New frmShowTableProperties
            ShowTableProperties.Show()
            ShowTableProperties.UpdateProperties(ds)
        Else
            ShowTableProperties.Show()
            ShowTableProperties.UpdateProperties(ds)
        End If
    End Sub

    Private Sub ShowTableProperties_FormClosed(sender As Object, e As FormClosedEventArgs) Handles ShowTableProperties.FormClosed
        ShowTableProperties = Nothing
    End Sub

    Private Sub btnCreateNewTable_Click(sender As Object, e As EventArgs) Handles btnCreateNewTable.Click
        'Show the Create New Table form:

        If IsNothing(CreateNewTable) Then
            CreateNewTable = New frmCreateNewTable
            CreateNewTable.Show()
        Else
            CreateNewTable.Show()
        End If
    End Sub

    Private Sub CreateNewTable_FormClosed(sender As Object, e As FormClosedEventArgs) Handles CreateNewTable.FormClosed
        CreateNewTable = Nothing
    End Sub

#End Region 'Open and Close Forms =============================================================================================================================================================


#Region " Form Methods - The main actions performed by this form." '---------------------------------------------------------------------------------------------------------------------------

    Private Sub btnProject_Click(sender As Object, e As EventArgs) Handles btnProject.Click
        Project.SelectProject()
    End Sub

    Private Sub btnAppInfo_Click(sender As Object, e As EventArgs) Handles btnAppInfo.Click
        ApplicationInfo.ShowInfo()
    End Sub

    Private Sub btnAndorville_Click(sender As Object, e As EventArgs) Handles btnAndorville.Click
        ApplicationInfo.ShowInfo()
    End Sub

#Region " Online/Offline code" '---------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub btnOnline_Click(sender As Object, e As EventArgs) Handles btnOnline.Click
        'Connect to or disconnect from the Application Network.
        If ConnectedToAppnet = False Then
            ConnectToAppNet()
        Else
            DisconnectFromAppNet()
        End If
    End Sub

    Private Sub ConnectToAppNet()
        'Connect to the Application Network. (Message Exchange)

        Dim Result As Boolean

        If IsNothing(client) Then
            client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
        End If

        If client.State = ServiceModel.CommunicationState.Faulted Then
            Message.SetWarningStyle()
            Message.Add("client state is faulted. Connection not made!" & vbCrLf)
        Else
            Try
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 8) 'Temporarily set the send timeaout to 8 seconds

                Result = client.Connect(ApplicationInfo.Name, ServiceReference1.clsConnectionAppTypes.Application, False, False) 'Application Name is "Application_Template"
                'appName, appType, getAllWarnings, getAllMessages

                If Result = True Then
                    Message.Add("Connected to the Application Network as " & ApplicationInfo.Name & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                    btnOnline.Text = "Online"
                    btnOnline.ForeColor = Color.ForestGreen
                    ConnectedToAppnet = True
                    SendApplicationInfo()
                Else
                    Message.Add("Connection to the Application Network failed!" & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                End If
            Catch ex As System.TimeoutException
                Message.Add("Timeout error. Check if the Application Network is running." & vbCrLf)
            Catch ex As Exception
                Message.Add("Error message: " & ex.Message & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
            End Try
        End If

    End Sub

    Private Sub DisconnectFromAppNet()
        'Disconnect from the Application Network.

        Dim Result As Boolean

        If IsNothing(client) Then
            Message.Add("Already disconnected from the Application Network." & vbCrLf)
            btnOnline.Text = "Offline"
            btnOnline.ForeColor = Color.Black
            ConnectedToAppnet = False
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted." & vbCrLf)
            Else
                Try
                    Message.Add("Running client.Disconnect(ApplicationName)   ApplicationName = " & ApplicationInfo.Name & vbCrLf)
                    client.Disconnect(ApplicationInfo.Name) 'NOTE: If Application Network has closed, this application freezes at this line! Try Catch EndTry added to fix this.
                    btnOnline.Text = "Offline"
                    btnOnline.ForeColor = Color.Black
                    ConnectedToAppnet = False
                Catch ex As Exception
                    Message.SetWarningStyle()
                    Message.Add("Error disconnecting from Application Network: " & ex.Message & vbCrLf)
                End Try
            End If
        End If
    End Sub

    Private Sub SendApplicationInfo()
        'Send the application information to the Administrator connections.

        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                'Create the XML instructions to send application information.
                Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                Dim applicationInfo As New XElement("ApplicationInfo")
                Dim name As New XElement("Name", Me.ApplicationInfo.Name)
                applicationInfo.Add(name)

                Dim exePath As New XElement("ExecutablePath", Me.ApplicationInfo.ExecutablePath)
                applicationInfo.Add(exePath)

                Dim directory As New XElement("Directory", Me.ApplicationInfo.ApplicationDir)
                applicationInfo.Add(directory)
                Dim description As New XElement("Description", Me.ApplicationInfo.Description)
                applicationInfo.Add(description)
                xmessage.Add(applicationInfo)
                doc.Add(xmessage)
                client.SendMessage("ApplicationNetwork", doc.ToString)
            End If
        End If

    End Sub


#End Region 'Online/Offline code ==============================================================================================================================================================

#Region " Process XMessages" '-----------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub XMsg_Instruction(Info As String, Locn As String) Handles XMsg.Instruction
        'Process an XMessage instruction.
        'An XMessage is a simplified XSequence. It is used to exchange information between Andorville (TM) applications.
        '
        'An XSequence file is an AL-H7 (TM) Information Vector Sequence stored in an XML format.
        'AL-H7(TM) is the name of a programming system that uses sequences of information and location value pairs to store data items or processing steps.
        'A single information and location value pair is called a knowledge element (or noxel).
        'Any program, mathematical expression or data set can be expressed as an Information Vector Sequence.

        'Add code here to process the XMessage instructions.
        'See other Andorville(TM) applciations for examples.

        Select Case Locn
            Case ""

        End Select

    End Sub

    Private Sub SendMessage()
        'Code used to send a message after a timer delay.
        'The message destination is stored in MessageDest
        'The message text is stored in MessageText
        Timer1.Interval = 100 '100ms delay
        Timer1.Enabled = True 'Start the timer.
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If IsNothing(client) Then
            Message.SetWarningStyle()
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.SetWarningStyle()
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                Try
                    Message.Add("Sending a message. Number of characters: " & MessageText.Length & vbCrLf)
                    client.SendMessage(MessageDest, MessageText)
                    Message.XAdd(MessageText & vbCrLf)
                    MessageText = "" 'Clear the message after it has been sent.
                Catch ex As Exception
                    Message.SetWarningStyle()
                    Message.Add("Error sending message: " & ex.Message & vbCrLf)
                End Try
            End If
        End If

        'Stop timer:
        Timer1.Enabled = False
    End Sub


#End Region 'Process XMessages ================================================================================================================================================================

    Private Sub Project_ErrorMessage(Msg As String) Handles Project.ErrorMessage
        'Display the Project error message:
        Message.SetWarningStyle()
        Message.Add(Msg)
        Message.SetNormalStyle()
    End Sub

    Private Sub Project_Message(Msg As String) Handles Project.Message
        'Display the Project message:
        Message.Add(Msg)
    End Sub

    Private Sub Project_Closing() Handles Project.Closing
        'The current project is closing.

        'Save the current project settings.
        'SaveProjectSettings() 'Define this subroutine if project settings need to be saved.

        'Save the current project usage information:
        Project.Usage.SaveUsageInfo()

    End Sub


    Private Sub Project_Selected() Handles Project.Selected
        'A new project has been selected.

        Project.ReadProjectInfoFile()
        Project.Usage.StartTime = Now

        ApplicationInfo.SettingsLocn = Project.SettingsLocn
        Message.SettingsLocn = Project.SettingsLocn

        'Restore the new project settings:
        RestoreProjectSettings() 'Define this subroutine if project settings need to be restored.
        txtDatabase.Text = DatabasePath
        FillLstTables()

        'Show the project information:
        txtProjectName.Text = Project.Name
        txtProjectDescription.Text = Project.Description
        Select Case Project.Type
            Case ADVL_Utilities_Library_1.Project.Types.Directory
                txtProjectType.Text = "Directory"
            Case ADVL_Utilities_Library_1.Project.Types.Archive
                txtProjectType.Text = "Archive"
            Case ADVL_Utilities_Library_1.Project.Types.Hybrid
                txtProjectType.Text = "Hybrid"
            Case ADVL_Utilities_Library_1.Project.Types.None
                txtProjectType.Text = "None"
        End Select

        txtCreationDate.Text = Format(Project.CreationDate, "d-MMM-yyyy H:mm:ss")
        txtLastUsed.Text = Format(Project.Usage.LastUsed, "d-MMM-yyyy H:mm:ss")
        Select Case Project.SettingsLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSettingsLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSettingsLocationType.Text = "Archive"
        End Select
        txtSettingsLocationPath.Text = Project.SettingsLocn.Path
        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtDataLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtDataLocationType.Text = "Archive"
        End Select
        txtDataLocationPath.Text = Project.DataLocn.Path

    End Sub

    Private Sub btnDatabase_Click(sender As Object, e As EventArgs) Handles btnDatabase.Click
        'Select the database file:

        If txtDatabase.Text <> "" Then
            Dim fInfo As New System.IO.FileInfo(txtDatabase.Text)
            OpenFileDialog1.InitialDirectory = fInfo.DirectoryName
            OpenFileDialog1.Filter = "Database |*.accdb"
            OpenFileDialog1.FileName = fInfo.Name
        Else
            OpenFileDialog1.InitialDirectory = System.Environment.SpecialFolder.MyDocuments
            OpenFileDialog1.Filter = "Database |*.accdb"
            OpenFileDialog1.FileName = ""
        End If

        If OpenFileDialog1.ShowDialog() = vbOK Then
            DatabasePath = OpenFileDialog1.FileName
            txtDatabase.Text = DatabasePath
            FillLstTables()
            FillCmbSelectTable()
        End If
    End Sub

    Private Sub btnUpdateTables_Click(sender As Object, e As EventArgs) Handles btnUpdateTables.Click
        FillLstTables()
    End Sub

    Public Sub FillLstTables()
        'Fill the lstSelectTable listbox with the available tables in the selected database.

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        lstTables.Items.Clear()
        lstFields.Items.Clear()

        'Specify the connection string:
        'Access 2003
        'connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + _
        '"data source = " + txtDatabase.Text

        If txtDatabase.Text = "" Then
            Exit Sub
        End If

        'Access 2007:
        'connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        '"data source = " + txtDatabase.Text
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + DatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)

        Try
            conn.Open()

            Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
            dt = conn.GetSchema("Tables", restrictions)

            'Fill lstSelectTable
            Dim dr As DataRow
            Dim I As Integer 'Loop index
            Dim MaxI As Integer

            MaxI = dt.Rows.Count
            For I = 0 To MaxI - 1
                dr = dt.Rows(0)
                lstTables.Items.Add(dt.Rows(I).Item(2).ToString)
            Next I

            conn.Close()
        Catch ex As Exception
            'Main.ShowMessage("Error opening database: " & Main.DatabasePath & vbCrLf, Color.Red)
            'Main.ShowMessage(ex.Message & vbCrLf & vbCrLf, Color.Blue)
            Message.Add("Error opening database: " & DatabasePath & vbCrLf)
            Message.Add(ex.Message & vbCrLf & vbCrLf)
        End Try

    End Sub

    Public Sub FillCmbSelectTable()
        'Fill the cmbSelectTable listbox with the available tables in the selected database.

        If DatabasePath = "" Then
            Message.AddWarning("No database selected!" & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        cmbSelectTable.Text = ""
        cmbSelectTable.Items.Clear()
        'DataGridView1.Rows.Clear()
        ds.Clear()
        ds.Reset()
        'DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        'Specify the connection string:
        'Access 2003
        'connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + _
        '"data source = " + txtDatabase.Text

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + DatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        'This error occurs on the above line (conn.Open()):
        'Additional information: The 'Microsoft.ACE.OLEDB.12.0' provider is not registered on the local machine.
        'Fix attempt: 
        'http://www.microsoft.com/en-us/download/confirmation.aspx?id=23734
        'Download AccessDatabaseEngine.exe
        'Run the file to install the 2007 Office System Driver: Data Connectivity Components.


        Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
        dt = conn.GetSchema("Tables", restrictions)

        'Fill lstSelectTable
        Dim dr As DataRow
        Dim I As Integer 'Loop index
        Dim MaxI As Integer

        MaxI = dt.Rows.Count
        For I = 0 To MaxI - 1
            dr = dt.Rows(0)
            cmbSelectTable.Items.Add(dt.Rows(I).Item(2).ToString)
        Next I

        conn.Close()

    End Sub

    Private Sub cmbDatabaseType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDatabaseType.SelectedIndexChanged

        Select Case cmbDatabaseType.SelectedItem
            Case "Access"
                DatabaseType = DatabaseTypes.Access
            Case "Unknown"
                DatabaseType = DatabaseTypes.Unknown
            Case Else '
                DatabaseType = DatabaseTypes.Unknown
        End Select
    End Sub

    Private Sub lstTables_Click(sender As Object, e As EventArgs) Handles lstTables.Click
        FillLstFields()
    End Sub

    Private Sub FillLstFields()
        'Fill the lstSelectField listbox with the availalble fields in the selected table.

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim commandString As String 'Declare a command string - contains the query to be passed to the database.
        Dim ds As DataSet 'Declate a Dataset.
        Dim dt As DataTable

        If lstTables.SelectedIndex = -1 Then 'No item is selected
            lstFields.Items.Clear()

        Else 'A table has been selected. List its fields:
            lstFields.Items.Clear()

            'Specify the connection string (Access 2007):
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
            "data source = " + txtDatabase.Text

            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            'Specify the commandString to query the database:
            'commandString = "SELECT * FROM " + lstTables.SelectedItem.ToString
            commandString = "SELECT Top 500 * FROM " + lstTables.SelectedItem.ToString
            Dim dataAdapter As New System.Data.OleDb.OleDbDataAdapter(commandString, conn)
            ds = New DataSet
            dataAdapter.Fill(ds, "SelTable") 'ds was defined earlier as a DataSet
            dt = ds.Tables("SelTable")

            Dim NFields As Integer
            NFields = dt.Columns.Count
            Dim I As Integer
            For I = 0 To NFields - 1
                lstFields.Items.Add(dt.Columns(I).ColumnName.ToString)
            Next

            conn.Close()

        End If
    End Sub

    Private Sub cmbSelectTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSelectTable.SelectedIndexChanged
        'Update DataGridView1:

        If IsNothing(cmbSelectTable.SelectedItem) Then
            Exit Sub
        End If

        ''Variables used to connect to a database and open a table:
        'Dim connString As String
        'Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        'Dim ds As DataSet = New DataSet
        'Dim da As OleDb.OleDbDataAdapter
        'Dim tables As DataTableCollection = ds.Tables

        TableName = cmbSelectTable.SelectedItem.ToString
        Query = "Select Top 500 * From " & TableName
        txtQuery.Text = Query

        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & DatabasePath
        myConnection.ConnectionString = connString
        myConnection.Open()

        da = New OleDb.OleDbDataAdapter(Query, myConnection)

        da.MissingSchemaAction = MissingSchemaAction.AddWithKey 'This statement is required to obtain the correct result from the statement: ds.Tables(0).Columns(0).MaxLength (This fixes a Microsoft bug: http://support.microsoft.com/kb/317175 )

        ds.Clear()
        ds.Reset()

        da.FillSchema(ds, SchemaType.Source, TableName)

        da.Fill(ds, TableName)

        DataGridView1.AutoGenerateColumns = True

        DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke

        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.AutoResizeColumns()

        DataGridView1.Update()
        DataGridView1.Refresh()
        myConnection.Close()

    End Sub

    Private Sub btnApplyQuery_Click(sender As Object, e As EventArgs) Handles btnApplyQuery.Click
        'Apply query on Table tab.
        'Update DataGridView1:

        If IsNothing(cmbSelectTable.SelectedItem) Then
            Exit Sub
        End If

        ''Variables used to connect to a database and open a table:
        'Dim connString As String
        'Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        'Dim ds As DataSet = New DataSet
        'Dim da As OleDb.OleDbDataAdapter
        'Dim tables As DataTableCollection = ds.Tables

        TableName = cmbSelectTable.SelectedItem.ToString

        'txtQuery.Text = "Select * From " & TableName

        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & DatabasePath
        myConnection.ConnectionString = connString
        myConnection.Open()

        da = New OleDb.OleDbDataAdapter(txtQuery.Text, myConnection)

        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        'Debug.Print("Start filling DataSet ds")
        ds.Clear()
        ds.Reset()
        Try
            da.Fill(ds, TableName)

            DataGridView1.AutoGenerateColumns = True

            DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke

            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.AutoResizeColumns()

            DataGridView1.Update()
            DataGridView1.Refresh()
        Catch ex As Exception
            Message.Add("Error applying query." & vbCrLf)
            Message.Add(ex.Message & vbCrLf & vbCrLf)
        End Try

        myConnection.Close()
    End Sub

    Public Sub ApplySqlCommand()
        'Apply the Sql Command to the selected database.
        'The command is stored in the property SqlCommandText

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.

        If txtDatabase.Text = "" Then
            Exit Sub
        End If

        ''Access 2007:
        'connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        '"data source = " + txtDatabase.Text

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + DatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim cmd As New OleDb.OleDbCommand
        cmd.CommandText = SqlCommandText
        cmd.Connection = conn

        Try
            cmd.ExecuteNonQuery()
            Message.Add("OK" & vbCrLf)
            SqlCommandResult = "OK"
        Catch ex As Exception
            'Message.SetWarningStyle()
            'Message.Add(ex.Message & vbCrLf)
            'Message.SetNormalStyle()
            Message.AddWarning(ex.Message & vbCrLf)
            'Message.AddWarning(vbCrLf & "ex.ToString: " & ex.ToString & vbCrLf)
            'Message.AddWarning(vbCrLf & "ex.Source: " & ex.Source & vbCrLf)
            'Message.AddWarning(vbCrLf & "ex.InnerException.Message: " & ex.InnerException.Message & vbCrLf) 'ERROR

            SqlCommandResult = "Error"
        End Try

    End Sub

    Private Sub btnMessages_Click(sender As Object, e As EventArgs) Handles btnMessages.Click
        'Show the Messages form.
        Message.ApplicationName = ApplicationInfo.Name
        Message.SettingsLocn = Project.SettingsLocn
        Message.Show()
        Message.MessageForm.BringToFront()
    End Sub

    Private Sub btnSaveChanges_Click(sender As Object, e As EventArgs) Handles btnSaveChanges.Click
        'Save the changes made to the data in DataGridView1 to the corresponding table in the database:

        ''Variables used to connect to a database and open a table:
        'Dim connString As String
        'Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        'Dim ds As DataSet = New DataSet
        'Dim da As OleDb.OleDbDataAdapter
        'Dim tables As DataTableCollection = ds.Tables

        Dim cb = New OleDb.OleDbCommandBuilder(da)
        Try
            DataGridView1.EndEdit()
            da.Update(ds.Tables(0))
            ds.Tables(0).AcceptChanges()
            UpdateNeeded = False
            DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke
        Catch ex As Exception
            Message.SetWarningStyle()
            Message.Add("Error saving changes." & vbCrLf)
            Message.SetNormalStyle()
            Message.Color = Color.Blue
            Message.Add(ex.Message & vbCrLf & vbCrLf)
            Message.Color = Color.Black
            'TDS_Finances.ShowMessage(ex.HelpLink.ToString & vbCrLf & vbCrLf, Color.Blue)
        End Try
    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        UpdateNeeded = True
    End Sub

    Private Sub btnViewSchema_Click(sender As Object, e As EventArgs) Handles btnViewSchema.Click
        'ViewSchema button pressed:
        'View the database schema specified by the selected restrictions.

        'View the Restrictions schema for a list of valid restrictions.

        ' Select Case cmbRestrictions.SelectedItem
        Select Case cmbRestrictions.Text
            Case "<No restrictions>"
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & DatabasePath
                myConnection.ConnectionString = connString
                myConnection.Open()
                DataGridView2.DataSource = myConnection.GetSchema
                myConnection.Close()

            Case "Tables"
                'Restrictions: TABLE_CATALOG TABLE_SCHEMA TABLE_NAME TABLE_TYPE
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & DatabasePath
                myConnection.ConnectionString = connString
                myConnection.Open()
                DataGridView2.DataSource = myConnection.GetSchema("Tables")
                myConnection.Close()

            Case "Restrictions"
                'This provides a list of restrictions that can be applied to the Columns, Indexes, Procedures, Tables and Views schema.
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & DatabasePath
                myConnection.ConnectionString = connString
                myConnection.Open()
                DataGridView2.DataSource = myConnection.GetSchema("Restrictions")
                myConnection.Close()

                'Case "Tables, <Nothing>, <Nothing>, <Nothing>, TABLE"
            Case "Tables (excluding System and Access tables)"
                'Restrictions: TABLE_CATALOG TABLE_SCHEMA TABLE_NAME TABLE_TYPE
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & DatabasePath
                myConnection.ConnectionString = connString
                myConnection.Open()
                DataGridView2.DataSource = myConnection.GetSchema("Tables", New String() {Nothing, Nothing, Nothing, "TABLE"})
                myConnection.Close()

            Case "Columns"
                'Restrictions: TABLE_CATALOG TABLE_SCHEMA TABLE_NAME COLUMN_NAME
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & DatabasePath
                myConnection.ConnectionString = connString
                myConnection.Open()
                DataGridView2.DataSource = myConnection.GetSchema("Columns")
                myConnection.Close()

            Case "Indexes"
                'http://msdn.microsoft.com/en-us/library/system.data.oledb.oledbschemaguid.indexes.aspx
                'Restrictions: TABLE_CATALOG TABLE_SCHEMA INDEX_NAME TYPE TABLE_NAME
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & DatabasePath
                myConnection.ConnectionString = connString
                myConnection.Open()
                DataGridView2.DataSource = myConnection.GetSchema("Indexes")
                myConnection.Close()

            Case "DataSourceInformation"
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & DatabasePath
                myConnection.ConnectionString = connString
                myConnection.Open()
                DataGridView2.DataSource = myConnection.GetSchema("DataSourceInformation")
                myConnection.Close()

            Case "DataTypes"
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & DatabasePath
                myConnection.ConnectionString = connString
                myConnection.Open()
                DataGridView2.DataSource = myConnection.GetSchema("DataTypes")
                myConnection.Close()

            Case "Table Relationships"
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & DatabasePath
                myConnection.ConnectionString = connString
                myConnection.Open()
                'Dim fkTable As DataTable
                'fkTable = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Foreign_Keys, New String() {Nothing, Nothing, Nothing, Nothing, Nothing, Nothing})
                'DataGridView1.DataSource = fkTable
                DataGridView2.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Foreign_Keys, New String() {Nothing, Nothing, Nothing, Nothing, Nothing, Nothing})
                myConnection.Close()


            Case "Index List"
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & DatabasePath
                myConnection.ConnectionString = connString
                myConnection.Open()
                'Dim fkTable As DataTable
                'fkTable = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Foreign_Keys, New String() {Nothing, Nothing, Nothing, Nothing, Nothing, Nothing})
                'DataGridView1.DataSource = fkTable

                'http://support.microsoft.com/kb/309488 List of OleDbSchema members:

                DataGridView2.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Indexes, New String() {Nothing, Nothing, Nothing, Nothing, Nothing})
                myConnection.Close()

            Case "Test"
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & DatabasePath
                myConnection.ConnectionString = connString
                myConnection.Open()
                'DataGridView1.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Foreign_Keys, New String() {Nothing, Nothing, Nothing, Nothing, Nothing, Nothing})
                'DataGridView1.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Views, New String() {Nothing, Nothing})
                ' DataGridView1.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Columns, New String() {Nothing})
                'DataGridView1.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Indexes, New String() {Nothing})
                'DataGridView1.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables_Info, New String() {Nothing})
                'DataGridView1.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Schemata, New String() {Nothing}) 'ERROR
                'DataGridView1.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Constraint_Column_Usage, New String() {Nothing})
                'DataGridView1.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Key_Column_Usage, New String() {Nothing})
                'DataGridView1.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Assertions, New String() {Nothing}) 'ERROR
                'DataGridView1.DataSource = myConnection.GetSchema()
                'DataGridView1.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Columns, New String() {Nothing, Nothing, Nothing, Nothing, Nothing})

                'DataGridView1.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Columns, New String() {Nothing, Nothing, Nothing, Nothing})
                'DataGridView1.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Constraint_Column_Usage, New String() {Nothing, Nothing, Nothing, Nothing})
                'DataGridView1.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Column_Privileges, New String() {Nothing, Nothing, Nothing, Nothing}) 'The Column_Privileges OleDbSchemaGuid is not a supported schema by the 'Microsoft.ACE.OLEDB.12.0' provider.
                'DataGridView1.DataSource = myConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Column_Domain_Usage, New String() {Nothing, Nothing, Nothing, Nothing}) 'The Column_Domain_Usage OleDbSchemaGuid is not a supported schema by the 'Microsoft.ACE.OLEDB.12.0' provider.

                myConnection.Close()
        End Select

        DataGridView2.AutoResizeColumns()

        'cmbRestrictions.Items.Add("<No restrictions>")
        'cmbRestrictions.Items.Add("Tables")
        'cmbRestrictions.Items.Add("Restrictions")
        'cmbRestrictions.Items.Add("Tables, <Nothing>, <Nothing>, <Nothing>, TABLE")
        'cmbRestrictions.Items.Add("Columns")
        'cmbRestrictions.Items.Add("Indexes")
        'cmbRestrictions.Items.Add("DataSourceInformation")
        'cmbRestrictions.Items.Add("DataTypes")
    End Sub

    Private Sub lstTables_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstTables.SelectedIndexChanged

    End Sub


    'Private Sub btnSaveTableDefinition_Click(sender As Object, e As EventArgs) Handles btnSaveTableDefinition.Click
    '    'Save the table defintion in an XML file:
    '    'TableDefinition
    '    '   Summary
    '    '   Primary Keys
    '    '   Columns
    '    '       Column1
    '    '       ...


    '    'SaveFileDialog1.Filter = "Table Definition |*.TableDef"

    '    Dim FilePath As String
    '    If Trim(Main.ProjectPath) <> "" Then 'Write the Form Settings file in the Project Directory
    '        FilePath = Main.ProjectPath
    '    Else 'Write the Form Settings file in the Application Directory
    '        FilePath = Main.ApplicationDir
    '    End If

    '    SaveFileDialog1.InitialDirectory = FilePath

    '    If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then

    '        'ds.Tables(0).WriteXmlSchema(SaveFileDialog1.FileName)
    '        'ds.Tables(0).WriteXml(SaveFileDialog1.FileName)
    '        'ds.Tables(0).WriteXmlSchema(SaveFileDialog1.FileName, True)

    '        Dim doc = New XDocument 'Create the XDocument to hold the XML data.

    '        'Add the XML declaration:
    '        Dim decl = New XDeclaration("1.0", "utf-8", "yes")
    '        doc.Declaration = decl

    '        doc.Add(New XComment(""))
    '        doc.Add(New XComment("Exported table definition."))

    '        Dim tableData As New XElement("TableDefinition")

    '        tableData.Add(New XComment(""))
    '        tableData.Add(New XComment("Table summary."))

    '        'Add table summary
    '        Dim summary = New XElement("Summary")
    '        summary.Add(New XElement("Database", Main.DatabasePath))
    '        summary.Add(New XElement("TableName", ds.Tables(0).TableName))
    '        summary.Add(New XElement("NumberOfColumns", ds.Tables(0).Columns.Count))
    '        summary.Add(New XElement("NumberOfPrimaryKeys", ds.Tables(0).PrimaryKey.Count))
    '        'summary.Add(New XElement("PrimaryKey", ds.Tables(0).PrimaryKey))

    '        tableData.Add(summary)

    '        tableData.Add(New XComment(""))
    '        tableData.Add(New XComment("Primary keys."))
    '        Dim NPrimaryKeys As Integer = ds.Tables(0).PrimaryKey.Count
    '        Dim I As Integer
    '        Dim primaryKeys = New XElement("PrimaryKeys")
    '        For I = 1 To NPrimaryKeys
    '            primaryKeys.Add(New XElement("Key", ds.Tables(0).PrimaryKey(I - 1)))
    '        Next

    '        tableData.Add(primaryKeys)

    '        'Add column definitions:
    '        Dim ColNo As Integer
    '        Dim NCols As Integer
    '        NCols = ds.Tables(0).Columns.Count
    '        Dim columns = New XElement("Columns")
    '        For ColNo = 1 To NCols
    '            Dim column = New XElement("Column")
    '            'For ColNo = 1 To NCols
    '            'Dim field = New XElement(FieldName(ColNo - 1), TDS_Finances.ViewTables.DataGridView1.Rows(RowNo - 1).Cells(ColNo - 1).Value.ToString)
    '            'record.Add(field)
    '            'Next
    '            Dim setting1 = New XElement("ColumnName", ds.Tables(0).Columns(ColNo - 1).ColumnName)
    '            column.Add(setting1)
    '            Dim setting2 = New XElement("DataType", ds.Tables(0).Columns(ColNo - 1).DataType)
    '            column.Add(setting2)
    '            'Dim setting3 = New XElement("MaxLength", ds.Tables(0).Columns(ColNo - 1).MaxLength) 'This returns -1
    '            'column.Add(setting3)
    '            'Dim setting3 = New XElement("Capton", ds.Tables(0).Columns(ColNo - 1).Caption)
    '            'column.Add(setting3)
    '            'Dim setting8 = New XElement("Attributes", ds.Tables(0).Columns(ColNo - 1).DataType.Attributes)
    '            'column.Add(setting8)
    '            Dim setting4 = New XElement("AllowDBNull", ds.Tables(0).Columns(ColNo - 1).AllowDBNull)
    '            column.Add(setting4)
    '            Dim setting5 = New XElement("AutoIncrement", ds.Tables(0).Columns(ColNo - 1).AutoIncrement)
    '            column.Add(setting5)
    '            Dim setting6 = New XElement("StringFieldLength", ds.Tables(0).Columns(ColNo - 1).MaxLength)
    '            column.Add(setting6)
    '            'Dim settings7 = New XElement("MaxLen", ds.Tables(0).Columns(ColNo - 1).MaxLength)
    '            columns.Add(column)
    '        Next

    '        'doc.Add(records)
    '        tableData.Add(New XComment(""))
    '        tableData.Add(New XComment("List of column definitions."))
    '        tableData.Add(columns)

    '        'Add Relations:
    '        Dim relations = New XElement("Relations")
    '        Dim RelCount As Integer
    '        'RelCount = ds.Tables(0).ChildRelations.Count
    '        RelCount = ds.Tables(0).DataSet.Relations.Count

    '        relations.Add(New XElement("NumberOfChildRelations", RelCount))
    '        For I = 1 To RelCount
    '            Dim relation = New XElement("Relation")
    '            Dim relName = New XElement("RelationName", ds.Tables(0).ChildRelations(I - 1).RelationName)
    '            relation.Add(relName)
    '            Dim childTable = New XElement("ChildTable", ds.Tables(0).ChildRelations(I - 1).ChildTable)
    '            relation.Add(childTable)
    '            Dim childColumn = New XElement("ChildColumn", ds.Tables(0).ChildRelations(I - 1).ChildColumns(0).ColumnName)
    '            relation.Add(childColumn)
    '            relations.Add(New XComment(""))
    '            relations.Add(relation)
    '        Next

    '        tableData.Add(New XComment(""))
    '        tableData.Add(New XComment("List of table relations."))
    '        tableData.Add(relations)


    '        doc.Add(tableData)

    '        Dim SaveFilePath As String
    '        SaveFilePath = SaveFileDialog1.FileName
    '        'If IO.Path.GetExtension = "
    '        'SaveFilePath = IO.Path.GetFileNameWithoutExtension(SaveFilePath) & ".TableDef"
    '        'SaveFilePath = IO.Path.GetFullPath(SaveFilePath) & IO.Path.GetFileNameWithoutExtension(SaveFilePath) & ".TableDef"
    '        SaveFilePath = IO.Path.GetDirectoryName(SaveFilePath) & "\" & IO.Path.GetFileNameWithoutExtension(SaveFilePath) & ".TableDef"

    '        'doc.Save(SaveFileDialog1.FileName)
    '        doc.Save(SaveFilePath)

    '    Else

    '    End If

    'End Sub




















#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Events - Events raised by this form." '-----------------------------------------------------------------------------------------------------------------------------------------

#End Region 'Form Events ----------------------------------------------------------------------------------------------------------------------------------------------------------------------



End Class
