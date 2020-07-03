Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports Extensibility
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop

''' <summary>
'''The Object For implementing an Add-In. Users don't need to change the "GUID"
'''for this add-in. But if you are trying to use this code and modify something 
'''to use it as well as using the provided .dll file as one of the add-ins then you must change it.
''' </summary>
''' <seealso class='IDTExtensibility2' />
<ComVisible(True), Guid("B3C60B32-6851-472F-A98E-99278ED7B539"), ProgId("SheetsCompatibilityAddIn.Connect")>
Public Class Connect
    Implements Extensibility.IDTExtensibility2
    'Interop VBE application object
    Private _VBE As VBE
    Private _AddIn As AddIn
    ' Buttons created by the add-in
    Private WithEvents _Button As CommandBarButton
     'Window created by add-in
    Private _ApiSummaryWindow As Window
    Dim count As Integer = 1
    Dim lines As List(Of String) = Nothing
    Dim fileId As String = ""
    
    ''' <summary>
    ''' Implements the OnConnection method of the IDTExtensibility2 interface.
    ''' Receives notification that the Add-in is being loaded.
    ''' </summary>
    ''' <param name="Application">
    ''' Root object of the host application.
    ''' </param>
    ''' <param name="ConnectMode">
    ''' Describes how the Add-in is being loaded.
    ''' </param>
    ''' <param name="AddInInst">
    ''' Object representing this Add-in.
    ''' </param>
    ''' <seealso class='IDTExtensibility2' />
    Private Sub OnConnection(Application As Object, ConnectMode As Extensibility.ext_ConnectMode,
      AddInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection
        Try
            _VBE = DirectCast(Application, VBE)
            _AddIn = DirectCast(AddInInst, AddIn)
            Select Case ConnectMode
                Case Extensibility.ext_ConnectMode.ext_cm_Startup
                ' OnStartupComplete will be called
                Case Extensibility.ext_ConnectMode.ext_cm_AfterStartup
                    InitializeAddIn()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub

    ''' <summary>
    ''' Implements the OnDisconnection method of the IDTExtensibility2 interface.
    ''' Receives notification that the Add-in is being unloaded.
    ''' </summary>
    ''' <param name="RemoveMode">
    ''' Describes how the Add-in is being unloaded.
    ''' </param>
    ''' <param name="custom">
    ''' Array of parameters that are host application-specific.
    ''' </param>
    ''' <seealso class='IDTExtensibility2' />
    Private Sub OnDisconnection(RemoveMode As Extensibility.ext_DisconnectMode,
      ByRef custom As System.Array) Implements IDTExtensibility2.OnDisconnection
        Try
            Select Case RemoveMode
                Case ext_DisconnectMode.ext_dm_HostShutdown, ext_DisconnectMode.ext_dm_UserClosed
                    If Not (_Button Is Nothing) Then
                        _Button.Delete()
                    End If
            End Select
        Catch e As System.Exception
            System.Windows.Forms.MessageBox.Show(e.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' Implements the OnStartupComplete method of the IDTExtensibility2 interface.
    ''' Receives notification that the host application has completed loading.
    ''' </summary>
    ''' <param name="custom">
    ''' Array of parameters that are host application-specific.
    ''' </param>
    ''' <seealso class='IDTExtensibility2' />
    Private Sub OnStartupComplete(ByRef custom As System.Array) _
      Implements IDTExtensibility2.OnStartupComplete
        InitializeAddIn()
    End Sub

    ''' <summary>
    ''' Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
    ''' Receives notification that the collection of Add-ins has changed.
    ''' </summary>
    ''' <param name="custom">
    ''' Array of parameters that are host application-specific.
    ''' </param>
    ''' <seealso class='IDTExtensibility2' />
    Private Sub OnAddInsUpdate(ByRef custom As System.Array) _
      Implements IDTExtensibility2.OnAddInsUpdate

    End Sub

    ''' <summary>
    ''' Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
    ''' Receives notification that the host application is being unloaded.
    ''' </summary>
    ''' <param name="custom">
    ''' Array of parameters that are host application-specific.
    ''' </param>
    ''' <seealso class='IDTExtensibility2' />
    Private Sub OnBeginShutdown(ByRef custom As System.Array) _
      Implements IDTExtensibility2.OnBeginShutdown

    End Sub

    ''' <summary>
    ''' Whenever the connection is made Add-in needs to initialize.
    ''' </summary>
    Private Sub InitializeAddIn()
        'On initialization of add-in, one button called "SheetsCompatibility" will create.
        'Private WithEvents _Button As CommandBarButton declare  as a member variable of class
        Dim menuCommandBar As CommandBar
        Dim commandBarControl As CommandBarControl
        Try
            'Decide where the button will create.
            menuCommandBar = _VBE.CommandBars.Item("Menu Bar")
            Dim toolsCommandBar As CommandBar = _VBE.CommandBars.Item("Tools")
            Dim toolsCommandBarControl As CommandBarControl
            Dim position As Integer
            ' Calculate the position of a new commandbarBarButton to the right of the "Tools"
            'option in the menu bar.
            toolsCommandBarControl = DirectCast(toolsCommandBar.Parent, CommandBarControl)
            position = toolsCommandBarControl.Index + 1
            commandBarControl = DirectCast(menuCommandBar.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing,
            position, True), CommandBarControl)
            'Assign control to the button which will be going to create.
            _Button = DirectCast(commandBarControl, CommandBarButton)
            'Personalize it as per your choice.
            _Button.Caption = "S&heetsCompatibility"
            _Button.Style = MsoButtonStyle.msoButtonCaption
            _Button.BeginGroup = False
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub
    
     ''' <summary>
    ''' Host a window by UserControlHost.
    ''' </summary>
    ''' <param name="toolWindowCaption">
    ''' String you need to put as a header on the window.
    ''' </param>
    ''' <param name="toolWindowGuid">
    ''' This uniquely identified a particular window and it is used to store the
    ''' information of it, like its size, position, etc. To create guid in visual studio
    ''' click 'Tools-->Create Guid' copy and use it.
    ''' </param>
    ''' <param name="toolWindowUserControl">
    ''' Windows to host.
    ''' </param>
    Private Function CreateToolWindow(ByVal toolWindowCaption As String, ByVal toolWindowGuid As String,
      ByVal toolWindowUserControl As UserControl) As Window
        Dim userControlObject As Object = Nothing
        Dim userControlHost As UserControlHost
        Dim _SummaryWindow As Window
        Dim progId As String
        'Ensure that you use the same ProgId value used in the ProgId attribute of the UserControlHost 
        progId = "MyVBAAddin.UserControlHost"
        _SummaryWindow = _VBE.Windows.CreateToolWindow(_AddIn, progId, toolWindowCaption, toolWindowGuid, userControlObject)
        userControlHost = DirectCast(userControlObject, UserControlHost)
        _SummaryWindow.Visible = True
        userControlHost.AddUserControl(toolWindowUserControl)
        Return _SummaryWindow
    End Function

    ''' <summary>
    ''' Implementation of the click event of the button created by add-in in "MENU BAR".
    ''' </summary>
    ''' <param name="Ctrl"></param>
    ''' <param name="CancelDefault"></param>
    Private Sub _Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton,
      ByRef CancelDefault As Boolean) Handles _Button.Click
        'If the button is clicked first time.
        If count = 1 Then
            '@todo If the drive 'D' is not available to make a file in it, then one should change this path.
            Const pathToCopyFile As String = "D:\GoogleDrive.xlsm"
            'Make a copy of the file active in excel.
            My.Computer.FileSystem.CopyFile(_VBE.ActiveVBProject.FileName,pathToCopyFile, True)
            Dim letsTry As uploadFileToDrive = New uploadFileToDrive()
            'Upload the file on drive.
            fileId = letsTry.UploadFile(pathToCopyFile)
            'Delete that file.
            My.Computer.FileSystem.DeleteFile(pathToCopyFile)
            'If no error ocurred then fileId is not empty.
            If fileId <> "" Then
                lines = hittingEndPoint.callSheetsAPI(fileId)
            Else
                Exit Sub
            End If
        End If
        'if the data list count item is not zero then all Api in the file is not supported.
        If lines.Count <> 0 Then
            Dim userControlObject As Object = Nothing
            Dim userControlToolWindow As UserControlToolWindow
            Try
                    'If button is clicked the first time then window needs to initialize.
                If _ApiSummaryWindow Is Nothing Then
                    userControlToolWindow = New UserControlToolWindow()
                    _ApiSummaryWindow = CreateToolWindow("Report of API used in this project", "0EB93108-D229-4F6F-82C5-0B96AFFBB9C5", userControlToolWindow)
                    userControlToolWindow.Initialize(_VBE, lines)
                    count += 1
                Else
                    'If the window is previously initialized then make it visible.
                    _ApiSummaryWindow.Visible = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
            'if the data list is empty and filed id is not empty then all API used in the file is supported.
        ElseIf fileId <> "" Then
            MessageBox.Show("Fully Compatible!!")
            count += 1
        End If
    End Sub

End Class
