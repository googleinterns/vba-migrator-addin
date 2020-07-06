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
    ''' Implementation of the click event of the button created by add-in in "MENU BAR".
    ''' </summary>
    ''' <param name="Ctrl"></param>
    ''' <param name="CancelDefault"></param>
    Private Sub _Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton,
      ByRef CancelDefault As Boolean) Handles _Button.Click
        Dim UserForm As userForm = New userForm()
        'Initialize the user form to get user credentials.
        UserForm.Initialize(_VBE, _AddIn)
    End Sub

End Class
