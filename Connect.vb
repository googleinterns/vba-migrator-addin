Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports Extensibility
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop


''' <summary>
''' Each VBA editor add-in made by vb.net class library project has same template only the change in GUID(Which identified the add-in uniquely) and program id 
'''So this his is the template of an add-in for an VB editor. i.e what happened when Startup is complete or OnConnection or OnDisconnection or OnUpdate or OnBeginShutDown.
'''When Connection is made we initialize the add-in by adding one button in the menu-bar and after disconnection we will delete the button from there otherwise what happen
'''our add-in got disconnected and button remains there so when the connection made next time our add-in will make same button so at the end our menu bar get conjusted with the same name button.
''' </summary>

<ComVisible(True), Guid("BDA9ECFF-0EDE-4C1D-81D1-51F6B4FF5F50"), ProgId("MyVBAAddin.Connect")>
Public Class Connect

    Implements Extensibility.IDTExtensibility2

    Private _VBE As VBE
    Private _AddIn As AddIn


    ' Buttons created by the add-in
    Private WithEvents _Button As CommandBarButton

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

    Private Sub OnStartupComplete(ByRef custom As System.Array) _
      Implements IDTExtensibility2.OnStartupComplete

        InitializeAddIn()

    End Sub

    Private Sub OnAddInsUpdate(ByRef custom As System.Array) _
      Implements IDTExtensibility2.OnAddInsUpdate

    End Sub

    Private Sub OnBeginShutdown(ByRef custom As System.Array) _
      Implements IDTExtensibility2.OnBeginShutdown

    End Sub

    Private Sub InitializeAddIn()
        'On initialization of add-in this button will create
        'Private WithEvents _Button As CommandBarButton declare  as a member variable of class
        Dim standardCommandBar As CommandBar
        Dim commandBarControl As CommandBarControl

        Try

            standardCommandBar = _VBE.CommandBars.Item("Menu Bar")     'There are many command bar like standard, Menu Bar etc. So to decide in which we want to make it. 

            commandBarControl = standardCommandBar.Controls.Add(MsoControlType.msoControlButton)
            _Button = DirectCast(commandBarControl, CommandBarButton)
            _Button.Caption = "SheetsCompatibility"              'Name of button
            _Button.Style = MsoButtonStyle.msoButtonIconAndCaption
            _Button.BeginGroup = True

        Catch ex As Exception

            MessageBox.Show(ex.ToString())

        End Try

    End Sub

End Class