'Copyright 2020 Google LLC
'
'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'   https://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.

Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Drawing

'Toolwindows are created by calling the VBE.Windows.CreateToolWindow method.
'This method returns a VBE.Window instance, which can be made visible setting its Window.Visible property to True
'Toolwindows host ActiveX UserDocuments inside, but UserDocuments don't exist in .NET. Fortunately,
'Usercontrols can be used instead, with one caveat: the toolwindow doesn't resize a UserControl 
'automatically as it does with a UserDocument when creating add-ins with Visual Basic 6.0.  
'So to handle the size, position of any Toolwindow we created in vb.net, we need one usercontrol class.
'To know more about this class please refer to this article https://www.mztools.com/articles/2012/MZ2012017.aspx.
'This usercontrol class needs to register in window registry, so to uniquely identify we provide Guid here.
'We can get a new Guid by the following procedure click(Tools->Create Guid) in visual studio.
'@todo Change the name of the add-in.
<ComVisible(True), Guid("150B0A49-A7CD-467B-BCF2-FB47A87E68C4"), ProgId("SheetsCompatibilityAddIn.UserControlHost")>
Public Class UserControlHost
    Private Class SubClassingWindow
        'The NativeWindow class provides the following properties and methods to manage handles: 
        'Handle, CreateHandle, AssignHandle, DestroyHandle, and ReleaseHandle.
        Inherits System.Windows.Forms.NativeWindow
        'An event is a signal that informs an application that something important has occurred.
        'For example, when a user clicks a control on a window, the window can raise a Click event 
        'and call a procedure that handles the event. Here it call '_subClassingWindow_CallBackProc()' procedure.
        Public Event CallBackProc(ByRef m As Message)
        'The IntPtr is the platform-specific type that is used to represent a pointer or a handle.
        'Here this new procedure is used to hold the handle of the platform on which it is running.
        Public Sub New(ByVal handle As IntPtr)
            MyBase.AssignHandle(handle)
        End Sub

        Protected Overrides Sub WndProc(ByRef m As Message)
            Const WM_SIZE As Integer = &H500
            If m.Msg = WM_SIZE Then
                RaiseEvent CallBackProc(m)
            End If
            MyBase.WndProc(m)
        End Sub

        Protected Overrides Sub Finalize()
            Me.ReleaseHandle()
            MyBase.Finalize()
        End Sub

    End Class

    <StructLayout(LayoutKind.Sequential)>
    Private Structure RECT
        Friend Left As Integer
        Friend Top As Integer
        Friend Right As Integer
        Friend Bottom As Integer
    End Structure

    Private Declare Function GetParent Lib "user32" (ByVal hWnd As IntPtr) As IntPtr
    Private Declare Function GetClientRect Lib "user32" Alias "GetClientRect" (ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Integer
    Private _parentHandle As IntPtr
    Private WithEvents _subClassingWindow As SubClassingWindow

    Friend Sub AddUserControl(ByVal control As UserControl)
        _parentHandle = GetParent(Me.Handle)
        _subClassingWindow = New SubClassingWindow(_parentHandle)
        control.Dock = DockStyle.Fill
        Me.Controls.Add(control)
        AdjustSize()
    End Sub

    'This class needs to do subclassing with the parent window to detect when the size is changed.
    'So this function is used to handle the event when the user changes the size.
    Private Sub _subClassingWindow_CallBackProc(ByRef m As System.Windows.Forms.Message) Handles _subClassingWindow.CallBackProc
        AdjustSize()
    End Sub

    Private Sub AdjustSize()
        Dim tRect As RECT
        If GetClientRect(_parentHandle, tRect) <> 0 Then
            Me.Size = New Size(tRect.Right - tRect.Left, tRect.Bottom - tRect.Top)
        End If
    End Sub

    'This function Previews a keyboard message.You can read more about this function from here
    'https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.control.processkeypreview?view=netcore-3.1
    Protected Overrides Function ProcessKeyPreview(ByRef m As System.Windows.Forms.Message) As Boolean
        Const WM_KEYDOWN As Integer = &H500
        Dim result As Boolean = False
        Dim pressedKey As Keys
        Dim hostedUserControl As UserControl
        Dim activeButton As Button
        hostedUserControl = DirectCast(Me.Controls.Item(0), UserControl)
        If m.Msg = WM_KEYDOWN Then
            pressedKey = CType(m.WParam, Keys)
            Select Case pressedKey
                Case Keys.Tab
                    If Control.ModifierKeys = Keys.None Then ' Tab
                        Me.SelectNextControl(hostedUserControl.ActiveControl, True, True, True, True)
                        result = True
                    ElseIf Control.ModifierKeys = Keys.Shift Then ' Shift + Tab
                        Me.SelectNextControl(hostedUserControl.ActiveControl, False, True, True, True)
                        result = True
                    End If
                Case Keys.Return
                    If TypeOf hostedUserControl.ActiveControl Is Button Then
                        activeButton = DirectCast(hostedUserControl.ActiveControl, Button)
                        activeButton.PerformClick()
                    End If
            End Select
        End If
        If result = False Then
            result = MyBase.ProcessKeyPreview(m)
        End If
        Return result
    End Function

End Class
