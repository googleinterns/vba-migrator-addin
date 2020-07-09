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

Imports Microsoft.Vbe.Interop
Imports System.Windows.Forms
''' <summary>
''' This show a user form when add-in button is clicked.
''' </summary>
Friend Class userForm
    Private _VBE As VBE
    Private _getClientId As String = Nothing
    Private _getClientSecretId As String = Nothing
    Dim lines As List(Of String) = Nothing
    Dim fileId As String = Nothing
    Private _AddIn As AddIn
    Private _pathToCopyFile As String = Nothing

    'Initialize the user form.
    Friend Sub Initialize(ByRef vbe As VBE, ByRef addIn As AddIn)
        _VBE = vbe
        _AddIn = addIn
        ClientIdInput.Clear()
        ClientSecretIdInput.Clear()
        pathInput.Clear()
        'To show the form dialog.
        Me.ShowDialog()
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
    Private Sub CreateToolWindow(ByVal toolWindowCaption As String, ByVal toolWindowGuid As String,
      ByVal toolWindowUserControl As UserControl)
        Dim userControlObject As Object = Nothing
        Dim userControlHost As UserControlHost
        Dim _SummaryWindow As Window
        Dim progId As String
        'Ensure that you use the same ProgId value used in the ProgId attribute of the UserControlHost 
        progId = "SheetsCompatibilityAddIn.UserControlHost"
        _SummaryWindow = _VBE.Windows.CreateToolWindow(_AddIn, progId, toolWindowCaption, toolWindowGuid, userControlObject)
        userControlHost = DirectCast(userControlObject, UserControlHost)
        _SummaryWindow.Visible = True
        userControlHost.AddUserControl(toolWindowUserControl)
    End Sub

    'Click event of submit button.
    Private Sub Submit_Click(sender As Object, e As EventArgs) Handles Submit.Click
        'Get the value in the text box of the form.
        _getClientId = ClientIdInput.Text
        _getClientSecretId = ClientSecretIdInput.Text
        _pathToCopyFile = pathInput.Text + "\GoogleDrive.xlsm"
        'If the client Id and client secret id is not empty. 
        If _getClientId IsNot Nothing And _getClientSecretId IsNot Nothing Then
            'Make a copy of the file active in excel.
            My.Computer.FileSystem.CopyFile(_VBE.ActiveVBProject.FileName, _pathToCopyFile, True)
            'Call the upload file function to upload the file to the drive.
            fileId = uploadFileToDrive.UploadFile(_pathToCopyFile, _getClientId, _getClientSecretId)
            'Delete the copy of file created above.
            My.Computer.FileSystem.DeleteFile(_pathToCopyFile)
            If fileId <> "" Then
               'Call the Sheets API to get the list of report.
               lines = hittingEndPoint.callSheetsAPI(fileId, _getClientId, _getClientSecretId, pathInput.Text)
            Else
                Exit Sub
            End If
            'If the data list count item Is Not zero Then all Api In the file Is Not supported.
            If lines.Count <> 0 Then
                Dim userControlObject As Object = Nothing
                Dim userControlToolWindow As UserControlToolWindow
                Try
                    'Create the Tool Window to show the result.
                    userControlToolWindow = New UserControlToolWindow()
                    'This("B9055551-73E5-4507-AB69-19FF25D00F2B") is unique id for tool window and it is explained in the CreateToolWindow() function.
                    CreateToolWindow("API Summary", "B9055551-73E5-4507-AB69-19FF25D00F2B", userControlToolWindow)
                    userControlToolWindow.Initialize(_VBE, lines)
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            'if the data list is empty and filed id is not empty then all API used in the file is supported.
            ElseIf fileId <> "" Then
                MessageBox.Show("Fully Compatible!!")
            End If
        End If
        'Close the form.
        Me.Close()
    End Sub

    'Closing the user form when the cancel button is clicked.
    Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub
End Class
