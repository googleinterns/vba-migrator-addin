Imports System.Windows.Forms
Imports System.Net
Imports System.IO
Imports System.Text
Module hittingEndPoint
    'For getAuthorizationToken() function see this:
    'https://www.example-code.com/vbnet/hmrc_oauth2_access_token.asp 
    'only need to change "scope" for which auth token needed
    'and "clientId" and "clientSecretId".
    Private Function getAuthorizationToken() As String
        Dim glob As New Chilkat.Global
        Dim success As Boolean = glob.UnlockBundle("Anything for 30-day trial")
        If (success <> True) Then
            MessageBox.Show(glob.LastErrorText)
            Return ""
        End If
        Dim oauth2 As New Chilkat.OAuth2
        oauth2.ListenPort = 55568
        oauth2.AuthorizationEndpoint = "https://accounts.google.com/o/oauth2/auth"
        oauth2.TokenEndpoint = "https://oauth2.googleapis.com/token"
        ' Replace these with actual values.
        oauth2.ClientId = "601010958158-ri1h9bipsbkfjip0qjhcnatfhdupnn08.apps.googleusercontent.com"
        oauth2.ClientSecret = "eeDXXNKGky5C7kX0SFwmsGXZ"
        oauth2.CodeChallenge = True
        oauth2.CodeChallengeMethod = "S256"
        oauth2.Scope = "https://www.googleapis.com/auth/spreadsheets"
        Dim url As String = oauth2.StartAuth()
        If (oauth2.LastMethodSuccess <> True) Then
            MessageBox.Show(oauth2.LastErrorText)
            Return ""
        End If
        System.Diagnostics.Process.Start(url)
        Dim numMsWaited As Integer = 0
        While (numMsWaited < 30000) And (oauth2.AuthFlowState < 3)
            oauth2.SleepMs(100)
            numMsWaited = numMsWaited + 100
        End While

        If (oauth2.AuthFlowState < 3) Then
            oauth2.Cancel()
            MessageBox.Show("No response from the browser!")
            Return ""
        End If
        If (oauth2.AuthFlowState = 5) Then
            MessageBox.Show("OAuth2 failed to complete.")
            MessageBox.Show(oauth2.FailureInfo)
            Return ""
        End If
        If (oauth2.AuthFlowState = 4) Then
            MessageBox.Show("OAuth2 authorization was denied.")
            MessageBox.Show(oauth2.AccessTokenResponse)
            Return ""
        End If
        If (oauth2.AuthFlowState <> 3) Then
            MessageBox.Show("Unexpected AuthFlowState:" & oauth2.AuthFlowState)
            Return ""
        End If
        Return oauth2.AccessToken
    End Function

    Public Function callSheetsAPI(ByRef fileId As String) As List(Of String)
        Dim lines As List(Of String) = Nothing
        'Get Authorize Token
        Dim _bearerToken As String = getAuthorizationToken()
        'If "_bearerToken" then get out of the function no need to request.
        If _bearerToken <> "" Then
            Dim myUri As New Uri("https://docs.google.com/spreadsheets/vbaprocessfile?fid=" + fileId)
            Dim myWebRequest = System.Net.HttpWebRequest.Create(myUri)
            Dim myHttpWebRequest = CType(myWebRequest, System.Net.HttpWebRequest)
            myHttpWebRequest.Method = "GET"
            myHttpWebRequest.PreAuthenticate = True
            myHttpWebRequest.Headers.Add("Authorization", "Bearer " & _bearerToken)
            Dim myWebResponse As HttpWebResponse = myWebRequest.GetResponse()
            'Collect response after request.
            Dim responseStream As Stream = myWebResponse.GetResponseStream()
            If responseStream Is Nothing Then
                MessageBox.Show("Api didn't respond.")
                Return lines
            End If
            'Create file to store data from response stream.
            Dim Json As New FileStream("D:\Json.txt", FileMode.Create)
            'Declare of local variable.
            Dim read As Byte() = New Byte(255) {}
            Dim count As Integer = responseStream.Read(read, 0, read.Length)
            'Write in file.
            While count > 0
                Json.Write(read, 0, count)
                count = responseStream.Read(read, 0, read.Length)
            End While
            'Close everything
            Json.Close()
            responseStream.Close()
            myWebResponse.Close()
            'Call the function parseTheFile() to get the information needed from the downloaded file.
            lines = parseTheFile()
        End If
        Return lines
    End Function

    Private Function parseTheFile() As List(Of String)
        'Declare location variable need to store and read the file.
        Dim lines As New List(Of String)
        Dim FileData As String = String.Empty
        Dim SReader As IO.StreamReader = Nothing
        Try
            'Read the file which is download after hitting the Endpoint.
            SReader = New IO.StreamReader("D:\Json.txt")
            'Read all data of the file and put it into a string.
            Do Until SReader.EndOfStream
                FileData &= SReader.ReadLine()
            Loop
            'Catch exception while reading the file
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
            Return lines
        Finally
            'Close Stream Reader.
            If SReader IsNot Nothing Then
                SReader.Close()
                SReader.Dispose()
            End If
        End Try
        'Define local variable need to parse the file. 
        Dim index As Integer
        Dim tempStr As String = Nothing
        For index = 0 To FileData.Length - 1 Step +1
            If FileData.Chars(index) = """" Then
                index += 1
                'Take all string present in double quote in file for example: "API_LINE/chart.visible/blad4 3".
                Do While FileData.Chars(index) <> """"
                    tempStr += FileData.Chars(index)
                    index += 1
                Loop
                'Define local variable array of string.
                Dim arrayStr() As String
                'Split the above string on the basis of '/'.
                arrayStr = Split(tempStr, "/")
                'Check the condition to get the only the information needed. 
                If arrayStr.Length >= 2 Then
                    If arrayStr(0) = "API_LINE" Then
                        'Last word in API_LINE contain module name and line number so need to split it.
                        Dim myModuleInfo() As String = Split(arrayStr(2), " ")
                        Dim putValue As String
                        putValue = arrayStr(1) + "," + myModuleInfo(0) + "," + myModuleInfo(1)
                        'Add the information in list.
                        lines.Add(putValue)
                    ElseIf arrayStr(0) = "SUPPORTED" Then
                        'Iterate through the list to delete the api which is supported.
                        For i As Integer = lines.Count - 1 To 0 Step -1
                            Dim getWord() As String
                            getWord = Split(lines.Item(i), ",")
                            If getWord(0) = arrayStr(1) Then
                                lines.RemoveAt(i)
                            End If
                        Next i
                    ElseIf arrayStr(0) = "POSSIBLY_SUPPORTED" Or arrayStr(0) = "MANUAL_SUPPORT" Or
                        arrayStr(0) = "DEFAULT_SUPPORTTYPE" Or arrayStr(0) = "ALMOST_SUPPORTED" Or
                        arrayStr(0) = "PARTIALLY_SUPPORTED" Or arrayStr(0) = "CAN_SUPPORT_LATER" Or
                        arrayStr(0) = "NOT_SUPPORTED" Or arrayStr(0) = "UNKNOWN" Or
                        arrayStr(0) = "SUPPORTED_AND_IMPLEMENTED" Or arrayStr(0) = "COM_API" Or
                        arrayStr(0) = "INVALIDATED" Then
                        'Iterate through the list and find that api whose support is this and append it into it.
                        For i As Integer = 0 To lines.Count - 1 Step +1
                            Dim getWord() As String
                            getWord = Split(lines.Item(i), ",")
                            If getWord(0) = arrayStr(1) Then
                                Dim putValue As String = lines.Item(i) + "," + arrayStr(0)
                                lines(i) = putValue
                            End If
                        Next i
                    End If
                End If
                'Reinitialize the temporary string to make it empty for next turn.
                tempStr = Nothing
            End If
        Next index
        'Return data to show in grid, if it is empty then one message box will be seen with proper message.
        Return lines
    End Function

End Module
