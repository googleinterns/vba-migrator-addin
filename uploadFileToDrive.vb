Imports Google.Apis.Drive.v2
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Services
Imports System.Threading
Imports Google.Apis.Drive.v2.Data

'' <summary>
'''To upload the file to Drive, "UploadFile("file path","userClientId","userClientSecretId")" member function of this class
''' have to call by passing the path of the file to upload and client credentials. Follow the instructions in README.md to get the credentials. 
''' </summary>
Public Class uploadFileToDrive
    Private Service As DriveService = New DriveService()
    'Create Drive API service.
    Private Sub CreateService(ByRef userClientId As String , ByRef userClientSecretId As String)
        'Setting your "clientId" and "clientSecretId".
        Dim ClientId = userClientId
        Dim ClientSecret = userClientSecretId
        'set application name you entered when you enabled the Drive API.
        Dim MyUserCredential As UserCredential = GoogleWebAuthorizationBroker.AuthorizeAsync(New ClientSecrets() With
            {.ClientId = ClientId, .ClientSecret = ClientSecret},{DriveService.Scope.Drive}, "user", CancellationToken.None).Result
        Service = New DriveService(New BaseClientService.Initializer() With {.HttpClientInitializer = MyUserCredential,
                 .ApplicationName = "Google Sheets Compatibility Report"})
    End Sub

    'Fuction to upload the file to drive using google drive api version-2.
    Public Function UploadFile(ByRef FilePath As String, ByRef userClientId As String, ByRef userClientSecretId As String) As String
        'set application name you entered when you enabled the Drive API.
        If Service.ApplicationName <> "Google Sheets Compatibility Report" Then CreateService(userClientId,userClientSecretId)
        'Define parameters of request.
        Const title As String = "internDemo" 
        Const description As String = "A test document"
        Const mimeType As String = "application/vnd.ms-excel.sheet.macroEnabled.12"
        Dim TheFile As Google.Apis.Drive.v2.Data.File = New Google.Apis.Drive.v2.Data.File() 
        TheFile.Title = title
        TheFile.Description = description
        TheFile.MimeType = mimeType
        'Reading the file to upload
        Dim ByteArray As Byte() = System.IO.File.ReadAllBytes(FilePath)
        Dim Stream As New System.IO.MemoryStream(ByteArray)
        'Creating upload Request
        Dim UploadRequest As FilesResource.InsertMediaUpload = Service.Files.Insert(TheFile, Stream, TheFile.MimeType)
        UploadRequest.SupportsTeamDrives = True
        UploadRequest.Fields = "id"
        'Call to upload
        UploadRequest.Upload()
        'Retrieving the response body to check our file is successfully uploaded or not
        Dim file As File = UploadRequest.ResponseBody
        If file Is Nothing Then
                MsgBox("Error: Upload request-response body is null.")
            Return ""
        Else
            Return file.Id
        End If
    End Function
End Class
