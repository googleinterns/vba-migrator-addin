Imports Google.Apis.Drive.v2
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Services
Imports System.Threading
Imports Google.Apis.Drive.v2.Data
''' <summary>
'''For this you need to have "google account" & "visual studio 2013 or later"
''' To import the above library firstly we have to install the "Google.Apis.Drive.v2" 
''' in the project using NuGet Package manager console "Install-Package Google.Apis.Drive.v2".
''' Then enable the google api for your google account to get your own "ClientId" & 
''' "ClientSecretId", using this link "https://console.developers.google.com/flows/enableapi?apiid=drive"
''' and you can follow this video to enable api "https://www.youtube.com/watch?v=xtqpWG5KDXY&t=1s"
''' To upload the file to Drive, "UploadFile("file path")" member function of this class
''' have to call by passing the path of the file to upload. 
''' </summary>
Public Class uploadFileToDrive
    Private Service As DriveService = New DriveService()
    'Create Drive API service.
    Private Sub CreateService()
        'Set the client id and clientSecret id you get after anabling the google drive api
        'for your google account.
        Dim ClientId =
            "601010958158-ri1h9bipsbkfjip0qjhcnatfhdupnn08.apps.googleusercontent.com"
        Dim ClientSecret = "eeDXXNKGky5C7kX0SFwmsGXZ"
        'Set the application name whatever you have entered during enabling the drive api
        Dim MyUserCredential As UserCredential =
        GoogleWebAuthorizationBroker.AuthorizeAsync(New ClientSecrets() With
            {.ClientId = ClientId, .ClientSecret = ClientSecret},
            {DriveService.Scope.Drive}, "user", CancellationToken.None).Result
        Service = New DriveService(New BaseClientService.Initializer() With {.HttpClientInitializer = MyUserCredential,
            .ApplicationName = "Google Sheets Compatibility Report"})
    End Sub


        'Fuction to upload the file to drive using google drive api version-2.
    Public Function UploadFile(FilePath As String) As String
        'Set the application name whatever you have entered during enabling the drive api.
        If Service.ApplicationName <> "Google Sheets Compatibility Report" Then CreateService()
        'Define parameters of request.
        Dim TheFile As Google.Apis.Drive.v2.Data.File =
            New Google.Apis.Drive.v2.Data.File()
        'Write the name of the file by which you want to save  
        TheFile.Title = "internDemo2"
        'Write whatever you want to write as description of your file
        TheFile.Description = "A test document"
        TheFile.MimeType = "application/vnd.ms-excel.sheet.macroEnabled.12"

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
            MsgBox("Something went wrong!! I am null body, So i can't able to upload the project to the drive")
            Return ""
        Else
            MsgBox("Upload Finished")
            Return file.Id
        End If
    End Function
End Class
