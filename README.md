## Upload file to drive
### Background
* Vb dot net 4.8 frameworks.
* Google Drive API is used to upload the file to drive.
### Deployment
* To use the google drive API one should have a google account.
* To import the libraries, the user needs to install  "Google.Apis.Drive.v2" in the project using the NuGet Package manager console using the command `Install-Package Google.Apis.Drive.v2`.
* To enable the drive API for their account, the user needs to refer to this [video](https://www.youtube.com/watch?v=xtqpWG5KDXY&t=1s") which will guide them to enable the drive API for their account. This is the drive API [Console](https://pantheon.corp.google.com/flows/enableapi?apiid=drive&pli=1&debugUI=DEVELOPERS), where the user will get the `clientId` and `clientSecretId` after following [this](https://www.youtube.com/watch?v=xtqpWG5KDXY&t=1s") video. 
* Now, the user can call the function `UploadFile(filePath,userClientId,userClientSecretId)` to upload the file to the drive, by providing the file path and user credentials as an argument.
