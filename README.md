## Upload file to drive
### Background
* Vb dot net 4.8 frameworks.
* Google Drive API is used to upload the file to drive.
### Deployment
* To use the google drive API one should have the google account.
* To import the libraries, you need to install  "Google.Apis.Drive.v2" in the project using the NuGet Package manager console using the command `Install-Package Google.Apis.Drive.v2`.
* To enable the drive API for your account, please refer to this [video](https://www.youtube.com/watch?v=xtqpWG5KDXY&t=1s") which will guide you to enable drive API for your account. This is the drive API [Console](https://pantheon.corp.google.com/flows/enableapi?apiid=drive&pli=1&debugUI=DEVELOPERS), where you will get the `clientId` and `clientSecretId` after following [this](https://www.youtube.com/watch?v=xtqpWG5KDXY&t=1s") video. 
* Change the `clientId` and `clientSecretId` in the code and save it.
* Now, you can call the function `UploadFile(filePath)` to upload the file to the drive providing file path as an argument.
