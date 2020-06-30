## Hitting The End Point

### About

* In this module the **`Sheets API`** is called, which will give the information about the support type of API used in the macros of `.xlsm` in  `.txt` format.
  
* Sheets API is not public, so it needs authorization when a user hits it. So `getAuthorizationToken(userClientId,userClientSecretId)` function is used to get the "Bearer Authorization Token". This function needs "clientId" and "clientSecretId" as an argument and it can be obtained from [this](https://github.com/googleinterns/vba-migrator-addin/blob/Upload-File/README.md) instructions.

 * `parseTheFile()` function parses the downloaded file after hitting the endpoint and provides the list of data needed.
