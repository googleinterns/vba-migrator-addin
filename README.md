## Hit-EndPoint & Parse-The-File

This module contain three function:

* getAuthorizationToken()
```
As this Api is not public so it needs some authorization when we call it. So this fuction provide us the 
"Bearer Authorization Token". In this function we only need to change "clientId" and "clientSecretId"
according to from which account we need to call it. And we can use the same id which we got for 
"Google-Drive-Api" right here.
```

* callSheetsAPI(fileId)
```
This is the main function in this module that need two remaining fuctions. It is the only function called from
the main "connect" class when the button is clicked. This function need filed id as argument and return list
of data that will be shown in the grid. When this function is called, It first call the 
"getAuthorizationToken()" function to get authorization token then hit the end-point and download the result
after that call the "parseTheFile()" function then return the result returned by this function.
```

* parseTheFile()
```
This fucntion is called at the last when end-point is hitted. It processes the downloaded file and return
the result to the function. 
```
So when all the API used in .xlsm file is "SUPPORTED" type then the return list of data is empty so from
here we can decide that the .xlsm file is fully compatible with "GOOGLE SHEETS" and we only need show a 
message-box pop-up not the window.
