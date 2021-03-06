### vba-migrator-addin
VBA Editor add-in to show compatibility with Google Sheets.

## Background

* This is a class library project implemented in the VB dot NET 4.8 framework.
* Google drive API is used to upload the file to drive. 
* Sheets API is used to generate the report of the support type of API used in the macros of the `'.xlsm'` file. 

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites

```
1. Having Windows OS
2. Microsoft Excel
```

> Tested on :
```
Windows :
  * Edition   Windows 10 Enterprise
  * Version   1803
  * Windows Registry Editor version 5.00

Microsoft Excel 2016 (16.0.5017.1000) MSO (16.0.5017.1000) 64bit
```

## Deployment

* Follow this [video](https://www.youtube.com/watch?v=y81Aq4bebZU) or [article](https://www.mztools.com/articles/2012/MZ2012013.aspx) to make a project in `visual studio` and use the code provided in the `Connect.vb` file and build the project which will give a  **.dll** file. Then open your favorite text editor and copy-paste the sample script given below and save it as `Registry.reg` extension. 
* This is a list of lines that need to be changed in the script:
     <!-- TODO : change the name of the project-->
     * Wherever it is written ~~MyVBAAddin.Connect~~ change it with `{your programId(project-name.class-name)}` also need to change in code but if you have created the project of this name, there is no need to change.
     * "Assembly"="~~MyVBAAddin~~, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null" - Your assembly name. i.e Your project name
     * "CodeBase"="~~file:///C:/Users/userName/source/repos/MyVBAAddin/bin/Debug/MyVBAAddin.DLL~~" - Path of your **".dll"** file.

_If VBA editor is 64bit then branch should also need to change i.e, **HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\``Addins64`\MyVBAAddin.Connect**._

```
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64\MyVBAAddin.Connect]
@="MyVBAAddin.Connect"
"Description"="Checking Compatibility with Google Sheets"
"FriendlyName"="SheetsCompatibilityAdd-in"
"LoadBehavior"=dword:00000002

[HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64\MyVBAAddin.Connect\Implemented Categories]

[HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64\MyVBAAddin.Connect\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}]

[HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64\MyVBAAddin.Connect\InprocServer32]
@="mscoree.dll"
"ThreadingModel"="Both"
"Class"="MyVBAAddin.Connect"
"Assembly"="MyVBAAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
"RuntimeVersion"="v4.0.30319"
"CodeBase"="file:///C:/Users/userName/source/repos/MyVBAAddin/bin/Debug/MyVBAAddin.DLL"  

[HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64\MyVBAAddin.Connect\InprocServer32\1.0.0.0]
"Class"="MyVBAAddin.Connect"
"Assembly"="MyVBAAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
"RuntimeVersion"="v4.0.30319"
"CodeBase"="file:///C:/Users/userName/source/repos/MyVBAAddin/bin/Debug/MyVBAAddin.DLL"

```

_If VBA editor is 32bit then branch should also need to change i.e, **HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\``Addins`\MyVBAAddin.Connect**._

```
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins\MyVBAAddin.Connect]
@="MyVBAAddin.Connect"
"Description"="Checking Compatibility with Google Sheets"
"FriendlyName"="SheetsCompatibilityAdd-in"
"LoadBehavior"=dword:00000002

[HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins\MyVBAAddin.Connect\Implemented Categories]

[HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins\MyVBAAddin.Connect\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}]

[HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins\MyVBAAddin.Connect\InprocServer32]
@="mscoree.dll"
"ThreadingModel"="Both"
"Class"="MyVBAAddin.Connect"
"Assembly"="MyVBAAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
"RuntimeVersion"="v4.0.30319"
"CodeBase"="file:///C:/Users/userName/source/repos/MyVBAAddin/bin/Debug/MyVBAAddin.DLL"  

[HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins\MyVBAAddin.Connect\InprocServer32\1.0.0.0]
"Class"="MyVBAAddin.Connect"
"Assembly"="MyVBAAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
"RuntimeVersion"="v4.0.30319"
"CodeBase"="file:///C:/Users/userName/source/repos/MyVBAAddin/bin/Debug/MyVBAAddin.DLL"
```

* Double click the `.reg` file, it will ask for your consent to make changes in your registry, after proceeding further required changes will be made in your registry.

* After running the `.reg` file these changes will be incorporated in the window registry under _"HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64"_ **OR** _"HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins"_ corresponding to 64bit or 32bit VBA editor `{MyVBAAddin.Connect(Your programId)}` will be present there. For any further issues refer [this](https://stackoverflow.com/questions/1942626/build-add-in-for-vba-ide-using-vb-net).

![alt text](/images/Registry.jpg)


* And after that, when the VBA editor is opened, this add-in will be available in the `Add-in Manager` section. 

![alt text](/images/Add-inManager.jpg)

* Then checking the Checkbox named **"Loaded/Unloaded"** and clicking **"OK"** will create a button named `SheetsCompatibility` in the **"Menu Bar"** of VBA editor.

![alt text](/images/button.png)

#### Upload file to drive
  * The `uploadFileToDrive.vb` uses the google drive API to upload the file to drive.
  
  * To use the google drive API one should have a google account.
  
  * To import the libraries in it, the user needs to install  "Google.Apis.Drive.v2" in the project using the NuGet Package manager console using the command `Install-Package Google.Apis.Drive.v2`.
  
  * To enable the drive API for their account, the user needs to refer to this [video](https://www.youtube.com/watch?v=xtqpWG5KDXY&t=1s") which will guide them to enable the drive API for their account. This is the drive API [Console](https://pantheon.corp.google.com/flows/enableapi?apiid=drive&pli=1&debugUI=DEVELOPERS), where the user will get the `clientId` and `clientSecretId` after following [this](https://www.youtube.com/watch?v=xtqpWG5KDXY&t=1s") video. 
  
  * Now, the user can call the function `UploadFile(filePath,userClientId,userClientSecretId)` to upload the file to the drive, by providing the file path and user credentials as an argument.

#### Hitting The End Point

  * In this module the **`Sheets API`** is called, which will give the information about the support type of API used in the macros of `.xlsm` in  `.txt` format.
  
  * Sheets API is not public, so it needs authorization when a user hits it. So `getAuthorizationToken(userClientId,userClientSecretId)` function is used to get the "Bearer Authorization Token". This function needs "clientId" and "clientSecretId" as an argument and it can be obtained from [this](https://github.com/googleinterns/vba-migrator-addin/blob/Upload-File/README.md) instructions.

  * `parseTheFile()` function parses the downloaded file after hitting the endpoint and provides the list of data needed.

#### ToolWindow for VBA editor

  * UserControlToolWindow.vb describes what happens when the window is initialized and when some action is performed in windows.

  * User-control tool-window design can be created manually like dragging-dropping and the property can be set in the property window, corresponding to that system generates the code in file UserControlHost.Designer.vb and UserControlToolWindow.Designer.vb.

  * For file UserControlHost.vb, refer to [this](https://www.mztools.com/articles/2012/MZ2012017.aspx).

  * To see the working of this add-in and how does the window look, refer these two links:- `.xlsm` file is fully [not-compatible](https://drive.google.com/file/d/11L-v_ym66W2XbvsDtJTX2QhFjnbWvidg/view?usp=sharing) and [compatible](https://drive.google.com/file/d/1cyYpA5mzLUfSRR8cKB-3H2nXAcutF0ZY/view?usp=sharing) with google sheets.

## Built With

* [Visual Studio](https://visualstudio.microsoft.com/vs/) - Vb.net Class Library Project
* [Registry](https://stackoverflow.com/questions/1942626/build-add-in-for-vba-ide-using-vb-net) - Dependency Management

## References

* [Add-in](https://www.mztools.com/articles/2012/MZ2012013.aspx) - How to make add-in for VBA Editor
* [Add-in](https://www.youtube.com/watch?v=y81Aq4bebZU) - YouTube link to make add-in project in Visual Studio
* [Button](https://www.mztools.com/articles/2012/MZ2012015.aspx) - How to make different types of button in VBA Editor
* [Tool Window](https://www.mztools.com/articles/2012/MZ2012017.aspx) - How to make Tool window in VBA Editor


