# Project Title

### vba-migrator-addin

VBA Editor add-in to show compatibility with Google Sheets.

## Background

This is a class library project implemented in the VB dot NET 4.8 framework. Drive API, as well as sheets API, is also used in this project.

## About

Macros in excel sheets are written in **VBA**, while Google sheets use **Apps Script**. This hinders the support of some APIs of VBA in Apps Script. This Add-in enables a user to open an excel `.xlsm` file with Google Sheets with the loss of fewer features. The add-in provides the list of APIs that are not supported by Google Sheets in a window through a data grid. This data grid enriches the user with pieces of information like; the name of the module and the particular line which contains the incompatible API. In this window, the user can simply click the particular row in the popped up data grid, and obtain control over the module, and then it can be easily modified.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites

```
1. Having Window OS
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

* Clone this project and open it in a visual Studio having .net installed in it. After that open the **UploadToDrive.vb** and **HittingEndPoint.vb** file and change the `clientId` and `clientSecretId` in it, then build the project. **.dll** got updated after building the project according to modification done. Then Open the **registry.reg** file in your favorite text editor and just change the path of **.dll** file which will be in project folder `bin-->Debug-->MyVBAAddin.dll`, copy this file path and paste it in the registry in the line `"CodeBase"="file:///PathToAssembly"`. A sample script is given below but use the `Registry.reg` file that is given in the project. One more point to consider is:-

> If VBA editor is 64bit then branch should also need to change i.e, **HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\``Addins64`\MyVBAAddin.Connect**.

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
"CodeBase"="file:///PathToAssembly"  
```

> If VBA editor is 32bit then branch should also need to change i.e, **HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\``Addins`\MyVBAAddin.Connect**.

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
"CodeBase"="file:///PathToAssembly"  
```

* Double click the `.reg` file, it will ask for your consent to make changes in your registry, after proceeding further required changes will be made in your registry.

* After running the `.reg` file these changes will be incorporated in the window registry under "HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64 or HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins" corresponding to 32bit or 64bit VBA editor `MyVBAAddin.Connect` will present. For any further issues refer [this](https://stackoverflow.com/questions/1942626/build-add-in-for-vba-ide-using-vb-net).

![Registry](/images/Registry.jpg)

* And after that, when the VBA editor is opened, this add-in will be available in the `Add-in Manager` section.

![Add-in Manager](/images/Add-inManager.jpg)

* Then checking the Checkbox named **"Loaded/Unloaded"** and clicking **"OK"** will create a button named `SheetsCompatibility` in the "Menu Bar" of VBA editor.  

![Button](/images/button.png)

* Event after 'button is clicked `fileUploadToDrive-->getAuthtoken-->hittingEndPoint-->windowInitialization` if the api is used in file is not supported, OR `fileUploadToDrive-->getAuthtoken-->hittingEndPoint-->messageBoxPop-Up` if all the api is supported. To see the live demo check these videos [Not-Supported](https://drive.google.com/file/d/11L-v_ym66W2XbvsDtJTX2QhFjnbWvidg/view) and [Supported](https://drive.google.com/file/d/1cyYpA5mzLUfSRR8cKB-3H2nXAcutF0ZY/view).

* Example images.

  Window
![window](/images/Window.jpg)

  Control
![control](/images/control.jpg)
## Built With

* [Visual Studio](https://visualstudio.microsoft.com/vs/) - Vb.net Class Library Project
* [Registry](https://stackoverflow.com/questions/1942626/build-add-in-for-vba-ide-using-vb-net) - Dependency Management

## References

* [Add-in](https://www.mztools.com/articles/2012/MZ2012013.aspx) - How to make add-in for VBA Editor
* [Button](https://www.mztools.com/articles/2012/MZ2012015.aspx) - How to make different types of button in VBA Editor
* [Tool Window](https://www.mztools.com/articles/2012/MZ2012017.aspx) - How to make Tool window in VBA Editor
* [Drive API](https://www.youtube.com/watch?v=xtqpWG5KDXY&t=1s") - Guide to use API console to enable `Drive API` for an account.
* [API console](https://pantheon.corp.google.com/flows/enableapi?apiid=drive&pli=1&debugUI=DEVELOPERS) - To generate `clientId` and `clientSecretId`
* [Auth token](https://www.example-code.com/vbnet/box_oauth2_json_web_token.asp) - To get authorization token of an account by which we hit the end-point 
