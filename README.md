# Project Title
### vba-migrator-addin
VBA Editor add-in to show compatibility with Google Sheets.

## Background
This is a class library project implemented in the VB dot NET 4.8 framework. Drive API, as well as sheets API, is also used in this project.

## About
The feature provided by this add-in makes excel .xlsm file flexible to open with google sheets. As we know, macros in excel sheets are written in VBA, and in google sheets, it is written in Apps Script. Due to which some APIs of VBA is not supported in Apps Script. This add-in provided the list of APIs in a window in a data grid form that is not supported by google sheets, this list also contains information about the name of the module and the line number containing this APIs When the data grid row is clicked in this window it will pass the control to the module and that line number got selected one can easily modify it.

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
* Clone this project and open it in a visual Studio having .net installed in it. Then change the `clientId` and `clientSecretId` in **UploadToDrive.vb** file and in **HittingEndPoint.vb** file, then build this project. Then Open the **registry.reg** file in your favorite text editor just change the path of your .dll file which will be in your project folder `bin-->Debug-->MyVBAAddin.dll` copy this file path and paste it in the registry in this line `"CodeBase"="file:///PathToAssembly"`. One more point to consider is:-

If VBA editor is 64bit.
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
If VBA editor is 32bit.
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

> `Double click your .reg file it will take your consent to make changes in your registry, give them then the changes occurred in your registry`

After running the .reg file you can see these changes in your window Registry under "HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64 or HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins" corresponding to 32bit or 64bit VBA editor "MyVBAAddin.Connect" is present. For any further issues refer [this](https://stackoverflow.com/questions/1942626/build-add-in-for-vba-ide-using-vb-net).

![alt text](/images/Registry.jpg)


After that when you open your VBA Editor in Add-In Manager this addin is present. Then when you check the checkBox called "Loaded/Unloaded" and click "OK", The button gets created in the "Menu Bar" of VBA Editor.  

![alt text](/images/Add-inManager.jpg)


![alt text](/images/button.png)


![window](/images/Window.jpg)

![control](/images/control.jpg)
## Built With

* [Visual Studio](https://visualstudio.microsoft.com/vs/) - Vb.net Class Library Project
* [Registry](https://stackoverflow.com/questions/1942626/build-add-in-for-vba-ide-using-vb-net) - Dependency Management

## References

* [Add-in](https://www.mztools.com/articles/2012/MZ2012013.aspx) - How to make add-in for VBA Editor
* [Button](https://www.mztools.com/articles/2012/MZ2012015.aspx) - How to make different types of button in VBA Editor
* [Tool Window](https://www.mztools.com/articles/2012/MZ2012017.aspx) - How to make Tool window in VBA Editor

