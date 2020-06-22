# Project Title
### vba-migrator-addin
VBA Editor add-in to show compatibility with Google Sheets.

## About
For now this Add-in will give you a button in your VBA editor.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites
```
Having Window OS
  * Edition   Windows 10 Enterprise
  * Version   1803
  * Windows Registry Editor version 5.00

Microsoft Excel 2016 (16.0.5017.1000) MSO (16.0.5017.1000) 64bit

```

## Deployment
* See the references first link to create a project in visual studio and built it, Instead of using their code in
connect class use the code provided here.
* After building the project you got .dll file. After that main part is make that add-in available in add-in manager of VBA Editor. So for this you need to Register .dll file in your window registry. Steps are Provided below:-

Here is the registry .key script to register the Addin, note you will need to change some of the settings in order to register it properly. So open your favourite text editor copy-paste the following text in it and save it as .reg file, don't forget to change the path of the your .dll file here "CodeBase"="file:///PathToAssembly".

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

After running the .reg file you can see this changes in your window Registry under "HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64 or HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins" corresponding to 32bit or 64bit VBA editor "MyVBAAddin.Connect" is present. For any further issues refer [this](https://stackoverflow.com/questions/1942626/build-add-in-for-vba-ide-using-vb-net).

![alt text](/images/Registry.png)


After that when you open your VBA Editor in Add-In Manager this addin is present. Then when you check the checkBox called "Loaded/Unloaded" and click "OK", The button get created in the "Menu Bar" of VBA Editor.  

![alt text](/images/Add-inManager.png)


![alt text](/images/button.png)
## Built With

* [Visual Studio](https://visualstudio.microsoft.com/vs/) - Vb.net Class Library Project
* [Registry](https://stackoverflow.com/questions/1942626/build-add-in-for-vba-ide-using-vb-net) - Dependency Management

## References

* [Add-in](https://www.mztools.com/articles/2012/MZ2012013.aspx) - How to make add-in for VBA Editor
* [Button](https://www.mztools.com/articles/2012/MZ2012015.aspx) - How to make different types of button in VBA Editor
* [Tool Window](https://www.mztools.com/articles/2012/MZ2012017.aspx) - How to make Tool window in VBA Editor

