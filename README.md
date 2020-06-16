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
```

## Deployment

Here is the registry .key script to register the Addin, note you will need to change some of the settings in order to register it properly. So open your favourite text editor copy-paste the following text in it and save it as .reg file, don't forget to change the path of the your .dll file here "CodeBase"="file:///PathToAssembly". One more important thing is if your VBA editor is 32 bit then the registry should be made in Addins so if it is the case change Addins64 ~ Addins.  
```
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64\MyVBAAddin.Connect]
@="MyVBAAddin.Connect"
"Description"="My VBA Add-in"
"FriendlyName"="My VBA Add-in"
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

After running the .reg file you can see this changes in your window Registry under "HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64 or HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins" corresponding to 32bit or 64bit VBA editor "MyVBAAddin.Connect" is present.

![alt text](/images/Registry.png)


After that when you open your VBA Editor in Add-In Manager this addin is present. Then when you check the checkBox called "Loaded/Unloaded" and click "OK", The button get created in the "Menu Bar" of VBA Editor.  

![alt text](/images/Add-inManager.png)


![alt text](/images/button.png)
## Built With

* [Visual Studio](https://visualstudio.microsoft.com/vs/) - Vb.net Class Library Project
* [Registry](https://stackoverflow.com/questions/1942626/build-add-in-for-vba-ide-using-vb-net) - Dependency Management

