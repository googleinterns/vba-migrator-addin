## Prerequisites
```
  1. Having Windows OS
  2. Microsoft Excel
```
> Tested on :
```
Windows :
  Edition   Windows 10 Enterprise
  Version   1803
  Windows Registry Editor Version 5.00
Microsoft Excel 2016 (16.0.5017.1000) MSO (16.0.5017.1000) 64bit
```
## Steps to install
* Download the `.DLL` and `.reg` from here and open the `.reg` file in text the editor.
* Given below is the list of lines, needed to be changed in the `.reg` file:
    * `"CodeBase"="file:///`~~`C:\Users\UserName\source\repos\SheetsCompatibilityAddIn\bin\Debug\`~~`SheetsCompatibilityAddIn.DLL"` - Path of ".dll" file.
    * To change the Runtime version(v4.0.30319) go to this path C:\Windows\Microsoft.NET\Framework64 or Framework(corresponding to 64bit and 32bit VBA Editor)  in the window and use the latest version available in your system.
    * It is also mentioned in the .reg file at the top as a comment after editing, delete that comment.
    * If the VBA editor is 32bit then the branch (“HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64\SheetsCompatibilityAddIn.Connect”) should changed to this (“HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins\SheetsCompatibilityAddIn.Connect”).
    * Change the version of your registry, for this refer to references.
* Double click the '.reg' file, it will ask for your consent to make changes in your registry, after proceeding further, required changes will be made in your registry.
* After running the '.reg' file these changes will be incorporated in the window registry under "HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64" OR "HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins" corresponding to 64bit or 32bit VBA editor, “SheetsCompatibilityAddIn.Connect” will be present there.

![alt text](/images/Registry.jpg)

* And after that, when the VBA editor is opened, this add-in will be available in the Add-in Manager section.

![alt text](/images/Add-inManager.jpg)

* If it prompts that the Add-In is missing in the Add-in Manager section. Then open your command prompt with administrator right and run this command:
 `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe /codebase C:\Users\UserName\source\repos\SheetsCompatibilityAddIn\bin\Debug\SheetsCompatibilityAddIn.dll` , change the Framework64/Framework according to your VBA editor 64/32bit, also change the latest version(v4.0.30319) available in your system and change your .dll file path.
* If it gets registered successfully then the "Add-In is missing" issue gets resolved.
* For any further issues refer to [this](https://stackoverflow.com/questions/1942626/build-add-in-for-vba-ide-using-vb-net).
* Then checking the Checkbox named "Loaded/Unloaded" and clicking "OK" will create a button named “SheetsCompatibility” in the "Menu Bar" of VBA editor.

![alt text](/images/button.png)

* Now to generate report click this button.
## References
* To know the VBA Editor is [64bit/32bit](https://help.xlstat.com/s/article/get-your-excel-version?language=en_US#:~:text=Start%20by%20clicking%20on%20the,top%20left%20corner%20of%20Excel.&text=Click%20on%20Account%2C%20on%20the,the%20dialog%20box%20that%20appears.).
* About [Windows](askvg.com/how-to-check-which-windows-version-is-installed-in-my-computer/).
* To know the version of [Registry](https://www.arclab.com/en/kb/cppmfc/get-windows-version-from-registry.html).
