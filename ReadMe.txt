The configuration file could be set

1- Key Locator (Default Value = “NO”)

If Key Locator is set to "NO" it will create the file extension from their legislative extension 
Example A751_IH.xml will be converted to A751.IH

If Key Locator is set to "Yes" it will only create files with locator extension.
Example A751_IH.xml will be converted to A751_IH.loc



2- Set Key "XmetalDirPath" to the xmetal(LEXA) installation Directory or any other location that you will have Utilities.dll, TableTool.dll,lxl.dll, and templates-system.xml. Default path will be "C:\GPO\XML2LOC".

NOTE: If the table can not be translated - check the registry key 

[HKEY_LOCAL_MACHINE\SOFTWARE\USGPO\TableTool]
"InstallPath"="C:\GPO\XML2LOC\" 
"LibPath"="C:\GPO\XML2LOC\"
"UserProfilePath"="%APPDATA%/GPO/TableTool"

InstallPath - must point to the TableTool.dll location in case if the user uses Xmetal DLLs to translate
OR to the local folder as shown above.

LibPath - must point to the TableTool.dll location in case if the user uses Xmetal DLLs to translate
OR to the local folder as shown above.

UserProfilePath must point exactly as shown above.

