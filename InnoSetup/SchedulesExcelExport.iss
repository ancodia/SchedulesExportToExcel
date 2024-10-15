#define Repository      "..\."
#define MyAppName       "Schedules Excel Export Revit Add-in"
#define MyAppVersion    "1.0"
#define MyAppPublisher  "ancodia"
#define MyAppURL        "https://github.com/ancodia/SchedulesExportToExcel"
#define MyAppExeName    "SchedulesExcelExport.Revit.Addin.Installer.exe"


#define RevitAppName      "SchedulesExcelExport"
#define RevitAddinFolder  "{sd}\ProgramData\Autodesk\Revit\Addins"
#define RevitFolder20     RevitAddinFolder+"\2020\"+RevitAppName
#define RevitAddin20      RevitAddinFolder+"\2020\"
#define RevitFolder21     RevitAddinFolder+"\2021\"+RevitAppName
#define RevitAddin21      RevitAddinFolder+"\2021\"
#define RevitFolder22     RevitAddinFolder+"\2022\"+RevitAppName
#define RevitAddin22      RevitAddinFolder+"\2022\"
#define RevitFolder23     RevitAddinFolder+"\2023\"+RevitAppName
#define RevitAddin23      RevitAddinFolder+"\2023\"
#define RevitFolder24     RevitAddinFolder+"\2024\"+RevitAppName
#define RevitAddin24      RevitAddinFolder+"\2024\"
#define RevitFolder25     RevitAddinFolder+"\2025\"+RevitAppName
#define RevitAddin25      RevitAddinFolder+"\2025\"

[Setup]
AppId=DAF7B9CA-9E11-41BA-9263-93FBCF68D74C
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={commonpf64}\{#MyAppName}
DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
DisableWelcomePage=no
OutputDir={#Repository}\output
OutputBaseFilename=SchedulesExcelExport.Revit.Addin.Installer
Compression=lzma
SolidCompression=yes
ChangesAssociations=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Components]
Name: revit20; Description: Addin for Autodesk Revit 2020;  Types: full
Name: revit21; Description: Addin for Autodesk Revit 2021;  Types: full
Name: revit22; Description: Addin for Autodesk Revit 2022;  Types: full
Name: revit23; Description: Addin for Autodesk Revit 2023;  Types: full
Name: revit24; Description: Addin for Autodesk Revit 2024;  Types: full
Name: revit25; Description: Addin for Autodesk Revit 2025;  Types: full

[Files]
; Include only the .dll files totally necessary for running within Revit
Source: "{#Repository}\output\Release-2020\EPPlus*.dll"; DestDir: "{#RevitFolder20}"; Flags: ignoreversion recursesubdirs; Components: revit20
Source: "{#Repository}\output\Release-2020\{#RevitAppName}.dll"; DestDir: "{#RevitFolder20}"; Flags: ignoreversion recursesubdirs; Components: revit20
Source: "{#Repository}\output\Release-2020\Microsoft.IO.RecyclableMemoryStream.dll"; DestDir: "{#RevitFolder20}"; Flags: ignoreversion; Components: revit20
Source: "{#Repository}\src\*.addin"; DestDir: "{#RevitAddin20}"; Flags: ignoreversion; Components: revit20

Source: "{#Repository}\output\Release-2021\EPPlus*.dll"; DestDir: "{#RevitFolder21}"; Flags: ignoreversion recursesubdirs; Components: revit21
Source: "{#Repository}\output\Release-2021\{#RevitAppName}.dll"; DestDir: "{#RevitFolder21}"; Flags: ignoreversion recursesubdirs; Components: revit21
Source: "{#Repository}\output\Release-2021\Microsoft.IO.RecyclableMemoryStream.dll"; DestDir: "{#RevitFolder21}"; Flags: ignoreversion; Components: revit21
Source: "{#Repository}\src\*.addin"; DestDir: "{#RevitAddin21}"; Flags: ignoreversion; Components: revit21

Source: "{#Repository}\output\Release-2022\EPPlus*.dll"; DestDir: "{#RevitFolder22}"; Flags: ignoreversion recursesubdirs; Components: revit22
Source: "{#Repository}\output\Release-2022\{#RevitAppName}.dll"; DestDir: "{#RevitFolder22}"; Flags: ignoreversion recursesubdirs; Components: revit22
Source: "{#Repository}\output\Release-2022\Microsoft.IO.RecyclableMemoryStream.dll"; DestDir: "{#RevitFolder22}"; Flags: ignoreversion; Components: revit22
Source: "{#Repository}\src\*.addin"; DestDir: "{#RevitAddin22}"; Flags: ignoreversion; Components: revit22

Source: "{#Repository}\output\Release-2023\EPPlus*.dll"; DestDir: "{#RevitFolder23}"; Flags: ignoreversion recursesubdirs; Components: revit23
Source: "{#Repository}\output\Release-2023\{#RevitAppName}.dll"; DestDir: "{#RevitFolder23}"; Flags: ignoreversion recursesubdirs; Components: revit23
Source: "{#Repository}\output\Release-2023\Microsoft.IO.RecyclableMemoryStream.dll"; DestDir: "{#RevitFolder23}"; Flags: ignoreversion; Components: revit23
Source: "{#Repository}\src\*.addin"; DestDir: "{#RevitAddin23}"; Flags: ignoreversion; Components: revit23

Source: "{#Repository}\output\Release-2024\EPPlus*.dll"; DestDir: "{#RevitFolder24}"; Flags: ignoreversion recursesubdirs; Components: revit24
Source: "{#Repository}\output\Release-2024\{#RevitAppName}.dll"; DestDir: "{#RevitFolder24}"; Flags: ignoreversion recursesubdirs; Components: revit24
Source: "{#Repository}\output\Release-2024\Microsoft.IO.RecyclableMemoryStream.dll"; DestDir: "{#RevitFolder24}"; Flags: ignoreversion; Components: revit24
Source: "{#Repository}\src\*.addin"; DestDir: "{#RevitAddin24}"; Flags: ignoreversion; Components: revit24

Source: "{#Repository}\output\Release-2025\EPPlus*.dll"; DestDir: "{#RevitFolder25}"; Flags: ignoreversion recursesubdirs; Components: revit25
Source: "{#Repository}\output\Release-2025\{#RevitAppName}.dll"; DestDir: "{#RevitFolder25}"; Flags: ignoreversion recursesubdirs; Components: revit25
Source: "{#Repository}\output\Release-2025\Microsoft.IO.RecyclableMemoryStream.dll"; DestDir: "{#RevitFolder25}"; Flags: ignoreversion; Components: revit25
Source: "{#Repository}\src\*.addin"; DestDir: "{#RevitAddin25}"; Flags: ignoreversion; Components: revit25