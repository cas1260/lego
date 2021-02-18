; 'C:\Cleber\Instaladores\lego\Setup.lst' imported by ISTool version 3.0.4

[Setup]
AppName=Lego 1.0
AppVerName=Lego 1.0
AdminPrivilegesRequired=true
DefaultDirName={pf}\Lego
DefaultGroupName=NS Lego
OutputBaseFilename=instalador
MessagesFile=compiler:brasil.isl
AppCopyright=Todos os direitos Reservados a Neo Software
WindowShowCaption=false
WindowStartMaximized=false
WindowResizable=false
WindowVisible=true
AppPublisher=Neo Softwarwe
AppPublisherURL=www.neobh.com.br
AppSupportURL=www.neobh.com.br
AppUpdatesURL=lego.neobh.com.br
AppVersion=1.0
UninstallDisplayIcon={app}\My PC.ico
UseSetupLdr=true
AllowUNCPath=true
Compression=zip/9
DisableStartupPrompt=true



[Files]
Source: C:\Cleber\Instaladores\lego\Support\COMCAT.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\Windows\System\MSVCRT40.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: C:\Windows\System\STDOLE2.TLB; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: C:\Windows\System\ASYCFILT.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: C:\Windows\System\OLEPRO32.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\Windows\System\OLEAUT32.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\WINDOWS\SYSTEM\MSVBVM60.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\Windows\System\MSSTDFMT.DLL; DestDir: {sys}; CopyMode: normal; Flags: regserver sharedfile
Source: C:\Windows\System\MSHFLXGD.OCX; DestDir: {sys}; CopyMode: normal; Flags: regserver sharedfile
Source: C:\Windows\System\RICHED32.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: C:\Windows\System\RICHTX32.OCX; DestDir: {sys}; CopyMode: normal; Flags: regserver sharedfile
Source: C:\Cleber\Fontes\Lego 1.1\TOOLBAR3.OCX; DestDir: {sys}; CopyMode: normal; Flags: regserver sharedfile
Source: C:\Windows\System\TABCTL32.OCX; DestDir: {sys}; CopyMode: normal; Flags: regserver sharedfile
Source: C:\Windows\System\COMDLG32.OCX; DestDir: {sys}; CopyMode: normal; Flags: regserver sharedfile
Source: C:\Windows\System\MSCOMCTL.OCX; DestDir: {sys}; CopyMode: normal; Flags: regserver sharedfile
Source: C:\Windows\System\VB5DB.DLL; DestDir: {sys}; CopyMode: normal; Flags: sharedfile
Source: C:\Windows\System\MSREPL35.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: C:\Windows\System\msrd2x35.dll; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\Windows\System\EXPSRV.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\Windows\System\VBAJET32.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: C:\Windows\System\msjint35.dll; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: C:\Windows\System\msjter35.dll; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: C:\Windows\System\MSJET35.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\Cleber\Instaladores\lego\Support\DAO350.DLL; DestDir: {dao}; CopyMode: normal; Flags: regserver sharedfile
Source: C:\Windows\System\Parse32.dll; DestDir: {sys}; CopyMode: normal; Flags: sharedfile
Source: Lego.exe; DestDir: {app}; CopyMode: alwaysoverwrite; Flags: sharedfile
Source: My PC.ico; DestDir: {app}

[Icons]
Name: {group}\NS - Lego; Filename: {app}\Lego.exe; WorkingDir: {app}; IconIndex: 0
Name: {userdesktop}\Lego; Filename: {app}\Lego.exe; IconFilename: {app}\Lego.exe; IconIndex: 0

[Registry]
Root: HKCR; SubKey: .exl; ValueType: string; ValueData: Arquivo executavel; Flags: uninsdeletekey
Root: HKCR; SubKey: Arquivo executavel; ValueType: string; ValueData: Arquivo executavel; Flags: uninsdeletekey
Root: HKCR; SubKey: Arquivo executavel\Shell\Open\Command; ValueType: string; ValueData: """{app}\Lego.exe"" %1 /RUN"; Flags: uninsdeletevalue
Root: HKCR; Subkey: Arquivo executavel\DefaultIcon; ValueType: string; ValueData: {app}\Lego.exe,0; Flags: uninsdeletevalue
Root: HKCR; SubKey: .afs; ValueType: string; ValueData: Fonte de dados de Lego; Flags: uninsdeletekey
Root: HKCR; SubKey: Fonte de dados de Lego; ValueType: string; ValueData: Fonte de dados de Lego; Flags: uninsdeletekey
Root: HKCR; SubKey: Fonte de dados de Lego\Shell\Open\Command; ValueType: string; ValueData: """{app}\Lego.exe"" %1 /OPEN"; Flags: uninsdeletevalue
Root: HKCR; Subkey: Fonte de dados de Lego\DefaultIcon; ValueType: string; ValueData: {app}\My PC.ico,0; Flags: uninsdeletevalue
