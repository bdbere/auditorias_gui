[Setup]
AppName=Gestor Encuestas ITQ
AppVersion=1.0.0
DefaultDirName={pf}\GestorEncuestasITQ
DefaultGroupName=Gestor Encuestas ITQ
OutputBaseFilename=GestorEncuestasITQ-Setup
Compression=lzma
SolidCompression=yes

[Files]
Source: "dist\main.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "carreras.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "servicios.json"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Gestor Encuestas ITQ"; Filename: "{app}\main.exe"
Name: "{commondesktop}\Gestor Encuestas ITQ"; Filename: "{app}\main.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el escritorio"; GroupDescription: "Accesos directos:"
