; ================================
; Inno Setup script - AgenteDisponibilidadPC
; ================================

#define AppName        "Agente Disponibilidad PC"
#define MyCompany      "Lubrisider Chile S.A."

; La versión llega desde la línea de comandos:
; ISCC.exe "/DMyAppVersion=0.1.0" "installer\AgenteDisponibilidadPC.iss"
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0-dev"
#endif

; Rutas relativas al repo
#define RootDir        ExtractFileDir(SourcePath)
#define ProjectRoot    AddBackslash(RootDir) + ".."
#define DistDir        AddBackslash(ProjectRoot) + "dist"
#define ReleaseDir     AddBackslash(ProjectRoot) + "releases"
#define AppExe         DistDir + "\\AgenteDisponibilidadPC.exe"
; (opcional) #define SetupIcon    ProjectRoot + "\\installer\\icon.ico"

[Setup]
; Usa un GUID propio y fijo para que Windows reconozca upgrades
AppId={{9A93B9B1-49E0-4F41-9E51-1C9D0F5E9C2A}}
AppName={#AppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyCompany}

; Carpeta destino por defecto
DefaultDirName={autopf}\Lubrisider\AgenteDisponibilidadPC
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes

; Dónde escribir el instalador generado
OutputDir={#ReleaseDir}
OutputBaseFilename=AgenteDisponibilidadPC_{#MyAppVersion}_Setup

; Compresión/arquitectura
Compression=lzma2/ultra64
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin
WizardStyle=modern

; Icono del desinstalador y del programa
UninstallDisplayIcon={app}\AgenteDisponibilidadPC.exe
; (opcional) SetupIconFile={#SetupIcon}

; Idiomas
[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

; Archivos a instalar
[Files]
; ejecutable principal
Source: "{#SourcePath}\..\dist\AgenteDisponibilidadPC.exe"; DestDir: "{app}"; Flags: ignoreversion

; carpetas opcionales (no fallan si no existen)
Source: "{#SourcePath}\..\assets\*"; DestDir: "{app}\assets"; Flags: recursesubdirs ignoreversion skipifsourcedoesntexist
Source: "{#SourcePath}\..\config\*";  DestDir: "{app}\config";  Flags: recursesubdirs ignoreversion skipifsourcedoesntexist


; Tareas opcionales
[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el Escritorio"; GroupDescription: "Accesos directos:"; Flags: checkedonce

; Accesos directos
[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\AgenteDisponibilidadPC.exe"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\AgenteDisponibilidadPC.exe"; Tasks: desktopicon

; Ejecutar al finalizar la instalación
[Run]
Filename: "{app}\AgenteDisponibilidadPC.exe"; Description: "Ejecutar {#AppName} al finalizar"; Flags: nowait postinstall skipifsilent
