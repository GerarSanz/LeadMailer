; ═══════════════════════════════════════════════════════════════════════════
;  LeadMailer — Script de instalador Inno Setup 6
;  Genera: installer-output\LeadMailer_Setup_v2.3.1.exe
;
;  Requisitos:
;    - Inno Setup 6  →  https://jrsoftware.org/isinfo.php
;    - Haber ejecutado previamente build-installer.ps1 (genera la carpeta publish\)
;
;  Para compilar manualmente:
;    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\LeadMailer.iss
; ═══════════════════════════════════════════════════════════════════════════

#define MyAppName        "LeadMailer"
#define MyAppVersion     "2.3.1"
#define MyAppPublisher   "LeadMailer"
#define MyAppURL         ""
#define MyAppExeName     "LeadMailer.exe"
#define MyAppDescription "Automatización de envío de correos a leads desde Excel"
#define SourceDir        "..\publish"
#define OutputDir        "..\installer-output"

; ── Configuración general ────────────────────────────────────────────────────
[Setup]
; IMPORTANTE: no cambies el AppId una vez publicado (identifica la app en el registro)
AppId={{F3A1C82D-4B87-4E2A-9C16-7D2F5E8B3041}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
AppComments={#MyAppDescription}

; Directorio de instalación (por defecto: C:\Program Files\LeadMailer)
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes

; Iconos del instalador y del wizard
SetupIconFile=..\LeadMailer\Resources\app.ico
UninstallDisplayIcon={app}\{#MyAppExeName}

; Salida del instalador
OutputDir={#OutputDir}
OutputBaseFilename=LeadMailer_Setup_v{#MyAppVersion}

; Compresión
Compression=lzma2/ultra64
SolidCompression=yes

; Apariencia
WizardStyle=modern
WizardResizable=yes

; Privilegios: instala sin necesitar administrador (por usuario)
; Si el usuario elige "todos los usuarios", pedirá elevación UAC
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; Solo Windows 10 1803+ y arquitectura x64
MinVersion=10.0.17763
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

; ── Idioma
[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

; ── Tareas opcionales (pantalla "Seleccionar tareas adicionales") ─────────────
[Tasks]
Name: "desktopicon"; \
  Description: "Crear icono en el &Escritorio"; \
  GroupDescription: "Iconos adicionales:"; \
  Flags: unchecked

; ── Archivos a instalar ───────────────────────────────────────────────────────
[Files]
; Ejecutable principal (single-file self-contained)
Source: "{#SourceDir}\{#MyAppExeName}"; \
  DestDir: "{app}"; \
  Flags: ignoreversion

; ── Accesos directos ─────────────────────────────────────────────────────────
[Icons]
; Menú Inicio
Name: "{group}\{#MyAppName}";                    Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Desinstalar {#MyAppName}";        Filename: "{uninstallexe}"
; Escritorio (solo si el usuario marcó la tarea)
Name: "{autodesktop}\{#MyAppName}";              Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

; ── Ejecutar al finalizar la instalación ─────────────────────────────────────
[Run]
Filename: "{app}\{#MyAppExeName}"; \
  Description: "Iniciar {#MyAppName} ahora"; \
  Flags: nowait postinstall skipifsilent

; ── Limpieza al desinstalar ───────────────────────────────────────────────────
; NOTA: Los datos del usuario (cursos, configuración) en %AppData%\LeadMailer\
;       NO se eliminan para preservar la configuración entre reinstalaciones.
;       Solo se eliminan los archivos de log.
[UninstallDelete]
Type: filesandordirs; Name: "{userappdata}\LeadMailer\logs"

; ── Mensajes personalizados ───────────────────────────────────────────────────
[Messages]
; Sobreescribe el mensaje de bienvenida
WelcomeLabel1=Bienvenido al asistente de instalación de [name]
WelcomeLabel2=Este asistente instalará [name/ver] en tu equipo.%n%nSe recomienda cerrar todas las demás aplicaciones antes de continuar.%n%nHaz clic en Siguiente para continuar.
