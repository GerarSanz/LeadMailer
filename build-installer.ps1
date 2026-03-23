#Requires -Version 5.1
param(
    [string]$Version       = "1.0.0",
    [string]$InnoSetupPath = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    [switch]$SkipInstaller
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$ProjectFile     = ".\LeadMailer\LeadMailer.csproj"
$PublishDir      = ".\publish"
$InstallerScript = ".\installer\LeadMailer.iss"
$InstallerOutDir = ".\installer-output"
$IconFile        = ".\LeadMailer\Resources\app.ico"
$IconScript      = ".\create-icon.ps1"

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   LeadMailer - Build de produccion" -ForegroundColor Cyan
Write-Host "   Version: $Version" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "[0/3] Verificando icono..." -ForegroundColor Yellow
if (-not (Test-Path $IconFile)) {
    if (Test-Path $IconScript) {
        Write-Host "      Generando icono desde app-icon.png..." -ForegroundColor DarkGray
        & powershell -ExecutionPolicy Bypass -File $IconScript
        if (-not (Test-Path $IconFile)) {
            Write-Host "  [!] No se pudo generar el icono. Continuando sin icono..." -ForegroundColor Yellow
        }
    } else {
        Write-Host "  [!] app.ico no encontrado. El ejecutable usara el icono generico." -ForegroundColor Yellow
    }
} else {
    Write-Host "      [OK] app.ico encontrado." -ForegroundColor DarkGray
}

Write-Host "[1/3] Limpiando carpeta publish..." -ForegroundColor Yellow
if (Test-Path $PublishDir) {
    Remove-Item $PublishDir -Recurse -Force
    Write-Host "      Carpeta eliminada." -ForegroundColor DarkGray
}
New-Item -ItemType Directory -Path $PublishDir | Out-Null

Write-Host ""
Write-Host "[2/3] Publicando aplicacion..." -ForegroundColor Yellow
Write-Host "      Configuracion : Release" -ForegroundColor DarkGray
Write-Host "      Runtime       : win-x64 (self-contained)" -ForegroundColor DarkGray
Write-Host ""

dotnet publish $ProjectFile `
    --configuration Release `
    --runtime win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:ReadyToRun=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -p:DebugType=none `
    -p:DebugSymbols=false `
    -p:Version=$Version `
    -p:FileVersion=$Version `
    --output $PublishDir `
    --nologo

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "  [ERROR] dotnet publish fallo (codigo: $LASTEXITCODE)" -ForegroundColor Red
    exit 1
}

$exe = Join-Path $PublishDir "LeadMailer.exe"
if (-not (Test-Path $exe)) {
    Write-Host "  [ERROR] No se encontro LeadMailer.exe en publish" -ForegroundColor Red
    exit 1
}

$sizeMB = [math]::Round((Get-Item $exe).Length / 1MB, 1)
Write-Host "  [OK] LeadMailer.exe generado ($sizeMB MB)" -ForegroundColor Green

if ($SkipInstaller) {
    Write-Host ""
    Write-Host "[3/3] Paso de instalador omitido (-SkipInstaller)." -ForegroundColor DarkGray
}
else {
    Write-Host ""
    Write-Host "[3/3] Compilando instalador con Inno Setup..." -ForegroundColor Yellow

    if (-not (Test-Path $InnoSetupPath)) {
        Write-Host ""
        Write-Host "  [!] Inno Setup 6 no encontrado en: $InnoSetupPath" -ForegroundColor Yellow
        Write-Host "      Descargalo desde: https://jrsoftware.org/isinfo.php" -ForegroundColor Cyan
        Write-Host "      El ejecutable publicado esta en: $PublishDir" -ForegroundColor Green
        exit 0
    }

    if (-not (Test-Path $InstallerOutDir)) {
        New-Item -ItemType Directory -Path $InstallerOutDir | Out-Null
    }

    & $InnoSetupPath $InstallerScript "/DMyAppVersion=$Version" "/Q"

    if ($LASTEXITCODE -ne 0) {
        Write-Host "  [ERROR] Inno Setup fallo (codigo: $LASTEXITCODE)" -ForegroundColor Red
        exit 1
    }

    $installer = Get-ChildItem $InstallerOutDir -Filter "LeadMailer_Setup_v*.exe" |
                 Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if ($installer) {
        $isizeMB = [math]::Round($installer.Length / 1MB, 1)
        Write-Host "  [OK] Instalador: $($installer.Name) ($isizeMB MB)" -ForegroundColor Green
        Write-Host "       Ruta: $($installer.FullName)" -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host "   [OK] Build completado correctamente" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host ""