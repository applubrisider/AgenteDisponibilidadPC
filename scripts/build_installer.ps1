param(
  [ValidateSet('patch','minor','major')] [string]$Increment = 'patch',
  [ValidateSet('gui','console')]         [string]$Type      = 'gui',
  # Publish 'actions' => solo tag/push y deja que GitHub Actions genere la release
  # Publish 'local'   => crea release con artefactos locales usando gh (opcional)
  [ValidateSet('actions','local')]       [string]$Publish   = 'actions'
)

$ErrorActionPreference = 'Stop'

# ── Rutas ──────────────────────────────────────────────────────────────────────
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root      = Resolve-Path (Join-Path $ScriptDir '..')
Set-Location $Root

$AppName     = 'AgenteDisponibilidadPC'
$VersionFile = Join-Path $Root 'version.txt'
$Venv        = Join-Path $Root '.venv'
$DistDir     = Join-Path $Root 'dist'
$ReleaseDir  = Join-Path $Root 'releases'
$MainPy      = Join-Path $Root 'src\main.py'
$IssFile     = Join-Path $Root 'installer\AgenteDisponibilidadPC.iss'

# ── Función versión ────────────────────────────────────────────────────────────
function Get-NextVersion([string]$cur, [string]$inc){
  if (-not $cur) { return '0.1.0' }
  $maj,$min,$pat = ($cur -split '\.')[0..2] | ForEach-Object {[int]$_}
  switch ($inc) {
    'major' { $maj++; $min=0; $pat=0 }
    'minor' { $min++; $pat=0 }
    default { $pat++ }
  }
  return "$maj.$min.$pat"
}

# ── Version bump ───────────────────────────────────────────────────────────────
if (-not (Test-Path $VersionFile)) { '0.1.0' | Out-File -Encoding utf8 $VersionFile }
$CurrentVersion = (Get-Content $VersionFile -Raw).Trim()
$NewVersion     = Get-NextVersion $CurrentVersion $Increment
$NewVersion | Out-File -Encoding utf8 $VersionFile
Write-Host "[*] Versión nueva: $NewVersion"

# ── Venv / deps ────────────────────────────────────────────────────────────────
if (-not (Test-Path (Join-Path $Venv 'Scripts\python.exe'))) {
  Write-Host "[*] Creando venv..."
  & py -3 -m venv $Venv
}
& "$Venv\Scripts\python.exe" -m pip install --upgrade pip wheel setuptools | Out-Null
& "$Venv\Scripts\python.exe" -m pip install pyinstaller        | Out-Null
if (Test-Path 'requirements.txt') {
  Write-Host "[*] Instalando requirements.txt..."
  & "$Venv\Scripts\python.exe" -m pip install -r requirements.txt
}

# ── Limpieza ───────────────────────────────────────────────────────────────────
if (Test-Path $DistDir)     { Remove-Item $DistDir -Recurse -Force }
if (Test-Path "$Root\build"){ Remove-Item "$Root\build" -Recurse -Force }
New-Item -ItemType Directory -Force -Path $ReleaseDir | Out-Null

# ── PyInstaller ────────────────────────────────────────────────────────────────
Write-Host "[*] Compilando con PyInstaller..."
$pyiArgs = @('--noconfirm','--clean','--onefile','--name', $AppName, $MainPy)
if ($Type -eq 'gui') { $pyiArgs = @('--windowed') + $pyiArgs } else { $pyiArgs = @('--console') + $pyiArgs }
& "$Venv\Scripts\pyinstaller.exe" $pyiArgs

$PortableExe = Join-Path $DistDir "$AppName.exe"
if (-not (Test-Path $PortableExe)) { throw "No se generó $PortableExe" }

# ── Inno Setup ────────────────────────────────────────────────────────────────
Write-Host "[*] Generando instalador con Inno Setup..."
$iscc = "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"
if (-not (Test-Path $iscc)) { $iscc = "${env:ProgramFiles}\Inno Setup 6\ISCC.exe" }
if (-not (Test-Path $iscc)) { throw "ISCC.exe no encontrado. Instala Inno Setup 6." }

& $iscc "/DMyAppVersion=$NewVersion" $IssFile
$innoExit = $LASTEXITCODE

# ── Colecta artefactos ────────────────────────────────────────────────────────
$PortableOut = Join-Path $ReleaseDir "$AppName`_$NewVersion.exe"
Copy-Item $PortableExe $PortableOut -Force

# Encuentra el _Setup generado por Inno (queda en releases por el .iss)
$SetupOut = Get-ChildItem $ReleaseDir -Filter "*_${NewVersion}_Setup.exe" | Select-Object -First 1
if (-not $SetupOut) { throw "No se encontró instalador *_${NewVersion}_Setup.exe en $ReleaseDir" }

Write-Host "[OK] Portable: $PortableOut"
Write-Host "[OK] Setup:    $($SetupOut.FullName)"

# ── Git: commit + tag + push ──────────────────────────────────────────────────
function InGitRepo { git rev-parse --is-inside-work-tree 2>$null | Out-Null; return ($LASTEXITCODE -eq 0) }
if (InGitRepo) {
  git add -A
  try { git commit -m "release: v$NewVersion" } catch { Write-Host "[i] Nada que commitear" }
  git tag -a "v$NewVersion" -m "Release v$NewVersion"
  git push origin HEAD
  git push origin "v$NewVersion"
  git push origin "v$NewVersion"
  Write-Host "[OK] Tag v$NewVersion creado y pusheado."
} else {
  Write-Host "[!] Carpeta no es repo git. Omitiendo tag/push."
}

# ── Publicación ───────────────────────────────────────────────────────────────
if ($Publish -eq 'local') {
  # Requiere GitHub CLI autenticado: gh auth login
  if (Get-Command gh -ErrorAction SilentlyContinue) {
    Write-Host "[*] Publicando release local con gh..."
    gh release create "v$NewVersion" `
      "$PortableOut" "$($SetupOut.FullName)" `
      --title "Agente Disponibilidad PC $NewVersion" `
      --notes "Release automática local" --latest
    Write-Host "[OK] Release publicada con artefactos locales."
  } else {
    Write-Host "[!] gh no está instalado. Instálalo con: winget install GitHub.cli  (o choco install gh)"
    Write-Host "    O deja Publish='actions' para que publique GitHub Actions."
  }
} else {
  Write-Host "[i] Publicación vía GitHub Actions: el tag v$NewVersion activó el workflow."
}
