#requires -version 7
$ErrorActionPreference = 'Stop'

# ========= Config =========
$Root       = (Resolve-Path "$PSScriptRoot\..").Path
$Venv       = Join-Path $Root ".venv"
$PythonExe  = Join-Path $Venv "Scripts\python.exe"
$AppName    = 'AgenteDisponibilidadPC'
$Main       = Join-Path $Root "src\main.py"
$VersionTxt = Join-Path $Root "version.txt"
$ReleaseDir = Join-Path $Root "releases"
$BuildDir   = Join-Path $Root "build"
$DistDir    = Join-Path $Root "dist"
$Iss        = Join-Path $Root "installer\AgenteDisponibilidadPC.iss"

Write-Host "`n── AgenteDisponibilidadPC - Build interactivo (PowerShell 7) ──`n"

# ========= Inputs (acepta numeros) =========
$bumpSel = Read-Host "Incremento de versión [1=patch / 2=minor / 3=major] (Enter=1)"
switch ($bumpSel) { '2' {$bump='minor'} '3' {$bump='major'} default {$bump='patch'} }

$modeSel = Read-Host "Tipo de app [1=GUI (sin consola) / 2=Consola] (Enter=1)"
$mode = if ($modeSel -eq '2') {'console'} else {'gui'}

$doTag = (Read-Host "Crear y subir tag git al terminar? [s/N]") -match '^[sS]$'
if ($doTag) { Write-Host "[*] Se creará y subirá tag tras el build." }

# ========= Version bump =========
if (-not (Test-Path $VersionTxt)) { Set-Content -NoNewline -Encoding utf8 $VersionTxt '0.1.0' }
$cur = (Get-Content -Raw $VersionTxt).TrimStart('v')
if (-not ($cur -match '^\d+(\.\d+){0,2}$')) { $cur = '0.1.0' }
$parts = $cur.Split('.'); while ($parts.Count -lt 3) { $parts += '0' }
switch ($bump) {
  'major' { $parts[0] = [int]$parts[0] + 1; $parts[1] = '0'; $parts[2] = '0' }
  'minor' { $parts[1] = [int]$parts[1] + 1; $parts[2] = '0' }
  default { $parts[2] = [int]$parts[2] + 1 }
}
$ver = ($parts -join '.')
Set-Content -NoNewline -Encoding utf8 $VersionTxt $ver
Write-Host "[*] Versión nueva: $ver`n"

# ========= Venv / deps =========
if (-not (Test-Path $PythonExe)) {
  Write-Host "[*] Creando venv..."
  try     { & py -3 -m venv $Venv }
  catch   { & python -m venv $Venv }
}
& $PythonExe -m pip install --upgrade pip wheel setuptools
& $PythonExe -m pip install pyinstaller
$req = Join-Path $Root "requirements.txt"
if (Test-Path $req) {
  Write-Host "[*] Instalando requirements.txt..."
  & $PythonExe -m pip install -r $req
}

# ========= Clean =========
Remove-Item -Recurse -Force $BuildDir,$DistDir -ErrorAction SilentlyContinue | Out-Null

# ========= PyInstaller =========
Write-Host "[*] Compilando con PyInstaller..."
$pyArgs = @(
  '-m','PyInstaller',
  '--noconfirm','--clean','--onefile',
  '--name', $AppName,
  '--distpath', $DistDir,
  '--workpath', $BuildDir,
  '--specpath', $BuildDir
)
if ($mode -eq 'gui') { $pyArgs += '--windowed' }
$assets = Join-Path $Root 'assets'
$config = Join-Path $Root 'config'
if (Test-Path $assets) { $pyArgs += @('--add-data', "$assets;assets") }
if (Test-Path $config) { $pyArgs += @('--add-data', "$config;config") }
$pyArgs += $Main
& $PythonExe @pyArgs

# ========= Inno Setup =========
if (-not (Test-Path $Iss)) { throw "No existe $Iss" }
$isccCandidates = @(
  'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
  'C:\Program Files\Inno Setup 6\ISCC.exe'
)
$ISCC = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $ISCC) { throw "No encuentro Inno Setup 6 (ISCC.exe). Instala: https://jrsoftware.org/isdl.php" }

Write-Host "[*] Generando instalador con Inno Setup..."
& $ISCC "/DMyAppVersion=$ver" $Iss
if ($LASTEXITCODE -ne 0) { throw "ISCC falló. Revisa el error anterior." }

# ========= Outputs =========
New-Item -ItemType Directory -Force -Path $ReleaseDir | Out-Null
$portable = Join-Path $ReleaseDir "${AppName}_$ver.exe"
$setup    = Join-Path $ReleaseDir "${AppName}_${ver}_Setup.exe"
Copy-Item (Join-Path $DistDir "$AppName.exe") $portable -Force
Write-Host "`n[OK] Portable: $portable"
Write-Host "[OK] Setup:    $setup`n"

# ========= Tag git (opcional) =========
if ($doTag) {
  if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Warning "git no está en PATH. Omitiendo tag."
  } else {
    Push-Location $Root
    git rev-parse --is-inside-work-tree *>$null
    if ($LASTEXITCODE -ne 0) {
      Write-Warning "Esta carpeta no es un repo git. Omitiendo tag."
    } else {
      Write-Host "[*] Creando tag v$ver y empujando a origin..."
      git tag -a "v$ver" -m "$AppName v$ver" 2>$null; if ($LASTEXITCODE -ne 0) { git tag -f "v$ver" }
      git push origin "v$ver" | Out-Host
    }
    Pop-Location
  }
}

Write-Host "Listo."
