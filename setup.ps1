# Configura o ambiente do PDF -> Markdown
# Uso:  setup.bat            (app base + pergunta sobre Docling/Marker)
#       setup.bat -Docling   (instala Docling)
#       setup.bat -Marker    (instala Marker)
#       setup.bat -All       (tudo, sem perguntas)
#
# Docling e Marker usam o Python 3.12, baixado automaticamente pelo `uv`: varias dependencias
# (ex.: Pillow 10.4 do Marker) nao tem pacotes prontos para Python 3.13/3.14.
param([switch]$Docling, [switch]$Marker, [switch]$All)
$ErrorActionPreference = "Continue"
Set-Location $PSScriptRoot
$EnginePython = "3.12"

function Find-Python {
    $cands = @(
        @{ exe = "py"; args = @("-3.12") }, @{ exe = "py"; args = @("-3.11") },
        @{ exe = "py"; args = @("-3.13") }, @{ exe = "py"; args = @("-3") },
        @{ exe = "python"; args = @() }
    )
    foreach ($c in $cands) {
        try {
            $out = & $c.exe @($c.args) --version 2>&1
            if ($LASTEXITCODE -eq 0 -and "$out" -match "Python 3\.(\d+)" -and [int]$Matches[1] -ge 10) { return $c }
        } catch { }
    }
    return $null
}

function Get-VenvMinor([string]$path) {
    $cfg = "$path\pyvenv.cfg"
    if (-not (Test-Path $cfg)) { return $null }
    $line = Get-Content $cfg | Where-Object { $_ -match "^\s*version(_info)?\s*=\s*3\.(\d+)" } | Select-Object -First 1
    if ($line -and $line -match "3\.(\d+)") { return [int]$Matches[1] }
    return $null
}

# venv do app (Flask + Word/Pandoc): qualquer Python 3.10+ serve
function New-AppVenv {
    if (-not (Test-Path ".venv\Scripts\python.exe")) {
        Write-Host ">> Criando .venv ..." -ForegroundColor Cyan
        & $script:py.exe @($script:py.args) -m venv .venv
        if ($LASTEXITCODE -ne 0) { throw "Falha ao criar .venv" }
    }
    & .venv\Scripts\python.exe -m pip install --upgrade pip --quiet
    Write-Host ">> pip install -r requirements.txt" -ForegroundColor Cyan
    & .venv\Scripts\python.exe -m pip install -r requirements.txt uv
    if ($LASTEXITCODE -ne 0) { throw "Falha ao instalar requirements.txt" }
}

# venv de motor (Docling/Marker): Python 3.12 via uv
function New-EngineVenv([string]$path, [string]$package, [int]$maxMinor) {
    $uv = ".venv\Scripts\uv.exe"
    $minor = Get-VenvMinor $path
    if ($null -ne $minor -and ($minor -lt 10 -or $minor -gt $maxMinor)) {
        Write-Host ">> $path usa Python 3.$minor (incompativel) - recriando com Python $EnginePython ..." -ForegroundColor Yellow
        Remove-Item -Recurse -Force $path
    }
    if (-not (Test-Path "$path\Scripts\python.exe")) {
        Write-Host ">> Criando $path com Python $EnginePython (o uv baixa se precisar) ..." -ForegroundColor Cyan
        & $uv venv $path --python $EnginePython
        if ($LASTEXITCODE -ne 0) { throw "Falha ao criar $path" }
    }
    Write-Host ">> Instalando $package em $path (varios GB, pode demorar)" -ForegroundColor Cyan
    & $uv pip install --python "$path\Scripts\python.exe" $package
    if ($LASTEXITCODE -ne 0) { throw "Falha ao instalar: $package" }
}

function Ask([string]$q) { return ((Read-Host "$q (s/N)") -match "^[sSyY]") }

$script:py = Find-Python
if (-not $script:py) { throw "Python 3.10+ nao encontrado. Instale em https://www.python.org/downloads/ (marque 'Add to PATH')." }
Write-Host "Python encontrado: $($script:py.exe) $($script:py.args -join ' ')" -ForegroundColor Green

New-AppVenv

if (Test-Path "Registry::HKEY_CLASSES_ROOT\Word.Application") {
    Write-Host "Microsoft Word detectado." -ForegroundColor Green
} else {
    Write-Host "AVISO: Microsoft Word nao detectado - o motor 'Word' ficara indisponivel." -ForegroundColor Yellow
}

$doDocling = $Docling -or $All
$doMarker = $Marker -or $All
if (-not ($Docling -or $Marker -or $All)) {
    $doDocling = Ask "Instalar Docling (~2-4 GB; melhor para artigos cientificos)?"
    $doMarker = Ask "Instalar Marker (~3-5 GB; boa ordem de leitura e equacoes)?"
}
# venvs separados: Docling e Marker fixam versoes diferentes de dependencias
if ($doDocling) { New-EngineVenv ".venv-docling" "docling" 14 }   # Docling funciona ate o 3.14
if ($doMarker) { New-EngineVenv ".venv-marker" "marker-pdf" 12 }   # Marker: Pillow 10.4 so ate o 3.12

Write-Host "`nPronto! Rode iniciar.bat (ou F5 no VS Code)." -ForegroundColor Green
