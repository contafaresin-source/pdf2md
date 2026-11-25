# Gera a versão PORTÁTIL (Windows x64, CPU) do PDF -> Markdown.
# Rode UMA vez, numa máquina COM internet:   build_portable.bat
# Resultado: dist\PDF2MD-Portable\  -> copie para o pendrive (exFAT/NTFS).
#
#   build_portable.bat                  tudo (Docling + Marker)
#   build_portable.bat -SkipMarker      só Docling
#   build_portable.bat -SkipDocling     só Marker
#   build_portable.bat -Dest D:\PDF2MD  destino customizado
param(
    [string]$Dest = "$PSScriptRoot\dist\PDF2MD-Portable",
    [switch]$SkipDocling,
    [switch]$SkipMarker,
    [string]$PythonVersion = "3.11"
)
$ErrorActionPreference = "Continue"   # uv/pip escrevem progresso em stderr
Set-Location $PSScriptRoot
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Step($t) { Write-Host "`n=== $t ===" -ForegroundColor Cyan }
function Check($what) { if ($LASTEXITCODE -ne 0) { throw "Falhou: $what (código $LASTEXITCODE)" } }

function New-SamplePdf([string]$path) {
    $content = "BT /F1 24 Tf 72 700 Td (Warm-up Document) Tj /F1 11 Tf 0 -36 Td (This is a sample paragraph used to download the conversion models.) Tj ET"
    $objs = @(
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        "<< /Length $($content.Length) >>`nstream`n$content`nendstream",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
    )
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("%PDF-1.4`n")
    $offs = @()
    for ($i = 0; $i -lt $objs.Count; $i++) {
        $offs += $sb.Length
        [void]$sb.Append("$($i + 1) 0 obj`n$($objs[$i])`nendobj`n")
    }
    $xref = $sb.Length
    [void]$sb.Append("xref`n0 $($objs.Count + 1)`n0000000000 65535 f `n")
    foreach ($o in $offs) { [void]$sb.Append(("{0:D10} 00000 n `n" -f $o)) }
    [void]$sb.Append("trailer`n<< /Size $($objs.Count + 1) /Root 1 0 R >>`nstartxref`n$xref`n%%EOF`n")
    [IO.File]::WriteAllText($path, $sb.ToString(), [Text.Encoding]::ASCII)
}

# ------------------------------------------------------------------ 1. arquivos do app
Step "1/6 Copiando o app para $Dest"
if (Test-Path $Dest) { throw "O destino já existe: $Dest (apague ou escolha outro com -Dest)" }
New-Item -ItemType Directory -Path $Dest | Out-Null
foreach ($f in "app.py", "engines.py", "engine_worker.py", "postprocess.py", "requirements.txt") { Copy-Item "$PSScriptRoot\$f" $Dest }
Copy-Item "$PSScriptRoot\static" "$Dest\static" -Recurse
foreach ($d in "cache\hf", "cache\torch", "cache\datalab", "cache\xdg", "home\AppData\Local", "home\AppData\Roaming", "saida", "logs", "temp", "libs") {
    New-Item -ItemType Directory -Path "$Dest\$d" -Force | Out-Null
}

# ------------------------------------------------------------------ 2. uv + Python standalone
Step "2/6 Baixando uv e Python $PythonVersion (standalone, relocável)"
$tools = Join-Path $env:TEMP "pdf2md-build-tools"
New-Item -ItemType Directory -Path $tools -Force | Out-Null
$zip = "$tools\uv.zip"
Invoke-WebRequest -UseBasicParsing "https://github.com/astral-sh/uv/releases/latest/download/uv-x86_64-pc-windows-msvc.zip" -OutFile $zip
Expand-Archive $zip -DestinationPath $tools -Force
$uv = (Get-ChildItem $tools -Recurse -Filter uv.exe | Select-Object -First 1).FullName
if (-not $uv) { throw "uv.exe não encontrado após o download" }

$env:UV_CACHE_DIR = "$tools\cache"
$env:UV_PYTHON_INSTALL_DIR = "$Dest\runtime"
$env:UV_PYTHON_INSTALL_BIN = "0"         # não cria atalhos fora da pasta
$env:UV_PYTHON_INSTALL_REGISTRY = "0"    # não registra no Windows
& $uv python install $PythonVersion
Check "uv python install"
$py = (Get-ChildItem "$Dest\runtime" -Directory | ForEach-Object { Join-Path $_.FullName "python.exe" } | Where-Object { Test-Path $_ } | Select-Object -First 1)
if (-not $py) { throw "python.exe não encontrado em $Dest\runtime" }
Write-Host "Python: $py"

# ------------------------------------------------------------------ 3. dependências
Step "3/6 Instalando app (Flask, pywin32, Pandoc)"
& $uv pip install --python $py --break-system-packages -r "$Dest\requirements.txt"
Check "dependências do app"

if (-not $SkipDocling) {
    Step "4/6 Instalando Docling (pode demorar e baixar vários GB)"
    & $uv pip install --python $py --break-system-packages --target "$Dest\libs\docling" docling
    Check "docling"
}
if (-not $SkipMarker) {
    Step "5/6 Instalando Marker"
    & $uv pip install --python $py --break-system-packages --target "$Dest\libs\marker" marker-pdf
    Check "marker-pdf"
}

# ------------------------------------------------------------------ 4. modelos (warm-up)
Step "6/6 Baixando modelos (uma conversão de teste por motor)"
$env:HF_HOME = "$Dest\cache\hf"; $env:TORCH_HOME = "$Dest\cache\torch"
$env:MODEL_CACHE_DIR = "$Dest\cache\datalab"; $env:XDG_CACHE_HOME = "$Dest\cache\xdg"
$env:USERPROFILE = "$Dest\home"; $env:HOME = "$Dest\home"
$env:LOCALAPPDATA = "$Dest\home\AppData\Local"; $env:APPDATA = "$Dest\home\AppData\Roaming"
$env:PYTHONNOUSERSITE = "1"; $env:PYTHONUTF8 = "1"

# Redes com proxy/antivirus que inspecionam HTTPS usam uma CA propria, instalada no Windows mas
# ausente do certifi do Python (erro "self-signed certificate in certificate chain").
# Exporta as CAs confiaveis do Windows para PEM e aponta requests/ssl/curl para ele.
$caPem = "$tools\windows-ca.pem"
$caSb = New-Object System.Text.StringBuilder
foreach ($store in "Cert:\LocalMachine\Root", "Cert:\CurrentUser\Root", "Cert:\LocalMachine\CA") {
    foreach ($cert in Get-ChildItem $store -ErrorAction SilentlyContinue) {
        [void]$caSb.Append("-----BEGIN CERTIFICATE-----`n")
        [void]$caSb.Append([Convert]::ToBase64String($cert.RawData, 'InsertLineBreaks'))
        [void]$caSb.Append("`n-----END CERTIFICATE-----`n")
    }
}
[IO.File]::WriteAllText($caPem, $caSb.ToString(), [Text.Encoding]::ASCII)
$env:REQUESTS_CA_BUNDLE = $caPem; $env:SSL_CERT_FILE = $caPem; $env:CURL_CA_BUNDLE = $caPem

$sample = "$tools\amostra.pdf"; New-SamplePdf $sample
$warn = @()

if (-not $SkipDocling) {
    $env:PYTHONPATH = "$Dest\libs\docling"
    $entry = "from docling.cli.main import app; app()"
    foreach ($ocr in "--no-ocr", "--ocr") {
        $o = "$tools\out-docling"; Remove-Item $o -Recurse -Force -ErrorAction SilentlyContinue
        & $py -X utf8 -c $entry $sample --from pdf --to md --output $o --image-export-mode referenced $ocr
        if ($LASTEXITCODE -ne 0) { $warn += "Docling ($ocr): warm-up falhou - os modelos serão baixados no 1º uso (precisa de internet)." }
    }
}
if (-not $SkipMarker) {
    $env:PYTHONPATH = "$Dest\libs\marker"
    $entry = "from marker.scripts.convert_single import convert_single_cli; convert_single_cli()"
    $o = "$tools\out-marker"; Remove-Item $o -Recurse -Force -ErrorAction SilentlyContinue
    & $py -X utf8 -c $entry $sample --output_dir $o --output_format markdown
    if ($LASTEXITCODE -ne 0) { $warn += "Marker: warm-up falhou - os modelos serão baixados no 1º uso (precisa de internet)." }
}
Remove-Item Env:PYTHONPATH -ErrorAction SilentlyContinue

# ------------------------------------------------------------------ 5. lançador e limpeza
@'
@echo off
cd /d "%~dp0"
set "PY="
for /d %%D in (runtime\cpython*) do set "PY=%%D\python.exe"
if not defined PY (
  echo Python embutido nao encontrado em runtime\
  pause
  exit /b 1
)
set PYTHONNOUSERSITE=1
set PYTHONDONTWRITEBYTECODE=1
"%PY%" app.py
pause
'@ | Set-Content "$Dest\iniciar.bat" -Encoding ASCII

@'
PDF -> Markdown (portatil)
==========================
1. De dois cliques em iniciar.bat. O navegador abre em http://127.0.0.1:5057
2. Arraste os PDFs, escolha o motor e converta. Saida: pasta "saida" (ou outra, definida na tela).
3. Erros: pasta "logs" (errors.log) ou botao "Log de erros" na tela.

- Docling/Marker funcionam offline (modelos estao na pasta "cache").
  Se algum modelo faltar, defina a variavel PDF2MD_ONLINE=1 antes de iniciar para permitir download.
- O motor Word so aparece disponivel em PCs com Microsoft Word instalado.
- Mantenha o caminho curto (ex.: E:\PDF2MD). Use pendrive USB 3 formatado em exFAT ou NTFS.
- Windows 64 bits, CPU (sem GPU).
'@ | Set-Content "$Dest\LEIA-ME.txt" -Encoding ASCII

Get-ChildItem $Dest -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $tools -Recurse -Force -ErrorAction SilentlyContinue

$size = [math]::Round(((Get-ChildItem $Dest -Recurse -File | Measure-Object Length -Sum).Sum) / 1GB, 1)
Write-Host "`nConcluído: $Dest  (~$size GB)" -ForegroundColor Green
if ($warn.Count) { Write-Host "`nAvisos:" -ForegroundColor Yellow; $warn | ForEach-Object { Write-Host " - $_" -ForegroundColor Yellow } }
Write-Host "`nCopie a pasta para o pendrive, por exemplo:`n  robocopy `"$Dest`" E:\PDF2MD /E" -ForegroundColor Green
