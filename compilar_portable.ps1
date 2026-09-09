$ErrorActionPreference = "Stop"

$proyecto = $PSScriptRoot
$codigo = Join-Path $proyecto "dividir_guias_remision.py"
$entorno = Join-Path $proyecto ".venv"
$pythonEntorno = Join-Path $entorno "Scripts\python.exe"

if (-not (Test-Path -LiteralPath $codigo -PathType Leaf)) {
    throw "No se encontro dividir_guias_remision.py en: $proyecto"
}

$candidatosTesseract = @(
    "C:\Program Files\Tesseract-OCR",
    "C:\Program Files (x86)\Tesseract-OCR"
)

$comandoTesseract = Get-Command tesseract.exe -ErrorAction SilentlyContinue
if ($comandoTesseract) {
    $candidatosTesseract = @((Split-Path $comandoTesseract.Source -Parent)) + $candidatosTesseract
}

$tesseractOrigen = $candidatosTesseract |
    Where-Object { Test-Path -LiteralPath (Join-Path $_ "tesseract.exe") -PathType Leaf } |
    Select-Object -First 1

if (-not $tesseractOrigen) {
    throw "No se encontro Tesseract OCR. Instalalo primero en C:\Program Files\Tesseract-OCR."
}

$idiomaSpa = Join-Path $tesseractOrigen "tessdata\spa.traineddata"
$idiomaEng = Join-Path $tesseractOrigen "tessdata\eng.traineddata"
if (-not (Test-Path -LiteralPath $idiomaSpa -PathType Leaf)) {
    throw "Falta el idioma espanol: $idiomaSpa"
}
if (-not (Test-Path -LiteralPath $idiomaEng -PathType Leaf)) {
    throw "Falta el idioma ingles: $idiomaEng"
}

if (-not (Test-Path -LiteralPath $pythonEntorno -PathType Leaf)) {
    $py = Get-Command py.exe -ErrorAction SilentlyContinue
    if (-not $py) {
        throw "No se encontro Python. Instalalo desde python.org y activa Add Python to PATH."
    }
    Write-Host "Creando entorno de compilacion..."
    & $py.Source -3 -m venv $entorno
}

Write-Host "Instalando o actualizando dependencias..."
& $pythonEntorno -m pip install --upgrade pip
& $pythonEntorno -m pip install --upgrade pypdf pymupdf pillow pytesseract pyinstaller

$temporal = Join-Path ([IO.Path]::GetTempPath()) ("DividirGRE_" + [guid]::NewGuid().ToString("N"))
$tesseractDestino = Join-Path $temporal "Tesseract-OCR"
$tessdataDestino = Join-Path $tesseractDestino "tessdata"

try {
    New-Item -ItemType Directory -Force -Path $tessdataDestino | Out-Null

    Write-Host "Preparando Tesseract portable..."
    Copy-Item -LiteralPath (Join-Path $tesseractOrigen "tesseract.exe") -Destination $tesseractDestino
    Get-ChildItem -LiteralPath $tesseractOrigen -Filter "*.dll" -File |
        Copy-Item -Destination $tesseractDestino

    Copy-Item -LiteralPath $idiomaSpa -Destination $tessdataDestino
    Copy-Item -LiteralPath $idiomaEng -Destination $tessdataDestino

    $idiomaOsd = Join-Path $tesseractOrigen "tessdata\osd.traineddata"
    if (Test-Path -LiteralPath $idiomaOsd -PathType Leaf) {
        Copy-Item -LiteralPath $idiomaOsd -Destination $tessdataDestino
    }

    $argumentos = @(
        "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--console",
        "--name", "DividirGuiasRemision",
        "--add-data", ($tesseractDestino + ";Tesseract-OCR"),
        $codigo
    )

    Write-Host "Generando ejecutable unico..."
    & $pythonEntorno @argumentos
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller termino con el codigo $LASTEXITCODE."
    }
}
finally {
    if (Test-Path -LiteralPath $temporal) {
        Remove-Item -LiteralPath $temporal -Recurse -Force
    }
}

$ejecutable = Join-Path $proyecto "dist\DividirGuiasRemision.exe"
if (-not (Test-Path -LiteralPath $ejecutable -PathType Leaf)) {
    throw "La compilacion termino, pero no se encontro el ejecutable esperado."
}

$tamanoMb = [math]::Round((Get-Item -LiteralPath $ejecutable).Length / 1MB, 1)
Write-Host ""
Write-Host "COMPILACION COMPLETADA"
Write-Host "Ejecutable: $ejecutable"
Write-Host "Tamano: $tamanoMb MB"
Write-Host "Puedes copiar solamente este EXE a otra PC con Windows de 64 bits."
