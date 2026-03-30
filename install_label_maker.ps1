Write-Host ""
Write-Host "==============================="
Write-Host " JTI Label Maker Installer"
Write-Host "==============================="
Write-Host ""

# Ensure Python exists
try {
    $pythonVersion = py -3 --version
    Write-Host "Python detected: $pythonVersion"
}
catch {
    Write-Host "Python not found. Install Python 3.10+ first."
    exit 1
}

Write-Host ""
Write-Host "Installing required packages..."
Write-Host ""

py -m pip install --upgrade pip

py -m pip install ^
    reportlab ^
    pdfplumber ^
    pillow ^
    pytesseract ^
    tkinterdnd2

Write-Host ""
Write-Host "Dependencies installed."
Write-Host ""

Write-Host "Checking Tesseract..."
if (Test-Path "C:\Program Files\Tesseract-OCR\tesseract.exe") {
    Write-Host "Tesseract found."
} else {
    Write-Host ""
    Write-Host "Tesseract NOT found."
    Write-Host "Download from:"
    Write-Host "https://github.com/UB-Mannheim/tesseract/wiki"
    Write-Host ""
    Write-Host "After installing, make sure it installs to:"
    Write-Host "C:\Program Files\Tesseract-OCR\"
}

Write-Host ""
Write-Host "Install complete."
Write-Host ""
Pause