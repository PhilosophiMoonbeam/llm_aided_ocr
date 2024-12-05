# PowerShell Script: LLM-Aided OCR Project Setup with .env Creation
# Automates setup of Python, Tesseract OCR, and dependencies and provides an option to create a .env file.

# Check Administrator Privileges
Write-Host "Checking administrator privileges..."
$adminCheck = (whoami /groups | Select-String "S-1-5-32-544") -ne $null
if (-NOT $adminCheck) {
    Write-Host "Error: This script must be run as an Administrator." -ForegroundColor Red
    Write-Host "- Right-click on PowerShell and select 'Run as Administrator'." -ForegroundColor Yellow
    Exit 1
}

# Verify Script Location
if (-NOT (Test-Path "requirements.txt")) {
    Write-Host "The script must be run from the root directory of the project files (where 'requirements.txt' is located)." -ForegroundColor Red
    Exit 1
}

# 1. Install Python 3.12 from the Microsoft Store (if not already installed)
Write-Host "Checking for Python installation..."
if (-NOT (Get-Command "python" -ErrorAction SilentlyContinue)) {
    Write-Host "Python 3.12 is not installed. Installing Python 3.12.x from the Microsoft Store..." -ForegroundColor Yellow
    $pythonPackage = "https://apps.microsoft.com/store/detail/python-312/9NRWMJP3717K"
    Start-Process "ms-windows-store:$pythonPackage" -Wait
    if (-NOT (Get-Command "python" -ErrorAction SilentlyContinue)) {
        Write-Host "Failed to detect Python installation. Please install Python 3.12 manually and ensure it is added to PATH." -ForegroundColor Red
        Exit 1
    }
} else {
    Write-Host "Python is already installed. Skipping installation..." -ForegroundColor Green
}

# 2. Install Tesseract OCR
$tesseractInstaller = "tesseract-ocr-w64-setup-5.5.0.20241111.exe"
$tesseractUrl = "https://github.com/tesseract-ocr/tesseract/releases/download/5.5.0/$tesseractInstaller"
$tesseractPath = "$env:TEMP\$tesseractInstaller"
$tesseractInstallDir = "C:\Program Files\Tesseract-OCR"

Write-Host "Checking for Tesseract installation..."
if (-NOT (Test-Path "$tesseractInstallDir\tesseract.exe")) {
    Write-Host "Downloading and installing Tesseract OCR..."
    Invoke-WebRequest -Uri $tesseractUrl -OutFile $tesseractPath -UseBasicParsing
    Start-Process -FilePath $tesseractPath -ArgumentList "/silent" -Wait

    if (-NOT (Test-Path "$tesseractInstallDir\tesseract.exe")) {
        Write-Host "Tesseract installation failed. Please install it manually from $tesseractUrl." -ForegroundColor Red
        Exit 1
    }
} else {
    Write-Host "Tesseract is already installed. Skipping installation..." -ForegroundColor Green
}

# Add Tesseract to PATH
$path = [System.Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::Machine)
if ($path -notlike "*$tesseractInstallDir*") {
    Write-Host "Adding Tesseract to the system PATH..."
    [System.Environment]::SetEnvironmentVariable("Path", "$path;$tesseractInstallDir", [System.EnvironmentVariableTarget]::Machine)
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::Machine)
    Write-Host "Tesseract successfully added to the system PATH." -ForegroundColor Green
}

# 3. Set Up Python Virtual Environment
Write-Host "Setting up Python virtual environment..."
if (-NOT (Test-Path ".\venv")) {
    python -m venv venv
    if (-NOT (Test-Path ".\venv")) {
        Write-Host "Virtual environment setup failed." -ForegroundColor Red
        Exit 1
    }
    Write-Host "Virtual environment created successfully." -ForegroundColor Green
} else {
    Write-Host "Virtual environment already exists. Skipping creation step..." -ForegroundColor Yellow
}

# Activate virtual environment and enforce UTF-8 for pip
Write-Host "Activating virtual environment and setting UTF-8 encoding..."
& ".\venv\Scripts\Activate.ps1"

# Upgrade pip, setuptools, and wheel
Write-Host "Upgrading pip, setuptools, and wheel..."
python -m pip install --upgrade pip setuptools wheel

# Enforce UTF-8 environment
$env:PYTHONUTF8 = 1

# Install dependencies with UTF-8 enabled
Write-Host "Installing dependencies from requirements.txt..."
if (Test-Path "requirements.txt") {
    python -m pip install -r requirements.txt
    Write-Host "Dependencies installed successfully." -ForegroundColor Green
} else {
    Write-Host "requirements.txt file not found. Ensure the project directory contains it." -ForegroundColor Red
    Exit 1
}

# Function to Create .env File
function Create-EnvFile {
    Write-Host "`n--- .env File Creation ---" -ForegroundColor Cyan

    # Prompt for LLM setup options
    $useLocalLLM = Read-Host "Would you like to use a local LLM? (True/False) [Default: False]"
    if (-not $useLocalLLM) { $useLocalLLM = "False" }

    if ($useLocalLLM -ieq "False") {
        # API-based LLM setup
        $apiProvider = Read-Host "Which cloud API provider would you like to use? (OPENAI/CLAUDE) [Default: OPENAI]"
        if (-not $apiProvider) { $apiProvider = "OPENAI" }
        if ($apiProvider -ieq "OPENAI") {
            $openaiApiKey = Read-Host "Enter your OpenAI API key (leave blank to skip)"
        } elseif ($apiProvider -ieq "CLAUDE") {
            $anthropicApiKey = Read-Host "Enter your CLAUDE API key (leave blank to skip)"
        } else {
            Write-Host "Invalid API provider selection. Defaulting to OPENAI..." -ForegroundColor Yellow
            $apiProvider = "OPENAI"
        }
    }

    # Write the .env file
    $envFilePath = ".\.env"
    Write-Host "Writing .env file to $envFilePath..."
    Set-Content -Path $envFilePath -Value "`n# LLM Configuration"
    Add-Content -Path $envFilePath -Value "USE_LOCAL_LLM=$useLocalLLM"

    if ($useLocalLLM -ieq "False") {
        Add-Content -Path $envFilePath -Value "API_PROVIDER=$apiProvider"
        if ($apiProvider -ieq "OPENAI" -and $openaiApiKey) {
            Add-Content -Path $envFilePath -Value "OPENAI_API_KEY=$openaiApiKey"
        }
        if ($apiProvider -ieq "CLAUDE" -and $anthropicApiKey) {
            Add-Content -Path $envFilePath -Value "ANTHROPIC_API_KEY=$anthropicApiKey"
        }
    }

    Write-Host ".env file created successfully." -ForegroundColor Green
}

# Call the Create-EnvFile function after setup
Write-Host "`nWould you like to create a .env file? (y/n) [Default: y]" -ForegroundColor Cyan
$createEnv = Read-Host
if (-not $createEnv -or $createEnv -ieq "y") {
    Create-EnvFile
} else {
    Write-Host "Skipping .env file creation. Remember to manually configure the .env file before running the script!" -ForegroundColor Yellow
}

Write-Host "`nSetup complete! Next steps:" -ForegroundColor Green
Write-Host "1. Verify your .env file or reconfigure its contents."
Write-Host "2. Place your PDF files in the project directory."
Write-Host "3. Activate the virtual environment: .\venv\Scripts\Activate.ps1"
Write-Host "4. Run the script: python llm_aided_ocr.py"
Write-Host "Enjoy using the LLM-Aided OCR project!"