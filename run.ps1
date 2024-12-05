# PowerShell Script: Activate Virtual Environment and Run the Program
# Automates virtual environment activation and execution of llm_aided_ocr.py.

# 1. Verify Virtual Environment Exists
if (-NOT (Test-Path ".\venv")) {
    Write-Host "The virtual environment 'venv' is not found. Please set it up by running the setup script." -ForegroundColor Red
    Exit 1
}

# 2. Verify Main Python Script Exists
if (-NOT (Test-Path ".\llm_aided_ocr.py")) {
    Write-Host "The main script 'llm_aided_ocr.py' is not found in the current directory." -ForegroundColor Red
    Exit 1
}

# 3. Activate Virtual Environment
Write-Host "Activating the virtual environment..."
& ".\venv\Scripts\Activate.ps1"
if (-NOT $?) {
    Write-Host "Failed to activate the virtual environment. Ensure the 'venv' folder is correctly set up." -ForegroundColor Red
    Exit 1
}

# 4. Execute the Python Program
Write-Host "Running the 'llm_aided_ocr.py' script..."
python llm_aided_ocr.py
if ($LASTEXITCODE -ne 0) {
    Write-Host "The script execution failed. Check for errors in your Python code or dependencies." -ForegroundColor Red
    Exit 1
}

# 5. Completion Message
Write-Host "Program executed successfully!" -ForegroundColor Green