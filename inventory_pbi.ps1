$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    Write-Output "/!\ ====> Python installation has not been found. Please install python or check system path and run the script again."
    exit
}

# Installing the module to query PBI
Install-Module -Name MicrosoftPowerBIMgmt -Force -Scope CurrentUser

# Installing virtual env
& python -m pip install virtualenv
& "$PSScriptRoot\venv\Scripts\activate.ps1"

# Setting up virtual env
& python -m ensurepip
& python -m pip install --upgrade pip
& python -m pip install -r requirements.txt
& python start.py
