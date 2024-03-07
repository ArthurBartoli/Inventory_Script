param (
    [Alias("e")]
    [string]$export,

    [Alias("h")]
    [switch]$help 
)

# Building python command
$pythonArgs = ""
if ($help) { $pythonArgs += " -h" }
if ($export) { $pythonArgs += " -e `"$export`"" }
$pythonCommand = "python start.py$pythonArgs"

# Check if python is accessible by the python command
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

Invoke-Expression $pythonCommand
