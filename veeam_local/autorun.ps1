# Path to the Python virtual environment activation script
$activateScript = "..\env\Scripts\Activate.ps1"

# Path to the Python script you want to run
$pythonScript = "veeam_local.py"

# Activate the Python virtual environment
. $activateScript

# Run the specific Python file
python $pythonScript

Read-Host -Prompt "Press Enter to exit"
