# Define the drive letter
$DriveLetter = 'C:'

# Create a namespace object
$namespace = (New-Object -ComObject Shell.Application).NameSpace($DriveLetter)

# Check for failure
if ($namespace -ne $null) {
    # Get the BitLocker protection status
    $bitLockerProtection = $namespace.Self.ExtendedProperty('System.Volume.BitLockerProtection')

    switch ($bitLockerProtection) {
        0 { $bit_status = "Unencrypted" }
        1 { $bit_status = "Encrypted" }
        2 { $bit_status = "Not Encrypted" }
        3 { $bit_status = "In Process" }
        default { $bit_status = "Unknown" }
    }

    # Output the results
    Write-Output "The $DriveLetter is $bit_status."
} else {
    Write-Output "Error"
}

# Define the applications to check
$appsToCheck = @(
    "Microsoft OneDrive",
    "Google Chrome"
)

# Checks if temp file is there 
$outputDir = "C:\temp"
if (-not (Test-Path -Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force
}

# Store the checklist results
$checklistResults = @()

# Retrieve the computer name
$ComputerName = $env:COMPUTERNAME

# Add the computer name and title to the checklist with formatting
$checklistResults += "Computer Name: $ComputerName"
$checklistResults += ""
$checklistResults += "BitLocker Status: $bit_status"
$checklistResults += ""
$checklistResults += "Date: $(Get-Date)"
$checklistResults += ""
$checklistResults += "Installed | Name"
$checklistResults += "-------------------"

# Writes the results from the application check into the checklist
foreach ($appName in $appsToCheck) {
    Write-Host "Checking: $appName"

    $installedApp = Get-Package -Name $appName -ErrorAction SilentlyContinue

    if ($installedApp) {
        Write-Host "$appName is installed."
        $status = "[X]"
    } else {
        Write-Host "$appName is not installed."
        $status = "[ ]"
    }

    $checklistResults += "$status $appName"
}

# Log the results of the checklist
$logPath = "C:\temp\AppChecklistResults.txt"
$checklistResults | Out-File -FilePath $logPath -Encoding UTF8

# Inform the user where the checklist is stored
Write-Host "`nThe checklist results have been saved to $logPath"

# Starts the notepad
Start-Process notepad.exe -ArgumentList $logPath

# Prompt the user to press any key to exit
Write-Host "`nPress any key to exit"
[System.Console]::ReadKey($true) | Out-Null
