function Show-Menu {
    Write-Host "Select a department to manage:"
    $index = 1
    $departments.Keys | ForEach-Object {
        Write-Host "$index. $_"
        $index++
    }
    Write-Host "Q. Quit"
}

function Show-Printers {
    param($department)
    $printServers = $departments[$department]
    Write-Host "Available printers for $department :"
    $index = 1
    $printServers | ForEach-Object {
        Write-Host "$index : $_"
        $index++
    }
}

function Add-PrinterOption { 
    param($department, $selectedPrinterIndex) 

    $printServers = $departments[$department] 
    $selectedPrinter = $printServers[$selectedPrinterIndex] 

    try { 
        # Try to add the selected shared printer 
        Add-Printer -ConnectionName $selectedPrinter 

        # Set the added printer as the default printer
        $defaultPrinter = Get-CimInstance -Class Win32_Printer | Where-Object { $_.Name -eq $selectedPrinter } 
        Invoke-CimMethod -InputObject $defaultPrinter -MethodName SetDefaultPrinter

        Write-Host "Shared printer $selectedPrinter added for department $department and set as the default printer successfully!" 
        
        # Remove printers connected to 'W2K8' print servers
        $w2k8Printers = Get-WmiObject -Class Win32_Printer | Where-Object { $_.PortName -match "W2K8" }

        if ($w2k8Printers.Count -gt 0) {
            Write-Host "The following printers connected to 'W2K8' print servers will be removed:"
            $w2k8Printers | ForEach-Object {
                Write-Host $_.Name
            }
            
            $w2k8Printers | ForEach-Object {
                Write-Host "Removing printer: $($_.Name)"
                $_.Delete()
            }
        } else {
            Write-Host "No printers connected to 'W2K8' print servers found."
        }
    }
    catch { 
        Write-Host "Failed to add the printer $selectedPrinter for department $department." 
        Write-Host "Error: $_" 
    } 
}

# Read the departments from the CSV file and create a hashtable
$departments = @{}
Import-Csv -Path "E:\Printer Project\departments.csv" | ForEach-Object {
    $department = $_.Department
    $printServer = $_.PrintServer

    if (-not $departments.ContainsKey($department)) {
        $departments[$department] = @()
    }

    $departments[$department] += $printServer
}

$continue = $true

while ($continue) {
    Show-Menu
    $choice = Read-Host "Enter the number of the department or 'Q' to quit"

    if ($choice -eq 'Q') {
        $continue = $false
        Write-Host "Exiting the script."
    }
    elseif ([int]$choice -ge 1 -and [int]$choice -le $departments.Count) {
        $selectedDepartment = $departments.Keys | Select-Object -Index ([int]$choice - 1)
        Show-Printers -department $selectedDepartment
        $printerChoice = Read-Host "Enter the number of the printer to add or 'Q' to go back to the main menu"

        if ($printerChoice -eq 'Q') {
            Write-Host "Going back to the main menu."
        }
        elseif ([int]$printerChoice -ge 1 -and [int]$printerChoice -le $departments[$selectedDepartment].Count) {
            $selectedPrinterIndex = [int]$printerChoice - 1
            Add-PrinterOption -department $selectedDepartment -selectedPrinterIndex $selectedPrinterIndex
        }
        else {
            Write-Host "Invalid selection. Please enter a valid number or 'Q' to go back to the main menu."
        }
    }
    else {
        Write-Host "Invalid selection. Please enter a valid number or 'Q' to quit."
    }
}

