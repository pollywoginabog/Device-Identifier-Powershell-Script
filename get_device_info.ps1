<#
    .NOTES
    ===========================================================================
     Created on:   	6/20/2024
     Created by:   	Rusty Harker
     Filename:     	get_device_info.ps1
    ===========================================================================
    .DESCRIPTION
    	Queries WMI for monitor serial numbers and general identifying device information.
#>

[CmdletBinding()]
param ()

Begin {
    # Function to convert ASCII codes to a string
    function ASCIItoString($asciiStr) {
        $charArray = $asciiStr -split ","
        $string = ""
        foreach ($char in $charArray) {
            if ([int]$char -ne 0) {  # Ignore NULL characters (ASCII 0)
                $string += [char][int]$char
            }
        }
        return $string
    }

    # Function to get the MAC Address
    function Get-MACAddress {
        $mac = Get-WmiObject -Class Win32_NetworkAdapter |
               Where-Object { $_.NetEnabled -eq $true } |
               Select-Object -First 1 -ExpandProperty MACAddress
        return $mac
    }

    # Function to get the Serial Number
    function Get-SerialNumber {
        $serial = Get-WmiObject -Class Win32_BIOS |
                  Select-Object -ExpandProperty SerialNumber
        return $serial
    }

    Write-Output "Gathering System Info..."

    # Create the output list
    $outputList = @()
}

Process {
    # Get the computer name
    $computerName = $env:COMPUTERNAME

    # Get the MAC Address
    $macAddress = Get-MACAddress

    # Get the Serial Number
    $serialNumber = Get-SerialNumber

    # Add system info to the output list
    $outputList += [PSCustomObject]@{
        Type = "System"
        ComputerName = $computerName
        MACAddress = $macAddress
        SerialNumber = $serialNumber
        Active = ""
        Manufacturer = ""
        ProductCodeID = ""
        ServiceTag_SerialNumber = ""
        ModelName = ""
        WeekOfManufacture = ""
        YearOfManufacture = ""
    }

    Write-Output "Gathering Monitor(s)..."
    # Query WMI for monitor information
    $query = "SELECT * FROM WmiMonitorID"
    $objWMIService = Get-WmiObject -Query $query -Namespace "root\wmi"
    
    foreach ($objItem in $objWMIService) {
       $counter = $counter + 1
        $outputList += [PSCustomObject]@{
            Type = "Monitor " + $counter
            ComputerName = ""
            MACAddress = ""
            SerialNumber = ""
            Active = $objItem.Active
            Manufacturer = (ASCIItoString -asciiStr ($objItem.ManufacturerName -join ","))
            ProductCodeID = (ASCIItoString -asciiStr ($objItem.ProductCodeID -join ","))
            ServiceTag_SerialNumber = (ASCIItoString -asciiStr ($objItem.SerialNumberID -join ","))
            ModelName = (ASCIItoString -asciiStr ($objItem.UserFriendlyName -join ","))
            WeekOfManufacture = $objItem.WeekOfManufacture
            YearOfManufacture = $objItem.YearOfManufacture
            
        }
    }
}

End {
    # Export the list to a CSV file
    $outputFile = "device_info.csv"
    if (Test-Path $outputFile)
			{
				$outputList += Import-CSV -Path $outputFile
			}
        $outputList | Export-Csv -Path $outputFile -NoTypeInformation
    
    
    

    Write-Output "System information has been written to $outputFile"
    Pause
}
