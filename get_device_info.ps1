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

    # Create the output file
    $outputFile = "device_info.txt"
    $fs = New-Object IO.StreamWriter($outputFile, $false)
}

Process {
    # Get the computer name
    $computerName = $env:COMPUTERNAME
    $fs.WriteLine($computerName)

    # Get the MAC Address
    $macAddress = Get-MACAddress
    $fs.WriteLine($macAddress)

    # Get the Serial Number
    $serialNumber = Get-SerialNumber
    $fs.WriteLine($serialNumber + "`n")

    Write-Output "Gathering Monitor(s)..."
    # Query WMI for monitor information
    $query = "SELECT * FROM WmiMonitorID"
    $objWMIService = Get-WmiObject -Query $query -Namespace "root\wmi"

    foreach ($objItem in $objWMIService) {
        $r = "Active: " + $objItem.Active + "`n"
        $r += "InstanceName: " + $objItem.InstanceName + "`n"
        $r += "Manufacturer: " + (ASCIItoString -asciiStr ($objItem.ManufacturerName -join ",")) + "`n"
        $r += "ProductCodeID: " + (ASCIItoString -asciiStr ($objItem.ProductCodeID -join ",")) + "`n"
        $r += "ServiceTag/SerialNumber: " + (ASCIItoString -asciiStr ($objItem.SerialNumberID -join ",")) + "`n"
        $r += "ModelName: " + (ASCIItoString -asciiStr ($objItem.UserFriendlyName -join ",")) + "`n"
        $r += "WeekOfManufacture: " + $objItem.WeekOfManufacture + "`n"
        $r += "YearOfManufacture: " + $objItem.YearOfManufacture + "`n"
        $r += "`n"

        $fs.WriteLine($r)
    }
}

End {
    # Close the file
    $fs.Close()

    Write-Output "System information has been written to $outputFile"
    Pause
}
