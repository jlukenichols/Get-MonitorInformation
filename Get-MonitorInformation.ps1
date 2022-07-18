[String]$ComputerName = "localhost"

#Start the WinRM service if it isn't already running
Start-Service "WinRM" -ErrorAction SilentlyContinue | Out-Null    

foreach ($Computer in $ComputerName.Split(",").Trim()) {
        Get-CimInstance -ComputerName $Computer -ClassName WmiMonitorID -Namespace root\wmi | Foreach-Object {
            if ((($_.ManufacturerName | ForEach-Object { [char]$_ }) -join "") -eq "DEL") {
                $ManufacturerName = "Dell"
            } elseif ((($_.ManufacturerName | ForEach-Object { [char]$_ }) -join "") -eq "ACI") {
                $ManufacturerName = "ASUS"
            } elseif ((($_.ManufacturerName | ForEach-Object { [char]$_ }) -join "") -eq "SEC") {
                $ManufacturerName = "Epson"
            } elseif ((($_.ManufacturerName | ForEach-Object { [char]$_ }) -join "") -eq "ACR") {
                $ManufacturerName = "Acer"
            } elseif ((($_.ManufacturerName | ForEach-Object { [char]$_ }) -join "") -eq "UGD") {
                $ManufacturerName = "XP-PEN"
            } elseif ((($_.ManufacturerName | ForEach-Object { [char]$_ }) -join "") -eq "SAM") {
                $ManufacturerName = "Samsung"
            } else {
                $ManufacturerName = ($_.ManufacturerName | ForEach-Object { [char]$_ }) -join ""
            }
            [PSCustomObject]@{
                'Active'              = $_.Active
                'Manufacturer'        = $ManufacturerName
                'Model'               = ($_.UserFriendlyName | ForEach-Object { [char]$_ }) -join ""
                'Serial Number'       = ($_.SerialNumberID | ForEach-Object { [char]$_ }) -join ""
                'Year Of Manufacture' = $_.YearOfManufacture
                'Week Of Manufacture' = $_.WeekOfManufacture
            }
        }        
}