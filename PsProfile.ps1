function cdp { Set-Location C:\Users\[User]\Documents\Code\_p }

function Get-PortProcess {
    param(
        [Parameter(Mandatory=$true)]
        [int]$LocalPort,
        
        [Alias("--with-usage")]
        [switch]$WithUsage
    )

    $processId = netstat -ano | 
        Select-String ":$LocalPort\s+" | 
        ForEach-Object {
            if ($_ -match '\s+(\d+)\s*$') {
                $matches[1]
            }
        }

    if ($processId) {
        $process = Get-Process -Id $processId
        $output = $process | Select-Object @{Name='Port'; Expression={$LocalPort}},  # Port Number
                                          @{Name='PID'; Expression={$_.Id}},  # Process ID
                                          @{Name='ProcessName'; Expression={$_.ProcessName}}  # Process Name

        if ($WithUsage) {
            $cpuCounter = New-Object System.Diagnostics.PerformanceCounter("Process", "% Processor Time", $process.ProcessName, $true)
            $ramCounter = New-Object System.Diagnostics.PerformanceCounter("Process", "Working Set - Private", $process.ProcessName, $true)
            
            # It may take a moment for the counters to provide accurate data
            Start-Sleep -Seconds 1
            
            $cpuUsage = $cpuCounter.NextValue() / [Environment]::ProcessorCount
            $ramUsage = $ramCounter.NextValue() / 1MB

            $output | Add-Member -MemberType NoteProperty -Name 'CPU Usage (%)' -Value ([math]::Round($cpuUsage, 2))
            $output | Add-Member -MemberType NoteProperty -Name 'RAM Usage (MB)' -Value ([math]::Round($ramUsage, 2))
        }
        
        $output | Format-Table -AutoSize
    } else {
        Write-Host "No process found on port $LocalPort"
    }
}

function unzip {
    param (
        [Parameter(Mandatory=$true)]
        [string]$zipFilePath,
        
        [Parameter(Mandatory=$false)]
        [switch]$Force
    )

    # Get the full path of the zip file in case a relative path is provided
    $fullZipPath = Resolve-Path $zipFilePath

    # Get the directory of the zip file
    $zipDirectory = [System.IO.Path]::GetDirectoryName($fullZipPath)

    # Get the name of the zip file without the extension
    $folderName = [System.IO.Path]::GetFileNameWithoutExtension($fullZipPath)

    # Create the full path of the new folder
    $newFolderPath = Join-Path -Path $zipDirectory -ChildPath $folderName

    # Check if the folder already exists
    if (Test-Path $newFolderPath) {
        if ($Force) {
            # If -Force is specified, remove the existing folder
            Remove-Item -Path $newFolderPath -Recurse -Force
        }
        else {
            # If -Force is not specified and the folder exists, throw an exception and exit
            throw "Folder '$newFolderPath' already exists. Use -Force to overwrite."
        }
    }

    # Create the new folder
    New-Item -Path $newFolderPath -ItemType Directory -Force

    # Extract the zip file to the new folder
    Expand-Archive -Path $fullZipPath -DestinationPath $newFolderPath
}

function vs {
    param (
        [string]$path = "."
    )

    $directory = Get-Item -Path $path
    $slnFile = Get-ChildItem -Path $directory.FullName -Filter *.sln -Recurse -Depth 0 | Select-Object -First 1
    $csprojFile = Get-ChildItem -Path $directory.FullName -Filter *.csproj -Recurse -Depth 0 | Select-Object -First 1

    if ($null -ne $slnFile) {
        Start-Process "devenv" -ArgumentList  "`"$($slnFile.FullName)`""
    }
    elseif ($null -ne $csprojFile) {
        Start-Process "devenv" -ArgumentList "`"$($csprojFile.FullName)`""
    }
    else {
        Start-Process "devenv" -ArgumentList "`"$($directory.FullName)`""
    }
}
