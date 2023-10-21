function cdp { Set-Location C:\Users\[User]\Documents\Code\_p }

function port {
    param(
        [Parameter(Mandatory=$true)]
        [int]$LocalPort
    )

    $processId = (Get-NetTCPConnection -LocalPort $LocalPort).OwningProcess
    Get-Process -Id $processId
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
