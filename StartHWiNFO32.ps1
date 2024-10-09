# Path to HWiNFO32.exe
$hwinfoPath = "C:\tools\hwi_804\HWiNFO32.exe"

# Path to NirCmd
$nircmdPath = "C:\tools\nircmd-x64\nircmd.exe"

# Start HWiNFO32.exe without blocking
Start-Process -FilePath $hwinfoPath

# Wait for 5 seconds to allow HWiNFO32.exe to start
Start-Sleep -Seconds 10

# Activate the window with the title "Warning"
Start-Process -FilePath $nircmdPath -ArgumentList "win activate stitle Warning"

# Send the Enter key press
Start-Process -FilePath $nircmdPath -ArgumentList "sendkey enter press"
