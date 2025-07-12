# ===============================
# Update Exchange Virtual Directories
# ===============================

$Server = "Konnash-EX-01"
$BaseUrl = "https://mail.konnash.net"

Write-Host "Updating Exchange virtual directories on server: $Server"
Write-Host "Using base URL: $BaseUrl"
Write-Host "---------------------------------------------------"

try {
    # OWA
    Get-OwaVirtualDirectory -Server $Server | ForEach-Object {
        Set-OwaVirtualDirectory -Identity $_.Identity -InternalUrl "$BaseUrl/owa" -ExternalUrl "$BaseUrl/owa"
        Write-Host "✔ OWA updated"
    }

    # ECP
    Get-EcpVirtualDirectory -Server $Server | ForEach-Object {
        Set-EcpVirtualDirectory -Identity $_.Identity -InternalUrl "$BaseUrl/ecp" -ExternalUrl "$BaseUrl/ecp"
        Write-Host "✔ ECP updated"
    }

    # EWS
    Get-WebServicesVirtualDirectory -Server $Server | ForEach-Object {
        Set-WebServicesVirtualDirectory -Identity $_.Identity -InternalUrl "$BaseUrl/EWS/Exchange.asmx" -ExternalUrl "$BaseUrl/EWS/Exchange.asmx"
        Write-Host "✔ EWS updated"
    }

    # ActiveSync
    Get-ActiveSyncVirtualDirectory -Server $Server | ForEach-Object {
        Set-ActiveSyncVirtualDirectory -Identity $_.Identity -InternalUrl "$BaseUrl/Microsoft-Server-ActiveSync" -ExternalUrl "$BaseUrl/Microsoft-Server-ActiveSync"
        Write-Host "✔ ActiveSync updated"
    }

    # OAB
    Get-OabVirtualDirectory -Server $Server | ForEach-Object {
        Set-OabVirtualDirectory -Identity $_.Identity -InternalUrl "$BaseUrl/OAB" -ExternalUrl "$BaseUrl/OAB"
        Write-Host "✔ OAB updated"
    }

    # PowerShell VDir
    Get-PowerShellVirtualDirectory -Server $Server | ForEach-Object {
        Set-PowerShellVirtualDirectory -Identity $_.Identity -InternalUrl "$BaseUrl/PowerShell" -ExternalUrl "$BaseUrl/PowerShell"
        Write-Host "✔ PowerShell VDir updated"
    }

    # MAPI VDir (ExternalUrl only)
    Get-MapiVirtualDirectory -Server $Server | ForEach-Object {
        Set-MapiVirtualDirectory -Identity $_.Identity -InternalUrl "$BaseUrl/mapi" -ExternalUrl "$BaseUrl/mapi"
        Write-Host "✔ MAPI Virtual Directory updated"
    }

    Write-Host "`n✅ All virtual directories updated successfully on $Server"
    Write-Host "Restarting IIS to apply changes..."
    Invoke-Command -ComputerName $Server -ScriptBlock { iisreset }

} catch {
    Write-Error "❌ Error occurred: $_"
}
