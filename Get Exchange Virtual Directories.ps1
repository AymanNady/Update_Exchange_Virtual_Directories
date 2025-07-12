# ===============================
# Get Exchange Virtual Directories
# ===============================

$Server = "Konnash-EX-01"

# OWA
Get-OwaVirtualDirectory -Server $Server | FL Name,InternalUrl,ExternalUrl

# ECP
Get-EcpVirtualDirectory -Server $Server | FL Name,InternalUrl,ExternalUrl

# EWS
Get-WebServicesVirtualDirectory -Server $Server | FL Name,InternalUrl,ExternalUrl

# ActiveSync
Get-ActiveSyncVirtualDirectory -Server $Server | FL Name,InternalUrl,ExternalUrl

# OAB
Get-OabVirtualDirectory -Server $Server | FL Name,InternalUrl,ExternalUrl

# PowerShell
Get-PowerShellVirtualDirectory -Server $Server | FL Name,InternalUrl,ExternalUrl

# MAPI
Get-MapiVirtualDirectory -Server $Server | FL Name,InternalUrl,ExternalUrl
