# Exchange Virtual Directory Configuration Script

This PowerShell script updates all major Exchange virtual directories on a specified Exchange server, including:

- **OWA** (`/owa`)
- **ECP** (`/ecp`)
- **EWS** (`/EWS/Exchange.asmx`)
- **ActiveSync** (`/Microsoft-Server-ActiveSync`)
- **OAB** (`/OAB`)
- **PowerShell** (`/PowerShell`)
- **MAPI** (`/mapi`) - external URL only

---

## 🔧 Purpose

To standardize and apply a unified base URL (internal and external) across all necessary Exchange services, ensuring consistent access and avoiding client connectivity issues, especially for Outlook clients, mobile devices, and external users.

---

## 💡 Prerequisites

- Exchange Management Shell or PowerShell with Exchange modules
- Administrative rights on the Exchange server
- Valid SSL certificate with SANs covering the base URL (`mail.example.com`)
- DNS A record pointing to the Exchange server for the chosen base URL

---

## ⚙️ Configuration

Before running, update the following variables in the script:

```powershell
$Server = "Konnash-EX-01"
$BaseUrl = "https://mail.konnash.net"
```

---

## ▶️ Usage

1. Copy the script into a `.ps1` file (e.g. `Set-ExchangeVDirs.ps1`).
2. Run the script in **Exchange Management Shell** or **elevated PowerShell**:

```powershell
.\Set-ExchangeVDirs.ps1
```

3. The script will:
    - Apply the URLs to all valid virtual directories
    - Restart IIS remotely to apply changes

---

## 🔍 What the Script Does

- Retrieves all matching virtual directories for the specified server
- Applies the `$BaseUrl` to both `-InternalUrl` and `-ExternalUrl` (except MAPI which uses only `-ExternalUrl`)
- Restarts IIS on the target server using `Invoke-Command`

---

## 🛡️ Example Output

```
✔ OWA updated
✔ ECP updated
✔ EWS updated
✔ ActiveSync updated
✔ OAB updated
✔ PowerShell VDir updated
✔ MAPI Virtual Directory updated

✅ All virtual directories updated successfully on Konnash-EX-01
```

---

## 📌 Important Notes

- **Autodiscover** virtual directory is intentionally excluded because Exchange 2016/2019 no longer supports setting `InternalUrl` or `ExternalUrl` for it.
- If you’re using a load balancer, ensure the external URL resolves correctly and routes traffic to the Exchange server.

---

## 📜 License

This script is provided as-is under the MIT License. Use at your own risk.
