# M365 Offboarding Monster

Enterprise-grade Microsoft 365 + Active Directory employee offboarding automation.

## Features

- Disable Active Directory account
- Move user to disabled OU
- Remove privileged group memberships
- Disable Microsoft Entra sign-in
- Revoke active M365 sessions
- Convert mailbox to shared mailbox
- Hide mailbox from GAL
- Configure email forwarding
- Grant manager mailbox access
- Remove M365 licenses
- Detailed logging
- WhatIf mode support

## Requirements

```powershell
Install-Module ActiveDirectory
Install-Module Microsoft.Graph.Users
Install-Module Microsoft.Graph.Identity.DirectoryManagement
Install-Module ExchangeOnlineManagement
```

## Usage

```powershell

.\M365-Offboarding-Monster.ps1
```

## Safe Test Mode

```powershell
.\M365-Offboarding-Monster.ps1 -WhatIfMode
```

## Notes

Requires:
- Exchange Admin
- User Administrator
- Active Directory permissions
