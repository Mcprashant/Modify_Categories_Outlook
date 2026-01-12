***

# Microsoft 365 Outlook Category Automation

This PowerShell script automates the creation, verification, and optional deletion of **Outlook master categories** for all users in a specified Microsoft 365 group.  
It uses **Microsoft Graph API** with **client credentials authentication** and provides an interactive configuration experience directly in PowerShell.

***

## ⚙️ Features

- Interactive input for all required configuration parameters  
- Secure handling of client secrets using `SecureString`  
- Automatic authentication against Microsoft Graph  
- Pagination support for fetching all group members  
- Automated category creation with customizable names and colors  
- Verification of categories per user  
- Optional cleanup (delete created categories after verification)

***

## 🧰 Prerequisites

Before running the script, ensure the following requirements are met:

- **PowerShell 5.1** or higher (or PowerShell Core)
- A **registered Azure AD App** with the following Graph API **Application permissions**:
  - `Group.Read.All`
  - `User.Read.All`
  - `MailboxSettings.ReadWrite`
- **Admin consent** granted for the above permissions
- The **Microsoft Graph API** must be accessible from your environment

***

## 🚀 Usage

### 1. Clone or Download

Download or clone this repository:

```bash
git clone https://github.com/<yourusername>/OutlookCategoryManager.git
cd OutlookCategoryManager
```

### 2. Run the Script

Run the script in PowerShell:

```powershell
.\Create-OutlookCategories.ps1
```

### 3. Provide Required Information

When prompted, enter:

- **Tenant ID** – Azure AD tenant ID  
- **Client ID** – Registered application ID  
- **Client Secret** – Securely entered value (won’t display)  
- **Group ID** – Target Microsoft 365 group ID  
- **Delete after verify? (y/n)** – Optional cleanup of created categories  

Example:

```
Enter Tenant ID: 00000000-0000-0000-0000-000000000000
Enter Client ID: 11111111-1111-1111-1111-111111111111
Enter Client Secret: ******
Enter Group ID: 22222222-2222-2222-2222-222222222222
Delete after verify? (y/n): n
```

***

## 🎨 Default Categories

The script creates the following Outlook master categories by default:

| Display Name        | Color     |
|----------------------|-----------|
| Vergadering intern   | preset3   |
| Vergadering extern   | preset7   |
| Focuswerk            | preset10  |
| Afwezigheid          | preset5   |

You can modify or extend this list in the `$Categories` array at the beginning of the script.

***

## 🔒 Security Notes

- The `Client Secret` is converted securely from `SecureString` and cleared from memory at the end of execution.  
- Avoid storing credentials directly in the script or repository.  
- Use **least privilege** when assigning Graph API permissions.

***

## 🧹 Cleanup Behavior

If you choose **“Delete after verify”**, the script removes all created categories after verifying them.  
Otherwise, the new categories remain visible in users’ Outlook.

***

## ⚠️ Disclaimer

This script is provided **as-is** without warranty of any kind.  
Use at your own risk, and test carefully in a **non-production tenant** before applying changes to live environments.

***
