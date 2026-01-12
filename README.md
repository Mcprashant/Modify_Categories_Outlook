# Microsoft 365 Outlook Category Automation

This PowerShell script automates the creation, verification, and optional deletion of **Outlook master categories** for all users in a specified Microsoft 365 group.[1]
It uses **Microsoft Graph API** with **client credentials authentication** and provides an interactive configuration experience directly in PowerShell.[2]

## ⚙️ Features

- Interactive input for all required configuration parameters  
- Secure handling of client secrets using `SecureString`  
- Automatic authentication against Microsoft Graph  
- Pagination support for fetching all group members  
- Automated category creation with customizable names and colors  
- Verification of categories per user  
- Optional cleanup (delete created categories after verification)[3]

## 🧰 Prerequisites

Ensure the following requirements are met before running:

- **PowerShell 5.1** or higher (or PowerShell Core)  
- A **registered Azure AD/Entra ID App** with **Application permissions**:
  - `Group.Read.All`[4]
  - `User.Read.All`
  - `MailboxSettings.ReadWrite`[5]
- **Admin consent** granted for the permissions[6]
- Microsoft Graph API accessible from your environment[7]

## 🚀 Quick Start

### 1. Clone Repository

```bash
git clone https://github.com/mcprashant/Modify_Categories_Outlook.git
cd Modify_Categories_Outlook
```

### 2. Run Script

```powershell
.\Powershell_Add_List_Verify_Group_Interactive.ps1
```

### 3. Enter Configuration

Provide values when prompted:
- **Tenant ID**
- **Client ID**  
- **Client Secret** (secure input)
- **Group ID**
- **Delete after verify? (y/n)**

## 🎨 Default Categories

| Display Name        | Color     |
|---------------------|-----------|
| Vergadering intern  | preset3   |
| Vergadering extern  | preset7   |
| Focuswerk           | preset10  |
| Afwezigheid         | preset5   |

Edit the `$Categories` array to customize.[8]

## 🔒 Security & Cleanup

Client secrets clear from memory automatically via garbage collection.  
Choose deletion option to remove test categories post-verification.[1]

## ⚠️ Important Notes

- Test in **non-production** first
- Categories apply to emails, calendar, tasks[8]
- Provided **as-is** – no warranty
