This modified PowerShell script replaces hardcoded configuration values with interactive prompts using  Read-Host , including secure input for the client secret. The script now asks for Tenant ID, Client ID, Client Secret (masked), Group ID, and confirmation for deletion interactively each time it runs.
Key Changes
	•	Added prompts at the top for all config values, converting secure string back to plain text for the token request as required by client credentials flow.
	•	Escaped  $  in the members query URL to prevent PowerShell variable expansion.
	•	Implemented optional deletion logic using category IDs fetched during verification, assuming  $DeleteAfterVerify  controls it (set via prompt).
	•	Clears sensitive variables and forces garbage collection at the end for basic security cleanup.
Usage Notes
Run the script in PowerShell; it will prompt sequentially for inputs without exposing secrets in plain text during entry. Ensure your app registration has  Group.Read.All ,  User.Read.All , and  MailboxSettings.ReadWrite  application permissions for Graph API access to groups, users, and Outlook categories. Test in a non-production group first to verify category creation/deletion behavior.