This updated PowerShell script for deleting Outlook master categories now uses interactive prompts for all configuration details, matching the style of the previous create script.
Key Changes
	•	Replaced hardcoded values with  Read-Host  prompts for Tenant ID, Client ID (masked secure input for secret), and Group ID.
	•	Added a loop to interactively collect category names to delete, one per line, ending with a blank input; validates at least one category is provided.
	•	Fixed URL escaping in the members query with  $ select.
	•	Included memory cleanup for sensitive data at the end.
Usage Notes
Run the script and provide inputs when prompted; category names must match exactly (case-sensitive) to existing  displayName  values. Requires Graph permissions:  Group.Read.All ,  User.Read.All ,  MailboxSettings.ReadWrite  (application scope). Use alongside the create script for full add/verify/delete workflow on M365 group members.