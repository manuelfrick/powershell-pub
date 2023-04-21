# PowerShell-Pub 
PowerShell-Scripts

Follow-SPO-Site prerequisites:

    1. Create a company app registration in Azure Active Directory
    2. Assign the following API-Permissions for GraphAPI:
        - Group.ReadWrite.All
        - Sites.ReadWrite.All
        - User.ReadWrite.All

Replace following lines with your values:

    13  -> Tenant-ID
    14  -> Application-ID
    15  -> Client-Secret
    86  -> Tenant-URL (replace contoso.sharepoint.com with your URL)
    151 -> SharePoint site URL
    157 -> Azure AD Group Name (Hybrid / Cloud Only)
    
