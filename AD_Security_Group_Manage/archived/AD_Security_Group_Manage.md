# AD Security Group Manager V2.4 - Quick Summary
        
**Script Name:** `AD_Security_Group_Checker_V.2.4.ps1`
**Author:** Albert Ng
**Type:** PowerShell GUI Application (WinForms)
        
## Prerequisites
- **Active Directory Module (RSAT)** is required.
- Automatically requests **Administrator privileges**.
- Warns users if they are not part of the support group: `GRP-RBC-R-ASI-WSP-ClientSideSupport`.
        
## Key Features (Tabs)
The application is split into four main functional tabs:
        
1. **ACL Viewer**
    - **Purpose:** Analyze folder permissions.
    - **Action:** Browse a folder path, click "Check Permissions", and view the owner and Access Control List (ACL) entries. Supports copying values (e.g., Owner, Group Name) via right-click menu.
        
2. **Access Checker**
    - **Purpose:** Check effective permissions for a specific user on a folder.
    - **Action:** Enter a folder path and a user's Email or SAM account name. Shows their access rights directly and via group memberships.
        
3. **Audit Members**
    - **Purpose:** Review Active Directory group memberships.
    - **Action:** Enter a Group Name or Email, fetch members (including nested sub-groups), and export the list to CSV. Allows right-clicking members to send them directly to the "Bulk Updates" tab.
        
4. **Bulk Updates**
    - **Purpose:** Bulk add or remove users from a target Security Group.
    - **Action:** Enter a target group, upload a CSV/TXT of users (or use imported ones from Audit tab), and perform "Bulk ADD" or "Bulk REMOVE" operations. Generates an execution log.
    
    
    
