# FastPM Helper Add-in - Microsoft Graph Authentication Setup

## Azure App Registration Requirements

This add-in uses Microsoft Graph API with delegated authentication (user login). Follow these steps to configure Azure AD:

### 1. Create App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations** → **New registration**
3. Configure:
   - **Name**: FastPM Helper Outlook Add-in
   - **Supported account types**: Accounts in this organizational directory only (Single tenant)
   - **Redirect URI**:
     - Platform: **Public client/native (mobile & desktop)**
     - URI: `http://localhost`

### 2. Configure API Permissions

1. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
2. Add these permissions:
   - `User.Read` (Sign in and read user profile)
   - `Sites.ReadWrite.All` (Read and write items in all site collections)
3. Click **Add permissions**
4. Request admin consent if required by your organization

### 3. Enable Public Client Flows

1. Go to **Authentication**
2. Scroll to **Advanced settings** → **Allow public client flows**
3. Set toggle to **Yes**
4. Click **Save**

### 4. Get Configuration Values

1. Go to **Overview** tab
2. Copy these values to your `.env` file:
   - **Application (client) ID** → `AZURE_CLIENT_ID`
   - **Directory (tenant) ID** → `AZURE_TENANT_ID`

### 5. Configure .env File

Create/update `bin\Debug\.env` file:

```env
GEMINI_API_KEY=your_gemini_api_key
SHAREPOINT_SITE_URL=https://yourcompany.sharepoint.com/personal/yourname
AZURE_TENANT_ID=your-tenant-id-here
AZURE_CLIENT_ID=your-client-id-here
```

### 6. First Run Authentication

On first launch:
1. The add-in will open a browser window for Microsoft login
2. Sign in with your organizational account
3. Consent to the requested permissions
4. Token will be cached for future use (no re-login required)

### Token Cache Location

MSAL stores tokens securely in: `%LOCALAPPDATA%\.IdentityService`

### Troubleshooting

**Error: AADSTS50194 (Application not configured as multi-tenant)**
- Ensure "Supported account types" is set correctly for your organization

**Error: AADSTS65001 (User or administrator has not consented)**
- Request admin consent for API permissions in Azure portal

**Error: Authentication dialog doesn't appear**
- Check Outlook window handle capture in GraphAuthService.cs
- Verify redirect URI is exactly `http://localhost`

**Error: Failed to resolve SharePoint site**
- Verify SHAREPOINT_SITE_URL points to site root, not a list URL
- Check Sites.ReadWrite.All permission is granted
