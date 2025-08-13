## Step 0: Get Application Client ID from Entra

1. Go to **Entra admin centre**
2. Navigate to **Microsoft Entra ID** > **App registrations**
3. Find and click on the SPFx extension's app registration
4. On the **Overview** page, copy the **Application (client) ID**
   - It's a GUID that looks like `12345678-1234-1234-1234-123456789012`
5. Keep this handy for Step 3

## Step 1: Update API Permissions

1. In the same app registration, go to **API permissions**
2. Find and **Remove** the `Files.ReadWrite.All` permission (click the ... menu > Remove permission)
3. Click **Add a permission** > **Microsoft Graph** > **Application permissions**
4. Search for and select **Sites.Selected**
5. Click **Add permissions**

## Step 2: Grant Admin Consent

1. Still in **API permissions**, click **Grant admin consent for [your tenant]**
2. Confirm when prompted
3. Verify the status shows "✓ Granted for [your tenant]"

## Step 3: Get Site ID in Graph Explorer

**URL:**

```
https://graph.microsoft.com/v1.0/sites/tbdevtenant.sharepoint.com:/sites/soteria
```

**Method:** GET

**Copy from response:** The `id` field (will look like `tbdevtenant.sharepoint.com,{guid},{guid}`)

## Step 4: Grant Site Permission in Graph Explorer

**URL:**

```
https://graph.microsoft.com/v1.0/sites/{site-id-from-step-3}/permissions
```

**Method:** POST

**Request Body:**

```json
{
  "roles": ["read", "write"],
  "grantedToIdentities": [
    {
      "application": {
        "id": "{application-client-id-from-step-0}",
        "displayName": "Soteria+ Connector"
      }
    }
  ]
}
```

## Step 5, Optional: regenerate client secret
