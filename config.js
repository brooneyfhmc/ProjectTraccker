// ─── CONFIGURATION ───────────────────────────────────────────────────────────
// After registering your Azure AD app, paste the values here.
// See README.md for step-by-step registration instructions.

const CONFIG = {
  // Azure AD App Registration values
  clientId:    "f7991857-dc4b-4a3b-a417-5479e4098691",       // Application (client) ID
  tenantId:    "5b3538ca-e4fc-4828-a305-8e9c303853bd",       // Directory (tenant) ID

  // SharePoint site & list
  sharePointHost: "gofirsthome.sharepoint.com",
  sitePath:       "/sites/Home",
  listName:       "First Home Project Tracker",

  // Column internal names in the SharePoint list.
  // If updates fail, open: <site>/Lists/<list>/_api/web/lists/getbytitle('<list>')/fields
  // and find the InternalName for "Work Item" and "Status Notes".
  workItemField:   "Work_x0020_Item",       // Display name: "Work Item"
  statusNotesField:"Status_x0020_Notes",    // Display name: "Status Notes"

  // SSO scope URI — must match the Application ID URI you set in "Expose an API"
  // Format: api://<your-static-web-app-domain>/<client-id>
  ssoScope: `api://salmon-forest-08b7a6d0f.7.azurestaticapps.net/f7991857-dc4b-4a3b-a417-5479e4098691/access_as_user`,

  // Microsoft Graph scopes needed
  scopes: [
    "https://graph.microsoft.com/Sites.ReadWrite.All",
    "User.Read",
  ],
};
