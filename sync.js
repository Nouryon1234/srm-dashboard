# Azure Setup Guide

Step-by-step setup to get the live dashboard running end-to-end: SharePoint site → Azure AD app registration → Function App → Static Web App.

## 1. Create the SharePoint site and lists

1. Create (or reuse) a SharePoint site, e.g. `https://nouryon.sharepoint.com/sites/SRM-Dashboard`.
2. Create the six lists described in `docs/sharepoint-schema.md`, with the exact internal column names given there. Column names are what the API's field mappers (`api/src/lib/mappers.js`) expect — renaming them breaks the mapping.

## 2. Register an Azure AD app for the API

In the Azure Portal → **Microsoft Entra ID → App registrations → New registration**:

1. Name: `SRM Dashboard API` (or similar).
2. Supported account types: single tenant.
3. No redirect URI needed (this is a daemon/service app, not an interactive sign-in app).
4. After creation, note the **Application (client) ID** and **Directory (tenant) ID**.
5. Go to **Certificates & secrets → New client secret**. Copy the secret value immediately — it's only shown once.

### Grant Graph permissions

Still on the app registration:

1. **API permissions → Add a permission → Microsoft Graph → Application permissions**.
2. Add **Sites.Selected**. (Preferred over `Sites.ReadWrite.All`, which grants write access to every SharePoint site in the tenant — `Sites.Selected` scopes access to just the one site you choose in step 3.)
3. Click **Grant admin consent** (requires a tenant admin).

### Scope the app to just the SRM site

`Sites.Selected` grants nothing by itself — you still need to explicitly authorize the app on the one site. An admin with SharePoint PowerShell access runs:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser  # if not already installed
Connect-PnPOnline -Url "https://nouryon.sharepoint.com/sites/SRM-Dashboard" -Interactive
Grant-PnPAzureADAppSitePermission -AppId "<client-id-from-step-1>" -DisplayName "SRM Dashboard API" -Site "https://nouryon.sharepoint.com/sites/SRM-Dashboard" -Permissions Write
```

(Or the SharePoint Online Management Shell equivalent, `Grant-SPOAzureADAppSitePermission`, referenced in `docs/sharepoint-schema.md`.)

## 3. Create the Azure Function App

1. Azure Portal → **Create a resource → Function App**.
2. Runtime stack: **Node.js 18 LTS** (or newer), Hosting: **Consumption** (serverless) plan is fine for this workload.
3. Once created, go to **Configuration → Application settings** and add:

   | Name | Value |
   |---|---|
   | `SP_TENANT_ID` | tenant ID from step 2 |
   | `SP_CLIENT_ID` | client ID from step 2 |
   | `SP_CLIENT_SECRET` | client secret from step 2 |
   | `SP_SITE_HOSTNAME` | `nouryon.sharepoint.com` |
   | `SP_SITE_PATH` | `/sites/SRM-Dashboard` |
   | `SUPPLIERS_LIST_NAME` | `Suppliers` |
   | `KPI_LIST_NAME` | `KPIRecords` |
   | `MEETINGS_LIST_NAME` | `Meetings` |
   | `ACTIONS_LIST_NAME` | `Actions` |
   | `RISKS_LIST_NAME` | `Actions` |
   | `HISTORY_LIST_NAME` | `PerformanceHistory` |
   | `CORS_ALLOWED_ORIGIN` | your frontend's URL once deployed, or `*` while testing |

   These correspond 1:1 to `api/local.settings.json.example` — copy that file to `api/local.settings.json` for local development (never commit it) and mirror the same keys into the Function App's settings for the deployed environment.

4. **Never store secrets in source control.** For production, prefer pulling `SP_CLIENT_SECRET` from **Azure Key Vault** via a Key Vault reference in the Function App's settings, rather than pasting the raw secret into the portal.

### Deploy the API code

From the `api/` folder:

```bash
npm install
npm install -g azure-functions-core-tools@4
func azure functionapp publish <your-function-app-name>
```

Or connect the Function App to a GitHub repo under **Deployment Center** for CI/CD.

### Verify it's working

```bash
curl https://<your-function-app-name>.azurewebsites.net/api/health
```

Should return `{"status":"ok","sharePointConnected":true,"siteId":"..."}`. If it returns a 503 or an auth error, double check the Graph permission grant (step 2) and that admin consent was actually clicked (not just requested).

## 4. Deploy the frontend to Azure Static Web Apps

1. Azure Portal → **Create a resource → Static Web App**.
2. Connect it to your GitHub repo (or deploy via the Static Web Apps CLI for a non-Git deploy).
3. Build details:
   - App location: `/frontend`
   - Api location: `/api` (this links the Function App you built in step 3 as the SWA's managed API — SWA will proxy `/api/*` calls to it automatically, so `js/api.js`'s default `apiBase: '/api'` just works with no CORS configuration needed)
   - Output location: *(leave blank — this is a static site, no build step)*
4. If you deployed the Function App separately (unmanaged, not linked via SWA), instead set `window.SRM_CONFIG.apiBase` in `frontend/index.html` to the full Function App URL (e.g. `https://your-function-app.azurewebsites.net/api`) and make sure `CORS_ALLOWED_ORIGIN` on the Function App matches your Static Web App's URL.

### Verify end-to-end

Open the Static Web App's URL. The header should show **● Live** within a few seconds (confirming `js/sync.js` successfully hit `/api/dashboard`). If it shows **● Offline**, open the browser console — the failed request's URL and status will tell you whether it's a CORS issue, a wrong `apiBase`, or the Function App itself failing (check `/api/health` directly).

## 5. Ongoing data entry

With no user authentication yet (per the current scope), anyone who can load the site can create/edit/delete data via the dashboard's UI, and anyone with SharePoint edit access to the site can also edit the lists directly. Both paths write to the same lists, so they stay in sync — SharePoint's own list UI works as a fallback data-entry method if the dashboard is ever down.

## 6. Adding authentication later

When you're ready to add sign-in (see the "no auth yet" note in this project's original scope), Azure Static Web Apps has built-in support for Microsoft Entra ID via its `/​.auth/login/aad` routes and a `staticwebapp.config.json` `routes` block requiring an authenticated role — at that point the Function App can also switch from validating nothing to validating the SWA-issued `x-ms-client-principal` header on each request. That's a separate, scoped follow-up; flag it explicitly when you want to start that work since it changes the API's auth model, not just the frontend.
