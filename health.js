// src/lib/graphClient.js
//
// Wraps Microsoft Graph access to SharePoint Lists behind app-only auth
// (client credentials). No user sign-in is involved yet — the Function App
// itself authenticates to Graph using the app registration's client secret.
//
// SETUP REQUIRED (see docs/azure-setup.md):
//   1. Register an Azure AD app.
//   2. Grant it Sites.Selected (preferred) or Sites.ReadWrite.All application
//      permission, with admin consent.
//   3. If using Sites.Selected, grant it access to the specific SharePoint
//      site via Grant-SPOAzureADAppSitePermission (see docs/sharepoint-schema.md).
//   4. Set SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET, SP_SITE_HOSTNAME,
//      SP_SITE_PATH as Function App settings (or in local.settings.json for
//      local dev).

const { ClientSecretCredential } = require('@azure/identity');
require('isomorphic-fetch');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

let _client = null;
let _siteId = null;

function getEnv(name, required = true) {
  const v = process.env[name];
  if (required && !v) {
    throw new Error(`Missing required app setting: ${name}`);
  }
  return v;
}

/**
 * Returns a cached Graph client authenticated as the Function App itself.
 */
function getGraphClient() {
  if (_client) return _client;

  const tenantId = getEnv('SP_TENANT_ID');
  const clientId = getEnv('SP_CLIENT_ID');
  const clientSecret = getEnv('SP_CLIENT_SECRET');

  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default'],
  });

  _client = Client.initWithMiddleware({ authProvider });
  return _client;
}

/**
 * Resolves and caches the SharePoint site's Graph siteId, e.g.
 * "nouryon.sharepoint.com,<guid>,<guid>".
 */
async function getSiteId() {
  if (_siteId) return _siteId;
  const hostname = getEnv('SP_SITE_HOSTNAME');
  const sitePath = getEnv('SP_SITE_PATH');
  const client = getGraphClient();
  const site = await client.api(`/sites/${hostname}:${sitePath}`).get();
  _siteId = site.id;
  return _siteId;
}

/**
 * Fetches every item from a list, expanding field values.
 * SharePoint paginates at 200 items by default; we page through with
 * @odata.nextLink until exhausted, so this always returns the full list
 * regardless of size.
 */
async function getAllListItems(listName, { filter } = {}) {
  const client = getGraphClient();
  const siteId = await getSiteId();
  let url = `/sites/${siteId}/lists/${listName}/items?expand=fields&$top=200`;
  if (filter) url += `&$filter=${encodeURIComponent(filter)}`;

  let items = [];
  let nextUrl = url;
  while (nextUrl) {
    const res = await client.api(nextUrl).get();
    items = items.concat(res.value || []);
    nextUrl = res['@odata.nextLink']
      ? res['@odata.nextLink'].replace(/^.*\/v1\.0/, '')
      : null;
  }
  return items;
}

async function getListItem(listName, itemId) {
  const client = getGraphClient();
  const siteId = await getSiteId();
  return client.api(`/sites/${siteId}/lists/${listName}/items/${itemId}?expand=fields`).get();
}

async function createListItem(listName, fields) {
  const client = getGraphClient();
  const siteId = await getSiteId();
  return client.api(`/sites/${siteId}/lists/${listName}/items`).post({ fields });
}

async function updateListItem(listName, itemId, fields) {
  const client = getGraphClient();
  const siteId = await getSiteId();
  return client.api(`/sites/${siteId}/lists/${listName}/items/${itemId}/fields`).patch(fields);
}

async function deleteListItem(listName, itemId) {
  const client = getGraphClient();
  const siteId = await getSiteId();
  return client.api(`/sites/${siteId}/lists/${listName}/items/${itemId}`).delete();
}

module.exports = {
  getGraphClient,
  getSiteId,
  getAllListItems,
  getListItem,
  createListItem,
  updateListItem,
  deleteListItem,
};
