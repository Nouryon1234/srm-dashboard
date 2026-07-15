// src/functions/health.js
//
// GET /api/health
//
// Verifies the Function App can authenticate to Graph and reach the
// configured SharePoint site. Useful for uptime checks and for diagnosing
// "is it the API or is it SharePoint" during setup.

const { app } = require('@azure/functions');
const { getSiteId } = require('../lib/graphClient');
const { ok, json, preflight } = require('../lib/http');

app.http('health', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'health',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return preflight();

    try {
      const siteId = await getSiteId();
      return ok({ status: 'ok', sharePointConnected: true, siteId });
    } catch (err) {
      context.error('health check failed', err);
      return json(503, { status: 'degraded', sharePointConnected: false, error: String(err.message || err) });
    }
  },
});
