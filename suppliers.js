// src/lib/http.js
//
// Small shared helpers so every function returns consistent JSON responses
// and CORS headers without repeating boilerplate.

function corsHeaders() {
  const origin = process.env.CORS_ALLOWED_ORIGIN || '*';
  return {
    'Access-Control-Allow-Origin': origin,
    'Access-Control-Allow-Methods': 'GET, POST, PUT, PATCH, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    'Content-Type': 'application/json',
  };
}

function json(status, body) {
  return {
    status,
    headers: corsHeaders(),
    jsonBody: body,
  };
}

function ok(body) {
  return json(200, body);
}

function created(body) {
  return json(201, body);
}

function noContent() {
  return { status: 204, headers: corsHeaders() };
}

function badRequest(message) {
  return json(400, { error: message });
}

function notFound(message = 'Not found') {
  return json(404, { error: message });
}

function serverError(err) {
  // Deliberately don't leak internal error detail (credentials, stack
  // traces referencing tenant/app IDs) to the client — log it server-side
  // via context.error in the caller and return a generic message here.
  return json(500, { error: 'Internal server error', detail: String(err && err.message || err) });
}

function preflight() {
  return { status: 204, headers: corsHeaders() };
}

module.exports = { corsHeaders, json, ok, created, noContent, badRequest, notFound, serverError, preflight };
