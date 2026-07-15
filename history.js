// src/functions/suppliers.js
//
// GET    /api/suppliers            -> list all suppliers
// GET    /api/suppliers/{id}       -> get one supplier
// POST   /api/suppliers            -> create a supplier
// PUT    /api/suppliers/{id}       -> update a supplier (full or partial)
// DELETE /api/suppliers/{id}       -> delete a supplier
//
// Deleting a supplier here does NOT cascade-delete its meetings/actions/
// history rows in SharePoint (Graph doesn't do cascades, and silently
// deleting a supplier's whole history is a bad default for an audit-minded
// SRM tool). Orphaned child rows are flagged instead — see
// docs/data-mapping.md "Deletion behavior".

const { app } = require('@azure/functions');
const { getAllListItems, getListItem, createListItem, updateListItem, deleteListItem } = require('../lib/graphClient');
const { supplierFromSP, supplierToSP } = require('../lib/mappers');
const { ok, created, noContent, badRequest, notFound, serverError, preflight } = require('../lib/http');

const LIST = () => process.env.SUPPLIERS_LIST_NAME || 'Suppliers';

app.http('suppliers', {
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'suppliers/{id?}',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return preflight();

    const id = request.params.id;

    try {
      if (request.method === 'GET') {
        if (id) {
          const item = await getListItem(LIST(), id);
          if (!item) return notFound('Supplier not found');
          return ok(supplierFromSP(item));
        }
        const items = await getAllListItems(LIST());
        return ok(items.map(supplierFromSP));
      }

      if (request.method === 'POST') {
        const body = await request.json();
        if (!body || !body.n) return badRequest('Field "n" (supplier name) is required.');
        const fields = supplierToSP(body);
        const item = await createListItem(LIST(), fields);
        return created(supplierFromSP(item));
      }

      if (request.method === 'PUT') {
        if (!id) return badRequest('Supplier id is required in the URL for updates.');
        const body = await request.json();
        const fields = supplierToSP(body);
        await updateListItem(LIST(), id, fields);
        const item = await getListItem(LIST(), id);
        return ok(supplierFromSP(item));
      }

      if (request.method === 'DELETE') {
        if (!id) return badRequest('Supplier id is required in the URL for deletion.');
        await deleteListItem(LIST(), id);
        return noContent();
      }

      return badRequest('Unsupported method');
    } catch (err) {
      context.error('suppliers function error', err);
      return serverError(err);
    }
  },
});
