// src/functions/meetings.js
//
// GET    /api/meetings                    -> all meetings (optionally ?supplierId=)
// GET    /api/meetings/{id}               -> one meeting
// POST   /api/meetings                    -> create a meeting
// PUT    /api/meetings/{id}                -> update a meeting
// DELETE /api/meetings/{id}                -> delete a meeting
//
// Creating/updating a meeting with an `act` (action items, newline separated)
// field does NOT automatically materialize rows in the Actions list from this
// endpoint alone — that sync step is explicit via POST /api/actions/materialize
// so the frontend can show the user what will be created before it happens,
// and so re-running it is idempotent rather than a hidden side effect buried
// in an unrelated save call.

const { app } = require('@azure/functions');
const { getAllListItems, getListItem, createListItem, updateListItem, deleteListItem } = require('../lib/graphClient');
const { meetingFromSP, meetingToSP } = require('../lib/mappers');
const { ok, created, noContent, badRequest, notFound, serverError, preflight } = require('../lib/http');

const LIST = () => process.env.MEETINGS_LIST_NAME || 'Meetings';

app.http('meetings', {
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'meetings/{id?}',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return preflight();

    const id = request.params.id;
    const supplierId = request.query.get('supplierId');

    try {
      if (request.method === 'GET') {
        if (id) {
          const item = await getListItem(LIST(), id);
          if (!item) return notFound('Meeting not found');
          return ok(meetingFromSP(item));
        }
        let items = await getAllListItems(LIST());
        let records = items.map(meetingFromSP);
        if (supplierId) records = records.filter((r) => String(r.sid) === String(supplierId));
        records.sort((a, b) => new Date(b.date) - new Date(a.date));
        return ok(records);
      }

      if (request.method === 'POST') {
        const body = await request.json();
        if (!body || body.sid == null || !body.date) {
          return badRequest('Fields "sid" (supplier id) and "date" are required.');
        }
        const fields = meetingToSP(body);
        const item = await createListItem(LIST(), fields);
        return created(meetingFromSP(item));
      }

      if (request.method === 'PUT') {
        if (!id) return badRequest('Meeting id is required in the URL for updates.');
        const body = await request.json();
        const fields = meetingToSP(body);
        await updateListItem(LIST(), id, fields);
        const item = await getListItem(LIST(), id);
        return ok(meetingFromSP(item));
      }

      if (request.method === 'DELETE') {
        if (!id) return badRequest('Meeting id is required in the URL for deletion.');
        await deleteListItem(LIST(), id);
        return noContent();
      }

      return badRequest('Unsupported method');
    } catch (err) {
      context.error('meetings function error', err);
      return serverError(err);
    }
  },
});
