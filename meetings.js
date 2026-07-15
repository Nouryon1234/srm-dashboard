// src/lib/mappers.js
//
// Converts between SharePoint list "fields" (flat, string-heavy, with
// JSON-in-a-text-column for anything array/object shaped) and the clean
// JSON shapes the frontend already expects (matching the original
// hardcoded S[], KPIS{}, NS{}, RISKS{} structures from the static version
// of the dashboard).
//
// Keeping this conversion in one place means the Functions and the
// frontend never have to know about SharePoint's field-naming quirks.

function safeJson(str, fallback) {
  if (!str) return fallback;
  try {
    return JSON.parse(str);
  } catch {
    return fallback;
  }
}

// ---------- Suppliers ----------

function supplierFromSP(item) {
  const f = item.fields || {};
  return {
    id: item.id,
    n: f.Title || '',
    ab: f.Abbreviation || '',
    type: f.SupplierType || 'ff',
    sub: f.Subtype || '',
    region: f.Region || '',
    country: f.Country || '',
    lead: f.Lead || '',
    tier: f.Tier || '',
    spend: f.Spend != null ? Number(f.Spend) : 0,
    website: f.Website || '',
    comm: safeJson(f.CommFlags, ['—', '—', '—']),
    focus: safeJson(f.FocusAreas, ['—', '—', '—', '—', '—', '—']),
    ecovadis: f.EcovadisScore != null ? Number(f.EcovadisScore) : undefined,
    ecovadisRecognition: f.EcovadisRecognition || undefined,
    ecovadisDate: f.EcovadisDate || undefined,
    ecovadisIndustry: f.EcovadisIndustry || undefined,
    extra: {
      revenue: f.Revenue || '',
      hq: f.HQ || '',
      marketPosition: f.MarketPosition || '',
      creditRating: f.CreditRating || '',
      logistics: f.Logistics || '',
      otherMetric: f.OtherMetric || '',
      news: safeJson(f.News, []),
      sustainHighlights: f.SustainHighlights || '',
      carbonInitiatives: f.CarbonInitiatives || '',
      esgCommitments: f.EsgCommitments || '',
      poc: {
        name: f.PocName || '',
        email: f.PocEmail || '',
        phone: f.PocPhone || '',
        position: f.PocPosition || '',
      },
    },
  };
}

function supplierToSP(body) {
  const fields = {
    Title: body.n,
    Abbreviation: body.ab || (body.n || '').slice(0, 3).toUpperCase(),
    SupplierType: body.type || 'ff',
    Subtype: body.sub || '',
    Region: body.region || '',
    Country: body.country || '',
    Lead: body.lead || '',
    Tier: body.tier || 'Preferred/Tier 2',
    Spend: body.spend != null ? Number(body.spend) : 0,
    Website: body.website || '',
    CommFlags: JSON.stringify(body.comm || ['—', '—', '—']),
    FocusAreas: JSON.stringify(body.focus || ['—', '—', '—', '—', '—', '—']),
  };
  if (body.ecovadis != null) fields.EcovadisScore = Number(body.ecovadis);
  if (body.ecovadisRecognition) fields.EcovadisRecognition = body.ecovadisRecognition;
  if (body.ecovadisDate) fields.EcovadisDate = body.ecovadisDate;
  if (body.ecovadisIndustry) fields.EcovadisIndustry = body.ecovadisIndustry;

  const ex = body.extra || {};
  if (ex.revenue !== undefined) fields.Revenue = ex.revenue;
  if (ex.hq !== undefined) fields.HQ = ex.hq;
  if (ex.marketPosition !== undefined) fields.MarketPosition = ex.marketPosition;
  if (ex.creditRating !== undefined) fields.CreditRating = ex.creditRating;
  if (ex.logistics !== undefined) fields.Logistics = ex.logistics;
  if (ex.otherMetric !== undefined) fields.OtherMetric = ex.otherMetric;
  if (ex.news !== undefined) fields.News = JSON.stringify(ex.news || []);
  if (ex.sustainHighlights !== undefined) fields.SustainHighlights = ex.sustainHighlights;
  if (ex.carbonInitiatives !== undefined) fields.CarbonInitiatives = ex.carbonInitiatives;
  if (ex.esgCommitments !== undefined) fields.EsgCommitments = ex.esgCommitments;
  if (ex.poc) {
    if (ex.poc.name !== undefined) fields.PocName = ex.poc.name;
    if (ex.poc.email !== undefined) fields.PocEmail = ex.poc.email;
    if (ex.poc.phone !== undefined) fields.PocPhone = ex.poc.phone;
    if (ex.poc.position !== undefined) fields.PocPosition = ex.poc.position;
  }
  return fields;
}

// ---------- KPI Records ----------

function kpiFromSP(item) {
  const f = item.fields || {};
  return {
    id: item.id,
    supplierId: f.SupplierId != null ? Number(f.SupplierId) : null,
    metricKey: f.MetricKey || '',
    actual: f.ActualValue || '',
    target: f.TargetValue || '',
    ok: !!f.IsOnTarget,
  };
}

function kpiToSP(body) {
  return {
    Title: `${body.supplierId}-${body.metricKey}`,
    SupplierId: Number(body.supplierId),
    MetricKey: body.metricKey,
    ActualValue: body.actual || '',
    TargetValue: body.target || '',
    IsOnTarget: !!body.ok,
  };
}

/** Groups a flat KPIRecords list into the KPIS{supplierId:{metricKey:{a,t,ok}}} shape the frontend expects. */
function groupKpisBySupplier(kpiRecords) {
  const out = {};
  kpiRecords.forEach((k) => {
    if (k.supplierId == null) return;
    out[k.supplierId] = out[k.supplierId] || {};
    out[k.supplierId][k.metricKey] = { a: k.actual, t: k.target, ok: k.ok };
  });
  return out;
}

// ---------- Meetings ----------

function meetingFromSP(item) {
  const f = item.fields || {};
  return {
    id: item.id,
    sid: f.SupplierId != null ? Number(f.SupplierId) : null,
    date: f.MeetingDate ? f.MeetingDate.slice(0, 10) : '',
    type: f.MeetingType || '',
    att: f.Attendees || '',
    creator: f.CreatedByName || '',
    dis: f.Discussion || '',
    act: f.ActionItemsRaw || '',
    sen: f.Sentiment || 'neutral',
    files: safeJson(f.Attachments, []),
    kpis: safeJson(f.KpiSnapshot, null),
  };
}

function meetingToSP(body) {
  const fields = {
    Title: `${body.sname || 'Supplier'} — ${body.type || 'Meeting'} — ${body.date || ''}`,
    SupplierId: Number(body.sid),
    MeetingDate: body.date,
    MeetingType: body.type || '',
    Attendees: body.att || 'Not specified',
    CreatedByName: body.creator || '',
    Discussion: body.dis || 'No notes.',
    ActionItemsRaw: body.act || '',
    Sentiment: body.sen || 'neutral',
    Attachments: JSON.stringify(body.files || []),
  };
  if (body.kpis) fields.KpiSnapshot = JSON.stringify(body.kpis);
  return fields;
}

// ---------- Actions / Risks ----------

function actionFromSP(item) {
  const f = item.fields || {};
  return {
    id: item.id,
    sid: f.SupplierId != null ? Number(f.SupplierId) : null,
    type: f.Title || '',
    severity: f.Severity || 'Medium',
    status: f.Status || 'Open',
    owner: f.Owner || '',
    due: f.DueDate ? f.DueDate.slice(0, 10) : '',
    notes: f.Notes || '',
    sourceMeetingId: f.SourceMeetingId != null ? Number(f.SourceMeetingId) : null,
  };
}

function actionToSP(body) {
  const fields = {
    Title: body.type,
    SupplierId: Number(body.sid),
    Severity: body.severity || 'Medium',
    Status: body.status || 'Open',
    Owner: body.owner || '',
    Notes: body.notes || '',
  };
  if (body.due) fields.DueDate = body.due;
  if (body.sourceMeetingId != null) fields.SourceMeetingId = Number(body.sourceMeetingId);
  return fields;
}

// ---------- Performance History ----------

function historyFromSP(item) {
  const f = item.fields || {};
  return {
    id: item.id,
    sid: f.SupplierId != null ? Number(f.SupplierId) : null,
    metricKey: f.MetricKey || '',
    date: f.RecordDate ? f.RecordDate.slice(0, 10) : '',
    value: f.Value != null ? Number(f.Value) : null,
    sourceMeetingId: f.SourceMeetingId != null ? Number(f.SourceMeetingId) : null,
  };
}

function historyToSP(body) {
  const fields = {
    Title: `${body.sid}-${body.metricKey}-${body.date}`,
    SupplierId: Number(body.sid),
    MetricKey: body.metricKey,
    RecordDate: body.date,
    Value: Number(body.value),
  };
  if (body.sourceMeetingId != null) fields.SourceMeetingId = Number(body.sourceMeetingId);
  return fields;
}

module.exports = {
  safeJson,
  supplierFromSP, supplierToSP,
  kpiFromSP, kpiToSP, groupKpisBySupplier,
  meetingFromSP, meetingToSP,
  actionFromSP, actionToSP,
  historyFromSP, historyToSP,
};
