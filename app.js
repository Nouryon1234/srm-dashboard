# SharePoint List Schema

Create these lists in a SharePoint site (e.g. `https://nouryon.sharepoint.com/sites/SRM-Dashboard`). List names below are exact — the API code references them by these internal names. For each column, "Internal Name" is what you set when creating it via **Site contents → New → List → + Add column → More options**, or via PnP provisioning (template included in `provisioning/`).

> SharePoint's own `Id`/`ID` column is used as the record's unique identifier everywhere. Every foreign key below (e.g. `SupplierId` on a Meeting) stores the numeric SharePoint `Id` of the related Supplier item.

---

## 1. Suppliers

| Internal Name | Type | Notes |
|---|---|---|
| Title | Single line text | Supplier name (`n`) |
| Abbreviation | Single line text | 3-letter code (`ab`) |
| SupplierType | Choice | `ff` (Freight Forwarder family), `oc` (Ocean Carrier) |
| Subtype | Choice | Freight Forwarder, Ocean Carrier, Container Shipping, Road, Rail, Tank / Bulk |
| Region | Choice | APAC, EMEIA, AMERICAS, GLOBAL |
| Country | Single line text | |
| Lead | Single line text | Relationship owner / internal lead |
| Tier | Choice | Strategic/Tier 1, Preferred/Tier 2 |
| Spend | Number | Annual spend, $M |
| Website | Single line text (Hyperlink) | |
| CommFlags | Multiple lines text (plain) | JSON array, e.g. `["Y","N","Y"]` |
| FocusAreas | Multiple lines text (plain) | JSON array of 6 values: H/M/L/— |
| EcovadisScore | Number | Optional |
| EcovadisRecognition | Choice | Platinum, Gold, Silver, Bronze, Committed, N/A |
| EcovadisDate | Date | |
| EcovadisIndustry | Single line text | |
| Revenue | Single line text | Extra profile field |
| HQ | Single line text | |
| MarketPosition | Multiple lines text (plain) | |
| CreditRating | Single line text | |
| Logistics | Multiple lines text (plain) | |
| OtherMetric | Single line text | |
| News | Multiple lines text (plain) | JSON array of `{date,text}` |
| SustainHighlights | Multiple lines text (plain) | |
| CarbonInitiatives | Multiple lines text (plain) | |
| EsgCommitments | Multiple lines text (plain) | |
| PocName | Single line text | |
| PocEmail | Single line text | |
| PocPhone | Single line text | |
| PocPosition | Single line text | |

---

## 2. KPIRecords

One row per supplier per KPI metric snapshot (the "current" scorecard values shown on the supplier profile — separate from historical trend points, which live on Meetings).

| Internal Name | Type | Notes |
|---|---|---|
| Title | Single line text | Auto: `"<SupplierId>-<MetricKey>"` |
| SupplierId | Number | FK → Suppliers.Id |
| MetricKey | Choice | onTime, claims, invoice, response |
| ActualValue | Single line text | Display string, e.g. `"94%"` |
| TargetValue | Single line text | Display string, e.g. `"≥95%"` |
| IsOnTarget | Yes/No | |

---

## 3. Meetings

| Internal Name | Type | Notes |
|---|---|---|
| Title | Single line text | Auto: `"<Supplier> — <Type> — <Date>"` |
| SupplierId | Number | FK → Suppliers.Id |
| MeetingDate | Date | |
| MeetingType | Single line text | e.g. QBR, Escalation Call |
| Attendees | Multiple lines text (plain) | |
| CreatedByName | Single line text | Free-text creator name (distinct from SharePoint's system Author) |
| Discussion | Multiple lines text (rich text off) | |
| ActionItemsRaw | Multiple lines text (plain) | Newline-separated action item text |
| Sentiment | Choice | positive, neutral, concern, escalation |
| Attachments | Multiple lines text (plain) | JSON array of `{name,url}` |
| KpiSnapshot | Multiple lines text (plain) | JSON map of metric→number captured at this meeting, e.g. `{"onTime":94,"claims":0.3}` |

---

## 4. Actions

Pending/resolved action items — seeded from meeting `ActionItemsRaw` but independently editable (owner, due date, status).

| Internal Name | Type | Notes |
|---|---|---|
| Title | Single line text | Action description (`type`) |
| SupplierId | Number | FK → Suppliers.Id |
| Severity | Choice | Low, Medium, High |
| Status | Choice | Open, Monitoring, Escalated, Resolved |
| Owner | Single line text | |
| DueDate | Date | |
| Notes | Multiple lines text (plain) | |
| SourceMeetingId | Number | FK → Meetings.Id, optional |

---

## 5. Risks

Distinct from Actions: standing risk register entries not necessarily tied to a single meeting/action.

| Internal Name | Type | Notes |
|---|---|---|
| Title | Single line text | Risk description |
| SupplierId | Number | FK → Suppliers.Id |
| Severity | Choice | Low, Medium, High |
| Status | Choice | Open, Monitoring, Escalated, Resolved |

*(In practice the current dashboard treats "Risks" and "Actions" as one materialized list — see `docs/data-mapping.md`. This SharePoint list exists so the two can be split cleanly later if governance requires separate registers; the API currently reads/writes both from the same `Actions` list unless `RISKS_LIST_NAME` is overridden in Function App settings.)*

---

## 6. PerformanceHistory

Historical KPI trend points for charts — one row per supplier/metric/date. This is what feeds the trend charts and the supplier comparison tool.

| Internal Name | Type | Notes |
|---|---|---|
| Title | Single line text | Auto |
| SupplierId | Number | FK → Suppliers.Id |
| MetricKey | Single line text | onTime, claims, invoice, response, otif, bookingConf, otd, transitTime, costSavings, nps, volume, sustainability, servicePerf |
| RecordDate | Date | Usually the related meeting's date |
| Value | Number | |
| SourceMeetingId | Number | FK → Meetings.Id, optional |

---

## Permissions

Grant the Azure AD app registration (used by the Function App, see `docs/azure-setup.md`) **Sites.ReadWrite.All** (application permission) scoped ideally via **SharePoint Application Access Policy** to only this one site, not the whole tenant. Ask your M365 admin to run:

```powershell
Grant-SPOAzureADAppSitePermission -AppId <clientId> -DisplayName "SRM Dashboard API" -Site "https://nouryon.sharepoint.com/sites/SRM-Dashboard" -Right "Write"
```

This avoids granting the app write access to every SharePoint site in the tenant.
