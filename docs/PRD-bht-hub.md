# PRD: BHT Hub — Multi-Project BHT Dashboard

**Status:** Implemented (v1), in internal use
**Owner:** Archie (Populix)
**Last updated:** 2026-07-28
**Repo:** [archiepro/bht-hub](https://github.com/archiepro/bht-hub) (private)
**Related docs:** [Design spec](superpowers/specs/2026-07-27-bht-multi-project-dashboard-design.md) · [Implementation plan](superpowers/plans/2026-07-27-bht-multi-project-dashboard.md)

## 1. Problem

Populix builds client-facing web dashboards for the BHT (Brand Health Tracker) survey study. Each new client (P. Hansel, P. Vertex, and future clients) previously required cloning an entire Next.js app, wiring up its own data pipeline, and deploying it as a separate project. This meant:

- A multi-hour-to-multi-day setup cost per new client, done by an engineer, before any data could be shown.
- No standard, repeatable way to hand a client a live link — each app's link/deploy process was bespoke.
- No separation between "this survey is still in the field" (progress tracking) and "this survey is done, here's the final report" (a locked, citable deliverable) — every dashboard mixed both concerns.

**BHT Hub** solves this for any client running the *same* BHT survey instrument: an internal team member creates a project, uploads the client's Excel export, and gets back a shareable, no-login link — in minutes, with no code changes.

## 2. Goals

- Internal (Populix) team creates a new client dashboard without engineering involvement, for any client running the standard BHT questionnaire.
- Two clear data lifecycles: **Progress** (fielding is ongoing, re-uploaded periodically, always shows the latest snapshot) and **Final** (fielding is done, one clean dataset, explicitly published and then locked as a stable client deliverable).
- The client-facing dashboard is a faithful continuation of the existing Hansel/Vertex dashboard experience (design system, chart types, KPI framing) — clients see the interface they already recognize.
- Uploads are validated against the known BHT question set before anything is saved, so a wrong file never silently corrupts a client's dashboard.
- No-login public links, consistent with how Populix has always shared these dashboards with clients.

## 3. Non-goals (v1)

- **Self-service for clients.** Only the internal Populix team creates projects and uploads data. Clients never get upload or admin access.
- **Arbitrary survey instruments.** BHT Hub assumes every project runs the *same* fixed BHT questionnaire (same question codes, same dimensions). A client whose survey structure genuinely differs (new question types, e.g. separate TOM/Aided/Unaided awareness questions not in the current instrument) is out of scope for this version — see [§8 Future Considerations](#8-future-considerations).
- **Concurrent-edit safety.** Two admins creating projects or uploading data at the exact same moment can race (last write wins). Acceptable given the internal, low-volume, one-admin-at-a-time usage pattern.
- **Per-project custom branding beyond the client name** (e.g. per-client logo upload) — deferred.
- **Client account management, notifications, or usage analytics.**

## 4. Users

- **Primary user: Populix internal team member** (research ops, project lead) who has fielded a BHT study for a client and needs to stand up a dashboard. Technical comfort: can use a web admin panel and knows their way around Excel exports; not expected to write code or touch a server.
- **Secondary "user": the client stakeholder**, who only ever sees the read-only, no-login link. They never log in, never see the admin console.

## 5. User Stories

1. *As an internal team member*, I can create a new project by naming the client and choosing whether this study is still in the field ("Progress") or complete ("Final"), so I don't have to ask engineering to set anything up.
2. *As an internal team member*, I can upload the client's BHT Excel export and immediately see whether it was accepted or what's wrong with it (missing question codes, bad file format), so I can catch mistakes before a client ever sees a broken link.
3. *As an internal team member managing a Progress-type project*, I can re-upload a fresh export at any point during fielding, and the live link updates immediately, so the client always sees current numbers without me managing multiple links.
4. *As an internal team member managing a Final-type project*, I upload the cleaned dataset once, review it, and only when I'm confident it's correct do I explicitly "publish" it — after which it's locked and can't be silently changed, so I can hand it to a client as a stable, citable report.
5. *As an internal team member*, I get a shareable link the moment I create a project (even before data is uploaded), so I can share it with the client contact early if needed, and the page shows a friendly "being prepared" message instead of an error.
6. *As a client stakeholder*, I open the link I was sent and see a live, branded (client-name-labeled) dashboard with no login required, matching the same chart types and KPIs Populix has always shown me.

## 6. Functional Requirements

### 6.1 Project lifecycle

| Data type | Behavior |
|---|---|
| `Progress` | Any upload replaces the current snapshot and goes live immediately. Re-upload anytime during fielding. Dashboard shows a KPI row (Target / Achievement / Remaining) plus a quota breakdown table grouped by category, with a "Last updated" timestamp. |
| `Final` | First upload lands as an internally-previewable draft. An explicit "Publish as Final" action locks the project (read-only, no further upload without an explicit "Unlock for correction"). Dashboard shows the full cleaned/respondent-level breakdown (Overview, Demographics, Brand, Awareness, NPS & CSAT, Media tabs), fully cross-filterable, labeled "Final Report." |

Project states: `empty` (created, no data yet) → `active` (has data; for Final, this is the unpublished-draft state) → `locked` (Final only, published).

### 6.2 Admin console (`/admin`)

- Single shared password (existing HMAC session-cookie pattern), 8-hour session.
- Project list: client name, data type, status, last updated, sorted by most recently updated. Copy-link, "Preview as client," and "Manage" actions per row.
- Create-project form: client name → auto-generated URL slug (editable before creation, immutable after) → data-type choice.
- Project detail page: upload form, upload result summary (rows/respondents detected, question codes detected vs. expected), Publish/Unlock controls for Final-type projects.

### 6.3 Upload & validation

- Accepts `.xlsx` only.
- Every upload is checked against the full BHT question-code set (currently 19 required codes) before anything is saved. Any missing code rejects the upload with the specific missing codes listed — the project's existing data is left untouched.
- Progress-type and Final-type projects both run the same underlying data pipeline; only the relevant dataset (quota-tracking vs. respondent-level) is persisted, matching the project's data type.

### 6.4 Public dashboard (`/d/{slug}`)

- No login required.
- Page title and header show the client's name (not a shared/generic title).
- States: no-data placeholder ("Data is being prepared"), Progress view, Final draft view (small "Draft — not yet published" indicator, since a draft link may exist before an admin has finished reviewing it), Final published view ("Final Report" label, no draft indicator).
- Final-type dashboard tabs: **Overview** (gender, age, domicile, SEC, marital status), **Demographics** (segment, occupation, employment, industry, education), **Brand** (insurance type owned, current brand usage), **Awareness** (brand awareness by category, source of awareness), **NPS & CSAT** (Net Promoter Score and satisfaction, overall and per-brand), **Media** (source of influence). Cross-filterable by demographic facets; results can be toggled between raw counts and percentages.

## 7. Non-Functional Requirements

- **Security:** every mutating and admin-listing action is gated behind the session check; the public dashboard and its data API are read-only with no path to a mutating operation. Uploaded files are validated for format before being parsed.
- **Data isolation:** each project's metadata and dataset are stored under a slug-namespaced key; one project's data can never be read or overwritten by another project's key.
- **Storage:** cloud key-value storage (Netlify Blobs) in production, with a zero-dependency local-filesystem fallback for local development — the site is never blank before a first upload.
- **Reliability of the pipeline:** the underlying Excel-parsing pipeline (question mapping, cell-value normalization) is reused unmodified from the existing, production-proven Hansel dashboard, including two previously-discovered and fixed spreadsheet-parsing edge cases (a zero-value formula bug and a merged-cell-value bug).

## 8. Future Considerations

Not committed, but the likely next asks based on how this is already being used:

- **Non-BHT survey instruments.** If a future client's questionnaire genuinely differs from the current BHT instrument (e.g., separate Top-of-Mind / Aided / Unaided awareness questions that aren't in today's fixed 19-code set), that requires defining a new instrument mapping (a small, one-time engineering task per new instrument shape — extending the existing question-code-map config, not a rebuild), not a fully generic self-service mapping tool. A generic drag-and-drop column-mapping UI was considered and explicitly deferred as disproportionate complexity for how rarely instrument shape actually varies.
- **Sheet name auto-detection.** Uploads currently require exact sheet names (`Progress`, `BHT`, `REFF`) matching the BHT template. A fuzzy-match fallback (detect the right sheet by column-pattern rather than name) is possible but trades away the current "silent wrong-file protection" guarantee unless paired with an explicit confirmation step before saving.
- **Per-project logo upload** for client branding beyond the name.
- **Project archiving** (hide finished projects from the default admin list without deleting them).
- **Production deployment** (a real Netlify/Vercel Blob store + site provisioning) — the app is deployment-ready but not yet deployed; this is a deliberate, separate go/no-go decision.

## 9. Success Metrics (informal, internal tool)

- Time from "client fielding starts" to "client has a live link": target under 15 minutes, admin-only, no engineering ticket.
- Zero incidents of a client seeing another client's data (isolation guarantee holds).
- Zero incidents of a corrupted/wrong file silently going live on a client-facing link (validation guarantee holds).

## 10. Open Questions

- Who owns the decision to provision production storage/deployment, and when?
- Should "Final, locked" projects ever be deletable, or retained indefinitely as an audit record?
- Is there an appetite for a lightweight instrument-mapping tool if a second BHT variant becomes common enough to justify it, versus continuing to hand-extend the config per instrument?
