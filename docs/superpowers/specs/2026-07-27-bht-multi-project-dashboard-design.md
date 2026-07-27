# BHT Multi-Project Dashboard — Design

Date: 2026-07-27
Status: Approved (pending user review of this doc)

## Context

Populix currently builds one dashboard per BHT client as a separate, hand-copied Next.js app (P. Hansel, P. Vertex). Each new client requires a manual clone-and-adapt. This spec designs a **new, standalone multi-project system**: one deployment that lets the internal team spin up a new client dashboard by uploading an Excel file, without touching code, and get a shareable link back.

This is a **new system**, built alongside the existing Hansel/Vertex apps — those are not migrated or touched by this work.

## Goals

- Internal team creates a new BHT client project in minutes: name it, pick a data type, upload the first Excel, get a link.
- Two data lifecycles per project: **Progress** (ongoing fielding, re-uploaded periodically) and **Final** (one-time cleaned dataset, explicitly published then locked).
- Same BHT chart/template experience as Hansel/Vertex, reused as-is — this system assumes a fixed BHT survey instrument (same question codes across all clients), not a general-purpose survey tool.
- Client-facing view stays no-login, same as today.

## Non-goals

- Self-service upload by stakeholders/clients — creation and upload stay internal-only (admin).
- Per-project custom column mapping — all projects share the one BHT `question-code-map.json`.
- Concurrent-upload locking — low volume, internal team; last-write-wins is acceptable.
- Migrating existing Hansel/Vertex apps into this system.

## Architecture: storage & slug scheme

Reuses the storage pattern proven in the Hansel app (Blob KV with local-fs fallback for dev, HMAC-signed admin session cookie), namespaced per project instead of global:

```
projects/{slug}/meta.json   → { slug, clientName, logo?, dataType: 'progress'|'final', status, createdAt, updatedAt }
projects/{slug}/data.json   → single dataset (progress-raw OR cleaned, per dataType — not both)
projects/index.json         → list of all slugs, for the admin project list
uploads/{slug}/{sessionId}/{index}  → chunked upload staging (existing chunk mechanism, namespaced by slug)
```

Local dev fallback mirrors this under `public/data/projects/{slug}/*.json`.

**Slug**: kebab-case of client name, admin-editable preview before create, uniqueness enforced (numeric suffix on collision). Immutable once a project leaves `empty` status, so shared links never break.

A database (e.g. Supabase/Postgres) was considered for project metadata but rejected as over-engineering for the expected scale (a few dozen internal-only BHT client projects) — Blob KV + an index file is sufficient and keeps the stack dependency-free.

## Data model & lifecycle

```
Project {
  slug: string
  clientName: string
  logo?: string
  dataType: 'progress' | 'final'
  status: 'empty' | 'active' | 'locked'
  createdAt, updatedAt: timestamp
}
```

- **Progress**: any upload replaces the current snapshot and goes live immediately (status `active`). Re-upload anytime.
- **Final**: first upload lands as a draft (status `active`, viewable via a "preview as client" link but bannered as draft) — admin must explicitly click "Publish as Final" to move to `locked` (read-only, no further upload without an explicit "Unlock for correction" action that reopens it to draft).

Explicit publish (not auto-lock on first upload) exists specifically to prevent a mistaken/test upload from permanently locking a project.

## Admin flow

- **Login**: unchanged — single shared `ADMIN_PASSWORD`, HMAC-signed session cookie, 8h TTL.
- **Project list** (`/admin`): table of all projects — client name, slug, dataType badge, status, last updated, sorted by last updated. Row actions: copy link, preview as client, upload.
- **Create wizard**: client name → slug preview (editable) → dataType radio (Progress/Final, one-line explanation each) → confirm. Lands on the new empty project page with the link already visible/copyable and a prominent "upload first Excel" CTA.
- **Upload**: drop `.xlsx` → runs the existing `build-data.ts` pipeline against the shared `question-code-map.json` → validates (see below) → shows a summary (respondent count, rows detected, question codes detected vs. expected) → admin confirms:
  - Progress: "Save & Publish" → live immediately.
  - Final: "Save Draft" first, then a separate "Publish as Final" action.

## Public dashboard behavior (`/d/[slug]`)

- **`empty`**: friendly "data is being prepared" placeholder (not a 404) — links may be shared before the first upload lands.
- **Progress / `active`**: same chart set as Hansel's progress view, single dataset (no Progress/Cleaned toggle — a project only has one type), footer shows "Last updated: {timestamp}".
- **Final / `locked`**: same chart set as Hansel's cleaned/respondent-level view, fully cross-filterable, header shows a "Final Report" badge instead of a freshness ticker.
- **Final / `active` (draft, pre-publish)**: same as locked view plus a small "Draft — not yet published" banner.

## Validation & error handling

- Required BHT question codes (S1, S2b, SEC5, SY, SZ, QA2, QA4, CO2, CX, CX2, SOI, SOA) must all be detected in the uploaded file's headers. Missing any → **reject the upload, do not save partial data** (existing data, if any, stays untouched); show which codes were found vs. missing so the admin can diagnose (wrong sheet, shifted columns, etc.).
- Accept `.xlsx` only, checked client- and server-side before entering the chunk-upload flow.
- Two known data-quality bugs already fixed in the existing pipeline must be preserved as-is when this is generalized: exceljs dropping a formula cell's cached `0` result, and merged-cell values needing to be read from the anchor cell only (openpyxl semantics) — see `dashboard/src/lib/server/build-data.ts`.

## Best-practice UX details

- Copy-link buttons everywhere a link is relevant (project list, project detail, post-upload screen).
- "Preview as client" opens `/d/[slug]` in a new tab for the admin to sanity-check before sharing.
- Upload summary card (respondent/row/question-code counts) shown before publish, not a bare "success" message.
- Archive instead of delete for finished/old projects — link keeps working, historical data stays reachable, just hidden from the default list.
- Slug becomes immutable once a project leaves `empty` status.

## Open items for the implementation plan

- Exact component split for reusing Hansel's progress-view / cleaned-view UI against arbitrary project data (likely: parameterize existing components by project slug instead of a hardcoded dataset path).
- Where the new system lives (new repo vs. new app within this monorepo-style working directory) — to be decided at plan time based on how "standalone" the user wants deployment/ops to be.
