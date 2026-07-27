# BHT Multi-Project Dashboard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build `bht-hub/`, a standalone Next.js app that lets the internal Populix team create a new BHT client dashboard (name + data type), upload the client's Excel export, and get back a shareable no-login link — without cloning a new repo per client the way Hansel/Vertex were built.

**Architecture:** Copy the proven foundation from `dashboard/` (the Hansel app) as-is — `build-data.ts`, `question-code-map.json`, shadcn/ui primitives, chart/tally helpers, the Blob-KV-with-local-fs-fallback storage pattern, the HMAC-signed single-passphrase admin session. Layer a project registry on top (slug → metadata → one dataset), namespace every storage key and API route by project slug, and add a data-type-aware public dashboard (`/d/[slug]`) plus an admin console (`/admin`) for creating projects and managing uploads.

**Tech Stack:** Next.js 16 (App Router) + React 19 + TypeScript, exceljs for spreadsheet parsing, `@netlify/blobs` for storage with local-fs dev fallback, Tailwind v4 + shadcn/ui, Recharts. Vitest added new for this app to unit-test the new pure-logic modules (nothing in `dashboard/` has tests today — the existing pipeline was validated by byte-diffing outputs, which stays true for the copied `build-data.ts`).

## Global Constraints

- Reuse `dashboard/config/question-code-map.json` and `dashboard/src/lib/server/build-data.ts` **unmodified except for one addition** (`validateQuestionCodes`, Task 2) — every dimension/label/normalization rule and the two known exceljs bugs (dropped cached `0` on formula cells, merged-cell value only on anchor) must carry over exactly. Do not "clean up" or refactor this file while porting it.
- A project has exactly one `dataType` (`"progress"` or `"final"`), fixed at creation — never both, unlike Hansel.
- All required question codes = every entry in `question-code-map.json`'s `dimensions` array (19 codes: S1, S2b, S3, SEC5, SY, SZ, QA2, QA4, S7, S7a, CO2, CX, CX2, SOI, SOA, A3, A3b, NPS1, CSAT1) — not a hand-picked subset. An upload missing any of them is rejected.
- Admin routes require the existing HMAC session cookie pattern (`ADMIN_COOKIE_NAME`, `verifySessionToken`) — copy `admin-auth.ts` verbatim, no new auth mechanism.
- Public dashboard routes (`/d/[slug]` and its data API) require no authentication, matching the existing no-login client-view requirement.
- Slugs are immutable once a project leaves `status: "empty"`.
- `bht-hub/` is its own nested git repository, gitignored from the parent repo — same pattern as `dashboard/` (which has its own `.git` and its own GitHub remote, `archiepro/p-hansel-bht-dashboard`, deployed independently). The parent repo's `.gitignore` already has a `bht-hub/` entry. **Every commit step in this plan runs `cd bht-hub` first and stages paths relative to `bht-hub/` (not prefixed with it)** — e.g. `git add src/lib/slug.ts`, not `git add bht-hub/src/lib/slug.ts`. Task 1 initializes this repo before any other task commits into it.

---

### Task 1: Scaffold `bht-hub/` and set up Vitest

**Files:**
- Create: `bht-hub/package.json`
- Create: `bht-hub/tsconfig.json`
- Create: `bht-hub/next.config.ts`
- Create: `bht-hub/postcss.config.mjs`
- Create: `bht-hub/eslint.config.mjs`
- Create: `bht-hub/vitest.config.ts`
- Create: `bht-hub/src/app/layout.tsx`
- Create: `bht-hub/src/app/globals.css`
- Create: `bht-hub/src/app/page.tsx`
- Create: `bht-hub/.gitignore`

**Interfaces:**
- Produces: a runnable `bht-hub` Next.js app (`npm run dev` on port 3100) and a runnable `npm run test` (Vitest) that later tasks add test files to.

- [ ] **Step 1: Create the directory, init it as its own git repo, and add package.json**

`bht-hub/` is a standalone app with its own history and (eventually) its own remote, exactly like `dashboard/` — not a subdirectory tracked by the parent repo. The parent's `.gitignore` already excludes `bht-hub/`.

```bash
mkdir -p "bht-hub/src/app"
cd bht-hub && git init
```

```json
{
  "name": "bht-hub",
  "version": "0.1.0",
  "private": true,
  "scripts": {
    "dev": "next dev -p 3100",
    "build": "next build",
    "start": "next start -p 3100",
    "lint": "eslint",
    "test": "vitest run"
  },
  "dependencies": {
    "@base-ui/react": "^1.6.0",
    "@netlify/blobs": "^10.7.9",
    "class-variance-authority": "^0.7.1",
    "clsx": "^2.1.1",
    "exceljs": "^4.4.0",
    "lucide-react": "^1.24.0",
    "next": "16.2.10",
    "react": "19.2.4",
    "react-dom": "19.2.4",
    "recharts": "^3.8.0",
    "tailwind-merge": "^3.6.0",
    "tw-animate-css": "^1.4.0"
  },
  "devDependencies": {
    "@tailwindcss/postcss": "^4",
    "@types/node": "^20",
    "@types/react": "^19",
    "@types/react-dom": "^19",
    "eslint": "^9",
    "eslint-config-next": "16.2.10",
    "tailwindcss": "^4",
    "typescript": "^5",
    "vitest": "^3.2.4"
  }
}
```

- [ ] **Step 2: Copy config files from `dashboard/` that need no changes**

```bash
cp dashboard/tsconfig.json bht-hub/tsconfig.json
cp dashboard/next.config.ts bht-hub/next.config.ts
cp dashboard/postcss.config.mjs bht-hub/postcss.config.mjs
cp dashboard/eslint.config.mjs bht-hub/eslint.config.mjs
cp dashboard/.gitignore bht-hub/.gitignore
cp dashboard/components.json bht-hub/components.json
cp dashboard/src/app/globals.css bht-hub/src/app/globals.css
cp dashboard/src/app/tokens.css bht-hub/src/app/tokens.css
cp dashboard/src/app/layout.tsx bht-hub/src/app/layout.tsx
```

- [ ] **Step 3: Write `vitest.config.ts`**

```typescript
import { defineConfig } from "vitest/config";
import path from "node:path";

export default defineConfig({
  test: {
    environment: "node",
  },
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "./src"),
    },
  },
});
```

- [ ] **Step 4: Write a placeholder home page (replaced by the real admin/public routes in later tasks)**

```typescript
// bht-hub/src/app/page.tsx
export default function Home() {
  return (
    <main className="flex min-h-screen items-center justify-center text-sm text-muted-foreground">
      BHT Hub — go to /admin to manage projects.
    </main>
  );
}
```

- [ ] **Step 5: Install and verify the app boots**

```bash
cd bht-hub && npm install
```

Run: `cd bht-hub && npm run build`
Expected: build succeeds (only `page.tsx`/`layout.tsx` exist so far, no routes yet to break).

- [ ] **Step 6: Verify Vitest runs with zero tests (sanity check before real tests exist)**

Run: `cd bht-hub && npm run test`
Expected: `No test files found` (not an error) — confirms the runner and alias resolve correctly. This will start reporting real results once Task 3 adds its first test.

- [ ] **Step 7: Commit**

```bash
cd bht-hub && git add package.json tsconfig.json next.config.ts postcss.config.mjs eslint.config.mjs vitest.config.ts components.json .gitignore src/app/globals.css src/app/tokens.css src/app/layout.tsx src/app/page.tsx package-lock.json
git commit -m "feat(bht-hub): scaffold new multi-project Next.js app"
```

---

### Task 2: Port the BHT data pipeline and add question-code validation

**Files:**
- Create: `bht-hub/config/question-code-map.json`
- Create: `bht-hub/src/lib/server/build-data.ts`
- Create: `bht-hub/src/lib/types.ts`
- Test: `bht-hub/src/lib/server/build-data.test.ts`

**Interfaces:**
- Consumes: nothing (pure port from `dashboard/`).
- Produces: `processWorkbook(buffer: Buffer): Promise<{progressRaw: ProgressRawOut; cleaned: CleanedOut; summary: ProcessSummary}>` (unchanged signature) and a new `validateQuestionCodes(buffer: Buffer): Promise<CodeDetection[]>` where `CodeDetection = {code: string; key: string; label: string; found: boolean}` — used by Task 8's upload route to reject files missing any required BHT question code before saving anything.

- [ ] **Step 1: Copy the pipeline, config, and types verbatim**

```bash
mkdir -p "bht-hub/config" "bht-hub/src/lib/server"
cp "dashboard/config/question-code-map.json" "bht-hub/config/question-code-map.json"
cp "dashboard/src/lib/server/build-data.ts" "bht-hub/src/lib/server/build-data.ts"
cp "dashboard/src/lib/types.ts" "bht-hub/src/lib/types.ts"
```

- [ ] **Step 2: Write the failing test for `validateQuestionCodes` (function doesn't exist yet)**

```typescript
// bht-hub/src/lib/server/build-data.test.ts
import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import { validateQuestionCodes } from "./build-data";
import codeMapJson from "../../../config/question-code-map.json";

const codeMap = codeMapJson as { dimensions: Array<{ code: string }> };

async function workbookBuffer(headers: string[]): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("BHT");
  ws.addRow(headers);
  const arrayBuffer = await wb.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}

describe("validateQuestionCodes", () => {
  it("marks every dimension as not found on an empty-ish sheet", async () => {
    const buffer = await workbookBuffer(["Respondent ID"]);
    const result = await validateQuestionCodes(buffer);
    expect(result).toHaveLength(codeMap.dimensions.length);
    expect(result.every((d) => !d.found)).toBe(true);
  });

  it("marks a dimension found when its header prefix is present", async () => {
    const buffer = await workbookBuffer(["S1. Domicile"]);
    const result = await validateQuestionCodes(buffer);
    const domicile = result.find((d) => d.code === "S1");
    expect(domicile?.found).toBe(true);
    const others = result.filter((d) => d.code !== "S1");
    expect(others.every((d) => !d.found)).toBe(true);
  });
});
```

- [ ] **Step 2b: Run it to confirm it fails**

Run: `cd bht-hub && npm run test -- build-data`
Expected: FAIL — `validateQuestionCodes is not exported` (or similar).

- [ ] **Step 3: Add `validateQuestionCodes` to `build-data.ts`**

Append to `bht-hub/src/lib/server/build-data.ts` (it already has `requireSheet`, `headerTexts`, `findSingleCol`, `findOnehotCols`, `findSoaCols`, `findScoreByBrandCols`, and `codeMap` in scope from the copy in Step 1 — this reuses them, no duplication):

```typescript
export interface CodeDetection {
  code: string;
  key: string;
  label: string;
  found: boolean;
}

export async function validateQuestionCodes(buffer: Buffer): Promise<CodeDetection[]> {
  const wb = new ExcelJS.Workbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  await wb.xlsx.load(buffer as any);
  const ws = requireSheet(wb, codeMap.primarySheet);
  const headers = headerTexts(ws);

  return codeMap.dimensions.map((dim) => {
    let found: boolean;
    switch (dim.type) {
      case "single":
        found = findSingleCol(headers, dim.code) !== null;
        break;
      case "onehot":
        found = findOnehotCols(headers, dim.code).length > 0;
        break;
      case "brand_matrix":
        found = findSoaCols(headers, dim.code).length > 0;
        break;
      case "score_by_brand":
        found = findScoreByBrandCols(headers, dim.code).length > 0;
        break;
      default:
        found = false;
    }
    return { code: dim.code, key: dim.key, label: dim.label, found };
  });
}
```

- [ ] **Step 4: Run the test again to confirm it passes**

Run: `cd bht-hub && npm run test -- build-data`
Expected: PASS (2 tests).

- [ ] **Step 5: Commit**

```bash
cd bht-hub && git add config/question-code-map.json src/lib/server/build-data.ts src/lib/server/build-data.test.ts src/lib/types.ts
git commit -m "feat(bht-hub): port BHT data pipeline, add question-code validation"
```

---

### Task 3: Slug utility

**Files:**
- Create: `bht-hub/src/lib/slug.ts`
- Test: `bht-hub/src/lib/slug.test.ts`

**Interfaces:**
- Produces: `slugify(name: string): string` and `uniqueSlug(name: string, existing: string[]): string` — consumed by Task 5's `createProject`.

- [ ] **Step 1: Write the failing tests**

```typescript
// bht-hub/src/lib/slug.test.ts
import { describe, expect, it } from "vitest";
import { slugify, uniqueSlug } from "./slug";

describe("slugify", () => {
  it("lowercases, replaces spaces with hyphens, strips punctuation", () => {
    expect(slugify("Great Eastern 2026")).toBe("great-eastern-2026");
    expect(slugify("P. Vertex, BHT!")).toBe("p-vertex-bht");
  });

  it("collapses repeated separators and trims leading/trailing hyphens", () => {
    expect(slugify("  Multi   Space -- Name  ")).toBe("multi-space-name");
  });
});

describe("uniqueSlug", () => {
  it("returns the base slug when there's no collision", () => {
    expect(uniqueSlug("great-eastern", [])).toBe("great-eastern");
  });

  it("appends -2, -3, ... on collision", () => {
    expect(uniqueSlug("great-eastern", ["great-eastern"])).toBe("great-eastern-2");
    expect(uniqueSlug("great-eastern", ["great-eastern", "great-eastern-2"])).toBe(
      "great-eastern-3"
    );
  });
});
```

- [ ] **Step 2: Run to confirm failure**

Run: `cd bht-hub && npm run test -- slug`
Expected: FAIL — module `./slug` not found.

- [ ] **Step 3: Implement**

```typescript
// bht-hub/src/lib/slug.ts
export function slugify(name: string): string {
  return name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

export function uniqueSlug(name: string, existing: string[]): string {
  const base = slugify(name);
  const taken = new Set(existing);
  if (!taken.has(base)) return base;

  let n = 2;
  while (taken.has(`${base}-${n}`)) n++;
  return `${base}-${n}`;
}
```

- [ ] **Step 4: Run to confirm pass**

Run: `cd bht-hub && npm run test -- slug`
Expected: PASS (4 tests).

- [ ] **Step 5: Commit**

```bash
cd bht-hub && git add src/lib/slug.ts src/lib/slug.test.ts
git commit -m "feat(bht-hub): add slug generation utility"
```

---

### Task 4: Generalized Blob/local-fs key-value storage

**Files:**
- Create: `bht-hub/src/lib/server/blob-store.ts`
- Test: `bht-hub/src/lib/server/blob-store.test.ts`

**Interfaces:**
- Consumes: nothing new (same `@netlify/blobs` dependency as `dashboard/`).
- Produces: `getJSON<T>(key: string): Promise<T | null>`, `setJSON(key: string, value: unknown): Promise<void>`, `saveChunk`, `assembleChunks`, `deleteChunks` (same signatures as `dashboard/src/lib/server/data-store.ts` but keyed by an arbitrary string instead of a fixed `DatasetName` union) — consumed by Task 5's `project-store.ts` and Task 8's upload routes.

This generalizes `dashboard/src/lib/server/data-store.ts`'s pattern (Blob when a token is configured, local-fs fallback otherwise, never blank pre-first-write) from two fixed dataset keys to arbitrary keys, since project data now lives at `projects/{slug}/...` paths instead of two hardcoded files.

- [ ] **Step 1: Write the failing test (exercises the local-fs fallback path, since no `NETLIFY_BLOBS_TOKEN` is set in the test env)**

```typescript
// bht-hub/src/lib/server/blob-store.test.ts
import { afterEach, describe, expect, it } from "vitest";
import fs from "node:fs/promises";
import path from "node:path";
import { getJSON, setJSON } from "./blob-store";

const TEST_KEY = "projects/__vitest-fixture__/meta.json";
const TEST_FILE = path.join(process.cwd(), "public", "data", "projects", "__vitest-fixture__", "meta.json");

afterEach(async () => {
  await fs.rm(TEST_FILE, { force: true });
});

describe("blob-store local-fs fallback", () => {
  it("returns null for a key that was never written", async () => {
    const result = await getJSON(TEST_KEY);
    expect(result).toBeNull();
  });

  it("round-trips a JSON value through setJSON/getJSON", async () => {
    await setJSON(TEST_KEY, { hello: "world" });
    const result = await getJSON(TEST_KEY);
    expect(result).toEqual({ hello: "world" });
  });
});
```

- [ ] **Step 2: Run to confirm failure**

Run: `cd bht-hub && npm run test -- blob-store`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement, adapting `dashboard/src/lib/server/data-store.ts`'s Blob-with-fallback pattern to arbitrary keys**

```typescript
// bht-hub/src/lib/server/blob-store.ts
import fs from "node:fs/promises";
import path from "node:path";
import { getStore } from "@netlify/blobs";

const LOCAL_DATA_DIR = path.join(process.cwd(), "public", "data");

// Not a secret -- identifies which site's store to use, meaningless without
// the token below. Distinct from the Hansel app's store so the two never collide.
const SITE_ID = "f4a7c112-9e3d-4b6a-9c1f-2d6e0a3b9c77";

function store() {
  const token = process.env.NETLIFY_BLOBS_TOKEN;
  if (token) {
    return getStore({ name: "bht-hub-data", siteID: SITE_ID, token });
  }
  return getStore("bht-hub-data");
}

function isMissingBlobsEnvironment(err: unknown): boolean {
  return err instanceof Error && err.name === "MissingBlobsEnvironmentError";
}

function localPathFor(key: string): string {
  return path.join(LOCAL_DATA_DIR, ...key.split("/"));
}

export async function setJSON(key: string, value: unknown): Promise<void> {
  try {
    await store().setJSON(key, value);
    return;
  } catch (err) {
    if (!isMissingBlobsEnvironment(err)) throw err;
  }

  const filePath = localPathFor(key);
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, JSON.stringify(value, null, 2));
}

export async function getJSON<T>(key: string): Promise<T | null> {
  try {
    const data = await store().get(key, { type: "json" });
    if (data !== null) return data as T;
  } catch (err) {
    if (!isMissingBlobsEnvironment(err)) throw err;
  }

  try {
    const raw = await fs.readFile(localPathFor(key), "utf-8");
    return JSON.parse(raw) as T;
  } catch {
    return null;
  }
}

const LOCAL_CHUNK_DIR = path.join(process.cwd(), ".upload-chunks");

function chunkKey(sessionId: string, index: number): string {
  return `uploads/${sessionId}/${index}`;
}

export async function saveChunk(sessionId: string, index: number, data: ArrayBuffer): Promise<void> {
  try {
    await store().set(chunkKey(sessionId, index), data);
    return;
  } catch (err) {
    if (!isMissingBlobsEnvironment(err)) throw err;
  }

  const dir = path.join(LOCAL_CHUNK_DIR, sessionId);
  await fs.mkdir(dir, { recursive: true });
  await fs.writeFile(path.join(dir, String(index)), Buffer.from(data));
}

export async function assembleChunks(sessionId: string, count: number): Promise<Buffer> {
  try {
    const buffers = await Promise.all(
      Array.from({ length: count }, (_, i) =>
        store()
          .get(chunkKey(sessionId, i), { type: "arrayBuffer" })
          .then((b) => Buffer.from(b))
      )
    );
    return Buffer.concat(buffers);
  } catch (err) {
    if (!isMissingBlobsEnvironment(err)) throw err;
  }

  const dir = path.join(LOCAL_CHUNK_DIR, sessionId);
  const buffers = await Promise.all(
    Array.from({ length: count }, (_, i) => fs.readFile(path.join(dir, String(i))))
  );
  return Buffer.concat(buffers);
}

export async function deleteChunks(sessionId: string, count: number): Promise<void> {
  try {
    await Promise.all(
      Array.from({ length: count }, (_, i) => store().delete(chunkKey(sessionId, i)))
    );
    return;
  } catch (err) {
    if (!isMissingBlobsEnvironment(err)) throw err;
  }

  const dir = path.join(LOCAL_CHUNK_DIR, sessionId);
  await fs.rm(dir, { recursive: true, force: true });
}
```

- [ ] **Step 4: Run to confirm pass**

Run: `cd bht-hub && npm run test -- blob-store`
Expected: PASS (2 tests).

- [ ] **Step 5: Commit**

```bash
cd bht-hub && git add src/lib/server/blob-store.ts src/lib/server/blob-store.test.ts
git commit -m "feat(bht-hub): add slug-namespaced blob/local-fs storage layer"
```

---

### Task 5: Project domain model and store

**Files:**
- Create: `bht-hub/src/lib/project.ts`
- Create: `bht-hub/src/lib/server/project-store.ts`
- Test: `bht-hub/src/lib/server/project-store.test.ts`

**Interfaces:**
- Consumes: `getJSON`/`setJSON` from `./blob-store` (Task 4), `uniqueSlug` from `../slug` (Task 3).
- Produces: `Project` type, `listProjects(): Promise<Project[]>`, `getProject(slug: string): Promise<Project | null>`, `createProject(clientName: string, dataType: DataType): Promise<Project>`, `updateProjectStatus(slug: string, status: ProjectStatus): Promise<Project>`, `saveProjectData(slug: string, data: unknown): Promise<void>`, `loadProjectData<T>(slug: string): Promise<T | null>` — consumed by every admin/public API route in Tasks 6-9 and 12.

- [ ] **Step 1: Define the domain types**

```typescript
// bht-hub/src/lib/project.ts
export type DataType = "progress" | "final";
export type ProjectStatus = "empty" | "active" | "locked";

export interface Project {
  slug: string;
  clientName: string;
  dataType: DataType;
  status: ProjectStatus;
  createdAt: string;
  updatedAt: string;
}
```

- [ ] **Step 2: Write the failing tests for the store**

```typescript
// bht-hub/src/lib/server/project-store.test.ts
import { afterEach, describe, expect, it } from "vitest";
import fs from "node:fs/promises";
import path from "node:path";
import {
  createProject,
  getProject,
  listProjects,
  loadProjectData,
  saveProjectData,
  updateProjectStatus,
} from "./project-store";

const PROJECTS_DIR = path.join(process.cwd(), "public", "data", "projects");

afterEach(async () => {
  // Removes projects/index.json too -- it lives inside this same directory.
  await fs.rm(PROJECTS_DIR, { recursive: true, force: true });
});

describe("project-store", () => {
  it("creates a project with status empty and a derived slug", async () => {
    const project = await createProject("Great Eastern 2026", "progress");
    expect(project.slug).toBe("great-eastern-2026");
    expect(project.status).toBe("empty");
    expect(project.dataType).toBe("progress");
  });

  it("de-duplicates slugs across repeated client names", async () => {
    const first = await createProject("Vertex", "final");
    const second = await createProject("Vertex", "final");
    expect(first.slug).toBe("vertex");
    expect(second.slug).toBe("vertex-2");
  });

  it("lists created projects and can fetch one by slug", async () => {
    await createProject("Alpha", "progress");
    await createProject("Beta", "final");
    const all = await listProjects();
    expect(all.map((p) => p.slug).sort()).toEqual(["alpha", "beta"]);
    expect((await getProject("alpha"))?.clientName).toBe("Alpha");
    expect(await getProject("does-not-exist")).toBeNull();
  });

  it("updates project status", async () => {
    const project = await createProject("Gamma", "progress");
    const updated = await updateProjectStatus(project.slug, "active");
    expect(updated.status).toBe("active");
    expect((await getProject(project.slug))?.status).toBe("active");
  });

  it("saves and loads project data round-trip", async () => {
    const project = await createProject("Delta", "final");
    await saveProjectData(project.slug, { rows: [{ id: 1 }] });
    const data = await loadProjectData<{ rows: { id: number }[] }>(project.slug);
    expect(data?.rows).toEqual([{ id: 1 }]);
  });
});
```

- [ ] **Step 3: Run to confirm failure**

Run: `cd bht-hub && npm run test -- project-store`
Expected: FAIL — module not found.

- [ ] **Step 4: Implement**

```typescript
// bht-hub/src/lib/server/project-store.ts
import { getJSON, setJSON } from "./blob-store";
import { uniqueSlug } from "../slug";
import type { DataType, Project, ProjectStatus } from "../project";

const INDEX_KEY = "projects/index.json";

function metaKey(slug: string): string {
  return `projects/${slug}/meta.json`;
}

function dataKey(slug: string): string {
  return `projects/${slug}/data.json`;
}

async function readIndex(): Promise<string[]> {
  return (await getJSON<string[]>(INDEX_KEY)) ?? [];
}

async function writeIndex(slugs: string[]): Promise<void> {
  await setJSON(INDEX_KEY, slugs);
}

export async function listProjects(): Promise<Project[]> {
  const slugs = await readIndex();
  const projects = await Promise.all(slugs.map((slug) => getJSON<Project>(metaKey(slug))));
  return projects.filter((p): p is Project => p !== null);
}

export async function getProject(slug: string): Promise<Project | null> {
  return getJSON<Project>(metaKey(slug));
}

export async function createProject(clientName: string, dataType: DataType): Promise<Project> {
  const existingSlugs = await readIndex();
  const slug = uniqueSlug(clientName, existingSlugs);
  const now = new Date().toISOString();

  const project: Project = {
    slug,
    clientName,
    dataType,
    status: "empty",
    createdAt: now,
    updatedAt: now,
  };

  await setJSON(metaKey(slug), project);
  await writeIndex([...existingSlugs, slug]);
  return project;
}

export async function updateProjectStatus(slug: string, status: ProjectStatus): Promise<Project> {
  const project = await getProject(slug);
  if (!project) throw new Error(`Project '${slug}' not found`);

  const updated: Project = { ...project, status, updatedAt: new Date().toISOString() };
  await setJSON(metaKey(slug), updated);
  return updated;
}

export async function saveProjectData(slug: string, data: unknown): Promise<void> {
  await setJSON(dataKey(slug), data);
}

export async function loadProjectData<T>(slug: string): Promise<T | null> {
  return getJSON<T>(dataKey(slug));
}
```

- [ ] **Step 5: Run to confirm pass**

Run: `cd bht-hub && npm run test -- project-store`
Expected: PASS (5 tests).

- [ ] **Step 6: Commit**

```bash
cd bht-hub && git add src/lib/project.ts src/lib/server/project-store.ts src/lib/server/project-store.test.ts
git commit -m "feat(bht-hub): add project domain model and storage"
```

---

### Task 6: Admin auth routes (copy) + login/session UI primitives

**Files:**
- Create: `bht-hub/src/lib/server/admin-auth.ts`
- Create: `bht-hub/src/app/api/admin/login/route.ts`
- Create: `bht-hub/src/app/api/admin/logout/route.ts`
- Create: `bht-hub/src/app/api/admin/session/route.ts`
- Create: `bht-hub/src/components/ui/button.tsx`, `card.tsx`, `input.tsx`, `label.tsx`, `select.tsx`, `badge.tsx`, `separator.tsx`, `tabs.tsx`, `switch.tsx`
- Create: `bht-hub/src/lib/utils.ts`

**Interfaces:**
- Produces: `ADMIN_COOKIE_NAME`, `verifySessionToken`, `verifyPassword`, `createSessionToken` (unchanged from `dashboard/`) — consumed by every route added in Tasks 7-9. Requires `ADMIN_PASSWORD` and `SESSION_SECRET` env vars, same as Hansel.

- [ ] **Step 1: Copy auth lib and shadcn UI primitives verbatim (no BHT-specific logic in these — safe to copy without a test)**

```bash
mkdir -p bht-hub/src/app/api/admin/login bht-hub/src/app/api/admin/logout bht-hub/src/app/api/admin/session bht-hub/src/components/ui
cp dashboard/src/lib/server/admin-auth.ts bht-hub/src/lib/server/admin-auth.ts
cp dashboard/src/lib/utils.ts bht-hub/src/lib/utils.ts
cp dashboard/src/components/ui/button.tsx bht-hub/src/components/ui/button.tsx
cp dashboard/src/components/ui/card.tsx bht-hub/src/components/ui/card.tsx
cp dashboard/src/components/ui/input.tsx bht-hub/src/components/ui/input.tsx
cp dashboard/src/components/ui/label.tsx bht-hub/src/components/ui/label.tsx
cp dashboard/src/components/ui/select.tsx bht-hub/src/components/ui/select.tsx
cp dashboard/src/components/ui/badge.tsx bht-hub/src/components/ui/badge.tsx
cp dashboard/src/components/ui/separator.tsx bht-hub/src/components/ui/separator.tsx
cp dashboard/src/components/ui/tabs.tsx bht-hub/src/components/ui/tabs.tsx
cp dashboard/src/components/ui/switch.tsx bht-hub/src/components/ui/switch.tsx
```

- [ ] **Step 2: Copy the three auth routes verbatim (identical logic to Hansel — same cookie/session mechanism)**

```typescript
// bht-hub/src/app/api/admin/login/route.ts
import { NextRequest, NextResponse } from "next/server";
import { cookies } from "next/headers";

import {
  ADMIN_COOKIE_NAME,
  SESSION_MAX_AGE_SECONDS,
  createSessionToken,
  verifyPassword,
} from "@/lib/server/admin-auth";

export async function POST(request: NextRequest) {
  let password = "";
  try {
    const body = await request.json();
    if (typeof body?.password === "string") password = body.password;
  } catch {
    return NextResponse.json({ error: "Invalid request" }, { status: 400 });
  }

  if (!verifyPassword(password)) {
    return NextResponse.json({ error: "Incorrect password" }, { status: 401 });
  }

  const cookieStore = await cookies();
  cookieStore.set(ADMIN_COOKIE_NAME, createSessionToken(), {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    sameSite: "strict",
    path: "/",
    maxAge: SESSION_MAX_AGE_SECONDS,
  });

  return NextResponse.json({ ok: true });
}
```

```typescript
// bht-hub/src/app/api/admin/logout/route.ts
import { NextResponse } from "next/server";
import { cookies } from "next/headers";

import { ADMIN_COOKIE_NAME } from "@/lib/server/admin-auth";

export async function POST() {
  const cookieStore = await cookies();
  cookieStore.delete(ADMIN_COOKIE_NAME);
  return NextResponse.json({ ok: true });
}
```

```typescript
// bht-hub/src/app/api/admin/session/route.ts
import { NextResponse } from "next/server";
import { cookies } from "next/headers";

import { ADMIN_COOKIE_NAME, verifySessionToken } from "@/lib/server/admin-auth";

export async function GET() {
  const cookieStore = await cookies();
  const token = cookieStore.get(ADMIN_COOKIE_NAME)?.value;
  return NextResponse.json({ authenticated: verifySessionToken(token) });
}
```

- [ ] **Step 3: Manual verification (no automated test — this is a verbatim copy of already-proven logic; the check is that it wires up in this new app)**

```bash
cd bht-hub && ADMIN_PASSWORD=test-pass SESSION_SECRET=test-secret npm run dev
```

In another terminal:
```bash
curl -i -c /tmp/bht-hub-cookies.txt -X POST http://localhost:3100/api/admin/login -H "Content-Type: application/json" -d '{"password":"test-pass"}'
curl -s -b /tmp/bht-hub-cookies.txt http://localhost:3100/api/admin/session
```
Expected: login returns `{"ok":true}` with a `Set-Cookie` header; session check returns `{"authenticated":true}`.

- [ ] **Step 4: Commit**

```bash
cd bht-hub && git add src/lib/server/admin-auth.ts src/lib/utils.ts src/components/ui src/app/api/admin/login src/app/api/admin/logout src/app/api/admin/session
git commit -m "feat(bht-hub): port admin session auth and shadcn ui primitives"
```

---

### Task 7: Project management API routes

**Files:**
- Create: `bht-hub/src/app/api/admin/projects/route.ts`

**Interfaces:**
- Consumes: `createProject`, `listProjects` (Task 5), `verifySessionToken`/`ADMIN_COOKIE_NAME` (Task 6).
- Produces: `GET /api/admin/projects` → `{ projects: Project[] }`; `POST /api/admin/projects` with body `{ clientName: string, dataType: "progress" | "final" }` → `{ project: Project }` — consumed by the admin UI in Task 10.

- [ ] **Step 1: Implement (route handlers are integration-tested manually in Step 2 — Task 5's unit tests already cover the create/list logic this route calls, so this task only needs to prove the HTTP wiring, not re-test the domain logic)**

```typescript
// bht-hub/src/app/api/admin/projects/route.ts
import { NextRequest, NextResponse } from "next/server";
import { cookies } from "next/headers";

import { ADMIN_COOKIE_NAME, verifySessionToken } from "@/lib/server/admin-auth";
import { createProject, listProjects } from "@/lib/server/project-store";
import type { DataType } from "@/lib/project";

async function requireAuth(): Promise<boolean> {
  const cookieStore = await cookies();
  return verifySessionToken(cookieStore.get(ADMIN_COOKIE_NAME)?.value);
}

export async function GET() {
  if (!(await requireAuth())) {
    return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
  }
  const projects = await listProjects();
  return NextResponse.json({ projects });
}

export async function POST(request: NextRequest) {
  if (!(await requireAuth())) {
    return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
  }

  let clientName: string;
  let dataType: DataType;
  try {
    const body = await request.json();
    clientName = body?.clientName;
    dataType = body?.dataType;
  } catch {
    return NextResponse.json({ error: "Invalid request" }, { status: 400 });
  }

  if (typeof clientName !== "string" || clientName.trim().length === 0) {
    return NextResponse.json({ error: "clientName is required" }, { status: 400 });
  }
  if (dataType !== "progress" && dataType !== "final") {
    return NextResponse.json({ error: "dataType must be 'progress' or 'final'" }, { status: 400 });
  }

  const project = await createProject(clientName.trim(), dataType);
  return NextResponse.json({ project }, { status: 201 });
}
```

- [ ] **Step 2: Manual verification**

```bash
curl -s -b /tmp/bht-hub-cookies.txt -X POST http://localhost:3100/api/admin/projects -H "Content-Type: application/json" -d '{"clientName":"Great Eastern 2026","dataType":"progress"}'
curl -s -b /tmp/bht-hub-cookies.txt http://localhost:3100/api/admin/projects
```
Expected: POST returns `{"project":{"slug":"great-eastern-2026","clientName":"Great Eastern 2026","dataType":"progress","status":"empty",...}}` with 201; GET returns that project in the `projects` array.

- [ ] **Step 3: Commit**

```bash
cd bht-hub && git add src/app/api/admin/projects
git commit -m "feat(bht-hub): add project create/list API"
```

---

### Task 8: Slug-namespaced upload flow with question-code validation

**Files:**
- Create: `bht-hub/src/app/api/admin/projects/[slug]/upload-chunk/route.ts`
- Create: `bht-hub/src/app/api/admin/projects/[slug]/upload-finish/route.ts`

**Interfaces:**
- Consumes: `saveChunk`, `assembleChunks`, `deleteChunks` (Task 4), `validateQuestionCodes`, `processWorkbook` (Task 2), `getProject`, `saveProjectData`, `updateProjectStatus` (Task 5).
- Produces: `POST /api/admin/projects/[slug]/upload-chunk` (same header protocol as Hansel: `x-upload-session`, `x-upload-index`, raw body) and `POST /api/admin/projects/[slug]/upload-finish` with body `{ sessionId, chunkCount }` → on success `{ ok: true, summary: ProcessSummary, codeDetections: CodeDetection[] }`, on missing codes `422 { error, missingCodes: string[] }` — consumed by the admin upload UI in Task 11.

Locked (`final`, published) projects reject new uploads with `409` until explicitly unlocked (Task 9). Progress-type projects always accept re-uploads. On a successful upload the project moves from `empty`/`active` to `active` and only the dataset matching its `dataType` is persisted (`progressRaw` for `"progress"`, `cleaned` for `"final"`) — never both, since Task 5's storage holds one `data.json` per project.

- [ ] **Step 1: Implement `upload-chunk` (identical to Hansel's, slug is path-only bookkeeping — chunks are already isolated by the client-generated `sessionId`, so no per-slug branching is needed here)**

```bash
mkdir -p "bht-hub/src/app/api/admin/projects/[slug]/upload-chunk" "bht-hub/src/app/api/admin/projects/[slug]/upload-finish"
```

```typescript
// bht-hub/src/app/api/admin/projects/[slug]/upload-chunk/route.ts
import { NextRequest, NextResponse } from "next/server";
import { cookies } from "next/headers";

import { ADMIN_COOKIE_NAME, verifySessionToken } from "@/lib/server/admin-auth";
import { saveChunk } from "@/lib/server/blob-store";

export async function POST(request: NextRequest) {
  const cookieStore = await cookies();
  if (!verifySessionToken(cookieStore.get(ADMIN_COOKIE_NAME)?.value)) {
    return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
  }

  const sessionId = request.headers.get("x-upload-session");
  const indexHeader = request.headers.get("x-upload-index");
  const index = indexHeader ? Number(indexHeader) : NaN;
  if (!sessionId || !/^[a-zA-Z0-9-]+$/.test(sessionId) || !Number.isInteger(index) || index < 0) {
    return NextResponse.json({ error: "Missing or invalid chunk metadata" }, { status: 400 });
  }

  try {
    const data = await request.arrayBuffer();
    await saveChunk(sessionId, index, data);
    return NextResponse.json({ ok: true });
  } catch (err) {
    console.error("Failed to save upload chunk", err);
    const message = err instanceof Error ? err.message : "Failed to save chunk.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
```

- [ ] **Step 2: Implement `upload-finish` with the validate-before-save gate and dataType-specific persistence**

```typescript
// bht-hub/src/app/api/admin/projects/[slug]/upload-finish/route.ts
import { NextRequest, NextResponse } from "next/server";
import { cookies } from "next/headers";

import { ADMIN_COOKIE_NAME, verifySessionToken } from "@/lib/server/admin-auth";
import { processWorkbook, validateQuestionCodes } from "@/lib/server/build-data";
import { assembleChunks, deleteChunks } from "@/lib/server/blob-store";
import { getProject, saveProjectData, updateProjectStatus } from "@/lib/server/project-store";

export async function POST(
  request: NextRequest,
  { params }: { params: Promise<{ slug: string }> }
) {
  const cookieStore = await cookies();
  if (!verifySessionToken(cookieStore.get(ADMIN_COOKIE_NAME)?.value)) {
    return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
  }

  const { slug } = await params;
  const project = await getProject(slug);
  if (!project) {
    return NextResponse.json({ error: `Project '${slug}' not found` }, { status: 404 });
  }
  if (project.status === "locked") {
    return NextResponse.json(
      { error: "This project is published as Final and locked. Unlock it first to upload a correction." },
      { status: 409 }
    );
  }

  let sessionId: string;
  let chunkCount: number;
  try {
    const body = await request.json();
    sessionId = body?.sessionId;
    chunkCount = body?.chunkCount;
  } catch {
    return NextResponse.json({ error: "Invalid request" }, { status: 400 });
  }

  if (
    typeof sessionId !== "string" ||
    !/^[a-zA-Z0-9-]+$/.test(sessionId) ||
    !Number.isInteger(chunkCount) ||
    chunkCount <= 0
  ) {
    return NextResponse.json({ error: "Missing or invalid session metadata" }, { status: 400 });
  }

  try {
    const buffer = await assembleChunks(sessionId, chunkCount);

    const codeDetections = await validateQuestionCodes(buffer);
    const missingCodes = codeDetections.filter((d) => !d.found).map((d) => d.code);
    if (missingCodes.length > 0) {
      return NextResponse.json(
        {
          error: "This file doesn't match the BHT question set. Check the sheet and column headers.",
          missingCodes,
          codeDetections,
        },
        { status: 422 }
      );
    }

    const { progressRaw, cleaned, summary } = await processWorkbook(buffer);
    await saveProjectData(slug, project.dataType === "progress" ? progressRaw : cleaned);
    if (project.status === "empty") {
      await updateProjectStatus(slug, "active");
    }

    return NextResponse.json({ ok: true, summary, codeDetections });
  } catch (err) {
    console.error("Failed to process uploaded workbook", err);
    const message = err instanceof Error ? err.message : "Failed to process the uploaded file.";
    return NextResponse.json({ error: message }, { status: 422 });
  } finally {
    await deleteChunks(sessionId, chunkCount).catch((err) =>
      console.error("Failed to clean up upload chunks", err)
    );
  }
}
```

- [ ] **Step 3: Manual verification against a real BHT export**

```bash
curl -s -b /tmp/bht-hub-cookies.txt -X POST "http://localhost:3100/api/admin/projects/great-eastern-2026/upload-chunk" \
  -H "x-upload-session: test-session-1" -H "x-upload-index: 0" -H "Content-Type: application/octet-stream" \
  --data-binary @"Data Source/P. Hansel BHT - Progress Report - V2 (16072026)-Checked RF.xlsx"
curl -s -b /tmp/bht-hub-cookies.txt -X POST "http://localhost:3100/api/admin/projects/great-eastern-2026/upload-finish" \
  -H "Content-Type: application/json" -d '{"sessionId":"test-session-1","chunkCount":1}'
```
Expected: `upload-finish` returns 200 with `summary.progressRows > 0` and every `codeDetections[].found === true` (this is the real Hansel export, so all 19 BHT codes must be present). Then re-run `GET /api/admin/projects` and confirm this project's `status` is now `"active"`.

Also verify the rejection path with a file that isn't the BHT format (any unrelated `.xlsx`) — expect `422` with a non-empty `missingCodes` array.

- [ ] **Step 4: Commit**

```bash
cd bht-hub && git add "src/app/api/admin/projects/[slug]/upload-chunk" "src/app/api/admin/projects/[slug]/upload-finish"
git commit -m "feat(bht-hub): add slug-scoped upload flow with question-code validation"
```

---

### Task 9: Publish / unlock routes for Final-type projects

**Files:**
- Create: `bht-hub/src/app/api/admin/projects/[slug]/publish/route.ts`
- Create: `bht-hub/src/app/api/admin/projects/[slug]/unlock/route.ts`

**Interfaces:**
- Consumes: `getProject`, `updateProjectStatus` (Task 5).
- Produces: `POST /api/admin/projects/[slug]/publish` (`active` → `locked`, only for `dataType: "final"`) and `POST /api/admin/projects/[slug]/unlock` (`locked` → `active`) — consumed by the admin project detail UI in Task 11.

- [ ] **Step 1: Implement `publish`**

```bash
mkdir -p "bht-hub/src/app/api/admin/projects/[slug]/publish" "bht-hub/src/app/api/admin/projects/[slug]/unlock"
```

```typescript
// bht-hub/src/app/api/admin/projects/[slug]/publish/route.ts
import { NextResponse } from "next/server";
import { cookies } from "next/headers";

import { ADMIN_COOKIE_NAME, verifySessionToken } from "@/lib/server/admin-auth";
import { getProject, updateProjectStatus } from "@/lib/server/project-store";

export async function POST(
  _request: Request,
  { params }: { params: Promise<{ slug: string }> }
) {
  const cookieStore = await cookies();
  if (!verifySessionToken(cookieStore.get(ADMIN_COOKIE_NAME)?.value)) {
    return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
  }

  const { slug } = await params;
  const project = await getProject(slug);
  if (!project) {
    return NextResponse.json({ error: `Project '${slug}' not found` }, { status: 404 });
  }
  if (project.dataType !== "final") {
    return NextResponse.json({ error: "Only Final-type projects can be published/locked" }, { status: 400 });
  }
  if (project.status !== "active") {
    return NextResponse.json(
      { error: `Cannot publish a project with status '${project.status}'. Upload data first.` },
      { status: 409 }
    );
  }

  const updated = await updateProjectStatus(slug, "locked");
  return NextResponse.json({ project: updated });
}
```

- [ ] **Step 2: Implement `unlock`**

```typescript
// bht-hub/src/app/api/admin/projects/[slug]/unlock/route.ts
import { NextResponse } from "next/server";
import { cookies } from "next/headers";

import { ADMIN_COOKIE_NAME, verifySessionToken } from "@/lib/server/admin-auth";
import { getProject, updateProjectStatus } from "@/lib/server/project-store";

export async function POST(
  _request: Request,
  { params }: { params: Promise<{ slug: string }> }
) {
  const cookieStore = await cookies();
  if (!verifySessionToken(cookieStore.get(ADMIN_COOKIE_NAME)?.value)) {
    return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
  }

  const { slug } = await params;
  const project = await getProject(slug);
  if (!project) {
    return NextResponse.json({ error: `Project '${slug}' not found` }, { status: 404 });
  }
  if (project.status !== "locked") {
    return NextResponse.json({ error: "Project is not locked" }, { status: 409 });
  }

  const updated = await updateProjectStatus(slug, "active");
  return NextResponse.json({ project: updated });
}
```

- [ ] **Step 3: Manual verification**

```bash
curl -s -b /tmp/bht-hub-cookies.txt -X POST http://localhost:3100/api/admin/projects -H "Content-Type: application/json" -d '{"clientName":"Vertex Final","dataType":"final"}'
# upload a real BHT file to vertex-final via upload-chunk/upload-finish (Task 8 steps, slug=vertex-final), then:
curl -s -b /tmp/bht-hub-cookies.txt -X POST http://localhost:3100/api/admin/projects/vertex-final/publish
curl -s -b /tmp/bht-hub-cookies.txt -X POST "http://localhost:3100/api/admin/projects/vertex-final/upload-chunk" -H "x-upload-session: s2" -H "x-upload-index: 0" --data-binary "x"
```
Expected: `publish` returns `{"project":{...,"status":"locked"}}`; the subsequent `upload-chunk` call is unaffected (chunk storage doesn't check status — this is fine, it's `upload-finish` that enforces the lock) but re-running Task 8's `upload-finish` verification against `vertex-final` now returns `409`. Then `unlock` returns `status: "active"` and uploads work again.

- [ ] **Step 4: Commit**

```bash
cd bht-hub && git add "src/app/api/admin/projects/[slug]/publish" "src/app/api/admin/projects/[slug]/unlock"
git commit -m "feat(bht-hub): add publish/unlock lifecycle for Final-type projects"
```

---

### Task 10: Admin UI — project list and create wizard

**Files:**
- Create: `bht-hub/src/app/admin/page.tsx`
- Create: `bht-hub/src/app/admin/admin-shell.tsx`
- Create: `bht-hub/src/components/admin/project-list.tsx`
- Create: `bht-hub/src/components/admin/create-project-form.tsx`

**Interfaces:**
- Consumes: `GET /api/admin/projects`, `POST /api/admin/projects`, `GET /api/admin/session`, `POST /api/admin/login`, `POST /api/admin/logout` (Tasks 6-7).
- Produces: the `/admin` page — login gate, then a project table with a copy-link button per row and a "New project" form. Links to `/admin/projects/[slug]` (Task 11).

- [ ] **Step 1: Build the login gate + project list/create shell**

```typescript
// bht-hub/src/app/admin/admin-shell.tsx
"use client";

import * as React from "react";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { CreateProjectForm } from "@/components/admin/create-project-form";
import { ProjectList } from "@/components/admin/project-list";
import type { Project } from "@/lib/project";

type Status = "checking" | "login" | "ready";

export function AdminShell() {
  const [status, setStatus] = React.useState<Status>("checking");
  const [password, setPassword] = React.useState("");
  const [loginError, setLoginError] = React.useState<string | null>(null);
  const [loggingIn, setLoggingIn] = React.useState(false);
  const [projects, setProjects] = React.useState<Project[]>([]);

  const loadProjects = React.useCallback(async () => {
    const res = await fetch("/api/admin/projects");
    if (res.ok) {
      const data = await res.json();
      setProjects(data.projects);
    }
  }, []);

  React.useEffect(() => {
    let cancelled = false;
    fetch("/api/admin/session")
      .then((res) => res.json())
      .then((data) => {
        if (cancelled) return;
        setStatus(data.authenticated ? "ready" : "login");
      })
      .catch(() => {
        if (!cancelled) setStatus("login");
      });
    return () => {
      cancelled = true;
    };
  }, []);

  React.useEffect(() => {
    if (status === "ready") void loadProjects();
  }, [status, loadProjects]);

  async function handleLogin(e: React.FormEvent) {
    e.preventDefault();
    setLoggingIn(true);
    setLoginError(null);
    try {
      const res = await fetch("/api/admin/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ password }),
      });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        setLoginError(data.error ?? "Incorrect password.");
        return;
      }
      setPassword("");
      setStatus("ready");
    } catch {
      setLoginError("Something went wrong. Try again.");
    } finally {
      setLoggingIn(false);
    }
  }

  async function handleLogout() {
    await fetch("/api/admin/logout", { method: "POST" }).catch(() => {});
    setStatus("login");
    setProjects([]);
  }

  if (status === "checking") {
    return <p className="text-center text-sm text-muted-foreground">Checking session…</p>;
  }

  if (status === "login") {
    return (
      <Card className="mx-auto max-w-sm border-border">
        <CardHeader>
          <CardTitle>Admin sign-in</CardTitle>
        </CardHeader>
        <CardContent>
          <form onSubmit={handleLogin} className="flex flex-col gap-3">
            <Input
              type="password"
              placeholder="Password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              autoFocus
            />
            {loginError && <p className="text-sm text-destructive">{loginError}</p>}
            <Button type="submit" disabled={loggingIn || password.length === 0}>
              {loggingIn ? "Signing in…" : "Sign in"}
            </Button>
          </form>
        </CardContent>
      </Card>
    );
  }

  return (
    <div className="flex flex-col gap-6">
      <div className="flex items-center justify-between">
        <h1 className="text-lg font-semibold">BHT Projects</h1>
        <Button variant="ghost" size="sm" onClick={handleLogout}>
          Log out
        </Button>
      </div>
      <CreateProjectForm onCreated={loadProjects} />
      <ProjectList projects={projects} />
    </div>
  );
}
```

- [ ] **Step 2: Create-project form**

```typescript
// bht-hub/src/components/admin/create-project-form.tsx
"use client";

import * as React from "react";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import type { DataType } from "@/lib/project";

export function CreateProjectForm({ onCreated }: { onCreated: () => void }) {
  const [clientName, setClientName] = React.useState("");
  const [dataType, setDataType] = React.useState<DataType>("progress");
  const [submitting, setSubmitting] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    setSubmitting(true);
    setError(null);
    try {
      const res = await fetch("/api/admin/projects", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ clientName, dataType }),
      });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        setError(data.error ?? "Failed to create project.");
        return;
      }
      setClientName("");
      onCreated();
    } catch {
      setError("Something went wrong. Try again.");
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <Card className="border-border">
      <CardHeader>
        <CardTitle>New project</CardTitle>
      </CardHeader>
      <CardContent>
        <form onSubmit={handleSubmit} className="flex flex-wrap items-end gap-3">
          <div className="flex flex-col gap-1">
            <Label htmlFor="client-name">Client name</Label>
            <Input
              id="client-name"
              value={clientName}
              onChange={(e) => setClientName(e.target.value)}
              placeholder="e.g. Great Eastern 2026"
              required
            />
          </div>
          <div className="flex flex-col gap-1">
            <Label htmlFor="data-type">Data type</Label>
            <select
              id="data-type"
              className="h-9 rounded-md border border-input bg-transparent px-3 text-sm"
              value={dataType}
              onChange={(e) => setDataType(e.target.value as DataType)}
            >
              <option value="progress">Progress (re-uploaded periodically)</option>
              <option value="final">Final (one-time, published then locked)</option>
            </select>
          </div>
          <Button type="submit" disabled={submitting || clientName.trim().length === 0}>
            {submitting ? "Creating…" : "Create project"}
          </Button>
          {error && <p className="w-full text-sm text-destructive">{error}</p>}
        </form>
      </CardContent>
    </Card>
  );
}
```

- [ ] **Step 3: Project list with copy-link**

```typescript
// bht-hub/src/components/admin/project-list.tsx
"use client";

import * as React from "react";
import Link from "next/link";

import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import type { Project } from "@/lib/project";

function CopyLinkButton({ slug }: { slug: string }) {
  const [copied, setCopied] = React.useState(false);
  return (
    <Button
      variant="outline"
      size="sm"
      onClick={() => {
        const url = `${window.location.origin}/d/${slug}`;
        navigator.clipboard.writeText(url).then(() => {
          setCopied(true);
          setTimeout(() => setCopied(false), 1500);
        });
      }}
    >
      {copied ? "Copied!" : "Copy link"}
    </Button>
  );
}

export function ProjectList({ projects }: { projects: Project[] }) {
  const sorted = [...projects].sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));

  if (sorted.length === 0) {
    return <p className="text-sm text-muted-foreground">No projects yet — create one above.</p>;
  }

  return (
    <Card className="border-border">
      <CardContent className="flex flex-col divide-y divide-border p-0">
        {sorted.map((project) => (
          <div key={project.slug} className="flex flex-wrap items-center justify-between gap-3 p-4">
            <div>
              <p className="font-medium">{project.clientName}</p>
              <p className="text-xs text-muted-foreground">/d/{project.slug}</p>
            </div>
            <div className="flex items-center gap-2">
              <Badge variant="outline">{project.dataType}</Badge>
              <Badge variant={project.status === "locked" ? "default" : "outline"}>
                {project.status}
              </Badge>
              <span className="text-xs text-muted-foreground">
                Updated {new Date(project.updatedAt).toLocaleString()}
              </span>
              <CopyLinkButton slug={project.slug} />
              <Button asChild variant="secondary" size="sm">
                <Link href={`/admin/projects/${project.slug}`}>Manage</Link>
              </Button>
              <Button asChild variant="ghost" size="sm">
                <Link href={`/d/${project.slug}`} target="_blank">
                  Preview
                </Link>
              </Button>
            </div>
          </div>
        ))}
      </CardContent>
    </Card>
  );
}
```

- [ ] **Step 4: Admin page entry**

```typescript
// bht-hub/src/app/admin/page.tsx
import type { Metadata } from "next";

import { AdminShell } from "@/app/admin/admin-shell";

export const metadata: Metadata = {
  title: "BHT Hub – Admin",
  robots: { index: false, follow: false },
};

export default function AdminPage() {
  return (
    <main className="mx-auto flex min-h-screen w-full max-w-4xl flex-col gap-6 px-6 py-12">
      <AdminShell />
    </main>
  );
}
```

- [ ] **Step 5: Manual verification in the browser**

```bash
cd bht-hub && ADMIN_PASSWORD=test-pass SESSION_SECRET=test-secret npm run dev
```
Open `http://localhost:3100/admin`, log in, confirm the projects created in Tasks 7-9 (`great-eastern-2026`, `vertex-final`) show up with correct badges, create one more via the form, click "Copy link" and confirm the clipboard holds `http://localhost:3100/d/<slug>`.

- [ ] **Step 6: Commit**

```bash
cd bht-hub && git add src/app/admin src/components/admin
git commit -m "feat(bht-hub): add admin project list and create-project UI"
```

---

### Task 11: Admin UI — project detail (upload, publish/unlock)

**Files:**
- Create: `bht-hub/src/app/admin/projects/[slug]/page.tsx`
- Create: `bht-hub/src/components/admin/project-detail.tsx`

**Interfaces:**
- Consumes: `GET /api/admin/projects` (to resolve the one project by slug, reusing the list endpoint — no new GET-by-slug endpoint needed for a few dozen projects), the upload/publish/unlock routes (Tasks 8-9).
- Produces: the `/admin/projects/[slug]` page — upload form with the same chunked-upload client logic as Hansel's `AdminPanel`, an upload-summary card that also lists rejected/missing question codes on failure, and Publish/Unlock buttons gated by `dataType`/`status`.

- [ ] **Step 1: Build the detail component**

```typescript
// bht-hub/src/components/admin/project-detail.tsx
"use client";

import * as React from "react";

import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import type { Project } from "@/lib/project";

const CHUNK_SIZE = 1.5 * 1024 * 1024;

interface UploadResult {
  summary?: { progressRows?: number; totalRespondents?: number; cleanedRespondents?: number };
  error?: string;
  missingCodes?: string[];
}

export function ProjectDetail({ project: initial }: { project: Project }) {
  const [project, setProject] = React.useState(initial);
  const [file, setFile] = React.useState<File | null>(null);
  const [uploading, setUploading] = React.useState(false);
  const [progressLabel, setProgressLabel] = React.useState<string | null>(null);
  const [result, setResult] = React.useState<UploadResult | null>(null);
  const [lifecycleBusy, setLifecycleBusy] = React.useState(false);
  const [lifecycleError, setLifecycleError] = React.useState<string | null>(null);

  async function handleUpload(e: React.FormEvent) {
    e.preventDefault();
    if (!file) return;
    setUploading(true);
    setResult(null);
    try {
      const sessionId = crypto.randomUUID();
      const chunkCount = Math.max(1, Math.ceil(file.size / CHUNK_SIZE));

      for (let index = 0; index < chunkCount; index++) {
        setProgressLabel(`Uploading part ${index + 1} of ${chunkCount}…`);
        const chunk = file.slice(index * CHUNK_SIZE, (index + 1) * CHUNK_SIZE);
        const res = await fetch(`/api/admin/projects/${project.slug}/upload-chunk`, {
          method: "POST",
          headers: {
            "x-upload-session": sessionId,
            "x-upload-index": String(index),
            "Content-Type": "application/octet-stream",
          },
          body: chunk,
        });
        if (!res.ok) {
          const data = await res.json().catch(() => ({}));
          setResult({ error: data.error ?? "Upload failed while sending file." });
          return;
        }
      }

      setProgressLabel("Processing spreadsheet…");
      const finishRes = await fetch(`/api/admin/projects/${project.slug}/upload-finish`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sessionId, chunkCount }),
      });
      const data = await finishRes.json().catch(() => ({}));
      if (!finishRes.ok) {
        setResult({ error: data.error, missingCodes: data.missingCodes });
        return;
      }
      setResult({ summary: data.summary });
      setProject((p) => ({ ...p, status: p.status === "empty" ? "active" : p.status }));
    } catch {
      setResult({ error: "Something went wrong while uploading. Try again." });
    } finally {
      setUploading(false);
      setProgressLabel(null);
    }
  }

  async function handlePublish() {
    setLifecycleBusy(true);
    setLifecycleError(null);
    try {
      const res = await fetch(`/api/admin/projects/${project.slug}/publish`, { method: "POST" });
      const data = await res.json();
      if (!res.ok) {
        setLifecycleError(data.error);
        return;
      }
      setProject(data.project);
    } finally {
      setLifecycleBusy(false);
    }
  }

  async function handleUnlock() {
    setLifecycleBusy(true);
    setLifecycleError(null);
    try {
      const res = await fetch(`/api/admin/projects/${project.slug}/unlock`, { method: "POST" });
      const data = await res.json();
      if (!res.ok) {
        setLifecycleError(data.error);
        return;
      }
      setProject(data.project);
    } finally {
      setLifecycleBusy(false);
    }
  }

  return (
    <div className="flex flex-col gap-6">
      <div className="flex items-center gap-3">
        <h1 className="text-lg font-semibold">{project.clientName}</h1>
        <Badge variant="outline">{project.dataType}</Badge>
        <Badge variant={project.status === "locked" ? "default" : "outline"}>{project.status}</Badge>
      </div>

      <Card className="border-border">
        <CardHeader>
          <CardTitle>Upload data</CardTitle>
        </CardHeader>
        <CardContent>
          {project.status === "locked" ? (
            <p className="text-sm text-muted-foreground">
              This project is published and locked. Unlock it below to upload a correction.
            </p>
          ) : (
            <form onSubmit={handleUpload} className="flex flex-col gap-3">
              <Input type="file" accept=".xlsx" onChange={(e) => setFile(e.target.files?.[0] ?? null)} />
              <Button type="submit" disabled={!file || uploading}>
                {uploading ? (progressLabel ?? "Processing…") : "Upload Excel"}
              </Button>
            </form>
          )}

          {result?.error && (
            <div className="mt-4 rounded-lg border border-destructive/40 bg-destructive/5 p-3 text-sm">
              <p className="font-medium text-destructive">{result.error}</p>
              {result.missingCodes && result.missingCodes.length > 0 && (
                <p className="mt-1 text-muted-foreground">
                  Missing question codes: {result.missingCodes.join(", ")}
                </p>
              )}
            </div>
          )}

          {result?.summary && (
            <div className="mt-4 rounded-lg border border-border bg-muted/40 p-3 text-sm">
              <p className="font-medium">Upload processed.</p>
              {project.dataType === "progress" ? (
                <p className="text-muted-foreground">{result.summary.progressRows} progress rows detected.</p>
              ) : (
                <p className="text-muted-foreground">
                  {result.summary.cleanedRespondents} of {result.summary.totalRespondents} respondents passed QC.
                </p>
              )}
            </div>
          )}
        </CardContent>
      </Card>

      {project.dataType === "final" && (
        <Card className="border-border">
          <CardHeader>
            <CardTitle>Publish</CardTitle>
          </CardHeader>
          <CardContent className="flex flex-col gap-2">
            {lifecycleError && <p className="text-sm text-destructive">{lifecycleError}</p>}
            {project.status === "active" && (
              <Button onClick={handlePublish} disabled={lifecycleBusy}>
                Publish as Final (locks the project)
              </Button>
            )}
            {project.status === "locked" && (
              <Button variant="outline" onClick={handleUnlock} disabled={lifecycleBusy}>
                Unlock for correction
              </Button>
            )}
            {project.status === "empty" && (
              <p className="text-sm text-muted-foreground">Upload data before publishing.</p>
            )}
          </CardContent>
        </Card>
      )}
    </div>
  );
}
```

- [ ] **Step 2: Page that resolves the project by slug and renders 404 if missing**

```typescript
// bht-hub/src/app/admin/projects/[slug]/page.tsx
import { notFound, redirect } from "next/navigation";
import { cookies } from "next/headers";

import { ProjectDetail } from "@/components/admin/project-detail";
import { ADMIN_COOKIE_NAME, verifySessionToken } from "@/lib/server/admin-auth";
import { getProject } from "@/lib/server/project-store";

export default async function AdminProjectPage({
  params,
}: {
  params: Promise<{ slug: string }>;
}) {
  const cookieStore = await cookies();
  if (!verifySessionToken(cookieStore.get(ADMIN_COOKIE_NAME)?.value)) {
    redirect("/admin");
  }

  const { slug } = await params;
  const project = await getProject(slug);
  if (!project) notFound();

  return (
    <main className="mx-auto flex min-h-screen w-full max-w-3xl flex-col gap-6 px-6 py-12">
      <ProjectDetail project={project} />
    </main>
  );
}
```

This page is gated the same way the Global Constraints require for every admin route: a missing/invalid session cookie redirects to `/admin` (which shows the login form) instead of rendering project details. The mutating actions inside `ProjectDetail` (upload, publish, unlock) are still independently protected by their own API routes' session checks — this page-level gate is defense-in-depth, not a substitute for it.

- [ ] **Step 3: Manual verification**

Visit `http://localhost:3100/admin/projects/great-eastern-2026`, upload the real BHT export, confirm the "progress rows detected" summary renders. Visit `http://localhost:3100/admin/projects/vertex-final`, publish it, confirm the Upload card switches to the locked message and the Unlock button appears.

- [ ] **Step 4: Commit**

```bash
cd bht-hub && git add "src/app/admin/projects/[slug]" src/components/admin/project-detail.tsx
git commit -m "feat(bht-hub): add project detail admin page with upload and publish/unlock"
```

---

### Task 12: Public data API and single-dataset dashboard context

**Files:**
- Create: `bht-hub/src/app/api/d/[slug]/route.ts`
- Create: `bht-hub/src/context/dashboard-context.tsx`
- Create: `bht-hub/src/lib/tally.ts`
- Create: `bht-hub/src/lib/chart-colors.ts`
- Create: `bht-hub/src/lib/progress-lookup.ts`

**Interfaces:**
- Consumes: `getProject`, `loadProjectData` (Task 5).
- Produces: `GET /api/d/[slug]` → `{ project: Project, data: ProgressData | CleanedData | null }` (data is `null` when `status === "empty"`); a `DashboardProvider`/`useDashboard()` context that fetches this single endpoint and exposes `{ loading, error, project, progressData, cleanedData, cleanedFilters, setCleanedFilter, clearFilters, filteredCleanedRows, kpi, showAsPercent, setShowAsPercent }` — consumed by Task 13's dashboard components.

Unlike Hansel's context (which always has both datasets), `progressData`/`cleanedData` are mutually exclusive here based on `project.dataType`. `kpi.target`/`kpi.remaining` are `null` whenever there's no progress dataset to derive a target from (i.e. always for `final`-type projects) — `KpiRow`'s existing `formatValue` already renders `null` as `"—"`, so no change needed there.

- [ ] **Step 1: Copy the two pure-logic helpers unchanged (no BHT-specific behavior differs here)**

```bash
mkdir -p "bht-hub/src/app/api/d/[slug]" bht-hub/src/context
cp dashboard/src/lib/tally.ts bht-hub/src/lib/tally.ts
cp dashboard/src/lib/chart-colors.ts bht-hub/src/lib/chart-colors.ts
cp dashboard/src/lib/progress-lookup.ts bht-hub/src/lib/progress-lookup.ts
```

- [ ] **Step 2: Public data route**

```typescript
// bht-hub/src/app/api/d/[slug]/route.ts
import { NextResponse } from "next/server";

import { getProject, loadProjectData } from "@/lib/server/project-store";

export const dynamic = "force-dynamic";

export async function GET(
  _request: Request,
  { params }: { params: Promise<{ slug: string }> }
) {
  const { slug } = await params;
  const project = await getProject(slug);
  if (!project) {
    return NextResponse.json({ error: "Not found" }, { status: 404 });
  }

  const data = project.status === "empty" ? null : await loadProjectData(slug);
  return NextResponse.json({ project, data }, { headers: { "Cache-Control": "no-store" } });
}
```

- [ ] **Step 3: Dashboard context, adapted from Hansel's to fetch one endpoint and branch on `dataType`**

```typescript
// bht-hub/src/context/dashboard-context.tsx
"use client";

import * as React from "react";

import { getCleanedTarget } from "@/lib/progress-lookup";
import type { Project } from "@/lib/project";
import type {
  CleanedData,
  CleanedFacet,
  CleanedFilters,
  CleanedRow,
  ProgressData,
} from "@/lib/types";

interface DashboardContextValue {
  loading: boolean;
  error: string | null;
  project: Project | null;

  progressData: ProgressData | null;
  cleanedData: CleanedData | null;

  cleanedFilters: CleanedFilters;
  setCleanedFilter: (facet: CleanedFacet, value: string | undefined) => void;
  clearFilters: () => void;
  filteredCleanedRows: CleanedRow[];

  kpi: { target: number | null; achievement: number; remaining: number | null };

  showAsPercent: boolean;
  setShowAsPercent: (value: boolean) => void;
}

const DashboardContext = React.createContext<DashboardContextValue | null>(null);

export function useDashboard(): DashboardContextValue {
  const ctx = React.useContext(DashboardContext);
  if (!ctx) throw new Error("useDashboard must be used within a DashboardProvider");
  return ctx;
}

export function DashboardProvider({ slug, children }: { slug: string; children: React.ReactNode }) {
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const [project, setProject] = React.useState<Project | null>(null);
  const [progressData, setProgressData] = React.useState<ProgressData | null>(null);
  const [cleanedData, setCleanedData] = React.useState<CleanedData | null>(null);

  const [cleanedFilters, setCleanedFilters] = React.useState<CleanedFilters>({});
  const [showAsPercent, setShowAsPercent] = React.useState(false);

  React.useEffect(() => {
    let cancelled = false;

    async function load() {
      try {
        const res = await fetch(`/api/d/${slug}`, { cache: "no-store" });
        if (!res.ok) throw new Error("Failed to load dashboard data.");
        const body = await res.json();
        if (cancelled) return;

        setProject(body.project);
        if (body.project?.dataType === "progress") {
          setProgressData(body.data);
        } else if (body.project?.dataType === "final") {
          setCleanedData(body.data);
        }
      } catch (err) {
        if (!cancelled) {
          setError(err instanceof Error ? err.message : "Unknown error loading data.");
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    }

    void load();
    return () => {
      cancelled = true;
    };
  }, [slug]);

  const setCleanedFilter = React.useCallback((facet: CleanedFacet, value: string | undefined) => {
    setCleanedFilters((prev) => {
      const next = { ...prev };
      if (value === undefined || value === "") delete next[facet];
      else next[facet] = value;
      return next;
    });
  }, []);

  const clearFilters = React.useCallback(() => setCleanedFilters({}), []);

  const filteredCleanedRows = React.useMemo(() => {
    const rows = cleanedData?.rows ?? [];
    const entries = Object.entries(cleanedFilters) as [CleanedFacet, string | undefined][];
    const active = entries.filter(([, v]) => v != null && v !== "");
    if (active.length === 0) return rows;
    return rows.filter((row) =>
      active.every(([facet, value]) => {
        const rowValue = row[facet];
        return Array.isArray(rowValue) ? rowValue.includes(value as string) : rowValue === value;
      })
    );
  }, [cleanedData, cleanedFilters]);

  const kpi = React.useMemo(() => {
    if (project?.dataType === "final") {
      return { target: null, achievement: filteredCleanedRows.length, remaining: null };
    }
    const achievement = progressData?.total.achievement ?? 0;
    const target = progressData?.total.target ?? null;
    return {
      target,
      achievement,
      remaining: target != null ? Math.max(target - achievement, 0) : null,
    };
  }, [project, progressData, filteredCleanedRows]);

  const value: DashboardContextValue = {
    loading,
    error,
    project,
    progressData,
    cleanedData,
    cleanedFilters,
    setCleanedFilter,
    clearFilters,
    filteredCleanedRows,
    kpi,
    showAsPercent,
    setShowAsPercent,
  };

  return <DashboardContext.Provider value={value}>{children}</DashboardContext.Provider>;
}
```

Note: this drops Hansel's `getCleanedTarget`-based per-filter target lookup (it depends on having *both* datasets to cross-reference) — a `"progress"`-type project's `kpi.target` comes straight from its own `progressRaw.total.target`, and a `"final"`-type project has no target/quota concept at all, matching the design spec's decision that KPI shows `"Total Respondents"`-style achievement-only framing for Final projects. `getCleanedTarget` stays copied in `progress-lookup.ts` only because `isNumeric`/`CLEANED_TO_PROGRESS_CATEGORY` there are harmless to keep; it is simply unused in this app — confirm no lint error is raised for the unused export (it's an export, not a local unused variable, so ESLint's `no-unused-vars` won't flag it).

- [ ] **Step 4: Manual verification**

```bash
curl -s http://localhost:3100/api/d/great-eastern-2026 | head -c 300
curl -s http://localhost:3100/api/d/does-not-exist
```
Expected: first returns `{"project":{...,"dataType":"progress","status":"active"},"data":{...}}`; second returns 404.

- [ ] **Step 5: Commit**

```bash
cd bht-hub && git add "src/app/api/d/[slug]" src/context src/lib/tally.ts src/lib/chart-colors.ts src/lib/progress-lookup.ts
git commit -m "feat(bht-hub): add public per-project data API and dashboard context"
```

---

### Task 13: Public dashboard — data-type-aware view (`/d/[slug]`)

**Files:**
- Create: `bht-hub/src/components/dashboard/dashboard-header.tsx`
- Create: `bht-hub/src/components/dashboard/kpi-row.tsx`
- Create: `bht-hub/src/components/dashboard/filter-bar.tsx`
- Create: `bht-hub/src/components/dashboard/percent-toggle.tsx`
- Create: `bht-hub/src/components/dashboard/tabs-section.tsx`
- Create: `bht-hub/src/components/dashboard/breakdown-card.tsx`
- Create: `bht-hub/src/components/dashboard/progress-breakdown-table.tsx`
- Create: `bht-hub/src/components/dashboard/dashboard-shell.tsx`
- Create: `bht-hub/src/app/d/[slug]/page.tsx`

**Interfaces:**
- Consumes: `useDashboard()` (Task 12).
- Produces: the full `/d/[slug]` experience: empty-state placeholder, Progress-type view (KPI row + new grouped breakdown table of quota rows), Final-type view (KPI row + filter bar + percent toggle + the existing four-tab cleaned breakdown, reused verbatim from Hansel), and a draft banner when a Final project is `active` (uploaded but not yet published).

This resolves a gap the design spec didn't spell out: Hansel's dashboard blends *both* datasets into one page (KPI target from Progress, breakdown tabs from Cleaned) — but a `bht-hub` project only ever has one dataset. So `"progress"`-type and `"final"`-type projects get two different chart layouts below the shared header/KPI row, not a toggle between "the same two views Hansel has."

- [ ] **Step 1: Copy the components that need no changes (cleaned-data breakdown UI, unchanged because Final-type projects use cleaned data exactly like Hansel does)**

```bash
mkdir -p bht-hub/src/components/dashboard "bht-hub/src/app/d/[slug]"
cp dashboard/src/components/dashboard/filter-bar.tsx bht-hub/src/components/dashboard/filter-bar.tsx
cp dashboard/src/components/dashboard/percent-toggle.tsx bht-hub/src/components/dashboard/percent-toggle.tsx
cp dashboard/src/components/dashboard/tabs-section.tsx bht-hub/src/components/dashboard/tabs-section.tsx
cp dashboard/src/components/dashboard/breakdown-card.tsx bht-hub/src/components/dashboard/breakdown-card.tsx
cp dashboard/src/components/dashboard/kpi-row.tsx bht-hub/src/components/dashboard/kpi-row.tsx
```

All five import `useDashboard` from `@/context/dashboard-context` and nothing else project-specific, so they work unchanged against Task 12's context (`KpiRow`'s `formatValue(null)` already renders `"—"`, which is exactly what a Final-type project's `target`/`remaining` need to show).

- [ ] **Step 2: New header, driven by the project instead of hardcoded Hansel branding**

```typescript
// bht-hub/src/components/dashboard/dashboard-header.tsx
"use client";

import { useDashboard } from "@/context/dashboard-context";

export function DashboardHeader() {
  const { project } = useDashboard();

  return (
    <header className="border-b border-border bg-[var(--surfaces-background-elevated)]">
      <div className="mx-auto flex max-w-7xl flex-col items-center gap-1 px-6 py-4 text-center">
        <h1 className="text-xl font-semibold leading-tight text-primary">
          {project?.clientName ?? "BHT Study"}
        </h1>
        <p className="text-sm text-muted-foreground">
          BHT Study
          {project?.dataType === "final" && " — Final Report"}
        </p>
      </div>
    </header>
  );
}
```

(Per-project logo upload was scoped out of this plan — the design spec lists it as optional; the header falls back to text-only branding, which is the "Random ID / no branding" fallback already covered by the approved design for projects that never set a logo. Revisit as a follow-up if a client project needs a logo.)

- [ ] **Step 3: New Progress-type breakdown table (Progress data can't be cross-filtered — surfaced honestly as a grouped table, not forced into the Cleaned-only tally charts)**

```typescript
// bht-hub/src/components/dashboard/progress-breakdown-table.tsx
"use client";

import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { useDashboard } from "@/context/dashboard-context";
import type { ProgressRow } from "@/lib/types";

function formatCell(value: number | string | null): string {
  if (value == null) return "—";
  return typeof value === "number" ? value.toLocaleString("en-US") : value;
}

export function ProgressBreakdownTable() {
  const { progressData } = useDashboard();
  const rows = progressData?.rows ?? [];

  const groups = new Map<string, ProgressRow[]>();
  for (const row of rows) {
    const bucket = groups.get(row.category) ?? [];
    bucket.push(row);
    groups.set(row.category, bucket);
  }

  return (
    <div className="flex flex-col gap-[var(--gap-normal-cards)]">
      {[...groups.entries()].map(([category, categoryRows]) => (
        <Card key={category} className="border-border">
          <CardHeader>
            <CardTitle className="text-base">{category}</CardTitle>
          </CardHeader>
          <CardContent>
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-xs text-muted-foreground">
                  <th className="py-1">Detail</th>
                  <th className="py-1 text-right">Target</th>
                  <th className="py-1 text-right">Achievement</th>
                  <th className="py-1 text-right">Shortfall</th>
                </tr>
              </thead>
              <tbody>
                {categoryRows.map((row) => (
                  <tr key={`${row.category}-${row.detail}`} className="border-t border-border">
                    <td className="py-1.5">{row.detail}</td>
                    <td className="py-1.5 text-right">{formatCell(row.target)}</td>
                    <td className="py-1.5 text-right">{formatCell(row.achievement)}</td>
                    <td className="py-1.5 text-right">{formatCell(row.shortfall)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </CardContent>
        </Card>
      ))}
    </div>
  );
}
```

- [ ] **Step 4: Dashboard shell, branching on `dataType`/`status`**

```typescript
// bht-hub/src/components/dashboard/dashboard-shell.tsx
"use client";

import { Badge } from "@/components/ui/badge";
import { DashboardHeader } from "@/components/dashboard/dashboard-header";
import { FilterBar } from "@/components/dashboard/filter-bar";
import { KpiRow } from "@/components/dashboard/kpi-row";
import { PercentToggle } from "@/components/dashboard/percent-toggle";
import { ProgressBreakdownTable } from "@/components/dashboard/progress-breakdown-table";
import { TabsSection } from "@/components/dashboard/tabs-section";
import { useDashboard } from "@/context/dashboard-context";

export function DashboardShell() {
  const { loading, error, project } = useDashboard();

  return (
    <div className="flex min-h-full flex-1 flex-col bg-background">
      <DashboardHeader />
      <main className="mx-auto flex w-full max-w-7xl flex-1 flex-col gap-[var(--gap-normal-cards)] px-6 py-6">
        {error ? (
          <div className="rounded-md border border-border bg-muted p-4 text-sm text-destructive">
            Failed to load dashboard data: {error}
          </div>
        ) : loading ? (
          <div className="flex flex-1 items-center justify-center text-sm text-muted-foreground">
            Loading dashboard data…
          </div>
        ) : project?.status === "empty" ? (
          <div className="flex flex-1 flex-col items-center justify-center gap-2 text-center">
            <p className="text-lg font-medium">Data is being prepared</p>
            <p className="text-sm text-muted-foreground">Check back soon.</p>
          </div>
        ) : (
          <>
            {project?.dataType === "final" && project.status === "active" && (
              <Badge variant="outline" className="w-fit">
                Draft — not yet published
              </Badge>
            )}
            <KpiRow />
            {project?.dataType === "final" ? (
              <>
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <FilterBar />
                  <PercentToggle />
                </div>
                <TabsSection />
              </>
            ) : (
              <ProgressBreakdownTable />
            )}
          </>
        )}
      </main>
    </div>
  );
}
```

- [ ] **Step 5: Page**

```typescript
// bht-hub/src/app/d/[slug]/page.tsx
import { DashboardShell } from "@/components/dashboard/dashboard-shell";
import { DashboardProvider } from "@/context/dashboard-context";

export default async function DashboardPage({
  params,
}: {
  params: Promise<{ slug: string }>;
}) {
  const { slug } = await params;
  return (
    <DashboardProvider slug={slug}>
      <DashboardShell />
    </DashboardProvider>
  );
}
```

- [ ] **Step 6: Browser verification (use the Browser preview tool against `bht-hub`'s dev server, port 3100)**

Check, for each state:
- `/d/does-not-exist-project` — should the empty-state text render gracefully even on a 404 project? (Currently `DashboardProvider` sets `error` on a non-OK response, so this correctly shows the "Failed to load" error box, not the empty-state — confirm that's what renders.)
- `/d/great-eastern-2026` (progress, active, has data from Task 8) — KPI row shows real target/achievement/remaining numbers, `ProgressBreakdownTable` renders grouped category cards.
- `/d/vertex-final` (final, currently locked from Task 9) — no draft badge, KPI row shows `—`/respondent-count/`—`, FilterBar + PercentToggle + TabsSection render and filtering works.
- Create one more Final-type project, upload data but don't publish — confirm the "Draft — not yet published" badge appears at `/d/<that-slug>`.
- Create a brand-new empty project and visit its `/d/<slug>` before any upload — confirm the "Data is being prepared" placeholder renders instead of an error or blank page.

- [ ] **Step 7: Commit**

```bash
cd bht-hub && git add src/components/dashboard "src/app/d/[slug]"
git commit -m "feat(bht-hub): add data-type-aware public dashboard"
```

---

### Task 14: Documentation and full-suite verification

**Files:**
- Create: `bht-hub/README.md`
- Create: `bht-hub/.env.example`

**Interfaces:**
- Consumes: nothing new — this task only documents and verifies what Tasks 1-13 built.

- [ ] **Step 1: Write `.env.example`**

```bash
# bht-hub/.env.example
ADMIN_PASSWORD=
SESSION_SECRET=
NETLIFY_BLOBS_TOKEN=
```

- [ ] **Step 2: Write `README.md`**

```markdown
# BHT Hub

Internal multi-project dashboard for the BHT survey template. Create a project, upload the client's Excel export, get a shareable no-login link at `/d/<slug>`.

## Local dev

\`\`\`bash
npm install
ADMIN_PASSWORD=changeme SESSION_SECRET=changeme npm run dev
\`\`\`

Visit `/admin` to create and manage projects.

## Data model

Each project has exactly one `dataType`:
- **progress** — re-uploaded periodically, always live immediately.
- **final** — uploaded once, then explicitly published (locked, read-only) from the project's admin page.

Every upload must match the BHT survey's fixed question-code set (`config/question-code-map.json`) — a file missing any expected code is rejected with a list of what wasn't found.

## Tests

\`\`\`bash
npm run test
\`\`\`
```

- [ ] **Step 3: Run the full automated test suite**

Run: `cd bht-hub && npm run test`
Expected: PASS — all suites from Tasks 2, 3, 4, 5 (build-data, slug, blob-store, project-store).

- [ ] **Step 4: Run a full production build**

Run: `cd bht-hub && npm run build`
Expected: build succeeds with no type errors across every route added in Tasks 6-13.

- [ ] **Step 5: Commit**

```bash
cd bht-hub && git add README.md .env.example
git commit -m "docs(bht-hub): add README and env example"
```

---

## Post-plan follow-ups (explicitly out of scope here)

- Per-project logo upload (design spec listed it as optional; header currently falls back to text-only).
- Deployment (Vercel project + `NETLIFY_BLOBS_TOKEN`/`ADMIN_PASSWORD`/`SESSION_SECRET` provisioning) — same deferred-until-go-ahead posture as the original Hansel app.
- Project archive/hide-from-list (design spec's "archive instead of delete" — the list currently always shows every project; add a `status`-adjacent `archived: boolean` field and a filter if the list grows unwieldy).
