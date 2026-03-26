# List Folder Organizer

An SPFx web part that organizes SharePoint list items into folder hierarchies based on configurable grouping rules.

## Features

- Displays all custom lists on the current site
- Supports three grouping strategies:
  - **By Date** — Year / Month / Day (1–3 levels)
  - **By Initial Letter** — First letter / First two / First three (1–3 levels)
  - **By Choice Value** — Groups items by choice field value (1 level)
- Auto-enables folder creation on lists that don't have it
- Creates missing folders automatically
- Moves items into target folders
- Shows real-time progress and activity log

## Project Structure

```
spfx-list-folder-organizer/
│
├── .env                          # SharePoint site URL (gitignored)
├── .env.example                  # Template for .env
├── .gitignore
├── gulpfile.js                   # Build pipeline, reads .env
├── package.json
├── tsconfig.json
│
├── config/                       # SPFx build configuration
│   ├── config.json               # Bundle entry points
│   ├── copy-assets.json
│   ├── deploy-azure-storage.json
│   ├── package-solution.json     # Solution ID, name, features
│   ├── serve.json                # Dev server (auto-updated from .env)
│   └── write-manifests.json
│
└── src/
    └── webparts/
        └── listFolderOrganizer/
            │
            ├── ListFolderOrganizerWebPart.ts           # Web part entry point
            ├── ListFolderOrganizerWebPart.manifest.json # Web part registration
            │
            ├── components/                             # React UI
            │   ├── ListFolderOrganizer.tsx              # Main component
            │   ├── ListFolderOrganizer.module.scss      # Scoped styles
            │   └── IListFolderOrganizerProps.ts          # Component props
            │
            ├── models/                                 # Data types
            │   ├── GroupingType.ts                      # Enums: Date, Text, Choice
            │   └── IListInfo.ts                         # List, field, status interfaces
            │
            ├── services/                               # Business logic
            │   └── SharePointService.ts                 # All SharePoint REST operations
            │
            └── loc/                                    # Localization
                ├── en-us.js
                └── mystrings.d.ts
```

### Key Files

| File | Responsibility |
|---|---|
| `ListFolderOrganizerWebPart.ts` | SPFx entry point — creates the React element and passes `spHttpClient` and `siteUrl` as props |
| `ListFolderOrganizer.tsx` | Main UI — renders a table of lists with dropdowns for grouping strategy, depth, and source field; handles the organize action |
| `SharePointService.ts` | All SharePoint REST API calls: fetch lists, fetch fields, enable folders, create folder paths, get items, move items |
| `GroupingType.ts` | Defines `GroupingStrategy` enum (Date, Text, Choice) and max depth per strategy |
| `IListInfo.ts` | Interfaces for list metadata, field metadata, and processing status |

## Setup

1. Clone the repository
2. Copy `.env.example` to `.env` and set your SharePoint site URL:
   ```
   SHAREPOINT_SITE_URL=https://yourtenant.sharepoint.com/sites/yoursite
   ```
3. Install dependencies:
   ```
   npm install
   ```
4. Run locally:
   ```
   gulp serve
   ```
   This starts a local dev server on `https://localhost:4321` and opens the SharePoint workbench in your browser.

### Using the Workbench

1. You must be **signed in** to your SharePoint Online tenant in the browser
2. The browser will warn about the self-signed certificate on first run — accept it
3. On the workbench page, click the **+** button to add a web part
4. Find **"List Folder Organizer"** under the **Advanced** category
5. Add it — you'll see the table of lists from that site with all the controls
6. The workbench talks to **real SharePoint data** — lists, fields, and items are all live
7. Keep `gulp serve` running — the workbench loads your code from `localhost`

## Build & Deploy

```bash
gulp bundle --ship
gulp package-solution --ship
```

The `.sppkg` package is output to `sharepoint/solution/spfx-list-folder-organizer.sppkg`.

Upload it to your SharePoint **App Catalog**, then add the **List Folder Organizer** web part to any page.

## How It Works

1. The web part loads all custom lists (BaseTemplate 100) from the current site
2. User selects lists, picks a grouping strategy, depth level, and source field
3. Source field dropdown filters by type (DateTime fields for date grouping, Text/Choice for others)
4. On **Organize Selected**:
   - Enables folder creation on the list if not already enabled
   - Fetches all list items
   - Calculates target folder path per item (e.g. `2024/03 - March/15`)
   - Creates any missing folders in the hierarchy
   - Moves items that aren't already in the correct folder
5. Progress and results are shown inline and in the activity log

## TypeScript vs JavaScript

This project uses TypeScript, which is the standard for SPFx development. Here's what changes between TypeScript (TSX) and JavaScript (JSX):

| TypeScript Feature | Example | In JSX |
|---|---|---|
| Type annotations | `listId: string` | `listId` |
| Interfaces | `IListInfo`, `IFieldInfo` | Gone (or use JSDoc comments) |
| Enums | `GroupingStrategy.Date` | Plain objects: `{ Date: "Date", ... }` |
| Access modifiers | `private _siteUrl` | Just `this._siteUrl` (convention only) |
| Generic types | `Promise<IListInfo[]>` | `Promise` (untyped) |
| `.tsx` files | `ListFolderOrganizer.tsx` | `ListFolderOrganizer.jsx` |
| `tsconfig.json` | Compiler config | Not needed |

### Side-by-side Example

**TypeScript (what we use):**
```typescript
public async getLists(): Promise<IListInfo[]> {
  const response: SPHttpClientResponse = await this._spHttpClient.get(
    apiUrl, SPHttpClient.configurations.v1
  );
  const data = await response.json();
  return (data.value || []).map((list: any): IListInfo => ({
    id: list.Id,
    title: list.Title,
    itemCount: list.ItemCount,
  }));
}
```

**JavaScript equivalent:**
```javascript
async getLists() {
  const response = await this._spHttpClient.get(
    apiUrl, SPHttpClient.configurations.v1
  );
  const data = await response.json();
  return (data.value || []).map((list) => ({
    id: list.Id,
    title: list.Title,
    itemCount: list.ItemCount,
  }));
}
```

The logic is identical — TypeScript just adds type safety. For SPFx specifically, TypeScript is the right choice because the entire SPFx toolchain (`@microsoft/sp-http`, `@microsoft/sp-webpart-base`, Fluent UI) ships type definitions that provide autocomplete and compile-time error checking. Writing SPFx in plain JS means losing all of that while still carrying the same build pipeline.
