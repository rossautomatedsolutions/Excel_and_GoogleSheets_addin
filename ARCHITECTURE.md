# Ross Spreadsheet Utilities – v1.0

## 1) Modular VBA Architecture Design

Use a **layered, feature-oriented VBA architecture** with a strict one-way dependency flow:

1. **Ribbon Entry Layer**
   - Receives Ribbon callbacks (`onAction`, `getLabel`, etc.).
   - Performs only lightweight validation/routing.
   - Delegates to application-facing procedures.

2. **Feature Modules (Business Logic)**
   - Grouped by capability:
     - Sheet Management
     - Formatting Utilities
     - Navigation Utilities
     - Data Utilities
   - Each public procedure represents a user-visible action.
   - Keep feature modules independent from each other whenever possible.

3. **Shared Utility Layer**
   - Cross-cutting helpers used by all feature modules:
     - Error handling/reporting
     - Application state toggling (performance guard)
     - Validation, workbook/sheet lookups, logging, constants

4. **Configuration/Contracts Layer**
   - Centralized constants, enum-like values, ribbon control IDs, and message templates.
   - Single source of truth for names used across modules.

### Recommended dependency rule
- `Ribbon Orchestrator` ➜ `Feature Modules` ➜ `Shared Utility Layer` ➜ `Config/Contracts`
- Feature modules should **not** call Ribbon callbacks directly.
- Avoid circular dependencies between feature modules.

---

## 2) Suggested Module Names

Use consistent prefixes for readability and maintenance.

### Ribbon / orchestration
- `modRibbonOrchestrator` (single dispatch point for callbacks)
- `modRibbonMap` (optional: mapping control IDs to handler names)

### Sheet Management
- `modSheetMgmt_Basic` (insert/delete/duplicate/rename/hide/unhide)
- `modSheetMgmt_Order` (move/reorder/grouping workflows)
- `modSheetMgmt_Protection` (protect/unprotect workflows)

### Formatting Utilities
- `modFormat_Core` (common formatting presets)
- `modFormat_Layout` (column widths, row heights, wrap/alignment)
- `modFormat_Visual` (borders, fills, theme-like quick styles)

### Navigation Utilities
- `modNav_Selection` (jump to regions, last used cell, named ranges)
- `modNav_Sheet` (next/previous sheet, index-based navigation)
- `modNav_Bookmarks` (optional lightweight location bookmarks)

### Data Utilities
- `modData_Cleanup` (trim, normalize, blanks handling, dedupe orchestration)
- `modData_Transform` (split/merge/convert data operations)
- `modData_AnalysisHelpers` (summary helpers, quick profiling)

### Utilities / common helpers
- `modAppState` (screen updating/calculation/events toggles)
- `modError` (standardized error handling entry points)
- `modValidate` (guard clauses, workbook/sheet/range checks)
- `modWorkbookContext` (active workbook/worksheet resolution rules)
- `modConstants` (all constants/control IDs/message keys)
- `modLogging` (debug trace + optional hidden-sheet/file logging)

> Optional for larger codebases: convert complex stateful helpers into class modules (e.g., `clsAppStateGuard`, `clsLogger`).

---

## 3) Separation of Responsibilities

### A. Sheet Management
Owns sheet lifecycle + organization only:
- Create/delete/duplicate sheets
- Rename and enforce naming rules
- Hide/unhide very hidden workflows
- Reorder sheets and protection-related sheet operations

**Must not include:** cell formatting presets, data cleanup, or navigation command logic.

### B. Formatting Utilities
Owns visual/layout changes only:
- Number/date/text style presets
- Border/fill/font/alignment application
- Row/column sizing, freeze panes helpers (if treated as layout)

**Must not include:** worksheet creation/deletion or data transformation semantics.

### C. Navigation Utilities
Owns movement and user context transitions:
- Jump commands (used range edges, named ranges, anchors)
- Sheet traversal and selection movement
- Optional bookmark capture/restore

**Must not include:** modifying workbook structure or rewriting dataset content.

### D. Data Utilities
Owns content-level operations:
- Cleaning and normalization
- Transforming values/shape
- Lightweight analysis helpers

**Must not include:** purely cosmetic formatting concerns except where explicitly part of data output contracts.

### E. Utilities/Common Helpers
Owns cross-cutting concerns:
- Error handling standard
- Performance toggles and restore logic
- Validation and context resolution
- Constants and logging

**No user-facing ribbon business action should live here.**

---

## 4) Orchestrator Module for Ribbon Entry Points

Use **`modRibbonOrchestrator`** as the only public Ribbon-facing module.

### Responsibilities
- Expose all callback signatures required by Ribbon XML.
- Translate `control.Id` into high-level action routing.
- Run shared preflight checks (e.g., workbook availability).
- Call exactly one feature-layer public procedure per action.
- Handle top-level fallback error display for unexpected exceptions.

### Suggested internal flow per callback
1. Receive callback (`control`, optional `pressed`/`selectedItem`)
2. Resolve action key from `control.Id`
3. Enter performance-safe execution wrapper (where appropriate)
4. Dispatch to feature module public procedure
5. Ensure cleanup/restore and standardized error response

### Naming convention for handlers
- Public callback wrappers: `Ribbon_OnAction_*`
- Private dispatch helpers: `Dispatch_*`
- Feature entry procedures: `Run*` or `Execute*` (e.g., `ExecuteSheetDuplicate`)

---

## 5) Error Handling Pattern (Standard Across Modules)

Adopt a **single structured pattern** in every public routine:

1. `On Error GoTo ErrHandler`
2. Guard clauses + input/context validation early
3. Main logic block
4. `CleanExit` label for deterministic cleanup
5. `ErrHandler` label that:
   - Captures `Err.Number`, `Err.Description`, source procedure name
   - Sends details to `modError`/`modLogging`
   - Shows user-friendly message when appropriate
   - Resumes `CleanExit` to guarantee state restoration

### Standard error metadata to capture
- Procedure name (constant in each procedure)
- Module name
- Workbook/worksheet context (if available)
- Optional action/control ID

### Error policy guidance
- **Feature modules:** raise contextualized errors upward after local enrichment.
- **Orchestrator:** final user-facing message boundary.
- **Helpers:** avoid message boxes unless explicitly designated; return/raise instead.

---

## 6) Performance Best Practices

Centralize performance toggles in `modAppState` and apply via a guarded pattern.

### Toggle set to manage
- `Application.ScreenUpdating`
- `Application.EnableEvents`
- `Application.Calculation`
- `Application.DisplayStatusBar` (optional)
- `Application.DisplayAlerts` (only when necessary)

### Recommended execution pattern
- Save current application state before heavy operation.
- Set fast-mode toggles.
- Execute operation.
- Restore original state in `CleanExit` regardless of success/failure.

### Additional best practices
- Minimize `.Select`/`.Activate`; work with object references.
- Batch read/write ranges using arrays rather than cell-by-cell loops.
- Use `With` blocks for repeated object access.
- Limit volatile worksheet function usage inside loops.
- Prefer explicit workbook/worksheet qualification (`wb.Worksheets("...")`).
- Avoid recalculation thrashing; set manual calc during large transformations and restore afterward.
- Keep status updates lightweight; avoid excessive `Debug.Print` in large loops unless logging level demands it.

---

## Suggested Next Step (when implementation begins)
Create a thin vertical slice first:
- One Ribbon button
- One feature action in each domain
- Shared `modAppState` + `modError`
- Validate architecture and naming consistency before scaling out.
