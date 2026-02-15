# YouTube Demo Script: Excel Add-in Development with Codex

## Video Goal
Show how we designed and implemented an Excel add-in using a requirements-first workflow, Codex-assisted scaffolding, modular prompt design, and performance-conscious engineering—plus how ribbon integration works today and how the architecture can be ported to Google Sheets next.

---

## 1) Opening (0:00–0:30)
**On screen:** Final add-in ribbon tab in Excel, quick preview of key features.

**Narration:**
"Hi everyone—today I’ll walk you through how we built this Excel add-in with a requirements-first process and Codex as a development accelerator. We’ll cover how we scoped the project, scaffolded modules, enforced modularity through prompts, baked in performance thinking early, integrated everything into the Excel ribbon, and set up a migration path to Google Sheets. If you’re an Excel power user or a developer building Office automation, this is for you."

---

## 2) Requirements Defined First (0:30–1:45)
**On screen:** Requirements doc with sections: user personas, workflows, functional/non-functional requirements.

**Narration:**
"Before writing any code, we defined requirements in plain language and made them testable. We started with two personas: power users who need one-click actions, and developers who need maintainable extension points.

From there, we wrote functional requirements such as:
- Commands must be discoverable from a dedicated ribbon tab
- Actions should operate on the active workbook context
- Error messages must be actionable for non-developers

Then non-functional requirements:
- Fast startup and responsive command execution
- Minimal workbook recalculation impact
- Clean module boundaries for future portability

This step reduced ambiguity, prevented overengineering, and gave Codex a precise target to scaffold against."

---

## 3) Using Codex to Scaffold Modules (1:45–3:00)
**On screen:** Folder structure and generated module stubs.

**Narration:**
"Once requirements were stable, we used Codex to scaffold the initial project modules instead of hand-writing boilerplate. We prompted for explicit boundaries, resulting in modules like:
- `core/` for domain logic
- `excel/` for Office/Excel-specific adapters
- `ui/` for ribbon command wiring
- `infra/` for logging, config, and telemetry

This gave us a working skeleton quickly, but more importantly it gave us consistent patterns—interfaces, naming conventions, and predictable wiring—that made review and iteration faster."

---

## 4) Prompt Design to Enforce Modularity (3:00–4:30)
**On screen:** Example prompt snippets and resulting code structure.

**Narration:**
"The key was not just using Codex—it was how we structured prompts. We used constraint-driven prompts to enforce modular architecture.

A typical prompt included:
1. **Module responsibility**: what this module owns and what it must not own
2. **Input/output contracts**: typed interfaces and error strategy
3. **Dependency rules**: allowed imports and forbidden cross-layer calls
4. **Test hooks**: where unit tests should attach

For example, we’d specify that ribbon handlers can call application services, but never directly perform workbook mutation logic—that belongs in core services. This kept business logic reusable and significantly reduced coupling."

---

## 5) Performance Considerations Included Early (4:30–5:40)
**On screen:** Performance checklist and before/after timing examples.

**Narration:**
"Performance was treated as a requirement, not a post-launch fix. We included guardrails during scaffolding:
- Batch workbook read/write operations instead of chatty cell-by-cell calls
- Cache static metadata where safe
- Avoid unnecessary recalculation triggers
- Add lightweight timing around high-traffic commands

For Excel add-ins, perceived performance matters as much as raw speed. So we designed for fast feedback: commands acknowledge immediately, then complete deterministic work in bounded steps. That keeps power-user workflows fluid even on large workbooks."

---

## 6) How Ribbon Integration Works (5:40–6:50)
**On screen:** Ribbon XML or config + handler mapping + live demo click.

**Narration:**
"Ribbon integration is the user-facing entry point. At a high level:
1. A custom ribbon tab defines groups and buttons
2. Each button maps to a command handler
3. The handler resolves dependencies and invokes application services
4. Services run domain logic and return user-facing results

We also keep command metadata centralized so labels, tooltips, and callbacks stay consistent. That gives us a clean UX for power users and a clean extension story for developers adding new commands later."

---

## 7) Future Plan: Port to Google Sheets (6:50–8:00)
**On screen:** Architecture diagram showing reusable core + platform adapters.

**Narration:**
"Because the architecture is modular, porting to Google Sheets becomes a platform-adapter project rather than a full rewrite.

The roadmap looks like this:
- Reuse `core/` business logic and validation rules
- Implement a `sheets/` adapter layer for Apps Script or Workspace APIs
- Recreate UI entry points in Sheets menus/sidebar
- Swap telemetry and auth integrations per platform constraints

In short, the requirements-first and modular approach pays off twice: faster delivery now in Excel, and lower migration cost later in Google Sheets."

---

## 8) Closing + CTA (8:00–8:30)
**On screen:** Final recap slide with bullets.

**Narration:**
"To recap: we started with requirements, used Codex to scaffold reliably, enforced modular design through structured prompts, included performance from day one, and built ribbon integration on clean service boundaries. That sets us up not just for maintainability, but for a practical Google Sheets port.

If you’d like, I can share a follow-up deep dive on prompt templates, module contracts, and a migration checklist for multi-platform spreadsheet add-ins. Thanks for watching."

---

## Optional B-Roll / Overlay Suggestions
- Requirements matrix screenshot
- Module dependency diagram
- Prompt template callouts
- Command latency chart
- Ribbon click-through demo
- Excel-to-Google-Sheets portability map
