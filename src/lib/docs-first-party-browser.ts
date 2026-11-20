/**
 * THE PART OF THE DECLARED-FORMATS PACKAGE A BROWSER BUNDLE MAY EMBED, first-party like `./docs-first-party-light.ts`:
 * the clock's shape and its two durations, and nothing else.
 *
 * WHY ITS OWN ENTRY. The app's clock module is embedded by the dashboard through three of the app's modules, and
 * whatever it imports travels to the browser with it. This entry reaches one leaf; the dashboard's guard pins it
 * (`src/core/import-boundaries.test.ts`).
 *
 * Decides: which names of the package a module embedded by the dashboard may take.
 * NOT here: anything else the app takes (`./docs-first-party-light.ts`, `./docs-first-party.ts`), the published doors,
 * nor any logic — every name is re-exported from the module that owns it.
 * Called by: `src/core/clock.ts`.
 */

export { DAY_MS, MINUTE_MS } from "./clock";
export type { Clock } from "./clock";
