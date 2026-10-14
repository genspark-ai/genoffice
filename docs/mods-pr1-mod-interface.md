# PR 1 — expose the interface to the mod assistant

Status: implementation plan. Supersedes the earlier "script function" framing
(docs/mods-kind-script.md keeps that material for its own later PR).

## 1. What ships

The interface that lets the mod assistant add features — and nothing else. No
feature kind is implemented in this PR: no script runs, no slot draws, no hook
fires. This PR delivers the **container and the contract**, so the kinds have
somewhere to land.

Four pieces:

1. **A feature registry** in the shell main process: features are records on
   disk (one JSON file each), with stable ids, schema-checked writes, and
   list/read/create/update/delete.
2. **The mod assistant** in Home: an AI surface composed from a new AgentSkill,
   so the assistant can actually call the interface.
3. **The interface itself** — the tools the mod assistant is handed:
   - `list_features()` — what exists
   - `get_feature(id)` — one record in full
   - `create_feature({ name, description, kind? })`
   - `update_feature({ id, name?, description?, kind? })`
   - `delete_feature({ id })`
   - `read_document({ path })` — read-only, documents only
4. **A read-only document reader** for the assistant: a document's parsed text.
   There is no writer, and no tool that can change a document.

## 2. Acceptance criteria

1. From Home, asking the mod assistant to add a feature results in a
   `create_feature` call and a record that `list_features` returns.
2. A feature can be listed, read, updated, and deleted; the change survives a
   restart.
3. The assistant cannot change a document: no tool writes one, and the reader
   only accepts document files.
4. The assistant never reports a feature as working: nothing executes a feature
   in this build, and the prompt plus the post-run check both say so.
5. Everything above is testable without a provider: the skill is unit-tested
   against an injected registry, and the registry is tested against a temp dir.

## 3. What a feature is, for now

    {
      "id": "weekly-report-formatting",
      "name": "Weekly report formatting",
      "description": "Apply the house style to a weekly report.",
      "kind": "",              // recorded, but no kind executes yet
      "createdAt": "2026-10-11T…",
      "updatedAt": "2026-10-11T…"
    }

`kind` is an open string on purpose: section 4 of docs/mods-design.md lists the
kinds a mod can add (script, event hook, slots, command, knowledge source), and
each becomes a value here as it is built. This PR executes none of them, so it
does not pick the first one.

## 4. Where things live

| Piece                            | Location                             | Why                                                                                              |
| -------------------------------- | ------------------------------------ | ------------------------------------------------------------------------------------------------ |
| Feature records                  | shell userData, `features/<id>.json` | same place as app-settings.json; no workspace concept to resolve yet                             |
| Registry + IPC                   | shell main                           | one process owns the files and validates every write                                             |
| The skill (tools, prompt, check) | `packages/agent-core`                | storage-agnostic and unit-testable, like the other skills                                        |
| The assistant panel              | Home (shell renderer)                | the shell main already registers the `ai:stream` handlers, so only the preload bridge is missing |

## 5. Deliberately not in this PR

- any feature kind, and therefore any execution of a feature;
- feature `preview` / dry-run: there is nothing to preview until a kind exists;
- per-workspace or per-document features (a later decision, section 10 of the
  design doc);
- installing, sharing or distributing features;
- any change to document rendering, save, export or print.

## 6. How this was verified

| Check                | Result                                                                                                                                         |
| -------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| features skill, unit | 11 tests: tool surface, create/update/delete/get, unknown ids, document reader failures, context flattening, the "does not run yet" correction |
| feature store, unit  | 8 tests: slug ids and uniqueness, partial updates, path traversal refusal, malformed records skipped, the count cap, no temp files             |
| interface end to end | 2 tests: a scripted model calls create_feature and the record is on disk for a fresh read; an unknown id is reported, nothing invented         |
| shell strings        | the 12 new keys in all 20 locales (the dictionary is type-checked per locale)                                                                  |
| gates                | typecheck 0 (agent-core and shell), theme colors, English-only, prettier                                                                       |
| panel as a picture   | e2e/panel-preview renders the real component with the real stylesheet in headless Chromium; see the light and dark captures                    |

The in-app Electron path (e2e/home-mods-assistant.spec.ts) is written but cannot
run on a machine where the shell stalls after boot: on this workstation the app
stops answering page and main-process calls about two seconds after launch,
before the panel is opened, so a click never lands. The spec stays for CI, where
the rest of the suite runs.

The preview harness is how the panel's look is reviewed:

    npx esbuild e2e/panel-preview/entry.tsx --bundle --outfile=e2e/panel-preview/out/app.js \
      --format=iife --platform=browser --jsx=automatic --define:process.env.NODE_ENV='"production"'
    node e2e/panel-preview/shot.mjs      # writes /tmp/mods-preview/*.png

It renders the real component against a small stub bridge, so what is reviewed is
the component and the stylesheet, not a mock-up.
