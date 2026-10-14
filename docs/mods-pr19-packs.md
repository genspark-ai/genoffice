# PR 19 — packs: what leaves this machine, and what arrives

Status: **implemented**. This is row 7 of docs/mods-design.md section 9.2: packaging
and trust.

## 1. What ships

Two buttons in the feature panel:

- **Export all** writes every installed feature to a JSON pack (the OS save dialog
  picks the place);
- **Import…** reads a pack and **shows what it declares before anything is written**:
  each feature's name, kind, and one line per thing it declares — where a slot draws,
  which event a hook observes, which folder a knowledge source reads, that a script
  runs code — followed by the parts worth a second look, and an Install button that
  only then writes.

## 2. Trust is the interesting half

A pack is untrusted input on the way back in, so:

- **every record is re-validated** with the validators the registry already uses
  (slot, appearance, script, hook, knowledge, command). A pack cannot smuggle a
  malformed slot, an invented appearance or a module-shaped script past the boundary;
- a pack from a **newer GenOffice** is refused by version, and one that is not a pack
  at all is refused by kind;
- the review **names the dangerous parts**: a hook on `document.will-save` _can stop
  a save_, a script _runs code_, a knowledge source _reads from a folder_;
- an import **never replaces an existing feature** unless the caller asks it to, so a
  pack cannot quietly overwrite what the user already has;
- the pack size is capped, and the description is what the renderer shows — the
  renderer never decides.

## 3. One bug the tests caught

An import must keep each feature's **own id**: identity has to survive the trip, or a
second import arrives as `tidy-headings-2` and the user ends up with two. The store
gained `putFeature` (write a record under its own id, validated), and the pack tests
assert that a second import _skips_ rather than duplicates, and that
`replace: true` replaces in place.

## 4. Files

| File                                                              | Change                                                                       |
| ----------------------------------------------------------------- | ---------------------------------------------------------------------------- |
| apps/shell/src/main/feature-packs.ts                              | new: the pack, its reader (re-validation), its description and the installer |
| apps/shell/src/main/feature-store.ts                              | `putFeature`: write a record under its own id                                |
| apps/shell/src/main/features-ipc.ts                               | export / preview / install, with the real file dialogs                       |
| apps/shell/src/shared/feature-channels.ts                         | the pack description shape, beside the channels                              |
| apps/shell/src/shared/home-api.ts, preload/index.ts               | the three bridge methods                                                     |
| apps/shell/src/renderer/src/ModsAssistant.tsx, mods-assistant.css | the two buttons and the review                                               |
| apps/shell/src/renderer/src/strings.ts                            | four keys, all 21 locale blocks                                              |
| apps/shell/tests/feature-packs.test.ts                            | seven cases                                                                  |
| e2e/panel-preview                                                 | the harness answers with a pack and drives the review                        |

## 5. Tests

1. a pack round-trips the records it was built from;
2. not-JSON, wrong-kind and newer-version packs are refused, with the reason;
3. a pack with a malformed record **or a malformed slot** is refused, not installed;
4. the size cap holds;
5. the description lists what each feature declares and warns about a save-stopping
   hook, code that runs, and a feature that would be replaced;
6. installing into an empty registry works; a second import skips; `replace: true`
   replaces in place;
7. a feature the user installed themselves is left alone.

## 6. Out of scope

- signatures or provenance: a pack is a file the user chose to open, and this slice
  says what it declares rather than who wrote it;
- uninstalling everything a pack added (features are removed one by one in the panel);
- a marketplace: that is distribution, a different problem from trust.
