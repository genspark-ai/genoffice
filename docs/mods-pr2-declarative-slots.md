# PR 2 — the declarative slot kind (K3)

Status: **implemented** (the home places). Follows docs/mods-pr1-mod-interface.md
(the registry and the assistant) and precedes the other kinds in
docs/mods-design.md section 9.2.

What landed: the slot contract and its validation (packages/agent-core/src/
feature-slots.ts), the payload through the skill, the store and the IPC boundary,
the host renderers for home.sidebar / home.card / home.panel, the action path
(open, ask-assistant, panel, run-feature), and the picture harness at
e2e/panel-preview (entry-slots.tsx renders the gallery). The editor places
(panel.header, ribbon.button, menu.item, status.item, editor.side) are still
closed: they wait for the shared surface contract.

## 1. What ships

A feature of kind **slot** stops being a bare record and starts being drawn: the
mod assistant can register a slot, the host renders it in a named place, and the
user can press what it draws.

    create_feature({
      name: "File switcher",
      description: "Quickly jump to a file",
      kind: "slot",
      slot: { at: "home.sidebar", element: { type: "button", label: "Files" } }
    })

Three pieces:

1. **A slot contract** (section 3): the names a slot may occupy, the element set,
   the props and events each element takes, and the limits (depth, count, text).
2. **Host renderers** for the first names: the home sidebar and home content, the
   AI panel header, a ribbon button and a context-menu item. Each renders the
   declared tree with the app's own components, so theme, dark mode, text scaling
   and i18n stay with the host.
3. **A press path**: an element that can be pressed reports back to the feature
   that declared it, and the host runs the feature's declared action.

## 2. Acceptance criteria

1. A slot registered by the assistant appears, unprompted, in the named place
   after the registry refreshes.
2. Pressing what it draws reports the press to the feature, and the host performs
   the declared action.
3. Deleting the feature removes its slot from the interface.
4. A slot that names an unknown place, an unknown element, too many children or
   too much text is rejected at registration, with a message naming the field.
5. Nothing here can touch the document: the element set has no writer, and the
   acceptance test of docs/mods-design.md section 1.3 still passes.

## 3. The slot contract

### 3.1 Names

| Name          | Where it lands                                            | Rendered on |
| ------------- | --------------------------------------------------------- | ----------- |
| home.sidebar  | a row in the home sidebar, under the built-in entries     | Home        |
| home.card     | a card in the home hero row                               | Home        |
| home.panel    | a full pane in the home content area (like the assistant) | Home        |
| panel.header  | the right side of an AI panel's header                    | six editors |
| ribbon.button | a button in the ribbon, in a named group                  | six editors |
| menu.item     | an item in a context menu, in a named section             | six editors |
| status.item   | an item in the status bar                                 | six editors |
| editor.side   | a collapsible side panel inside an editor                 | six editors |

Only the home names land in this PR; the editor names follow once the app panels
share the surface contract (a later PR).

### 3.2 Elements

A deliberately small set, drawn by host components:

| Element  | Props                                                      | Events |
| -------- | ---------------------------------------------------------- | ------ |
| text     | value, tone (default, muted, danger), size (body, caption) | —      |
| button   | label, tone (default, primary), disabled                   | press  |
| input    | name, placeholder, value                                   | submit |
| select   | name, options, value                                       | change |
| list     | items (each: label, value)                                 | select |
| divider  | —                                                          | —      |
| markdown | value (sanitised by the host)                              | link   |

Limits: at most 4 levels deep, at most 24 elements, at most 200 characters per
text, at most 32 list items. Anything over is rejected at registration rather
than truncated at paint, so the author sees the error.

### 3.3 Actions

What a press does is declared, not coded: the feature names an action the host
already knows how to perform.

    { "kind": "open", "path": "/Users/me/report.docx" }
    { "kind": "run-feature", "id": "weekly-report-formatting" }
    { "kind": "ask-assistant", "prompt": "Add a weekly report formatting feature" }

The set starts with those three; a script action arrives with K1.

## 4. Where the work goes

| Piece                               | File                                        | Change                                                  |
| ----------------------------------- | ------------------------------------------- | ------------------------------------------------------- |
| Slot declaration types + validation | packages/agent-core/src/feature-slots.ts    | new: element/props/limits, validateSlot()               |
| Registration                        | packages/agent-core/src/features-skill.ts   | kind slot carries a slot payload; the tool validates it |
| Storage                             | apps/shell/src/main/feature-store.ts        | records keep the optional slot payload verbatim         |
| IPC                                 | apps/shell/src/main/features-ipc.ts         | validate the slot payload at the boundary               |
| Host rendering                      | apps/shell/src/renderer/src/slots/          | new: element renderers + the name registry              |
| Home names                          | apps/shell/src/renderer/src/Home.tsx        | home.sidebar / home.card / home.panel                   |
| Tests                               | packages/agent-core/tests, apps/shell/tests | see 5                                                   |

## 5. Tests

1. validation: unknown place, unknown element, depth, count, text length, bad
   action — each rejected with the field named;
2. the skill: create_feature carries the payload; update replaces it; an invalid
   payload fails the tool call without writing;
3. the store: a record round-trips the payload;
4. the flow: a scripted model registers a slot and the registry holds it;
5. e2e: the slot appears in the home sidebar and pressing it reports the press.

## 6. Out of scope

- the three editor surface names (they need the shared surface contract);
- canvas overlay and native view (K4, K5 — their own PRs);
- any element that can write a document;
- mod packaging: slots registered through the assistant are per-user records,
  not yet shippable files.
