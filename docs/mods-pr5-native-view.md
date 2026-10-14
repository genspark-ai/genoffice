# PR 5 — the native view slot (K5): a browser a feature declares

Status: **implemented**. This is the second of the three slots; the declarative one
shipped in docs/mods-pr2-declarative-slots.md, the canvas overlay comes next.

## 1. What ships

A slot element that is a real view, not a drawing:

    create_feature({
      name: "Headlines",
      description: "A news page, next to the files.",
      slot: { at: "home.panel", element: { type: "view", view: "browser",
              url: "https://news.ycombinator.com", title: "News", minHeight: 320 } }
    })

The shell mounts a WebContentsView over the rectangle the slot reserves and keeps
its own frame around it: the dot, the title, the address.

## 2. The contract

| Field     | Rule                                                                  |
| --------- | --------------------------------------------------------------------- |
| view      | "browser" — the only native view the host mounts                      |
| url       | https anywhere; http only on the loopback (localhost, 127.0.0.1, ::1) |
| title     | optional, shown in the frame                                          |
| minHeight | optional, 80–1200 px                                                  |

file:, data:, about:, javascript: and custom schemes are refused **before a view
exists**, with the field named. The rule lives in the contract
(packages/agent-core, isAllowedViewUrl) so the tool, the store, the IPC boundary
and the mount path all judge addresses the same way.

## 3. What the host enforces

- the renderer holds **no webview tag and no webContents id**: it reports a
  rectangle and an address; the shell owns the view;
- the view is sandboxed with context isolation, no node integration and no preload;
- navigation to a non-allowed address is cancelled, and a window.open goes to the
  user's own browser instead of opening a second, unreported window;
- the rectangle is clamped into the window's content area; a collapsed panel
  unmounts rather than leaving a stray view;
- one view per slot instance, unmounted when the slot unmounts, and all views are
  torn down on dispose.

## 4. Files

| File                                                      | Change                                                              |
| --------------------------------------------------------- | ------------------------------------------------------------------- |
| packages/agent-core/src/feature-slots.ts                  | the view element, its validation and isAllowedViewUrl               |
| apps/shell/src/shared/feature-view-geometry.ts            | new: clamping a reported rectangle, and the host shown in the frame |
| apps/shell/src/shared/feature-channels.ts                 | new: the mount/unmount channels and their payload types             |
| apps/shell/src/main/features-view.ts                      | new: the view host (create, navigate, clamp, tear down)             |
| apps/shell/src/main/index.ts                              | register it against the shell window                                |
| apps/shell/src/preload/index.ts, shared/home-api.ts       | the two bridge methods                                              |
| apps/shell/src/renderer/src/slots/SlotView.tsx, slots.css | the frame and the rectangle reporting                               |

## 5. Tests

1. the contract accepts a browser view, refuses file:/data:/about:/javascript:,
   refuses another view kind and a silly minHeight;
2. the geometry: rounding and clamping into the window, refusing a rectangle that
   is too small or not a rectangle, and naming the host for the frame;
3. the picture harness renders the real frame and the state before the view
   attaches (there is no main process in the harness, which is exactly that state).

The mount itself needs a running Electron shell: this workstation's shell stops
answering about two seconds after launch (see the note in
docs/mods-pr1-mod-interface.md), so the frame and the rules are verified here, and
the mount is verified by the shape of the IPC payload and the clamping unit tests.

## 6. Out of scope

- other native views (pdf, image, a second editor pane) — the element already has a
  view field so they are additive;
- devtools, a per-view session partition, a per-view preload: a mod gets a page,
  not a runtime;
- making the browser the user's tab: a slot view never takes over the tab strip.
