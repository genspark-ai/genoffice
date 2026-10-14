# PR 4 — the script kind (K1), first slice: read-only

Status: **implemented**. Follows the slot and appearance kinds; this is row 2 of
the order in docs/mods-design.md section 9.2 (a read-only, dry-run host).

## 1. What ships

A feature of kind **script** carries code:

    create_feature({
      name: "Tidy headings",
      description: "unfold every heading that was wrapped",
      script: { source: "var b = $.docs.blocks()\n$.docs.propose([...])" }
    })

The mod assistant writes it; the **native** assistant (or the user, through it)
runs it read-only with the new tool:

    dry_run_feature({ id: "tidy-headings", path: "/Users/me/report.docx" })

and reads back the plan: which operations would apply, which the engine refused,
and anything the script logged. **Nothing is written** — not the document, not a
temporary file.

## 2. The contract

| Field         | Rule                                                                          |
| ------------- | ----------------------------------------------------------------------------- |
| script.source | the body of a function receiving ($, input); at most 16 KiB; no import/export |

A script may reach: `$.docs.blocks()`, `$.docs.text()`, `$.docs.propose(ops)`,
`$.log(...)`, `$.input`, `$.budget`. It may not reach: `process`, `require`,
timers, `Promise`, the filesystem, the network, or any host object — the API is
built _inside_ the sandbox from plain JSON, code generation is off, and the run is
synchronous under the vm timeout. This is a capability boundary, not a hardened
security boundary: a script is still local code the user or their assistant wrote.

## 3. Where it runs

The host lives in the shell main process (apps/shell/src/main/script-host.ts)
because that is where a document can be opened headlessly and where a registry
entry can be run on demand. It is the same sandbox the command line used, revived
from docs/mods-kind-script.md, and it opens documents through the command line's
own headless path (packages/cli/src/headless-docs.ts), so a plan is judged by the
engine that will one day apply it.

The renderer reaches it through one channel, features:run, and the registry is the
only source of the source: a caller cannot run code that was never recorded.

## 4. What the tests prove

1. the sandbox reads a snapshot, proposes a plan and writes a log;
2. nothing but the API is reachable: process, require and timers are undefined,
   and the constructor escape is blocked;
3. a runaway script stops at its budget;
4. a real run against fixtures/generated/simple.docx produces a plan and leaves
   the file **byte-identical** — the read-only claim, checked on the bytes;
5. an operation the engine does not know comes back as a reported failure, not a
   thrown error, and the file is still unchanged;
6. the skill dry-runs a feature, names a non-script, and surfaces a refused run.

## 5. Out of scope

- applying a script (writing the document) — the next slice, and it goes through
  the native assistant's own checked apply path;
- per-script inputs beyond a plain object, and any scheduling or background run;
- letting a slot run a script directly: a slot's run-feature action still asks the
  assistant, which is what keeps a press from silently editing a file.
