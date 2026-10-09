# Connecting a coding agent

GenOffice speaks the Model Context Protocol, so a coding agent can read, write and render your documents through the same engines the app uses. The agent is not guessing at a file format: it gets the typed op schemas from the same definitions the executor validates against.

## Registering it from the app

**Settings ▸ Integrations** is the place to do this. The pane has two halves, and you can use either or both.

**The skill.** One row per coding agent found on this machine — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf — each with **Install**, **Update** and **Uninstall**, plus **Install into another folder…**, **Download skill (zip)** and **Copy path**. If your assistant is not listed, point GenOffice at a folder it reads `SKILL.md` from, or save the zip and let the assistant install it. The skill and MCP can sit side by side: the assistant picks one, and they do exactly the same things.

**MCP.** Two routes are offered: **Started by the assistant (recommended)**, where you add the shown config to your client and the assistant launches the server itself, and **Local HTTP server**, which the app runs for you. Either way the assistant ends up talking to GenOffice and you never type a command.

## Registering it from the command line

The same thing from a terminal — this is the advanced path, and the one to reach for when the agent lives somewhere the pane cannot find:

```sh
genoffice mcp install all
```

It finds the coding agents on this machine and writes the stdio server entry into each one's own config, leaving the rest of that file as it found it.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

An agent you installed somewhere unusual takes `--dir <path>`; `--force` rewrites an entry that is already there.

Everything the server takes is on one screen — the register, remove and list forms, serving over HTTP, and the two schema flags:

![The real output of genoffice mcp --help: the install, uninstall and list forms, with the --http, --host, --token, --compact-schemas, --dir and --force options](img/mcp.png)

## Running it without an assistant

For a client on another machine, serve it over HTTP instead:

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` changes where it listens. Pass the same token to the client.

Over HTTP files travel too: `PUT /files/<name>` uploads one, every tool takes an `http(s)` URL in place of a path, and outputs come back as download URLs — and, when they are small enough, as embedded resources. Ops, specs and Markdown are passed inline either way.

## What the agent gets

Every command is a tool. The interesting ones:

- **`docs`, `sheet`, `slides`** — read and edit a file through the app's own write path, one **op** at a time. A new deck goes `deck_start`, `deck_page`, `deck_build`.
- **`render`** — one PNG per page, laid out by the app's renderer, so the agent can look at a slide rather than guess at it.
- **`pdf`** — a PDF's text layer, page by page, with no app process.
- **`info`** — metadata and a structure summary, which is usually the cheapest first call on an unfamiliar file.
- **`search`, `image`, `media`** — the providers configured in the app, so the agent does not need its own keys.
- **`merge`** — fill a `{{key}}` template.

## Schemas, and a smaller budget

`apply` and `create` advertise their `ops`, `cells` and `data` parameters with the typed per-op schema, generated from `genoffice guide <domain> --json`. That is precise and not small. A client with a tight context window can ask for plain arrays instead:

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## Why the skill

An agent that does not know the op vocabulary will guess. The skill carries the reference and the design guides — the same material `genoffice guide` prints — so the assistant writes ops whose spec it has actually read. Install it from the pane at the top, or with `genoffice skill` from a terminal.

## What it can reach in the app

The server is not limited to files on disk. While GenOffice is running, the agent can also work through the window:

- **`open_in_genoffice`** opens a file in a tab and focuses it.
- **`open_documents`** lists every document you have open — id, type, path, and whether it has unsaved changes — then reads one's live content or closes it, saving first unless you ask it to discard.
- **The content tools** take that id (or path) as their `document` argument, so an edit lands in the tab you already have open and the window switches to show it.

Two things stay out of reach: there is no AI panel, and the in-app update dialog does not apply.

## The Local HTTP server panel

Under **Local HTTP server** the app can run the server itself instead of leaving it to the assistant: an enable switch, a **Port** field, a **Running / Not running** indicator, **Background generation** (write documents straight to a path without opening the UI), and the **Client configuration example** to copy into your client. Opening **Advanced** adds the two connection URLs — Streamable HTTP and the legacy SSE one — a **Health check** URL, and a **Logging** switch that records server and tool activity to a local file you can **Open**, **Refresh** or **Clear** from there. It listens on localhost only.
