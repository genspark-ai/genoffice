# Connecting a coding agent

GenOffice speaks the Model Context Protocol, so a coding agent can read, write and render your documents through the same engines the app uses. The agent is not guessing at a file format: it gets the typed op schemas from the same definitions the executor validates against.

## Registering it

The usual case is one command:

```sh
genoffice mcp install all
```

It finds the coding agents on this machine — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf — and writes the stdio server entry into each one's own config, leaving the rest of that file as it found it.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

An agent you installed somewhere unusual takes `--dir <path>`; `--force` rewrites an entry that is already there.

## Running it yourself

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

## The skill

An agent that does not know the op vocabulary will guess. `genoffice skill` installs a GenOffice skill into the agents it finds, carrying the reference and the design guides — the same material `genoffice guide` prints.

## What it is not

The MCP server reads and writes files. It is not the window: there is no AI panel, and the in-app update dialog does not apply. If a step needs the window, open the file.
