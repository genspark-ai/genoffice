# GenOffice Enterprise (`ee/`)

This directory is reserved for future enterprise modules (for example
private deployment and offline license verification). It is intentionally
empty today except for this notice and the license.

## License boundary

Everything under `ee/` is covered by the
[GenOffice Enterprise License](LICENSE), not the Apache-2.0 license that
covers the rest of the repository. Keeping all enterprise code behind
this single top-level directory keeps the license boundary auditable and
lets the open-source core stay plain Apache-2.0 permanently.

## Contributions

`ee/` does not accept external contributions. Pull requests from outside
the maintainer team must not modify files in this directory (enforced
via CODEOWNERS). See [CONTRIBUTING.md](../CONTRIBUTING.md).
