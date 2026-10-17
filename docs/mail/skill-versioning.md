# Skill version-bump procedure (`skills/*/SKILL.md`)

Source of truth: `tools/check-skill-version.mjs`. Skill files are copied
into users' agent directories (Claude Code, Codex, …) and only replaced when
`metadata.version` grows — a body change that keeps the version would never
reach anyone.

## Rule

- Any body change to `skills/*/SKILL.md` MUST bump `metadata.version`
  (semver `major.minor.patch`; the check only requires the new version to be
  greater than the base revision's).
- Frontmatter-only version bumps with an unchanged body are pointless but
  harmless; body changes without a bump fail the gate.

## How the check works

- Diff-based like the other checks: it compares each changed `SKILL.md`
  against its base revision (`--base <ref>`, `FORMAT_BASE_REF`, else HEAD)
  and also picks up new untracked skill files.
- Only paths matching `skills/*/SKILL.md` are examined.

## Running it

```bash
node tools/check-skill-version.mjs --base <merge-base-or-parent>
```

## History

- H10 (t_86121039): added the Mail app-only note to
  `skills/genoffice/SKILL.md` and bumped `2.49.0` → `2.50.0`.
