# Changesets

This folder is managed by [`@changesets/cli`](https://github.com/changesets/changesets). Each markdown file represents one user-facing change about to ship in the next release.

## Add a changeset

After making a change worth releasing:

```bash
pnpm changeset
```

The CLI asks which packages changed and at what semver level (major / minor / patch), then writes a `*.md` file here with your description. Commit it alongside your code change.

## How releases happen

1. PR merges to `main` with a changeset file.
2. The Release workflow (`.github/workflows/release.yml`) sees pending changesets and opens (or updates) a "Version Packages" PR that bumps versions and rewrites `CHANGELOG.md` per package.
3. When that PR is merged, the workflow runs `pnpm build && changeset publish` — every package whose version was just bumped is published to GitHub Packages (`@alavida-ai/*` scope).

So a release is just: open changeset PR → merge → merge the auto-opened version PR → published.

## What's "user-facing"?

- Anything that affects the published API or behaviour → changeset required.
- Internal refactors, tests, docs, CI tweaks → no changeset.

If a change isn't worth a changeset, no action needed — the workflow simply won't publish anything new.

## Workspace deps

When you bump `@alavida-ai/outlook-core`, the `cli` and `openclaw-plugin` packages auto-bump too (via the `updateInternalDependencies: patch` setting in `config.json`). You don't need to add separate changesets for each package — describe the change at the level it occurs and changesets handles the dependency graph.
