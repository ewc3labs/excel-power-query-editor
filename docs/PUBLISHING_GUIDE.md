# Publishing Guide

How a release is cut, and what stops one happening by accident.

## The short version

```bash
git tag v0.6.0
git push origin v0.6.0
```

That builds, packages, and creates a **draft** GitHub release with the `.vsix` attached. Nothing is
published to the Marketplace and nothing is visible to users until you edit the draft and press
publish.

## Marketplace publishing is OFF

**This is deliberate, and it is enforced structurally rather than by remembering.** Publishing to
the Marketplace requires all three of:

1. a **stable** tag — anything with a suffix (`v0.6.0-rc.1`) is a prerelease and never publishes
2. the repository variable **`MARKETPLACE_PUBLISH`** set to `enabled`
3. the repository secret **`VSCE_PAT`**

Neither the variable nor the secret exists right now. So the worst outcome of pushing a tag — any
tag, including a wrong one — is a draft release you delete.

That matters more than it sounds. It is what makes the pipeline safe to test with real tags, and
testing with real tags is the only way to know a tag-triggered pipeline works.

To publish for real: set `MARKETPLACE_PUBLISH` to `enabled` and add `VSCE_PAT`, or run the workflow
manually from the Actions tab with **publish_marketplace** checked.

## Why the pipeline was rebuilt

Worth knowing, because the failure was invisible for a year.

The previous release workflow was triggered by CI finishing, using `workflow_run`. A `workflow_run`
job executes **in the context of the default branch** — `github.ref` is `refs/heads/main` regardless
of what was actually pushed. Every `refs/tags/v*` condition in its gating logic was therefore
unreachable, and the workflow could never classify a run as a release.

It did not fail loudly. It ran, decided there was nothing to release, and went green. Its last
successful release was 2025-07-21, and a finished 0.5.2 sat unshipped behind it for a year.

The current workflow is triggered by the tag directly. That is the whole fix.

## What the workflow does

`.github/workflows/release.yml`, on `push` of a `v*` tag or manual dispatch:

| Job | What it does |
| --- | --- |
| **build** | types, lint, package the `.vsix`, and verify the package contains what it must |
| **github-release** | creates a **draft** release with the `.vsix` attached |
| **marketplace** | publishes — only when all three conditions above are met |
| **summary** | reports what happened, including what was skipped and why |

Two details that are easy to trip over:

- **A prerelease is any tag with a suffix.** `v0.6.0-rc.1` and `v0.6.0-beta.2` are prereleases;
  `v0.6.0` is not. The GitHub release is marked accordingly.
- **`vsce` rejects a prerelease suffix in the manifest.** An extension version must be plain
  `x.y.z`, so the packaged version is the base version even when the tag carries a suffix. The tag
  keeps the full name; the `.vsix` inside cannot.

## Cutting a release

**1. Make sure the tree is releasable.** CI green on the branch, `CHANGELOG.md` updated, and the
version in `package.json` matching the tag you are about to push.

**2. Tag and push.**

```bash
git tag v0.6.0
git push origin v0.6.0
```

**3. Watch it.**

```bash
gh run watch
```

**4. Edit the draft release.** The generated body is a starting point, not the release notes. Say
what changed for a user, and link the relevant documentation.

**5. Publish the draft** when you are happy with it.

For a release candidate, tag `v0.6.0-rc.1` instead. Same pipeline, marked as a prerelease, and
Marketplace publishing is skipped even if it has been enabled.

## Versioning

```bash
npm version patch     # or minor / major - commits and tags
npm run bump-version  # EWC3 script: sets the version in package.json only
```

**Be careful with `npm version`.** It creates a git tag, and pushing that tag now **fires the
release pipeline**. That was harmless when the pipeline was broken; it is not harmless now. If you
only want to change the number, use `bump-version`, which performs no git operations.

## Testing the pipeline

The pipeline can be exercised safely, and should be after any change to it:

- **Push a real prerelease tag.** `v0.6.0-rc.N` produces a draft prerelease and nothing else. Delete
  the tag and the draft afterwards.
- **Run it manually** from the Actions tab, leaving **publish_marketplace** unchecked.

Do not "test" by reading the YAML. That is how the `workflow_run` bug survived a year.

## Manual publishing

If the pipeline is unavailable and something must ship:

```bash
npm run package-vsix
npx vsce publish --packagePath excel-power-query-editor-0.6.0.vsix --pat "$VSCE_PAT"
```

This bypasses every check the pipeline performs. Prefer fixing the pipeline.

## Before you tag

- [ ] CI green on the branch
- [ ] `CHANGELOG.md` has an entry for this version
- [ ] `package.json` version matches the tag
- [ ] `npm run docs:check` passes ([@ewc3labs/docs-tools][docs-tools]; `npm run docs:fix` fixes most
      of it)
- [ ] `npm run package-vsix` succeeds, and the `.vsix` contains what you expect (`npx vsce ls`)
- [ ] Installed the `.vsix` locally and used it — into the VS Code you actually run, Insiders or
      stable, not whichever one `code` happens to point at

---

Contributing is in [CONTRIBUTING](CONTRIBUTING.md); the documentation index is in
[Overview](Overview.md).

[docs-tools]: https://github.com/ewc3labs/ewc3-docs-tools
