# Publishing Guide

How a release is cut, and what stops one happening by accident.

## The short version

```bash
git tag v0.6.0
git push origin v0.6.0
```

That builds, tests, packages, and creates a **draft** GitHub release with the `.vsix` attached.

**Publishing the draft and publishing to the Marketplace are two independent acts.** Editing the
draft and pressing publish makes the GitHub release visible; it does nothing to the Marketplace. The
Marketplace gate is the `marketplace` environment approval described below.

## Marketplace publishing is OFF

**This is deliberate, and it is enforced structurally rather than by remembering.** Publishing
requires all three of:

1. a **plain `vX.Y.Z` tag** — anything with a suffix (`v0.6.0-rc.1`) never publishes to either
   marketplace
2. the repository variable **`MARKETPLACE_PUBLISH`** set to `enabled`, or a manual run with
   **publish_marketplace** checked
3. a **trusted publishing policy** on the Marketplace naming this repo and workflow — see [Setting
   up publishing](#setting-up-publishing)

None of the three is configured right now. So the worst outcome of pushing a tag — any tag,
including a wrong one — is a draft release you delete.

That matters more than it sounds. It is what makes the pipeline safe to test with real tags, and
testing with real tags is the only way to know a tag-triggered pipeline works.

## Two channels, decided by the minor version

The Marketplace has no concept of a semver prerelease suffix. A version is **either** the stable one
**or** the pre-release one, and the same version cannot be both. VS Code's convention is therefore:

| Tag | Channel | Who receives it |
| --- | --- | --- |
| `v0.6.0` | **stable** — even minor | everyone, including auto-update |
| `v0.7.0` | **pre-release** — odd minor | only users who opted in via the extension pane |
| `v0.6.0-rc.2` | **neither** | nobody; a VSIX on a draft release, for handing to someone |

**The channel is derived from the version, so the wrong one is not possible.** It is, however,
*silent* — believing you are shipping `0.7.0` to stable gets you a pre-release, having consumed the
version number permanently. So a manual run can **assert** the channel, and a disagreement fails
before anything is published:

```text
You asked for the stable channel, but v0.7.0 is prerelease.
```

**Once a version is published it is consumed forever.** There is no republishing `0.7.0` as
something else, which is the whole reason the assertion exists.

## Setting up publishing

**There is no PAT in the implemented publishing path, and no Azure at all.** Azure DevOps personal
access tokens are retired on **2026-12-01**, and this publishes with **trusted publishing** instead:
`vsce` asks GitHub for an OIDC token scoped to the `marketplace.visualstudio.com` audience,
exchanges it at `POST /_apis/gallery/token` for a short-lived Marketplace credential, and publishes
with that.

**Nothing long-lived is stored in this repository**, and there is no Entra tenant, app registration,
federated credential, or `azure/login` step anywhere in the path.

> **Correcting an earlier version of this guide.** It said `vsce publish --oidc` does not exist.
> That was true of stable 3.9.2 — which answers `unknown option '--oidc'` — and wrong as a general
> claim. Trusted publishing landed upstream on 2026-07-23 and shipped in **3.9.3-5** on the `next`
> tag. The workflow pins that version exactly.

### Marketplace side (manual, one time)

Configure a **trusted publishing policy** on the `ewc3labs` publisher at
<https://marketplace.visualstudio.com/manage/publishers/ewc3labs>, naming:

| | |
| --- | --- |
| Repository | `ewc3labs/excel-power-query-editor` |
| Workflow | `release.yml` |
| Environment | `marketplace` — if the policy form offers it |

That is the entire setup. The policy is what makes the Marketplace trust a token minted by this
repository's workflow, and it replaces every step an Entra route would have needed.

### GitHub side

| Kind | Name | Value |
| --- | --- | --- |
| Environment | `marketplace` | must exist, **with a required reviewer** — see below |
| Variable | `MARKETPLACE_PUBLISH` | `enabled`, when you want tags to publish |
| Secret | `OVSX_PAT` | Open VSX token — optional, see below |

### The human gate is the environment, not the draft

**Make `marketplace` a protected environment requiring a reviewer.** Settings → Environments →
`marketplace` → Required reviewers → add yourself.

This is stronger than it looks. A protected environment does not pause a running job — **the job
never starts**, so the OIDC token is never minted. The *capability* to publish does not exist until
a person approves it, rather than existing continuously and being politely unused.

```text
push v0.6.0
  ├─ build, test, package
  ├─ draft GitHub release          (visible to nobody until you publish it)
  ▼
marketplace environment
  ├─ REQUIRED HUMAN APPROVAL  ◄── the gate
  ▼
OIDC token minted → Marketplace publish → Open VSX
```

**Do not enable "prevent self-review."** With a single maintainer it deadlocks the pipeline
permanently: the only person who can approve is the person who triggered it.

**Two gates, doing different jobs.** `MARKETPLACE_PUBLISH` decides whether the job is *reachable* —
unset, it is skipped entirely, so testing the pipeline with real tags produces no approval prompts
to dismiss. The environment approval decides whether a reachable job *runs*. Keep both.

An approval request expires after 30 days and the run fails, which is the correct outcome for a
release nobody remembered to approve.

**No `AZURE_CLIENT_ID`, no `AZURE_TENANT_ID`, no secrets for the Marketplace at all.** The
`id-token: write` permission on the job is what lets `vsce` request the token; without it the error
is explicit about the missing permission.

The `marketplace` environment is kept as the **release gate**, and as something the trust policy can
name. It is no longer carrying an OIDC subject, because trusted publishing does not use one.

### The pinned version, and when to unpin

```yaml
env:
  VSCE_VERSION: 3.9.3-5
```

Pinned **exactly**, not to `@next`, which moves. `--oidc` is hidden from `--help` in that build
because it is still preview. **Drop the pin once `--oidc` reaches `latest`**, which is checkable in
one command:

```bash
npx --yes @vscode/vsce@latest publish --oidc --pat dummy
# "cannot be used with option '-p, --pat'"  -> it has landed; unpin
# "unknown option '--oidc'"                 -> not yet; leave the pin
```

That test publishes nothing — the conflict is rejected during argument parsing.

### If trusted publishing is not available

The fallback is Entra ID workload identity federation with `vsce publish --azure-credential`, which
works on stable 3.9.2. It needs an app registration, a federated credential on subject
`repo:ewc3labs/excel-power-query-editor:environment:marketplace`, that identity added to the
publisher, and an `azure/login@v3` step with `allow-no-subscriptions: true`.

**It is documented here rather than implemented** because it is materially more setup for the same
result, and a fallback in the workflow is the thing that quietly keeps the worse path alive.

## Open VSX

Cursor, Windsurf, and VSCodium install from [Open VSX][open-vsx], not the Microsoft Marketplace — so
those users currently cannot install this extension at all. That is a strange gap for a tool whose
pitch is editing Power Query M with an AI coding agent.

Publishing there needs a free account at <https://open-vsx.org>, an `ewc3labs` namespace, and an
access token stored as the **`OVSX_PAT`** secret.

**The job is deliberately non-blocking.** A second registry being down, rate-limiting, or rejecting
a token must not turn a successful Marketplace publish into a red release. Without `OVSX_PAT` it
warns and skips.

**Open VSX is not passed `--pre-release`, on purpose.** Given an already-packaged VSIX it warns
*"Ignoring option '--pre-release' for prepackaged extension"* and carries on — the channel is baked
into the manifest at package time. Passing the flag would suggest the upload decides something it
does not. `ovsx` is pinned to `1.1.1` for the same reason `vsce` is pinned.

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
| **build** | types, lint, **the full test suite**, documentation checks, package the `.vsix`, and verify the package contains what it must |
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

The pipeline now runs the tests itself, so a tag cannot outrun them - but finding out from a failed
release is a worse way to learn it than finding out from CI.

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
**neither marketplace is touched even if publishing has been enabled** — an rc is for handing
somebody a VSIX, and 5,450 installs with auto-update are not the audience for a release candidate.

## Versioning

```bash
npm version patch     # or minor / major - commits and tags
npm run bump-version  # EWC3 script: sets the version in package.json only
```

**The minor version now carries meaning.** Even is the stable channel, odd is pre-release, so
`npm version minor` off `0.6.0` gives you `0.7.0` — a **pre-release**, not the next stable. The next
stable after `0.6.x` is `0.8.0`.

**Be careful with `npm version`.** It creates a git tag, and pushing that tag now **fires the
release pipeline**. That was harmless when the pipeline was broken; it is not harmless now. If you
only want to change the number, use `bump-version`, which performs no git operations.

## Testing the pipeline

The pipeline can be exercised safely, and should be after any change to it:

- **Push a real prerelease tag.** `v0.6.0-rc.N` produces a draft prerelease and nothing else. Delete
  the tag and the draft afterwards.
- **Run it manually** from the Actions tab. It asks for a **tag** - a manual run has no tag of its
  own, and without one the version derivation produces nonsense. Leave **publish_marketplace**
  unchecked.

Do not "test" by reading the YAML. That is how the `workflow_run` bug survived a year.

## Manual publishing

If the pipeline is unavailable and something must ship:

```bash
npm run package-vsix
npx vsce publish --packagePath excel-power-query-editor-0.6.0.vsix --pat "$PAT"
```

Trusted publishing only works **from GitHub Actions** — it needs the runner's OIDC token — so a
local publish still needs a credential of its own until PATs are retired.

Add `--pre-release` for an odd minor, and package it that way too — `vsce publish --pre-release`
refuses a VSIX that was not built with the flag.

This bypasses every check the pipeline performs: the tests, the version/tag agreement, the package
contents check, and the channel assertion. Prefer fixing the pipeline.

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
[open-vsx]: https://open-vsx.org
[pq-34]: project/slices/PQ-34_Marketplace_Prerelease_Channel.md
