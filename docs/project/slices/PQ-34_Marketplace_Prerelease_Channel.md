# PQ-34 — Marketplace pre-release channel, and the semver it requires

**State:** ⬜ planned · **Est:** M · Minted 2026-08-16

## What this buys

The **"Switch to Pre-Release Version"** button in the VS Code Extensions view. A tester opts in once
and then receives pre-release builds automatically, in the editor, without hunting for a `.vsix` on a
releases page.

That is the whole value, and it is real — but it only matters once there is a *stream* of testers.
One person downloading a VSIX from a GitHub prerelease, which is what we do today, does not need it.

## The constraint that makes this a policy change, not a workflow change

**The Marketplace has no concept of a pre-release suffix.** `vsce` rejects `0.6.0-rc.2` outright; an
extension version must be plain `major.minor.patch`. This is why the release pipeline already strips
the suffix and packages a tag like `v0.6.0-rc.2` as version `0.6.0`.

Which leads to the trap:

> **A version number can be published once.** Publish `0.6.0` as a pre-release and `0.6.0` can never
> be published as stable.

So the RC we ship today cannot simply be pushed to the Marketplace with `--pre-release`. Doing that
would consume the version number the stable release needs.

Microsoft's documented workaround is a versioning convention:

| | minor version | example |
| --- | --- | --- |
| **stable** | even | `0.6.0`, `0.8.0`, `0.10.0` |
| **pre-release** | odd | `0.7.0`, `0.9.0`, `0.11.0` |

Stable minor versions then advance **two at a time, permanently**. That is the actual cost of this
feature: not the pipeline work, but a numbering rule that can never be violated afterwards without
burning a version.

## Decisions already made

Settled during design on 2026-08-16, so they do not need relitigating:

**Manual trigger only.** A `workflow_dispatch` input, the same shape as the existing
`publish_marketplace`. Never automatic on an rc tag. A pre-release reaches real users' editors on its
own once they have opted in, and that deserves a deliberate hand on the lever rather than being a
side effect of tagging something.

**Enforce the convention in the pipeline, do not document it.** The job must refuse to publish an
even minor as pre-release, and refuse to publish an odd minor as stable. A rule that lives only in a
document is a rule that rots — this repository spent a day in August 2026 proving that. Make the
violation fail the build instead of silently consuming a version number.

**One job in `release.yml`, not a second workflow.** It shares the built VSIX, the version
derivation, and the `VSCE_PAT` check that already exist. A separate workflow duplicates all three and
drifts from them, which is exactly how the previous `workflow_run` pipeline died.

## Implementation sketch

```yaml
marketplace-prerelease:
  needs: [build, github-release]
  if: inputs.publish_prerelease == true
  steps:
    - # refuse an even minor - that number belongs to a stable release
    - # vsce publish --pre-release --packagePath *.vsix --pat "$VSCE_PAT"
```

Plus:

- `PUBLISHING_GUIDE.md` gains the odd/even table and the reason for it.
- `bump-version` should know the convention, so nobody hand-bumps into the wrong lane.
- `Config_Changes.md` is not involved — this changes no user setting.

## When to do it

**Not yet, deliberately.** Trigger conditions, either one:

1. **More than one person is testing pre-release builds.** At that point manual VSIX handoff becomes
   the bottleneck and the button pays for itself.
2. **A pre-release cycle lasts long enough that testers need updates during it.** A single RC that
   ships in a week does not; a month of iteration does.

Until then GitHub prereleases do the job — `v0.6.0-rc.2` was published this way and the one person
who asked for the feature has it.

## Origin

Raised by Wilson on 2026-08-16 immediately after replying to
[@namgaw in discussion #3](https://github.com/ewc3labs/excel-power-query-editor/discussions/3) with
the 0.6.0 RC: *"shall we add the vs marketplace prerelease pipeline? And how automated do we make
it?"*

Deferred in the same conversation, on the grounds that we had just told a tester the version is
`0.6.0` and should see how the first real feedback lands before spending a version number on
machinery for an audience of one.
