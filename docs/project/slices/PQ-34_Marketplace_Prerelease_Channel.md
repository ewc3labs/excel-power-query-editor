# PQ-34 — Marketplace pre-release channel, and the semver it requires

**State:** ⬜ planned · **Est:** M · Minted 2026-08-16 · Scope changed 2026-08-17

## What this buys

The **"Switch to Pre-Release Version"** button in the VS Code Extensions view. A tester opts in once
and then receives pre-release builds automatically, in the editor, without hunting for a `.vsix` on
a releases page.

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

## 2026-08-17: PAT authentication is being retired, so this is no longer optional

**Microsoft retires global Azure DevOps Personal Access Tokens on 2026-12-01.** Verbatim from the
[publishing docs][publishing]:

> On December 1, 2026, global Personal Access Tokens (PATs) in Azure DevOps are retired.

That is roughly three and a half months away, and it changes this slice from "a nice channel to
have" into "the authentication this project publishes with has an expiry date". Any pipeline built
around `VSCE_PAT` is being built on something already scheduled for removal.

It also answers a question from the same day - *what happened to my VSCE_PAT?* Nothing: there never
was one in GitHub. The old workflow's publish step was a placeholder that printed setup
instructions, and 0.4.3 and 0.5.1 were published by hand. Whatever token was used locally has since
expired on its own; Azure DevOps PATs live at most a year.

### What the replacement actually is

Microsoft's recommendation, verbatim:

> We recommend that extension publishing use Microsoft Entra ID–based authentication with workload
> identity federation and managed identities.

The flag is **`--azure-credential`**, and it is already present in the version we have (3.6.0) as
well as latest (3.9.2).

**Verified by reading `node_modules/@vscode/vsce/out/auth.js`** rather than assuming: the flag
builds a `ChainedTokenCredential` and asks it for a token scoped to
`499b84ac-1321-427f-aa17-267ca6975798`, the Azure DevOps resource:

```text
EnvironmentCredential
  -> AzureCliCredential          <- the link that makes GitHub Actions work
  -> ManagedIdentityCredential
  -> AzurePowerShellCredential
  -> AzureDeveloperCliCredential
```

`azure/login@v2` with OIDC authenticates the Azure CLI, and `AzureCliCredential` then satisfies the
chain. No secret is stored anywhere.

### Corrections to the plan that prompted this

The suggestion to build on `vsce publish --oidc` and "Marketplace Trusted Publishing" was checked
against the tool and the docs:

| Claim | Verified |
| --- | --- |
| PATs retired 2026-12-01 | **True** |
| Do not build on `VSCE_PAT` | **True** |
| `vsce publish --oidc` exists | **False** - absent from 3.6.0 and from 3.9.2, the latest published |
| Marketplace "Trusted Publishing" | **False** - that is npm's OIDC feature, not a Marketplace concept |
| Upgrade vsce to get OIDC | **Moot** - `--azure-credential` is already in the version we have |
| `permissions: id-token: write` | **True in substance**, for `azure/login`, not for a vsce flag |
| Even minor stable / odd minor prerelease | **True** |

The strategy was right and two implementation details were wrong. Recorded here so nobody builds the
version that does not exist.

### 2026-08-17: superseded - trusted publishing removes the Azure route entirely

**`vsce publish --oidc` exists after all, and it is a better path than everything below.**

The earlier claim here - that `--oidc` does not exist, verified against 3.6.0 and 3.9.2 - was true
of those *published* versions and wrong as a general statement. Trusted publishing was merged
upstream on 2026-07-23 and shipped in **3.9.3-5** on the `next` dist-tag on 2026-08-11. The flag is
`hideHelp(true)`, so it does not appear in `--help` and a version check alone will not find it:

```bash
npx @vscode/vsce@3.9.2 publish --oidc --pat x    # unknown option '--oidc'
npx @vscode/vsce@next  publish --oidc --pat x    # cannot be used with option '-p, --pat'
```

What it does: requests a GitHub Actions OIDC token for the `marketplace.visualstudio.com` audience,
exchanges it at `POST /_apis/gallery/token` for a short-lived Marketplace credential, and never
falls back to a PAT. **No Entra tenant, no app registration, no federated credential, no
`azure/login`.** The trust is a policy on the Marketplace publisher naming the repository and
workflow file.

Everything below about the Entra route remains accurate and is kept as the documented fallback. It
is simply no longer the plan, and the tenant question it raised does not need answering.

### One caveat, and it is the risky part

**Microsoft documents `--azure-credential` for Azure DevOps Pipelines, not GitHub Actions.** Their
example uses an `AzureCLI@2` task and a service connection. The GitHub Actions route should work -
`AzureCliCredential` is second in the chain and `azure/login@v2` satisfies it - but it is inference
from the source, not a documented path, and it has not been run.

Treat the first successful publish as the proof. Until then this is a well-founded expectation.

### Azure-side setup, which is manual and Wilson's to do

1. A **user-assigned managed identity** or Entra ID app registration.
2. A **federated credential** on it, bound to this repository and ideally to the tag ref or a
   protected environment - not to the whole repo.
3. That identity added as a **member of the Marketplace publisher** with a role that can publish.
4. `AZURE_CLIENT_ID` and `AZURE_TENANT_ID` as repository **variables** - they are identifiers, not
   secrets, and there is no token to leak.

## Decisions already made

Settled during design on 2026-08-16, so they do not need relitigating:

**Manual trigger only.** A `workflow_dispatch` input, the same shape as the existing
`publish_marketplace`. Never automatic on an rc tag. A pre-release reaches real users' editors on
its own once they have opted in, and that deserves a deliberate hand on the lever rather than being
a side effect of tagging something.

**Enforce the convention in the pipeline, do not document it.** The job must refuse to publish an
even minor as pre-release, and refuse to publish an odd minor as stable. A rule that lives only in a
document is a rule that rots — this repository spent a day in August 2026 proving that. Make the
violation fail the build instead of silently consuming a version number.

**One job in `release.yml`, not a second workflow.** It shares the built VSIX, the version
derivation, and the credential setup. A separate workflow duplicates all three and drifts from them,
which is exactly how the previous `workflow_run` pipeline died.

**The EXISTING stable `marketplace` job needs the same treatment.** It currently reads `VSCE_PAT`,
which stops working on 2026-12-01 like everything else. Both publish paths move to
`--azure-credential` together, or we will fix this twice.

## Implementation sketch

```yaml
marketplace-prerelease:
  needs: [build, github-release]
  if: inputs.publish_prerelease == true
  permissions:
    contents: read
    id-token: write          # for azure/login OIDC, NOT for a vsce flag
  steps:
    - uses: azure/login@v2
      with:
        client-id: ${{ vars.AZURE_CLIENT_ID }}
        tenant-id: ${{ vars.AZURE_TENANT_ID }}
        allow-no-subscriptions: true
    - # refuse an even minor - that number belongs to a stable release
    - # vsce publish --pre-release --packagePath *.vsix --azure-credential
```

**No `VSCE_PAT` anywhere, and no fallback to it.** A fallback would be the thing that quietly keeps
working until 2026-12-01 and then does not.

Plus:

- `PUBLISHING_GUIDE.md` gains the odd/even table and the reason for it.
- `bump-version` should know the convention, so nobody hand-bumps into the wrong lane.
- `Config_Changes.md` is not involved — this changes no user setting.

## When to do it

**Two clocks now, and they point in opposite directions.**

The pre-release *channel* is still discretionary - it earns its keep once more than one person is
testing, and today that is one person with a VSIX. That has not changed.

**The authentication is not discretionary and has a date on it.** Publishing of any kind - stable or
pre-release - stops working when PATs are retired on **2026-12-01**. The Marketplace currently
serves 0.5.1 from 2025-07-21, so nothing is being published right now anyway, which means there is
no emergency. There is a deadline.

Sensible order:

1. **Move authentication to `--azure-credential` first**, for the stable path that already exists.
   That is the part with a deadline, and it is worth proving on a real publish well before December.
2. **Add the pre-release job second**, when there is an audience for it. It is a small addition once
   the credential story works, because it differs only by `--pre-release` and the version check.

Doing (1) also settles the caveat above: whether `azure/login` plus `AzureCliCredential` actually
works from GitHub Actions is currently an inference from reading vsce's source. One successful
publish converts it into a fact.

## Origin

Raised by Wilson on 2026-08-16 immediately after replying to [@namgaw in discussion
#3][namgaw-in-discussion] with the 0.6.0 RC: *"shall we add the vs marketplace prerelease pipeline?
And how automated do we make it?"*

Deferred in the same conversation, on the grounds that we had just told a tester the version is
`0.6.0` and should see how the first real feedback lands before spending a version number on
machinery for an audience of one.

[namgaw-in-discussion]: https://github.com/ewc3labs/excel-power-query-editor/discussions/3
