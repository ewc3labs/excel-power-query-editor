# Marketplace identity setup

The manual, one-time identity work behind Marketplace publishing — and the record of what was
created, so nobody has to reverse-engineer it later.

**No credentials belong in this file.** Client ID and tenant ID are identifiers, not secrets, and
they will live in repository variables. Nothing here should ever be a secret, because this design
has no long-lived credential to record.

## Two authentication modes, selected by configuration

| `MARKETPLACE_AUTH` | Path | Status |
| --- | --- | --- |
| `entra` | Entra ID + workload identity federation → `vsce publish --azure-credential` | **Microsoft's documented production route today** |
| `oidc` | GitHub OIDC → Marketplace directly → `vsce publish --oidc` | implemented, waiting on the Marketplace to expose the trust policy |

**There is no default and no fallback.** An unset or misspelled value fails the publish job rather
than guessing, and a failure in one mode never retries in the other. The mode selects one
authentication contract; silently satisfying it a different way would mean nobody ever knows which
one actually published the extension.

Switching later is one repository variable:

```text
MARKETPLACE_AUTH: entra → oidc
```

No workflow redesign. Everything else — the protected environment, human approval,
`id-token: write`, the `MARKETPLACE_PUBLISH` gate, the channel logic, the prebuilt VSIX — is shared
by both modes.

## Why `entra` first

Trusted publishing (`--oidc`) is real, implemented, and shipped in `vsce` 3.9.3-5 — but the
**Marketplace side has no configuration surface for it yet**. The trust policy that would name this
repository and workflow is documented in the codebase and absent from the UI. Until it appears,
`--oidc` has nothing to trust it.

`--azure-credential` is the route Microsoft's own publishing guide walks through today, and it works
on stable `vsce`. So the interim path runs **no preview code at all** — the prerelease pin applies
only to the `oidc` mode.

## Setup for `entra`

### 1. Establish which tenant you actually control

Microsoft Entra admin center → **Identity → Overview**, signed in as the identity that owns the
`ewc3labs` Marketplace publisher. Record tenant name, tenant ID, and primary domain below **before
creating anything.**

**An app registration lives in the tenant that created it and cannot be moved to another one.** If
the sign-in lands somewhere that is not a tenant you control — a directory named for an employer, or
`Microsoft Services` — stop. Creating the publishing identity in someone else's tenant is not a
mistake you fix, it is one you redo.

### 2. Create an app registration with a federated credential

Microsoft's walkthrough uses a **user-assigned managed identity**, which requires an Azure
subscription. This project uses an **app registration federated directly to the GitHub environment**
instead — a supported pattern, and an adaptation of Microsoft's pieces rather than their exact
recipe. Note the difference if their steps and these ever disagree.

Federated credential:

| | |
| --- | --- |
| Issuer | `https://token.actions.githubusercontent.com` |
| Subject | `repo:ewc3labs/excel-power-query-editor:environment:marketplace` |
| Audience | `api://AzureADTokenExchange` |

**The subject names the environment, not a branch or tag.** A ref-based subject would need a new
federated credential for every release, or a wildcard, which defeats the purpose.

### 3. Find the profile id — this is the step everyone gets wrong

**The publisher member is not the client ID.** Microsoft's procedure is to authenticate *as* the
identity, call the Azure DevOps profile endpoint, and use the `id` from the response:

```bash
az rest -u https://app.vssps.visualstudio.com/_apis/profile/profiles/me \
        --resource 499b84ac-1321-427f-aa17-267ca6975798
```

**You cannot run that locally.** The credential is federated to a GitHub environment, so it only
works from a job bound to that environment — and creating a client secret to work around this
reintroduces exactly the long-lived credential the design exists to remove.

So run the **Identity** workflow instead: Actions → Identity → **Run workflow from `main`**. It
signs in with the federated credential, calls the endpoint, and prints the id to the run summary.

**It must be `main`.** The `marketplace` environment allows only `main` and `v*`, so a run launched
from another branch is refused by the environment before it starts — correctly, but the message
points at deployment rules rather than at the branch you picked.

That run doubles as the **first real test of the federated credential**. If it prints an id, the
trust relationship is correct — proven before a release depends on it.

### 4. Add the member and set the variables

Add the profile id as a member of the `ewc3labs` publisher with the **Contributor** role, at
<https://marketplace.visualstudio.com/manage/publishers/ewc3labs>.

The `marketplace` environment already exists and is configured correctly:

| | |
| --- | --- |
| Required reviewer | `Wilson421` |
| Prevent self-review | **off** — with one maintainer, on would deadlock the pipeline |
| Allowed refs | `main` (branch), `v*` (tag) |

| Kind | Name | Value |
| --- | --- | --- |
| Variable | `MARKETPLACE_AUTH` | `entra` |
| Variable | `AZURE_CLIENT_ID` | the app registration's client ID |
| Variable | `AZURE_TENANT_ID` | the directory (tenant) ID |
| Variable | `MARKETPLACE_PUBLISH` | `enabled`, when you want tags to publish |
| Environment | `marketplace` | with a required reviewer — see the [Publishing Guide][guide] |

## Setup for `oidc`, when it becomes available

Configure a trusted publishing policy on the publisher naming `ewc3labs/excel-power-query-editor`
and `release.yml`, then set `MARKETPLACE_AUTH` to `oidc`. `AZURE_CLIENT_ID` and `AZURE_TENANT_ID`
become unused, the `azure/login` step skips itself, and the app registration can be deleted.

## The record

Established 2026-08-17. Identifiers only — this design has no secret to record, and the app
registration deliberately has **no client secret**.

### Tenant

| | |
| --- | --- |
| Name | `EWC3 Labs` |
| Tenant ID | `2d4be3ba-e29f-4a9a-b19d-711e9b07539a` |
| Primary domain | `wilsonewc3.onmicrosoft.com` |

### App registration

| | |
| --- | --- |
| Name | `EPQE Marketplace Publisher` |
| Client ID | `0096515e-d1d0-462d-882f-66fd5fadfb02` |
| Tenancy | single-tenant |
| Redirect URI | none |
| Client secret | **none, deliberately** — federation is the whole point |

### Federated credential

| | |
| --- | --- |
| Name | `github-epqe-marketplace` |
| Issuer | `https://token.actions.githubusercontent.com` |
| Subject | `repo:ewc3labs/excel-power-query-editor:environment:marketplace` |
| Audience | `api://AzureADTokenExchange` |

> **The subject form is not the one Entra offered, and that matters.** Entra generated GitHub's newer
> **immutable-ID** subject; it was corrected by hand to the name-based form above, because this
> repository is configured with `use_immutable_subject: false`:
>
> ```json
> {"use_default":true,"use_immutable_subject":false,"sub_claim_prefix":"repo:ewc3labs/excel-power-query-editor"}
> ```
>
> **If that setting is ever flipped to `true`, this credential stops working** and the failure will
> look like a permissions problem rather than a subject mismatch. Change the federated credential
> subject at the same time, or do not flip it.

### GitHub configuration

| Kind | Name | Value | State |
| --- | --- | --- | --- |
| Variable | `AZURE_CLIENT_ID` | `0096515e-…fb02` | ✅ set |
| Variable | `AZURE_TENANT_ID` | `2d4be3ba-…539a` | ✅ set |
| Variable | `MARKETPLACE_AUTH` | `entra` | ⬜ unset — publishing fails closed |
| Variable | `MARKETPLACE_PUBLISH` | `enabled` | ⬜ unset — publish job unreachable |
| Environment | `marketplace` | reviewer `Wilson421`, refs `main` and `v*` | ✅ configured |
| Secret | `OVSX_PAT` | Open VSX | ⬜ unset — that job warns and skips |

**Both remaining variables are intentionally unset.** Nothing can publish, which is what makes the
next step safe to run.

### Azure DevOps profile

Retrieved 2026-08-17 by the Identity workflow, run 32082093157.

| | |
| --- | --- |
| Profile id | `76fbca24-39b1-6b8e-94c5-af13dee0a54b` |

**This is the value to add as a publisher member**, with the **Contributor** role — not the client
ID. It is an identifier, not a credential.

That run also proved the federation end to end: GitHub minted a token for the `marketplace`
environment, Entra accepted it, and the identity authenticated to Azure DevOps. **The name-based
subject was correct** — the immutable-ID form Entra generated by default would have been rejected.

### Still to do

| | |
| --- | --- |
| Publisher member added | not yet — add the profile id above, Contributor |
| `MARKETPLACE_AUTH` → `entra` | after the member is added |
| `MARKETPLACE_PUBLISH` → `enabled` | when you want tags to publish |

## What is still unproven

**Proven.** The federated credential works: GitHub → Entra → Azure DevOps, through the protected
environment, on run 32082093157. The release pipeline itself is proven on `v0.6.0-rc.3` by both tag
push and manual dispatch.

**Not proven.** Neither publish job has ever run. Two things remain untested: whether the publisher
membership grants what `vsce` needs, and whether `vsce`'s credential chain picks up the Azure CLI
session `azure/login` leaves behind — `AzureCliCredential` is second in the chain it builds, so it
should, but that is reading the source rather than observing it.

### Prove the auth on the pre-release channel first

**Marketplace versions only go up**, so publishing `0.7.0` means the next stable is `0.8.0`, not
`0.6.0`. That is not a cost — it is the cadence the even/odd convention already describes, and it
buys a genuinely safe first publish:

```text
0.7.0  pre-release   nobody has opted in -> effectively zero blast radius
0.8.0  stable        after the auth path is proven, to 5,450 users
```

The alternative — making the first ever use of an unproven publish path a stable release to 5,450
auto-updating installs — has the same irreversibility and none of the safety.

**Nothing is orphaned by the renumber.** `v0.6.0` was never released; only `v0.6.0-rc.1` through
`rc.3` exist. The version number is cheap and the release content is unchanged.

The README version badge is derived from `package.json`, so it follows a bump on its own.

[guide]: PUBLISHING_GUIDE.md
