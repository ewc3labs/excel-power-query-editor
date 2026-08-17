# Marketplace identity setup

The manual, one-time identity work behind Marketplace publishing — and the record of what was
created, so nobody has to reverse-engineer it later.

**No credentials belong in this file.** Client ID and tenant ID are identifiers, not secrets, and
they live in repository variables already. Nothing here should ever be a secret, because this design
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

So run the **Identity** workflow instead: Actions → Identity → Run workflow. It signs in with the
federated credential, calls the endpoint, and prints the id to the run summary.

That run doubles as the **first real test of the federated credential**. If it prints an id, the
trust relationship is correct — proven before a release depends on it.

### 4. Add the member and set the variables

Add the profile id as a member of the `ewc3labs` publisher with the **Contributor** role, at
<https://marketplace.visualstudio.com/manage/publishers/ewc3labs>.

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

Fill this in as the work is done. Identifiers only.

| | |
| --- | --- |
| Tenant name | *(not yet established)* |
| Tenant ID | *(not yet established)* |
| Primary domain | *(not yet established)* |
| App registration name | *(not yet created)* |
| Client ID | *(not yet created)* |
| Azure DevOps profile id | *(run the Identity workflow)* |
| Publisher member added | *(no)* |
| `MARKETPLACE_AUTH` | *(unset — publishing fails closed)* |

## What is still unproven

The release pipeline has been exercised end to end on `v0.6.0-rc.3`, by tag push and by manual
dispatch. **Neither publish job has ever run.** `--azure-credential` from GitHub Actions is an
adaptation of a documented Azure Pipelines path, and `--oidc` has no Marketplace policy to talk to
yet. The first successful publish is the proof for whichever mode goes first.

[guide]: PUBLISHING_GUIDE.md
