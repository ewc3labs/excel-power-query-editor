<div align="center">
<table>
<tr>
<td style="padding: 0 40px; vertical-align: middle;">
<img src="../../assets/EWC3LabsLogo-blue-128x128.png" alt="EWC3 Labs" width="72" height="72">
</td>
<td style="vertical-align: middle;">
<h1 style="margin: 0;">Excel Power Query Editor</h1>
<h3 style="margin: 5px 0;"><strong>Check-in Punchlist — stuff we noticed, and what is still open</strong></h3>
</td>
</tr>
</table>
</div>

---

Started: 2026-08-14

## How this works

Things we saw and need to address **sometime**. No evidence required to write one down — that is the
point. An item can sit here indefinitely without guilt.

An item becomes a roadmap slice once someone has done enough analysis to **document the problem with
receipts**. Before minting a new slice, check whether an existing one already covers it.

---

## 2026-08-14 — first pass, adopting the HQ conventions

**Items:**

- [x] No `AGENTS.md` and no line-ending policy in the repo that every other repo is told to copy.
      Added, and 33 tracked files renormalised from CRLF.
- [ ] `release.yml:260` contains a corrupted byte — `EF BF BD` (replacement character) in a step name.
      Cosmetic on its own; a fair signal about how carefully the file has been reviewed. [PQ-04]
- [ ] Three `.vsix` files sit in the repo root. Untracked and gitignored, so harmless, but they are
      build output living where a human looks first.
- [ ] `docs/archive/` holds 14 tracked files of v0.4.3 documentation. Worth deciding whether that is
      history worth carrying or clutter to delete.
- [ ] `generate-expected-results.js` sits at the repo root while everything else of its kind is in
      `scripts/`.
- [ ] `docs/analysis/` and `docs/design/` do not exist here, though the org conventions expect them.
      Not worth creating empty — create when there is something to put in them.
- [ ] Org-wide: `RAG_sessions` (lowercase) in 8 repos, `RAG_Sessions` in 2. HQ declares the capital
      canonical and is outvoted. Cheaper to move the standard than eight folders — needs a decision.
- [ ] The extension has real users and no telemetry, deliberately. That means we learn about breakage
      only from issues. Worth deciding whether that stays true forever, in writing, either way.
