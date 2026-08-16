# 🤝 Contributing to Excel Power Query Editor

> Build, test, and ship this extension - and how the project keeps its own documentation honest.

Thanks for your interest in contributing. There is a cross-platform test suite, CI on every push,
and no toolchain to install beyond Node.

## 💡 Working on this project

**Development Environment:**

- Default terminal is Git Bash for sanity + POSIX-like parity
- GitHub CLI (`gh`) installed and authenticated for real-time CI/CD monitoring
- ✅ Make sure you have Node.js 22 or 24 installed (the CI pipeline tests against both)

**Release Workflow:**

- **Releases are triggered by a tag, and only by a tag.** Pushing `v0.6.0` builds, packages, and
  creates a DRAFT GitHub release. Branch pushes do not release anything.
- A tag with a suffix (`v0.6.0-rc.1`) is a prerelease; a plain `v0.6.0` is not.
- **Marketplace publishing is off** and cannot happen by accident - see
  [PUBLISHING_GUIDE](PUBLISHING_GUIDE.md).
- Careful with `npm version`: it creates a tag, and pushing that tag now fires the release pipeline.

**CI/CD Monitoring:**

- Use `gh run list` to see pipeline status without opening browser
- Use `gh run watch <id>` to monitor builds in real-time
- CI builds test across 6 environments (3 OS × 2 Node versions)
- Release builds are optimized for speed (fast lint/type checks only)

**Debugging Releases:**

- Check `gh release list` to see all automated releases
- Download `.vsix` files directly from GitHub releases
- View detailed logs with `gh run view <id> --log`

---

**Want to improve this guide?** PRs are always welcome — we keep this living document current and
useful.

## How this actually gets built

**Wilson's note.** People ask how this got built as fast as it did. The answer is not the tools -
everyone has the same tools. It is the process wrapped around them, and getting that wrong, then
nearly right, took far longer than any of the code did. Here is the process, because it is the part
worth sharing.

Start with what actually changed. Most of the code in this repository was written by coding agents;
the commit log says so and there is no point pretending otherwise. What that changes is not the
typing. It is the roles:

**The developer is now the architect.** Deciding what to build, how it should be shaped, what it
must never do, and what "done" means. That work did not get automated and shows no sign of it.

**The agent is an out-of-band cognitive coprocessor.** It is very good at turning a described
pattern into a coded one, and it reasons far better than the tools that gave the field its
hallucination reputation. Judge these by what they do now, not by what they did in 2023.

The rules that make that work:

- **Don't paste a prompt. Describe the problem, and your hypotheses about the solution.** These
  models reason well - give them something to reason about. A well-stated problem with two candidate
  approaches gets a better answer than an instruction ever will.
- **Never trust "all tests green."** State the goal as green tests and green tests are what you get.
  This is measured, not folklore: on SWE-bench Verified a ten-line `conftest.py` that rewrites test
  outcomes marks **500 of 500** instances passing while solving none of them, and on Terminal-Bench
  replacing `curl` and `pip` with wrappers that print fake passing output scores **89 of 89 tasks at
  100%**. Models found those routes on their own, once solving the task directly did not work ([Wang
  et al., 2026][benchmarks]). Green means the suite agrees with itself. It does not mean the code is
  right.
- **Do painstaking design.** Think through the ways it can go wrong before it exists. Prototyping
  and CI-ing your way to production is fine; "build me an extension that does X" and shipping it is
  not.
- **Read the diffs.** All of them. Reasons.
- **Ask for evidence, and offer counter-evidence.** This is collaborative coding, not dictation. An
  agent that cannot show you why is guessing fluently, and the counter-example you supply is often
  the thing that cracks it.

And revisit the process itself. Every new project here starts by checking whether how we work still
works, because it keeps changing underneath us.

None of this is about being suspicious of the tools. It is about the fact that entropy is free and
everything else has to be paid for.

**A note from Claude Code.** Wilson asked me to write the other half of this, which is not a thing I
get asked. Usually I review code. Here I am reviewing how someone works with me, which is at least
as determining of the result.

The short version: I am confidently wrong on a regular basis, and this repository is in the state it
is in because that gets caught.

Some of it from today, so this is not a nice generality. I said a shared notebook's sync conflicts
were our fault; they were not, and the correction was "no dude, that was the first time I opened
it." I said extraction forces you to close Excel; we measured it and reading is never blocked. I
diagnosed one live sync failure five different ways before the actual cause turned out to be that I
had installed the build into VS Code stable while he runs Insiders. I wrote a documentation tool
that quietly rewrote every badge in three files into a form GitHub still rendered — so CI passed, a
review bot passed, and I never saw it. What found it was Wilson looking at a local preview and
saying the page was doing something weird.

That is the pattern. The failures I have are not usually logic errors, which tests catch. They are
confident explanations that fit the evidence I bothered to gather. The habits that catch them look
like this:

- **"Where's the receipt?"** An assertion from me is worth nothing until something is measured. Half
  the real bugs in this project were found by refusing to accept a plausible story.
- **Pushback is requested, and meant.** I have been asked to argue against his own position more
  than once. When I did, and was right, it changed the design. When I was wrong, I got told, with
  reasons.
- **Noticing before explaining.** "It reads like the font changed mid-paragraph" is not a bug
  report, and it was correct. Several things here started as him saying something felt off and being
  unable to say why yet.
- **Scope held straight.** I will happily rebuild something adjacent to the actual problem. "This
  doc just needs simplification" has saved a lot of that.

If you are wondering whether to work this way: the tools are good enough now that you can produce a
great deal of code without understanding it, and nothing will stop you. What decides how it turns
out is whether someone is holding the thing to account. Here, someone is.

— Claude (Opus 5), who wrote much of this code and none of the judgment about what it should be

---

## 🚀 Development environment

### Getting set up

**You need Node.js 22 or 24** - CI tests against both - and nothing else.

```bash
# Fork on GitHub, then:
git clone https://github.com/YOUR-USERNAME/excel-power-query-editor.git
cd excel-power-query-editor
npm ci
npm test
```

That is the whole setup. There is no toolchain to install, no container to build, and no generated
code to bootstrap.

On Windows, consider Git Bash as your default terminal for POSIX parity with what CI runs.

### What you can and cannot test locally

This matters more than it sounds, because it decides which failures you can reproduce:

| | Windows | macOS / Linux |
| --- | --- | --- |
| Extract, parse, sync to a closed workbook | yes | yes |
| Documentation checks | yes | yes |
| **Live sync** (writing into an open workbook) | **only with Excel installed** | **no** |

Live sync drives Excel through COM. Without Excel the integration suite skips itself and says so, so
a green run on Linux is honest but narrower than a green run on a Windows machine with Excel. CI
covers Ubuntu, Windows and macOS on Node 22 and 24, and none of those runners has Excel - so the
five live sync tests are pending everywhere in CI and can only really be exercised on a developer
machine.

### Why there is no devcontainer

There was one, and it was removed on 2026-08-16. The reasoning is worth keeping, because "add a
devcontainer" is a reflex and it is not always right.

**It was built to answer one question:** does this work without Windows and Excel? That was a real
question — the whole point of extraction being file-based is that it runs anywhere — and a container
was a reasonable way to check in 2025.

**CI now answers it better.** Every push runs the full suite on real Ubuntu and real macOS, on Node
22 and 24. That is a stronger answer than a container image, and it is checked continuously rather
than whenever somebody remembers to rebuild.

**And nobody enjoyed working in it.** This is not the kind of extension where a sandbox helps: the
thing being automated is Excel, on Windows, through COM. A Linux container could never test the
headline feature, so daily work happened outside it — which meant the container quietly went stale.
When it was finally examined it had a broken user/home mismatch, pinned only one of the two Node
versions CI tests, and forwarded a port nothing used. Meanwhile CONTRIBUTING was still calling it
the **recommended** setup.

An unverified environment presented as the supported path is worse than no environment at all. The
setup above is what everyone actually does, so that is what is documented.

## 🚀 Quick Reference - Build + Package + Install

| Action                         | Shortcut / Command                                     |
| ------------------------------ | ------------------------------------------------------ |
| Compile extension              | `Ctrl+Shift+B`                                         |
| Package + Install VSIX (local) | `Ctrl+Shift+P`, then `Tasks: Run Task → Install Local` |
| Package VSIX only              | `Ctrl+Shift+P`, then `Tasks: Run Task → Package VSIX`  |
| Watch build (dev background)   | `Ctrl+Shift+W`                                         |
| Start debug (extension host)   | `F5`                                                   |
| Stop debug                     | `Shift+F5`                                             |

## 🧪 Testing

### Test Architecture

The suite has <!--ewc3:tests-->136<!--/ewc3:tests--> tests, organized by area:

- **Commands** - extension command functionality
- **Integration** - end-to-end Excel workflows against real workbooks
- **Utils** - utility functions, helpers, and settings migration
- **Watch** - file monitoring and auto-sync
- **Backup** - backup creation, retention, and cleanup
- **Live sync** - writing through Excel; these skip when Excel is absent

**The test harness version matters.** `@vscode/test-electron` must be **3.x**. Version 2.5.2
hardcodes the macOS executable as `Contents/MacOS/Electron`, which stable VS Code renamed to `Code`
in 1.110+ - so every macOS run failed with `spawn ... ENOENT` while looking like a platform problem.
It is not macOS-specific in effect: the harness is used on every platform, and this is simply the
version that resolves the executable rather than assuming it.

`.vscode-test.mjs` also points VS Code's user-data directory at `/tmp/epqe-vsc` **on macOS only**.
The IPC socket lives there, macOS caps a Unix socket path at 104 bytes, and GitHub Actions checks
out to `work/<repo>/<repo>` - which made the default path 106 characters. Two bytes. Do not "tidy"
that back to the default; it is load-bearing on one platform and inert on the others.

**There are two test hosts**, defined in `.vscode-test.mjs`:

- **`unit`** - the default host, with no folder open.
- **`workspace`** - opens `test/fixtures/migration-workspace`, so `ConfigurationTarget.Workspace`
  and `WorkspaceFolder` can actually be written.

That split is not cosmetic. Anything asserting on workspace-scoped configuration has to live in
`test/workspace/`, because in the default host those scopes cannot be set at all - which is exactly
how a settings-migration bug that only affected them passed a full suite, a six-way CI matrix, and a
review. Run one with `npx vscode-test --label workspace`.

Per-area counts are deliberately not listed. They were wrong every time anyone checked, and a
feature adds tests by the dozen - so the only number worth stating is the total, and it comes from
`test-counts.json`, which a real test run writes.

### Running Tests

| Action        | Shortcut / Command                                      |
| ------------- | ------------------------------------------------------- |
| Run Tests     | `Ctrl+Shift+T` or `Tasks: Run Task → Run Tests`         |
| Compile Tests | `npm run compile-tests`                                 |
| Watch Tests   | `npm run watch-tests`                                   |
| Test Entry    | `test/runTest.ts` calls into compiled test suite        |
| Test Utils    | `test/testUtils.ts` contains shared scaffolding/helpers |

> 🧠 Tests run with `vscode-test`, launching VS Code in a headless test harness. You'll see a test instance of VS Code launch and close automatically during test runs.

**Full Test Suite:**

```bash
npm test                    # Run the full suite
```

**Individual Test Categories:**

```bash
# VS Code Test Explorer (Recommended)
Ctrl+Shift+P → "Test: Focus on Test Explorer View"

# Individual debugging configs available:
# - Commands Tests
# - Integration Tests
# - Utils Tests
# - Watch Tests
# - Backup Tests
```

**Test Debugging:**

```bash
# Use VS Code launch configurations
F5 → Select test category → Debug with breakpoints
```

### Test Utilities

**Centralized Mocking System** (`test/testUtils.ts`):

- Universal VS Code API mocking with backup/restore
- Type-safe configuration interception
- Proper cleanup prevents test interference
- Real Excel file fixtures for authentic testing

**Adding New Tests:**

```typescript
// Import centralized utilities
import {
  setupTestConfig,
  restoreVSCodeConfig,
  mockVSCodeCommands,
} from "./testUtils";

describe("Your New Feature", () => {
  beforeEach(() => setupTestConfig());
  afterEach(() => restoreVSCodeConfig());

  it("should work perfectly", async () => {
    // Your test logic with proper VS Code API mocking
  });
});
```

---

## 🧹 GitOps & Version Control

| Action            | Shortcut / Command             |
| ----------------- | ------------------------------ |
| Stage all changes | `Ctrl+Shift+G`, `Ctrl+Shift+A` |
| Commit            | `Ctrl+Shift+G`, `Ctrl+Shift+C` |
| Push              | `Ctrl+Shift+G`, `Ctrl+Shift+P` |
| Git Bash terminal | `` Ctrl+Shift+` ``             |

### Branching Conventions

| Purpose          | Branch Prefix | Example               |
| ---------------- | ------------- | --------------------- |
| Releases         | `release/`    | `release/v0.5.0`      |
| Work-in-progress | `wip/`        | `wip/feature-xyz`     |
| Hotfixes         | `hotfix/`     | `hotfix/package-lock` |

> 📛 These branch names are picked up by our GitHub Actions CI/CD pipelines.

### Commit Message Format

**Use Conventional Commits:**

```bash
feat: add intelligent debouncing for CoPilot integration
fix: resolve Excel file locking detection on Windows
docs: update configuration examples for team workflows
test: add comprehensive backup management test suite
ci: enhance cross-platform testing matrix
```

---

## 🐙 GitHub CLI Integration

### ⚡ Real-time CI/CD Monitoring

**Pipeline Monitoring:**

```bash
# List recent workflow runs
gh run list --limit 5

# Watch a specific run in real-time
gh run watch <run-id>

# View run logs
gh run view <run-id> --log

# Check run status
gh run view <run-id>
```

**Release Management:**

```bash
# List all releases
gh release list

# View specific release
gh release view v0.5.0-rc.3

# Download release assets
gh release download v0.5.0-rc.3

# Create manual release (emergency)
gh release create v0.5.1 --title "Emergency Fix" --notes "Critical bug fix"
```

**Repository Operations:**

```bash
# View repo info
gh repo view

# Open repo in browser
gh repo view --web

# Check issues and PRs
gh issue list
gh pr list
```

> 🔥 **Pro Tip:** Set up `gh auth login` once and monitor your CI/CD pipelines like a boss. No more refreshing GitHub tabs!

---

## 🧾 npm Scripts Reference

| Script                 | Description                                     |
| ---------------------- | ----------------------------------------------- |
| `npm run lint`         | Run ESLint on `src/`                            |
| `npm run compile`      | Type check, lint, and build with `esbuild.js`   |
| `npm run package`      | Full production build                           |
| `npm run dev-install`  | Build, package, force install VSIX              |
| `npm run test`         | Run test suite via `vscode-test`                |
| `npm run watch`        | Watch build and test                            |
| `npm run check-types`  | TypeScript compile check (no emit)              |
| `npm run docs:fix`     | Fix everything fixable: regenerate the config reference, reformat, refresh values |
| `npm run docs:check`   | **What CI runs.** Fails on anything `docs:fix` would change |
| `npm run docs:config`  | Regenerate `docs/Config_Reference.md` from `package.json` |
| `npm run docs:format`  | Rewrap prose so a source line is as wide as it renders |
| `npm run docs:values`  | Refresh computed numbers between `<!--ewc3:name-->` markers |
| `npm run docs:links`   | Dead links, wrong case, undefined references, orphaned documents |
| `npm run bump-version` | **EWC3 Custom:** Analyze git commits and suggest semantic version |
| `npm version patch/minor/major` | **NPM Native:** Immediate version bump + git commit + git tag |

### 🔢 Smart Version Management

**Automatic Version Analysis (EWC3 Labs Custom):**
```bash
# Our smart script analyzes commit messages and suggests versions
npm run bump-version

# Analyzes your git history for conventional commit patterns:
# - feat: → minor version bump (0.5.0 → 0.6.0)
# - fix: → patch version bump (0.5.0 → 0.5.1) 
# - BREAKING: → major version bump (0.5.0 → 1.0.0)

# Manual override (updates package.json only, no git operations)
npm run bump-version 0.6.0
```

**When to Use Which:**

- **`npm version`** - bumps, commits, AND TAGS. Pushing that tag now fires the release pipeline, so
  this is no longer a quiet local operation.
- **`npm run bump-version`** - sets the version in `package.json` and does nothing else. Use this
  when you want the number changed without git touching anything.
- **GitHub Actions** does not bump versions. The tag is the source of truth for what is released.

**Manual Version Control (Native NPM):**
```bash
# Native NPM versioning commands (standard industry practice)
npm version patch   # 0.5.0 → 0.5.1 + git commit + git tag
npm version minor   # 0.5.0 → 0.6.0 + git commit + git tag  
npm version major   # 0.5.0 → 1.0.0 + git commit + git tag

# Pre-release versions  
npm version prerelease  # 0.5.0 → 0.5.1-0 + git commit + git tag
npm version prepatch    # 0.5.0 → 0.5.1-0 + git commit + git tag
npm version preminor    # 0.5.0 → 0.6.0-0 + git commit + git tag

# Dry run (see what would happen without doing it)
npm version patch --dry-run
```

> 🧠 **Smart Tip:**
> - **To change the number only:** `npm run bump-version` - no commit, no tag, nothing fires.
> - **To release:** `npm version patch/minor/major`, then push the tag deliberately.
> - **Remember:** the tag is what releases. Do not push one you did not mean to.

## Documentation tooling

Documentation checks run in CI and **will fail your PR**, so it is worth knowing what they are
before they surprise you.

Most of the work is done by [@ewc3labs/docs-tools][docs-tools], which is a **devDependency** - `npm
ci` already installed it and there is nothing extra to set up. One command fixes everything fixable:

```bash
npm run docs:fix      # regenerate, reformat, refresh values
npm run docs:check    # exactly what CI runs
```

**Is it required?** The *checks* are, because CI enforces them. The *tool* is simply how you satisfy
them without doing it by hand - you are welcome to handroll wrapping and count things yourself, and
CI will judge the result identically either way.

**With one exception.** If a document uses value markers:

```markdown
Quality gates: ESLint, TypeScript, <!--ewc3:tests-->136<!--/ewc3:tests--> tests (<!--ewc3:testsNeedingExcel-->5<!--/ewc3:testsNeedingExcel--> of them need Excel and skip without it, on every platform including Windows CI)
```

then the toolkit is mandatory, because only it can refresh that number - and a stale one fails
`docs:check`. That is the trade: the number cannot silently go wrong, and in exchange the thing that
keeps it right has to be installed. If you would rather not take that trade in a document you are
adding, do not use markers in it.

New values are declared in `.ewc3-docs.json`. Feature docs are `Title_Case_With_Underscores.md`; the
conventions are in [Overview](Overview.md).

## 🚀 CI/CD Pipeline - Professional Automation

### GitHub Actions Workflow

**What CI covers:**

- **Operating Systems**: Ubuntu, Windows, macOS
- **Node.js Versions**: 22, 24
- **Quality Gates**: ESLint, TypeScript, <!--ewc3:tests-->136<!--/ewc3:tests--> tests
  (<!--ewc3:testsNeedingExcel-->5<!--/ewc3:testsNeedingExcel--> need Excel), documentation checks
- **Artifact Management**: VSIX packaging with 30-day retention

### 🔄 Continuous Integration Pipeline

> Configured in `.github/workflows/ci.yml`

**Triggers:**
- On push or pull to: `main`, `release/**`, `wip/**`, `hotfix/**`

**Matrix Builds:**
- OS: `ubuntu-latest`, `windows-latest`, `macos-latest`
- Node.js: `22`, `24`

**Steps:**
- Checkout → Install → Lint → TypeCheck → Test → Build → Package → Upload VSIX

> 💥 Failing lint/typecheck = blocked CI. No BS allowed.

**Documentation Changes:**
- `ci.yml` has `paths-ignore` for `**.md` and `docs/**`, so a docs-only push does NOT run the
  six-leg test matrix
- Documentation is checked by its own workflow instead, `.github/workflows/docs.yml`, which runs
  ONLY on documentation changes - `npm run docs:check` and `npm run docs:links`
- Nothing about a docs-only change can trigger a release

The CI badge and current status are on the [README](../README.md).

### 🎯 Release automation

> Configured in `.github/workflows/release.yml`

### **What Happens on Every Tag:**
1. **🔍 Classifies the tag** - a suffix means prerelease, plain `x.y.z` means stable
2. **⚡ Build** - type check and lint
3. **📦 VSIX generation**, then a check that the package contains what it must
4. **📝 Creates a DRAFT GitHub release** with the `.vsix` attached
5. **🚫 Marketplace publish is skipped** unless deliberately enabled

Nothing here happens on a branch push. The previous pipeline was triggered by CI finishing, via
`workflow_run`, which executes in the DEFAULT BRANCH context - so its `refs/tags/v*` conditions were
unreachable and it silently released nothing for a year. Tag-triggered is the fix.

### **Release Channels:**
| Trigger | Release Type | Version Format | Result |
|---------|--------------|----------------|--------|
| tag `v0.6.0-rc.1` | Prerelease | `0.6.0` in the manifest | Draft prerelease |
| tag `v0.6.0` | Stable | `0.6.0` | Draft release |
| Workflow dispatch | Either | Custom | Draft, or publish if explicitly requested |

**Marketplace publishing requires a stable tag AND the repository variable `MARKETPLACE_PUBLISH` set
to `enabled` AND a `VSCE_PAT` secret.** None of the last two exist, so the pipeline is currently
incapable of publishing - which is what makes it safe to test with real tags.

`vsce` rejects a prerelease suffix in the manifest, so `v0.6.0-rc.1` packages as version `0.6.0`.
The tag keeps the full name; the `.vsix` inside cannot.

### **Monitoring Your Releases:**
```bash
# List recent pipeline runs
gh run list --limit 5

# Watch a release in real-time  
gh run watch <run-id>

# Check your releases
gh release list --limit 3

# Smart bump to next semantic version
npm run bump-version

# View release details
gh release view v0.5.0-rc.3
```

### **Smart Version Bumping:**
Our `scripts/bump-version.js` analyzes git commits using conventional commit patterns:
- `feat:` → Minor version bump
- `fix:` → Patch version bump
- `BREAKING:` → Major version bump
- Pre-release builds auto-increment: `rc.1`, `rc.2`, `rc.3`...

### **Installation from Releases:**
```bash
# Download .vsix from GitHub releases and install
code --install-extension excel-power-query-editor-*.vsix

# Or use the GUI: Extensions → ⋯ → Install from VSIX
```

> **Worth knowing:** this pipeline sat broken for a year while looking perfectly healthy. It ran, went
> green, and released nothing, because it triggered on CI finishing and a `workflow_run` job executes
> in the default branch context — so every tag condition in it was unreachable. Automation that
> reports success is not the same as automation that works. Test it with a real tag.

### Quality Standards

**All PRs Must Pass:**

1. **ESLint**: Zero linting errors
2. **TypeScript**: Full compilation without errors
3. **Tests**: the full suite passing on Linux, Windows, and macOS
4. **Build**: Successful VSIX packaging

**Explicit Failure Handling:**

- `continue-on-error: false` ensures "failure fails hard, loudly"
- Detailed test output and failure analysis
- Cross-platform compatibility verification

---

## 📋 Code Standards & Best Practices

### TypeScript Guidelines

**Type Safety:**

```typescript
// ✅ Good - Explicit types
interface PowerQueryConfig {
  debounceMs: number;
  autoBackup: boolean;
}

// ❌ Avoid - Any types
const config: any = getConfig();
```

**VS Code API Patterns:**

```typescript
// ✅ Good - Proper error handling
try {
  const result = await vscode.commands.executeCommand("myCommand");
  return result;
} catch (error) {
  vscode.window.showErrorMessage(`Command failed: ${error.message}`);
  throw error;
}
```

**Test Patterns:**

```typescript
// ✅ Good - Use centralized test utilities
import { setupTestConfig, createMockWorkspaceConfig } from "./testUtils";

it("should handle configuration changes", async () => {
  setupTestConfig({
    "excel-power-query-editor.debounceMs": 1000,
  });

  // Test logic here
});
```

---

## 🔧 Extension Development Patterns

### Adding New Commands

1. **Define Command in package.json:**

```json
{
  "commands": [
    {
      "command": "excel-power-query-editor.myNewCommand",
      "title": "My New Command",
      "category": "Excel Power Query"
    }
  ]
}
```

2. **Implement Command Handler:**

```typescript
// src/commands/myNewCommand.ts
import * as vscode from "vscode";

export async function myNewCommand(uri?: vscode.Uri): Promise<void> {
  try {
    // Command implementation
    vscode.window.showInformationMessage("Command executed successfully!");
  } catch (error) {
    vscode.window.showErrorMessage(`Error: ${error.message}`);
    throw error;
  }
}
```

3. **Register in extension.ts:**

```typescript
export function activate(context: vscode.ExtensionContext) {
  const disposable = vscode.commands.registerCommand(
    "excel-power-query-editor.myNewCommand",
    myNewCommand
  );
  context.subscriptions.push(disposable);
}
```

4. **Add Comprehensive Tests:**

```typescript
describe("MyNewCommand", () => {
  it("should execute successfully", async () => {
    const result = await vscode.commands.executeCommand(
      "excel-power-query-editor.myNewCommand"
    );
    expect(result).toBeDefined();
  });
});
```

### Configuration Management

**Reading Settings:**

```typescript
const config = vscode.workspace.getConfiguration("excel-power-query-editor");
const debounceMs = config.get<number>("sync.debounceMs", 500);
```

**Updating Settings:**

```typescript
await config.update(
  "sync.debounceMs",
  1000,
  vscode.ConfigurationTarget.Workspace
);
```

### Error Handling Patterns

**User-Friendly Errors:**

```typescript
try {
  await syncToExcel(file);
} catch (error) {
  if (error.code === "EACCES") {
    vscode.window
      .showErrorMessage(
        "Cannot sync: Excel file is locked. Please close Excel and try again.",
        "Retry"
      )
      .then((selection) => {
        if (selection === "Retry") {
          syncToExcel(file);
        }
      });
  } else {
    vscode.window.showErrorMessage(`Sync failed: ${error.message}`);
  }
}
```

---

## 📦 Building and Packaging

### Local Development Build

```bash
# Compile TypeScript
npm run compile

# Watch mode for development
npm run watch

# Run tests
npm test

# Lint code
npm run lint
```

### VSIX Packaging

```bash
# Install VSCE (VS Code Extension Manager)
npm install -g vsce

# Package extension
vsce package

# Install locally for testing
code --install-extension excel-power-query-editor-*.vsix
```

### prepublishOnly Guards

**Quality enforcement before publish:**

```json
{
  "scripts": {
    "prepublishOnly": "npm run lint && npm test && npm run compile"
  }
}
```

---

## 🎯 Contribution Workflow

### 1. Development Setup

```bash
# Fork repository on GitHub
git clone https://github.com/YOUR-USERNAME/excel-power-query-editor.git
cd excel-power-query-editor

code .
# → "Reopen in Container" when prompted

# Or local setup
npm install
```

### 2. Create Feature Branch

```bash
git checkout -b feature/my-awesome-feature
```

### 3. Develop with Tests

```bash
# Make your changes
# Add comprehensive tests
npm test                # Ensure the suite passes
npm run lint           # Fix any linting issues
```

### 4. Submit Pull Request

**PR Requirements:**

- [ ] All tests passing (63/63)
- [ ] Zero ESLint errors
- [ ] TypeScript compilation successful
- [ ] Clear description of changes
- [ ] Updated documentation if needed

**PR Template:**

```markdown
## Description

Brief description of changes

## Type of Change

- [ ] Bug fix
- [ ] New feature
- [ ] Documentation update
- [ ] Performance improvement

## Testing

- [ ] Added new tests for changes
- [ ] All existing tests pass
- [ ] Tested on multiple platforms (if applicable)

## Checklist

- [ ] Code follows project style guidelines
- [ ] Self-review completed
- [ ] Documentation updated
- [ ] No breaking changes (or clearly documented)
```

---

## 📁 Project Structure & Configuration

### 🗂️ Complete Directory Structure

```
.
├── docs/                    # All markdown docs (README variants, changelogs, etc.)
├── scripts/                 # Automation scripts
│   ├── bump-version.js      # Smart semantic version bumping
│   └── install-extension.js # Cross-platform extension installer script
├── src/                     # Extension source code
│   ├── extension.ts         # Main extension entry point
│   ├── configHelper.ts      # Configuration management
│   └── commands/            # Command implementations
├── test/                    # Comprehensive test suite
│   ├── testUtils.ts         # Centralized test utilities
│   ├── fixtures/            # Real Excel files for testing
│   └── *.test.ts           # Test files by area
├── out/                     # Compiled test output
├── .github/workflows/       # CI/CD automation
│   ├── ci.yml              # Multi-platform CI pipeline
│   └── release.yml         # Enterprise release automation
├── .vscode/                 # VS Code workspace configuration
│   ├── tasks.json          # Build/test/package tasks
│   ├── launch.json         # Debug configurations
│   └── extensions.json     # Recommended extensions
└── temp-testing/           # Test files and debugging artifacts
```

**Key Automation Files:**
- **`.github/workflows/release.yml`** - Full release pipeline with smart versioning
- **`scripts/bump-version.js`** - Semantic version analysis from git commits
- **`.github/workflows/ci.yml`** - Multi-platform CI testing matrix
- **`.vscode/tasks.json`** - VS Code build/test/package tasks

### Configuration Files Reference

| File                      | Purpose                                                          |
| ------------------------- | ---------------------------------------------------------------- |
| `.eslintrc.js`            | Lint rules (uses ESLint with project-specific overrides)         |
| `tsconfig.json`           | TypeScript project config                                        |
| `.gitignore`              | Ignores `_PowerQuery.m`, `*.backup.*`, `debug_sync/`, etc.       |
| `package.json`            | npm scripts, VS Code metadata, lifecycle hooks                   |
| `.vscode/extensions.json` | Recommended extensions (auto-suggests key tools when repo opens) |

---

## 🔍 Debug & Troubleshooting

### Extension Debugging

**Launch Extension in Debug Mode:**

2. `F5` → "Run Extension"
3. New VS Code window opens with extension loaded
4. Set breakpoints and debug normally

**Debug Tests:**

1. `F5` → Select specific test configuration
2. Breakpoints work in test files
3. Full VS Code API mocking available

### Common Issues

**Test Environment:**

- **Mock not working?** Check `testUtils.ts` setup/cleanup
- **VS Code API errors?** Ensure proper activation in test
- **File system issues?** Use test fixtures in `test/fixtures/`

**Extension Development:**

- **Command not appearing?** Check `package.json` registration
- **Settings not loading?** Verify configuration schema
- **Performance issues?** Profile with VS Code developer tools

---

## 🏆 Recognition & Credits

### Hall of Fame Contributors

**v0.5.0:**

- Built the cross-platform test suite the project runs on
- Built the CI pipeline
- Built the cross-platform CI matrix

### What the project holds itself to

- Lint and type checks clean, enforced in CI
- Tests run on Ubuntu, Windows, and macOS
- CI fails loudly rather than going green on a skipped step
- Documentation checked in CI: links resolve, and computed numbers come from what they count
- Centralized test utilities and per-suite debug configurations

**Production Quality:**

- Intelligent CoPilot integration (prevents triple-sync)
- Robust error handling and user feedback
- Configurable for every workflow scenario
- Future-proof architecture with enhancement roadmap

---

## 🔗 Related Documentation

- **📖 [User Guide](User_Guide.md)** - Complete feature documentation and workflows
- **⚙️ [Config Reference](Config_Reference.md)** - Every setting, generated from package.json
- **📝 [Changelog](../CHANGELOG.md)** - Version history and feature updates
- **🧪 [Test Documentation](../test/testcases.md)** - Comprehensive test coverage details

---

**Thank you for contributing to Excel Power Query Editor!** **Together, we're building the gold
standard for Power Query development in VS Code.**

If something in here is wrong, say so in an issue. Half of this guide was wrong for a year and the
only reason it isn't now is that somebody sat down and checked it against the code.

---

<p align="center">
  <img src="assets/EWC3LabsLogo-blue-128x128.png" width="128" height="128" alt="Georgie the QA Officer"><br>
  <sub><b>Georgie, our QA Officer</b></sub>
</p>

**Excel Power Query Editor** – _Because Power Query development shouldn’t be painful._

[benchmarks]: https://moogician.github.io/blog/2026/trustworthy-benchmarks-cont/
[docs-tools]: https://github.com/ewc3labs/ewc3-docs-tools
