# Excel Power Query Editor
A modern, reliable VS Code extension for editing Power Query M code directly from Excel files

---

# 🤝 Contributing to Excel Power Query Editor

> **Complete Developer Guide** - Build, test, commit, package, and ship VS Code extensions like a pro with EWC3 Labs' enterprise-grade development platform.

**Welcome to the most professional VS Code extension development environment you'll ever see!**

Thanks for your interest in contributing! This project has achieved **enterprise-grade quality**
with 63 comprehensive tests, cross-platform CI/CD, and a world-class development experience.

## 📋 Table of Contents

- [🚀 Development Environment](#-development-environment---devcontainer-excellence)
- [🚀 Quick Reference](#-quick-reference---build--package--install)
- [🧪 Testing](#-testing---enterprise-grade-test-suite)
- [🧹 GitOps & Version Control](#-gitops--version-control)
- [🐙 GitHub CLI Integration](#-github-cli-integration)
- [🧾 npm Scripts Reference](#-npm-scripts-reference)
- [🚀 CI/CD Pipeline](#-cicd-pipeline---professional-automation)
- [📋 Code Standards](#-code-standards--best-practices)
- [🔧 Extension Development](#-extension-development-patterns)
- [📦 Building and Packaging](#-building-and-packaging)
- [🎯 Contribution Workflow](#-contribution-workflow)
- [📁 Project Structure](#-project-structure--configuration)
- [🔍 Debug & Troubleshooting](#-debug--troubleshooting)
- [🏆 Recognition & Credits](#-recognition--credits)

**Want to jump to a specific section?** Use the GitHub-style anchors above or bookmark specific
sections like `#testing` or `#release-automation`.

---

<details>
<summary><strong>💡 Pro Developer Workflow Tips</strong> (click to expand)</summary>

**Development Environment:**

- DevContainers optional, but fully supported if Docker + Remote Containers is installed
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

</details>

---

**Want to improve this guide?** PRs are always welcome — we keep this living document current and
useful.

🔥 **Wilson's Note:** This is my first extension, first public repo, first devcontainer (first time
even using Docker), first automated test suite, and first time using Git Bash — so I'm drinking from
the firehose here and often learning as I go. That said, I **do** know how this stuff should work,
and EWC3 Labs is about building it right. Our goal is an enterprise-grade DX platform for VS Code
extension development. We went from manual builds to automated releases with smart versioning,
multi-channel distribution, and real-time monitoring. It's modular, CI-tested, scriptable, and
optimized for contributors. If you're reading this — welcome to the automation party. **From a
simple commit/push to professional releases. Shit works when you work it.**

---

## 🚀 Development Environment - DevContainer Excellence

### Quick Start (Recommended)

**Prerequisites:** Docker Desktop and VS Code with Remote-Containers extension

1. **Clone and Open:**

   ```bash
   git clone https://github.com/ewc3labs/excel-power-query-editor.git
   cd excel-power-query-editor
   code .
   ```

2. **Automatic DevContainer Setup:**

   - VS Code will prompt: "Reopen in Container" → **Click Yes**
   - Or: `Ctrl+Shift+P` → "Dev Containers: Reopen in Container"

3. **Everything is Ready:**
   - Node.js 22 with all dependencies pre-installed
   - TypeScript compiler and ESLint configured
   - Test environment with VS Code API mocking
   - Power Query syntax highlighting auto-installed
   - 63 comprehensive tests ready to run

### DevContainer Features

<a id="devcontainer-setup"></a>
**Pre-installed & Configured:**

- Node.js 22 LTS with npm
- TypeScript compiler (`tsc`)
- ESLint with project rules
- Git with full history
- VS Code extensions: Power Query language support
- Complete test fixtures (real Excel files)

**VS Code Tasks Available:**

```bash
Ctrl+Shift+P → "Tasks: Run Task"
```

- **Run Tests** - Execute full 63-test suite
- **Compile TypeScript** - Build extension
- **Lint Code** - ESLint validation
- **Package Extension** - Create VSIX file

### Alternative Setup (Local Development)

**Without DevContainer:**

```bash
# Fork repository on GitHub
git clone https://github.com/YOUR-USERNAME/excel-power-query-editor.git
cd excel-power-query-editor

# Install dependencies
npm install
```

Optional: use Git Bash as your default terminal for POSIX parity with Linux/macOS. This repo is
fully devcontainer-compatible out of the box.

> You can run everything without the container too, but it's the easiest way to mirror the CI pipeline.

---

## 🚀 Quick Reference - Build + Package + Install

| Action                         | Shortcut / Command                                     |
| ------------------------------ | ------------------------------------------------------ |
| Compile extension              | `Ctrl+Shift+B`                                         |
| Package + Install VSIX (local) | `Ctrl+Shift+P`, then `Tasks: Run Task → Install Local` |
| Package VSIX only              | `Ctrl+Shift+P`, then `Tasks: Run Task → Package VSIX`  |
| Watch build (dev background)   | `Ctrl+Shift+W`                                         |
| Start debug (extension host)   | `F5`                                                   |
| Stop debug                     | `Shift+F5`                                             |

## 🧪 Testing - Enterprise-Grade Test Suite

### Test Architecture

<a id="test-suite"></a>
**63 Comprehensive Tests** organized by category:

- **Commands**: 10 tests - Extension command functionality
- **Integration**: 11 tests - End-to-end Excel workflows
- **Utils**: 11 tests - Utility functions and helpers
- **Watch**: 15 tests - File monitoring and auto-sync
- **Backup**: 16 tests - Backup creation and management

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
npm test                    # Run all 63 tests
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

<details>
<summary><strong>⚡ Real-time CI/CD Monitoring</strong> (click to expand)</summary>

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

</details>

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

<details>
<summary><strong>🔢 Smart Version Management</strong> (click to expand)</summary>

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

</details>

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
Quality gates: ESLint, TypeScript, <!--ewc3:tests-->119<!--/ewc3:tests--> tests (<!--ewc3:testsNeedingExcel-->5<!--/ewc3:testsNeedingExcel--> of them need Excel and skip without it, on every platform including Windows CI)
```

then the toolkit is mandatory, because only it can refresh that number - and a stale one fails
`docs:check`. That is the trade: the number cannot silently go wrong, and in exchange the thing that
keeps it right has to be installed. If you would rather not take that trade in a document you are
adding, do not use markers in it.

New values are declared in `.ewc3-docs.json`. Feature docs are `Title_Case_With_Underscores.md`; the
conventions are in [Overview](Overview.md).

## 🚀 CI/CD Pipeline - Professional Automation

<a id="release-automation"></a>
### GitHub Actions Workflow

**Cross-Platform Excellence:**

- **Operating Systems**: Ubuntu, Windows, macOS
- **Node.js Versions**: 22, 24
- **Quality Gates**: ESLint, TypeScript, <!--ewc3:tests-->119<!--/ewc3:tests--> tests
  (<!--ewc3:testsNeedingExcel-->5<!--/ewc3:testsNeedingExcel--> need Excel), documentation checks
- **Artifact Management**: VSIX packaging with 30-day retention

<details>
<summary><strong>🔄 Continuous Integration Pipeline</strong> (click to expand)</summary>

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

**View CI/CD Status:**

- [![CI/CD][ci-cd]](https://github.com/ewc3labs/excel-power-query-editor/actions/workflows/ci.yml)
- [![Tests][tests]](https://github.com/ewc3labs/excel-power-query-editor/actions/workflows/ci.yml)

</details>

<details>
<summary><strong>🎯 Enterprise-Grade Release Automation</strong> (click to expand)</summary>

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

> 🔥 **Wilson's Note:** This is the same automation infrastructure used by enterprise software companies. From a simple commit/push to professional releases with changelogs, versioning, and distribution. No manual BS required.

</details>

### Quality Standards

**All PRs Must Pass:**

1. **ESLint**: Zero linting errors
2. **TypeScript**: Full compilation without errors
3. **Tests**: All 63 tests passing across all platforms
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

# Open in DevContainer (recommended)
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
npm test                # Ensure all 63 tests pass
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

<details>
<summary><strong>🗂️ Complete Directory Structure</strong> (click to expand)</summary>

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
│   └── *.test.ts           # Test files by category (63 tests total)
├── out/                     # Compiled test output
├── .devcontainer/           # Docker container configuration
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

</details>

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

1. Open in DevContainer
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

**v0.5.0 Excellence Achievement:**

- Achieved 63 comprehensive tests with 100% passing rate
- Implemented enterprise-grade CI/CD pipeline
- Created professional development environment
- Delivered all ChatGPT 4o recommendations

### What Makes This Project Special

**Technical Excellence:**

- Zero linting errors across entire codebase
- Full TypeScript compliance with type safety
- Cross-platform validation (Ubuntu, Windows, macOS)
- Professional CI/CD with explicit failure handling

**Developer Experience:**

- World-class DevContainer setup
- Centralized test utilities with VS Code API mocking
- Individual test debugging configurations
- Comprehensive documentation and examples

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

🔥 **Wilson's Note:** This platform is now CI-tested, Docker-ready, GitHub-integrated, and
script-powered. First release or fiftieth — this guide's got you covered.

---

<p align="center">
  <img src="assets/EWC3LabsLogo-blue-128x128.png" width="128" height="128" alt="Georgie the QA Officer"><br>
  <sub><b>Georgie, our QA Officer</b></sub>
</p>

**Excel Power Query Editor** – _Because Power Query development shouldn’t be painful._

<!--ewc3:badgeTests-->
[tests]: https://img.shields.io/badge/tests-119-brightgreen.svg
<!--/ewc3:badgeTests-->

[ci-cd]: https://github.com/ewc3labs/excel-power-query-editor/actions/workflows/ci.yml/badge.svg
[docs-tools]: https://github.com/ewc3labs/ewc3-docs-tools
