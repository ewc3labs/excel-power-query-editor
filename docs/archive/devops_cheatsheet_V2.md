# 🎯 VS Code Extension DevOps Cheat Sheet (EWC3 Labs Style)

> This cheat sheet is for **any developer** working on an EWC3 Labs project using VS Code. It’s your one-stop reference for building, testing, committing, packaging, and shipping VS Code extensions like a badass.

---

<details>
<summary><strong>💡 Pro Developer Workflow Tips</strong> (click to expand)</summary>

**Development Environment:**

- DevContainers optional, but fully supported if Docker + Remote Containers is installed
- Default terminal is Git Bash for sanity + POSIX-like parity
- GitHub CLI (`gh`) installed and authenticated for real-time CI/CD monitoring
- ✅ Make sure you have Node.js 22 or 24 installed (the CI pipeline tests against both)

**Release Workflow:**

- Push to `release/v0.5.0` branch triggers automatic pre-release builds
- Push to `main` creates stable releases (when marketplace is configured)
- Manual tags `v*` trigger official marketplace releases
- Every release includes auto-generated changelog from git commit messages

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

**Want to improve this cheat sheet?** PRs are always welcome — we keep this living document current and useful.

🔥 **Wilson's Note:** This is my first extension, first public repo, first devcontainer (first time even using Docker), first automated test suite, and first time using Git Bash — so I'm drinking from the firehose here and often learning as I go. That said, I **do** know how this stuff should work, and EWC3 Labs is about building it right. Our goal is an enterprise-grade DX platform for VS Code extension development. We went from manual builds to automated releases with smart versioning, multi-channel distribution, and real-time monitoring. It's modular, CI-tested, scriptable, and optimized for contributors. If you're reading this — welcome to the automation party. **From a simple commit/push to professional releases. Shit works when you work it.**

---

## 🧊 DevContainer setup 🐋
- ✅ Install [Docker](https://www.docker.com/)
- ✅ Install the VS Code extension: `ms-vscode-remote.remote-containers`
- ✅ Clone the repo and open it in VS Code — it will prompt to reopen in the container.

Optional: use Git Bash as your default terminal for POSIX parity with Linux/macOS. This repo is fully devcontainer-compatible out of the box.

> You can run everything without the container too, but it's the easiest way to mirror the CI pipeline.

## 🚀 Build + Package + Install

| Action                         | Shortcut / Command                                     |
| ------------------------------ | ------------------------------------------------------ |
| Compile extension              | `Ctrl+Shift+B`                                         |
| Package + Install VSIX (local) | `Ctrl+Shift+P`, then `Tasks: Run Task → Install Local` |
| Package VSIX only              | `Ctrl+Shift+P`, then `Tasks: Run Task → Package VSIX`  |
| Watch build (dev background)   | `Ctrl+Shift+W`                                         |
| Start debug (extension host)   | `F5`                                                   |
| Stop debug                     | `Shift+F5`                                             |

## 🧪 Testing

| Action        | Shortcut / Command                                      |
| ------------- | ------------------------------------------------------- |
| Run Tests     | `Ctrl+Shift+T` or `Tasks: Run Task → Run Tests`         |
| Compile Tests | `npm run compile-tests`                                 |
| Watch Tests   | `npm run watch-tests`                                   |
| Test Entry    | `test/runTest.ts` calls into compiled test suite        |
| Test Utils    | `test/testUtils.ts` contains shared scaffolding/helpers |

> 🧠 Tests run with `vscode-test`, launching VS Code in a headless test harness. You’ll see a test instance of VS Code launch and close automatically during test runs.

## 🧹 GitOps

| Action            | Shortcut / Command             |
| ----------------- | ------------------------------ |
| Stage all changes | `Ctrl+Shift+G`, `Ctrl+Shift+A` |
| Commit            | `Ctrl+Shift+G`, `Ctrl+Shift+C` |
| Push              | `Ctrl+Shift+G`, `Ctrl+Shift+P` |
| Git Bash terminal | `` Ctrl+Shift+` ``             |

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

## 🌱 Branching Conventions

| Purpose          | Branch Prefix | Example               |
| ---------------- | ------------- | --------------------- |
| Releases         | `release/`    | `release/v0.5.0`      |
| Work-in-progress | `wip/`        | `wip/feature-xyz`     |
| Hotfixes         | `hotfix/`     | `hotfix/package-lock` |

> 📛 These branch names are picked up by our GitHub Actions CI/CD pipelines.

## 🧾 npm Scripts

| Script                 | Description                                     |
| ---------------------- | ----------------------------------------------- |
| `npm run lint`         | Run ESLint on `src/`                            |
| `npm run compile`      | Type check, lint, and build with `esbuild.js`   |
| `npm run package`      | Full production build                           |
| `npm run dev-install`  | Build, package, force install VSIX              |
| `npm run test`         | Run test suite via `vscode-test`                |
| `npm run watch`        | Watch build and test                            |
| `npm run check-types`  | TypeScript compile check (no emit)              |
| `npm run bump-version` | Smart semantic version bumping from git commits |

<details>
<summary><strong>🔢 Smart Version Management</strong> (click to expand)</summary>

**Automatic Version Bumping:**
```bash
# Analyze commits and bump version automatically
npm run bump-version

# The script analyzes your git history for:
# - feat: → minor version bump (0.5.0 → 0.6.0)
# - fix: → patch version bump (0.5.0 → 0.5.1) 
# - BREAKING: → major version bump (0.5.0 → 1.0.0)
```

**Manual Version Control:**
```bash
# Bump specific version types
npm version patch   # 0.5.0 → 0.5.1
npm version minor   # 0.5.0 → 0.6.0  
npm version major   # 0.5.0 → 1.0.0

# Pre-release versions
npm version prerelease  # 0.5.0 → 0.5.1-0
npm version prepatch    # 0.5.0 → 0.5.1-0
npm version preminor    # 0.5.0 → 0.6.0-0
```

> 🧠 **Smart Tip:** The release pipeline automatically handles version bumping, but you can use `npm run bump-version` locally to preview what version would be generated.

</details>

## 🔁 README Management

| Task                          | Script                                                              |
| ----------------------------- | ------------------------------------------------------------------- |
| Set README for GitHub         | `node scripts/set-readme-gh.js`                                     |
| Set README for VS Marketplace | `node scripts/set-readme-vsce.js`                                   |
| Automated pre/post-publish    | Hooked via `prepublishOnly` and `postpublish` npm lifecycle scripts |

> `vsce package` **must** see a clean Marketplace README. Run `set-readme-vsce.js` right before packaging.

## 📦 CI/CD (GitHub Actions)

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

> 💥 Failing lint/typecheck = blocked CI. No bullshit allowed.

**Documentation Changes:**
- Pushes that only modify `docs/**` or `*.md` files skip the release pipeline
- CI still runs to validate documentation quality  
- No version bumps or releases triggered for docs-only changes

</details>

## 🚀 Release Automation Pipeline

<details>
<summary><strong>🎯 Enterprise-Grade Release Automation</strong> (click to expand)</summary>

> Configured in `.github/workflows/release.yml`

### **What Happens on Every Push:**
1. **🔍 Auto-detects release type** (dev/prerelease/stable)
2. **🔢 Smart version bumping** in `package.json` using semantic versioning
3. **⚡ Fast optimized build** (lint + type check, skips heavy integration tests)
4. **📦 Professional VSIX generation** with proper naming conventions
5. **🎉 Auto-creates GitHub release** with changelog, assets, and metadata

### **Release Channels:**
| Branch/Trigger | Release Type | Version Format | Auto-Publish |
|----------------|--------------|----------------|--------------|
| `release/**`   | Pre-release  | `v0.5.0-rc.X`  | GitHub only  |
| `main`         | Stable       | `v0.5.0`       | GitHub + Marketplace* |
| Manual tag `v*`| Official     | `v0.5.0`       | GitHub + Marketplace* |
| Workflow dispatch | Emergency  | Custom         | Configurable |

*Marketplace publishing requires `VSCE_PAT` secret

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

> 🔥 **Wilson's Note:** This is the same automation infrastructure used by enterprise software companies. From a simple commit/push to professional releases with changelogs, versioning, and distribution. No manual bullshit required.

</details>

## 📁 Folder Structure Highlights

<details>
<summary><strong>🗂️ Project Structure Overview</strong> (click to expand)</summary>

```
.
├── docs/                    # All markdown docs (README variants, changelogs, etc.)
├── scripts/                 # Automation scripts
│   ├── set-readme-gh.js     # GitHub README switcher
│   ├── set-readme-vsce.js   # VS Marketplace README switcher  
│   └── bump-version.js      # Smart semantic version bumping
├── src/                     # Extension source code (extension.ts, configHelper.ts, etc.)
├── test/                    # Mocha-style unit tests + testUtils scaffolding
├── out/                     # Compiled test output
├── .devcontainer/           # Dockerfile + config for remote containerized development
├── .github/workflows/       # CI/CD automation
│   ├── ci.yml              # Continuous integration pipeline
│   └── release.yml         # Enterprise release automation
├── .vscode/                 # Launch tasks, keybindings, extensions.json
└── temp-testing/           # Test files and debugging artifacts
```

**Key Automation Files:**
- **`.github/workflows/release.yml`** - Full release pipeline with smart versioning
- **`scripts/bump-version.js`** - Semantic version analysis from git commits
- **`.github/workflows/ci.yml`** - Multi-platform CI testing matrix
- **`.vscode/tasks.json`** - VS Code build/test/package tasks

## 🔧 Misc Configs

| File                      | Purpose                                                          |
| ------------------------- | ---------------------------------------------------------------- |
| `.eslintrc.js`            | Lint rules (uses ESLint with project-specific overrides)         |
| `tsconfig.json`           | TypeScript project config                                        |
| `.gitignore`              | Ignores `_PowerQuery.m`, `*.backup.*`, `debug_sync/`, etc.       |
| `package.json`            | npm scripts, VS Code metadata, lifecycle hooks                   |
| `.vscode/extensions.json` | Recommended extensions (auto-suggests key tools when repo opens) |

</details>

---

🔥 **Wilson’s Note:** This platform is now CI-tested, Docker-ready, GitHub-integrated, and script-powered. First release or fiftieth — this cheatsheet’s got you. 