# 🎯 VS Code Extension DevOps Cheat Sheet (EWC3 Labs Style)

> This cheat sheet is for **any developer** working on an EWC3 Labs project using VS Code. It’s your one-stop reference for building, testing, committing, packaging, and shipping extensions like a badass.

## 🧰 Dev Environment Setup

To match the full EWC3 Labs development environment:

- ✅ Install [Docker](https://www.docker.com/) (for devcontainers)
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

> 🧠 Tests run with `vscode-test`, launching VS Code in a headless test harness. You’ll see VS Code flash briefly on execution.

## 🧹 GitOps

| Action            | Shortcut / Command             |
| ----------------- | ------------------------------ |
| Stage all changes | `Ctrl+Shift+G`, `Ctrl+Shift+A` |
| Commit            | `Ctrl+Shift+G`, `Ctrl+Shift+C` |
| Push              | `Ctrl+Shift+G`, `Ctrl+Shift+P` |
| Git Bash terminal | \`Ctrl+Shift+\`\`              |

## 🌱 Branching Conventions

| Purpose          | Branch Prefix | Example               |
| ---------------- | ------------- | --------------------- |
| Releases         | `release/`    | `release/v0.5.0`      |
| Work-in-progress | `wip/`        | `wip/feature-xyz`     |
| Hotfixes         | `hotfix/`     | `hotfix/package-lock` |

> 📛 These branch names are picked up by our GitHub Actions CI/CD pipelines.

## 🧾 npm Scripts

| Script                | Description                                   |
| --------------------- | --------------------------------------------- |
| `npm run lint`        | Run ESLint on `src/`                          |
| `npm run compile`     | Type check, lint, and build with `esbuild.js` |
| `npm run package`     | Full production build                         |
| `npm run dev-install` | Build, package, force install VSIX            |
| `npm run test`        | Run test suite via `vscode-test`              |
| `npm run watch`       | Watch build and test                          |
| `npm run check-types` | TypeScript compile check (no emit)            |

## 🔁 README Management

| Task                          | Script                                                              |
| ----------------------------- | ------------------------------------------------------------------- |
| Set README for GitHub         | `node scripts/set-readme-gh.js`                                     |
| Set README for VS Marketplace | `node scripts/set-readme-vsce.js`                                   |
| Automated pre/post-publish    | Hooked via `prepublishOnly` and `postpublish` npm lifecycle scripts |

> `vsce package` **must** see a clean Marketplace README. Run `set-readme-vsce.js` right before packaging.

## 📦 CI/CD (GitHub Actions)

> Configured in `.github/workflows/ci.yml`

**Triggers:**

- On push or pull to: `main`, `release/**`, `wip/**`, `hotfix/**`

**Matrix Builds:**

- OS: `ubuntu-latest`, `windows-latest`, `macos-latest`
- Node.js: `18`, `20`, `22`, `24`

**Steps:**

- Checkout → Install → Lint → TypeCheck → Test → Build → Package → Upload VSIX

> 💥 Failing lint/typecheck = blocked CI. No bullshit allowed.

## 📁 Folder Structure Highlights

```
.
├── docs/                    # All markdown docs (README variants, changelogs, etc.)
├── scripts/                 # Automation: prepublish, postpublish, readme switchers
├── src/                     # Extension source code (extension.ts, configHelper.ts, etc.)
├── test/                    # Mocha-style unit tests + testUtils scaffolding
├── out/                     # Compiled test output
├── .devcontainer/           # Dockerfile + config for remote containerized development
├── .github/workflows/       # CI/CD config
├── .vscode/                 # Launch tasks, keybindings, extensions.json
```

## 🔧 Misc Configs

| File                      | Purpose                                                     |
| ------------------------- | ----------------------------------------------------------- |
| `.eslintrc.js`            | Lint rules (uses ESLint with project-specific overrides)    |
| `tsconfig.json`           | TypeScript project config                                   |
| `.gitignore`              | Ignores `_PowerQuery.m`, `*.backup.*`, `debug_sync/`, etc.  |
| `package.json`            | npm scripts, VS Code metadata, lifecycle hooks              |
| `.vscode/extensions.json` | Recommended extensions (auto-suggests them when repo opens) |

## 🧠 Bonus Tips

- DevContainers optional, but fully supported if Docker + Remote Containers is installed.
- Default terminal is Git Bash for sanity + POSIX-like parity.
- CI/CD will auto-build your branch on push to `release/**` and others.
- The Marketplace README build status badge is tied to GitHub Actions CI.

---

This is my first extension, first public repo, first devcontainer (and first time even using Docker), first automated test suite, and first time using Git Bash — so I'm drinking from the firehose here and often learning as I go. That said, I *do* know how this stuff should work, and EWC3 Labs is about building it right.

PRs improving this cheat sheet are always welcome.

🔥 **Wilson’s Note:** This is now a full DX platform for VS Code extension development. It's modular, CI-tested, scriptable, and optimized for contributors. If you're reading this — welcome to the code party. Shit works when you work it.

