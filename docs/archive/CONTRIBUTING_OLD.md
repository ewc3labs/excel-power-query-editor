<!-- HEADER_TABLE -->
<table align="center">
<tr>
  <td width="112" align="center" valign="middle">
    <img src="assets/excel-power-query-editor-logo-128x128.png" width="128" height="128"><br>
    <strong>E · P · Q · E</strong>
  </td>

  <td align="center" valign="middle">
    <h1 align="center">Excel Power Query Editor</h1>
    <p align="left">
      <b>Edit Power Query M code directly from Excel files in VS Code. No Excel needed. No BS. It Just Works™.</b><br>
      <sub>
        Built by <strong>EWC3 Labs</strong> — where we rage-build the tools everyone needs, but nobody <del>cares to build</del>
        <em>is deranged enough to spend days perfecting until it actually works right.</em>
      </sub>
    </p>
  </td>

  <td width="112" align="center" valign="middle">
    <img src="assets/EWC3LabsLogo-blue-128x128.png" width="128" height="128"><br>
    <strong><em>QA Officer</em></strong>
  </td>
</tr>
</table>
<!-- /HEADER_TABLE -->

---

## Contributing Guide

> **Welcome to the most professional VS Code extension development environment you'll ever see!**

---

Thanks for your interest in contributing! This project has achieved **enterprise-grade quality** with 63 comprehensive tests, cross-platform CI/CD, and a world-class development experience.

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

## 🧪 Testing - Enterprise-Grade Test Suite

### Test Architecture

**63 Comprehensive Tests** organized by category:

- **Commands**: 10 tests - Extension command functionality
- **Integration**: 11 tests - End-to-end Excel workflows
- **Utils**: 11 tests - Utility functions and helpers
- **Watch**: 15 tests - File monitoring and auto-sync
- **Backup**: 16 tests - Backup creation and management

### Running Tests

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

## 🚀 CI/CD Pipeline - Professional Automation

### GitHub Actions Workflow

**Cross-Platform Excellence:**

- **Operating Systems**: Ubuntu, Windows, macOS
- **Node.js Versions**: 18.x, 20.x
- **Quality Gates**: ESLint, TypeScript, 63-test validation
- **Artifact Management**: VSIX packaging with 30-day retention

**Workflow Triggers:**

- Push to `main` branch
- Pull requests to `main`
- Manual workflow dispatch

**View CI/CD Status:**

- [![CI/CD](https://github.com/ewc3labs/excel-power-query-editor/actions/workflows/ci.yml/badge.svg)](https://github.com/ewc3labs/excel-power-query-editor/actions/workflows/ci.yml)
- [![Tests](https://img.shields.io/badge/tests-63%20passing-brightgreen.svg)](https://github.com/ewc3labs/excel-power-query-editor/actions/workflows/ci.yml)

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

### Code Organization

**File Structure:**

```
src/
├── extension.ts          # Main extension entry point
├── commands/            # Command implementations
├── utils/              # Utility functions
├── types/              # TypeScript type definitions
└── config/             # Configuration handling

test/
├── testUtils.ts        # Centralized test utilities
├── fixtures/           # Real Excel files for testing
└── *.test.ts          # Test files by category
```

### Commit Message Format

**Use Conventional Commits:**

```bash
feat: add intelligent debouncing for CoPilot integration
fix: resolve Excel file locking detection on Windows
docs: update configuration examples for team workflows
test: add comprehensive backup management test suite
ci: enhance cross-platform testing matrix
```

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

## 🔗 Related Documentation

- **📖 [User Guide](USER_GUIDE.md)** - Complete feature documentation and workflows
- **⚙️ [Configuration Reference](CONFIGURATION.md)** - All settings with examples and use cases
- **📝 [Changelog](../CHANGELOG.md)** - Version history and feature updates
- **🧪 [Test Documentation](../test/testcases.md)** - Comprehensive test coverage details

---

**Thank you for contributing to Excel Power Query Editor!**  
**Together, we're building the gold standard for Power Query development in VS Code.**
