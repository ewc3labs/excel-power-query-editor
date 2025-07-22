import os
import re

# Mapping of old keys to new keys
RENAME_MAP = {
    'autoBackupBeforeSync': 'backup.autoBackupBeforeSync',
    'backupLocation': 'backup.location',
    'customBackupPath': 'backup.customPath',
    'autoCleanupBackups': 'backup.autoCleanup',
    'backup.maxFiles': 'backup.maxFiles',
    'logLevel': 'log.level',
    'showStatusBarInfo': 'log.showStatusBarInfo',
    'verboseMode': 'log.verboseMode',
    'debugMode': 'log.debugMode',
    'syncTimeout': 'sync.timeout',
    'syncDeleteAlwaysConfirm': 'sync.deleteAlwaysConfirm',
    'watchAlways': 'watch.always',
    'watchAlwaysMaxFiles': 'watch.maxFiles',
    'watchOffOnDelete': 'watch.offOnDelete',
    'sync.openExcelAfterWrite': 'sync.openExcelAfterWrite',
    'sync.debounceMs': 'sync.debounceMs',
    'watch.checkExcelWriteable': 'watch.checkExcelWriteable',
    'symbols.installLevel': 'symbols.installLevel',
    'symbols.autoInstall': 'symbols.autoInstall'
}

# File extensions to scan
EXTENSIONS = ['.ts', '.json', '.md']

# Folders to include (relative to project root)
INCLUDE_FOLDERS = ['src', 'test', 'docs', '.', 'scripts']

# Folders to exclude
EXCLUDE_FOLDERS = {'node_modules', 'dist', '.git', '.vscode', '__pycache__'}

# Set to True for dry run (no changes), False to actually replace
testmode = True

def should_scan(subdir):
    parts = os.path.relpath(subdir, os.getcwd()).split(os.sep)
    # Only scan if the first part is in INCLUDE_FOLDERS and not in EXCLUDE_FOLDERS
    return parts[0] in INCLUDE_FOLDERS and not any(part in EXCLUDE_FOLDERS for part in parts)

def scan_and_replace(root_dir):
    for subdir, _, files in os.walk(root_dir):
        if not should_scan(subdir):
            continue
        for file in files:
            if any(file.endswith(ext) for ext in EXTENSIONS):
                filepath = os.path.join(subdir, file)
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
                original_content = content
                for old, new in RENAME_MAP.items():
                    # Match in quotes, dot notation, or as a property
                    pattern = re.compile(rf'([\'\"]){old}([\'\"])|{old}')
                    matches = list(pattern.finditer(content))
                    for match in matches:
                        print(f"{'Would change' if testmode else 'Changing'}: {old} → {new} in {filepath}")
                    if not testmode:
                        content = pattern.sub(lambda m: m.group(0).replace(old, new), content)
                if not testmode and content != original_content:
                    with open(filepath, 'w', encoding='utf-8') as f:
                        f.write(content)

if __name__ == '__main__':
    scan_and_replace(os.getcwd())