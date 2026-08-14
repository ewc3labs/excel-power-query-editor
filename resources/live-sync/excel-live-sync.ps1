<#
.SYNOPSIS
    Read and write Power Query M in a workbook Excel already has open.

.DESCRIPTION
    Excel holds an open workbook with an exclusive WRITE lock, so the extension's normal path -
    rewrite the DataMashup part inside the .xlsx zip - fails while the user is looking at the file.
    Reading is fine; only writing is blocked.

    This helper does not fight the lock. It goes through Excel's object model instead:
    Workbook.Queries(name).Formula is read/write in Excel 2016+, so the running application does the
    write and the file on disk is never touched. The workbook is left DIRTY, exactly as if the user
    had typed the change, and they save when they choose.

    MUST run under Windows PowerShell (powershell.exe). Marshal.GetActiveObject does not exist in
    PowerShell 7 / .NET Core, and attaching to a running Excel is the entire point.

    Every response is a single line of JSON on stdout, including failures. The caller never has to
    parse an error message or interpret an exit code.

.PARAMETER Action
    status - is Excel running, is this workbook open, what queries does it hold
    write  - set the formulas listed in the payload file

.PARAMETER Path
    Full path of the workbook, as the user knows it on disk.

.PARAMETER PayloadFile
    For 'write': a UTF-8 JSON file of [{ name, formula }, ...].

    A FILE, not an argument, deliberately. M code contains quotes, newlines and backslashes, and
    putting it on a command line is a quoting bug waiting to corrupt somebody's query.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet('status', 'write')][string] $Action,
    [Parameter(Mandatory)][string] $Path,
    [string] $PayloadFile
)

$ErrorActionPreference = 'Stop'

function Respond($object) {
    # -Compress so the caller reads exactly one line; -Depth because queries nest.
    Write-Output ($object | ConvertTo-Json -Depth 6 -Compress)
    exit 0
}

function Fail([string] $code, [string] $message) {
    Respond @{ ok = $false; error = $code; message = $message }
}

# --- attach to a RUNNING Excel; never start one --------------------------------------------
# Starting Excel to write a file the user is not looking at would be a surprise, and the on-disk
# path already handles that case properly.
$excel = $null
try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    Fail 'excel-not-running' 'No running Excel instance is available to this session.'
}

# --- find the workbook -----------------------------------------------------------------------
$target = $null
try {
    $wanted = [System.IO.Path]::GetFullPath($Path)
    foreach ($wb in $excel.Workbooks) {
        if ([string]::Equals($wb.FullName, $wanted, 'OrdinalIgnoreCase')) { $target = $wb; break }
    }
} catch {
    Fail 'enumerate-failed' $_.Exception.Message
}

if ($null -eq $target) {
    # Not an error. It is the normal case, and it means "use the on-disk writer".
    Respond @{ ok = $true; open = $false; excelVersion = $excel.Version }
}

# --- status ----------------------------------------------------------------------------------
if ($Action -eq 'status') {
    $names = @()
    try {
        for ($i = 1; $i -le $target.Queries.Count; $i++) { $names += $target.Queries.Item($i).Name }
    } catch {
        Fail 'queries-unavailable' $_.Exception.Message
    }
    Respond @{
        ok           = $true
        open         = $true
        excelVersion = $excel.Version
        workbook     = $target.Name
        saved        = [bool]$target.Saved
        queries      = $names
    }
}

# --- write -----------------------------------------------------------------------------------
if (-not $PayloadFile -or -not (Test-Path $PayloadFile)) {
    Fail 'payload-missing' "Payload file not found: $PayloadFile"
}

try {
    # NOT @(... | ConvertFrom-Json). ConvertFrom-Json emits a JSON array as ONE pipeline object,
    # so @() wraps it into a single-element array holding the array - and the foreach below then
    # runs once with $item bound to the whole thing. $item.name silently becomes a COLLECTION of
    # every name, which PowerShell happily stringifies, and the first live write created a query
    # called "StudentResults BrandNewQuery". Convert first, normalise shape second.
    $payload = ConvertFrom-Json (Get-Content -Raw -Encoding UTF8 $PayloadFile)
    if ($null -eq $payload) { $payload = @() }
    elseif ($payload -isnot [object[]]) { $payload = @($payload) }
} catch {
    Fail 'payload-unreadable' $_.Exception.Message
}

# Index the workbook's queries once, by name.
$existing = @{}
try {
    for ($i = 1; $i -le $target.Queries.Count; $i++) {
        $q = $target.Queries.Item($i)
        $existing[$q.Name] = $q
    }
} catch {
    Fail 'queries-unavailable' $_.Exception.Message
}

$updated = @(); $added = @(); $failures = @()

foreach ($item in $payload) {
    if (-not $item.name) { continue }
    if ($item.name -isnot [string]) {
        $failures += @{ name = "$($item.name)"; message = 'Query name was not a string; payload shape is wrong.' }
        continue
    }
    try {
        if ($existing.ContainsKey($item.name)) {
            # Skip a write that would change nothing. Excel marks the workbook dirty on any set,
            # and making somebody's file dirty for no reason is rude.
            if ($existing[$item.name].Formula -cne $item.formula) {
                $existing[$item.name].Formula = $item.formula
                $updated += $item.name
            }
        } else {
            [void]$target.Queries.Add($item.name, $item.formula)
            $added += $item.name
        }
    } catch {
        # One bad query must not strand the others.
        $failures += @{ name = $item.name; message = $_.Exception.Message }
    }
}

Respond @{
    ok       = ($failures.Count -eq 0)
    open     = $true
    workbook = $target.Name
    updated  = $updated
    added    = $added
    failures = $failures
    dirty    = (-not [bool]$target.Saved)
}
