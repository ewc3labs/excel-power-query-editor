<#
.SYNOPSIS
    Read and write Power Query M in a workbook Excel already has open.

.DESCRIPTION
    Excel holds an open workbook with an exclusive WRITE lock, so rewriting the DataMashup part
    inside the .xlsx zip fails while the user is looking at the file. (Reading is fine - only
    writing is blocked, which is why extraction never needed the file closed.)

    This helper does not fight the lock. It asks the running Excel to make the change through
    Workbook.Queries(name).Formula, so the file on disk is never touched and the workbook is left
    DIRTY for the user to save.

    HOW IT FINDS THE WORKBOOK, and why not the obvious ways:

      Marshal.GetActiveObject("Excel.Application") returns ONE instance, whichever registered
      first. With two Excel instances the workbook can be in the other one, and we would report
      "not open" while the user is looking at it.

      Marshal.BindToMoniker(path) binds across instances, which sounds ideal. Measured: on a file
      that EXISTS BUT IS CLOSED it OPENS THE FILE. Silently opening somebody's workbook because we
      were checking whether it was open is unacceptable.

    So it enumerates the Running Object Table and binds only to a path ALREADY registered there.
    Excel registers every open workbook under its full path, so presence in the table IS the
    question "is this open, in any instance?", and binding an existing entry cannot start anything.

    MUST run under Windows PowerShell (powershell.exe). Marshal.GetActiveObject and the COM interop
    types used here do not exist in PowerShell 7 / .NET Core.

    Every response is one line of JSON on stdout, including failures. The caller never parses an
    error message or interprets an exit code.

.PARAMETER Action
    status - is this workbook open, and what queries does it hold
    write  - set the formulas supplied on stdin

.PARAMETER Path
    Full path of the workbook as it exists on disk.

.NOTES
    For 'write', the payload arrives on STDIN as JSON: [{ name, formula }, ...].
    Not a command-line argument and not a temp file - M contains quotes, newlines and backslashes,
    and it is the user's source code, which has no business being written to disk by us.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet('status', 'write')][string] $Action,
    [Parameter(Mandatory)][string] $Path
)

$ErrorActionPreference = 'Stop'

function Respond($object) {
    Write-Output ($object | ConvertTo-Json -Depth 6 -Compress)
    exit 0
}

function Fail([string] $code, [string] $message) {
    Respond @{ ok = $false; open = $false; error = $code; message = $message }
}

# --- Excel says "busy" far more often than anyone expects ------------------------------------
# A modal dialog, a recalculation, a cell in edit mode: any of these make COM calls fail with
# RPC_E_CALL_REJECTED (0x80010001) or RPC_E_SERVERCALL_RETRYLATER (0x8001010A). They are transient
# by definition, and retrying is the documented answer. Without this, live sync would appear to
# fail at random - which is exactly how flaky automation gets a reputation it cannot shake.
$script:BusyHResults = @(0x80010001, 0x8001010A, 0x80010005)

function Invoke-WithRetry {
    param([Parameter(Mandatory)][scriptblock] $Action, [int] $Attempts = 5)

    # $__retryAttempt, not $i. PowerShell resolves variables in a scriptblock through the SCOPE
    # CHAIN AT INVOCATION, so a loop variable in here SHADOWS the caller's variable of the same
    # name. With a plain $i, every `Invoke-WithRetry { $book.Queries.Item($i).Name }` read
    # Item(1) - a 29-query workbook reported the same query 29 times, which would then have been
    # diffed as 28 missing queries and fed to Queries.Add on names that already existed.
    for ($__retryAttempt = 1; $__retryAttempt -le $Attempts; $__retryAttempt++) {
        try {
            return & $Action
        } catch {
            $hr = $null
            $ex = $_.Exception
            while ($ex -and $null -eq $hr) {
                if ($ex -is [System.Runtime.InteropServices.COMException]) { $hr = $ex.HResult }
                $ex = $ex.InnerException
            }
            $busy = $hr -ne $null -and ($script:BusyHResults -contains ([uint32]$hr))
            if (-not $busy -or $__retryAttempt -eq $Attempts) { throw }
            Start-Sleep -Milliseconds (150 * $__retryAttempt)   # 150, 300, 450, 600 - Excel dialogs are brief
        }
    }
}

# --- find the workbook, without side effects --------------------------------------------------
try {
    $csPath = Join-Path $PSScriptRoot 'RunningObjects.cs.txt'
    if (-not (Test-Path $csPath)) { Fail 'helper-incomplete' "Missing $csPath" }
    Add-Type -TypeDefinition (Get-Content -Raw $csPath) -ErrorAction Stop
} catch {
    Fail 'interop-unavailable' $_.Exception.Message
}

try {
    $wanted = [System.IO.Path]::GetFullPath($Path)
} catch {
    Fail 'bad-path' $_.Exception.Message
}

# A workbook synced from OneDrive or SharePoint is NOT registered under the local path the user
# sees. Measured on a real machine:
#
#   the user's file : D:\OneDrive - medarms.com\Scripts\PowerQuery\PowerQueryFunctions.xlsx
#   the ROT entry   : https://<tenant>-my.sharepoint.com/personal/<user>/Documents/Scripts/
#                     PowerQuery/PowerQueryFunctions.xlsx
#
# Excel knows the file by its cloud identity; Explorer and VS Code know it by the synced path.
# An exact string match can never join those, so live sync silently declined for every file in a
# synced folder - which for a lot of people is every file they own.
#
# There is no API that maps one to the other, so match on what the two forms genuinely share: the
# trailing path segments. Score each candidate by how many segments agree from the right, require
# the filename itself to agree, and refuse a tie rather than guess at somebody's workbook.
function Get-OneDriveSyncRoots {
    # OneDrive records every sync root as MountPoint (local) + UrlNamespace (cloud). That IS the
    # mapping, maintained by the sync client itself, so there is no need to guess at one.
    $roots = @()
    $key = 'HKCU:\Software\SyncEngines\Providers\OneDrive'
    if (-not (Test-Path $key)) { return $roots }

    foreach ($sub in Get-ChildItem $key -ErrorAction SilentlyContinue) {
        $p = Get-ItemProperty $sub.PSPath -ErrorAction SilentlyContinue
        if ($p.MountPoint -and $p.UrlNamespace) {
            $roots += [pscustomobject]@{
                Mount = $p.MountPoint.TrimEnd('\')
                Url   = $p.UrlNamespace.TrimEnd('/')
            }
        }
    }
    # Longest mount point first, so a nested sync root wins over the parent it sits inside.
    @($roots | Sort-Object { $_.Mount.Length } -Descending)
}

function ConvertTo-CloudUrl {
    # The local path of a synced workbook, expressed as the URL Excel registers it under.
    param([Parameter(Mandatory)][string] $LocalPath)

    foreach ($r in Get-OneDriveSyncRoots) {
        if ($LocalPath.StartsWith($r.Mount, [StringComparison]::OrdinalIgnoreCase)) {
            $rest = $LocalPath.Substring($r.Mount.Length).TrimStart('\')
            return ($r.Url + '/' + ($rest -replace '\\', '/'))
        }
    }
    return $null
}

# Find the workbook by IDENTITY. Never by resemblance.
#
# A synced workbook is registered under its cloud URL rather than the local path the user sees, so
# both forms have to be tried - but both are EXACT comparisons.
#
# THE PREVIOUS VERSION SCORED TRAILING PATH SEGMENTS, AND IT COULD NOT BE MADE SAFE. Measured on a
# real machine: C:\Users\...\Documents\PowerQuery\PowerQueryFunctions.xlsx matched an unrelated
# workbook open from OneDrive at .../Scripts/PowerQuery/PowerQueryFunctions.xlsx with a score of 2,
# because both end "PowerQuery\<same name>". The legitimate match scored 3. No threshold separates
# those two, and being wrong means writing somebody's queries into a different workbook.
#
# A false negative is an inconvenience. A false positive is data corruption. There is no scoring
# here on purpose.
function Resolve-Workbook {
    param([Parameter(Mandatory)][string] $FullPath)

    $exact = [RunningObjects]::Get($FullPath)
    if ($null -ne $exact) { return @{ book = $exact; how = 'exact-path' } }

    $url = ConvertTo-CloudUrl -LocalPath $FullPath
    if ($null -eq $url) { return $null }

    $byUrl = [RunningObjects]::Get($url)
    if ($null -ne $byUrl) { return @{ book = $byUrl; how = 'exact-cloud-url'; registeredAs = $url } }

    return $null
}

$book = $null
$matchedHow = 'exact'
$registeredAs = $null
try {
    $resolved = Invoke-WithRetry { Resolve-Workbook -FullPath $wanted }
    if ($null -ne $resolved) {
        $book = $resolved.book
        $matchedHow = $resolved.how
        $registeredAs = $resolved.registeredAs
    }
} catch {
    Fail 'lookup-failed' $_.Exception.Message
}

if ($null -eq $book) {
    # Usually the ordinary answer: nobody has this file open, so use the on-disk writer.
    #
    # But there is a failure that LOOKS identical and is not. COM's Running Object Table is
    # partitioned by integrity level as well as by session, so an ELEVATED process cannot see the
    # registrations of a normal one. Run VS Code as administrator with Excel started normally (or
    # the reverse) and every workbook is invisible: live sync reports "not open" while the user is
    # staring at the file, and nothing in the message hints why.
    #
    # Measured here: an elevated shell saw zero ROT entries while Excel was plainly running.
    #
    # So when nothing matched, check whether Excel is running at all. If it is, say so - the caller
    # can tell the user something useful instead of "not open".
    $processes = @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count
    $elevated = ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
            [Security.Principal.WindowsBuiltInRole]::Administrator)

    Respond @{
        ok             = $true
        open           = $false
        excelProcesses = $processes
        elevated       = [bool]$elevated
    }
}

# A ROT entry under a workbook path should be a Workbook, but a hostile or unexpected registration
# should not turn into a confusing COM error three lines later.
try {
    $null = Invoke-WithRetry { $book.Queries }
} catch {
    Fail 'not-a-workbook' "The object registered for this path does not expose Queries: $($_.Exception.Message)"
}

# --- status -----------------------------------------------------------------------------------
if ($Action -eq 'status') {
    try {
        $names = @()
        $count = Invoke-WithRetry { $book.Queries.Count }
        for ($i = 1; $i -le $count; $i++) {
            $names += (Invoke-WithRetry { $book.Queries.Item($i).Name })
        }
        Respond @{
            ok           = $true
            open         = $true
            workbook     = (Invoke-WithRetry { $book.Name })
            fullName     = (Invoke-WithRetry { $book.FullName })
            saved        = [bool](Invoke-WithRetry { $book.Saved })
            # AutoSave changes what a live write MEANS: with it on, Excel commits the change within
            # seconds and closing without saving no longer undoes it. Measured, not assumed.
            # Unreadable on older builds, so absent rather than false when we cannot tell.
            autoSaveOn   = $(try { [bool](Invoke-WithRetry { $book.AutoSaveOn }) } catch { $null })
            queries      = $names
            matchedHow   = $matchedHow
            registeredAs = $registeredAs
        }
    } catch {
        Fail 'status-failed' $_.Exception.Message
    }
}

# --- write ------------------------------------------------------------------------------------
try {
    $raw = [Console]::In.ReadToEnd()
} catch {
    Fail 'payload-unreadable' $_.Exception.Message
}

if ([string]::IsNullOrWhiteSpace($raw)) { Fail 'payload-missing' 'No payload on stdin.' }

try {
    # NOT @(... | ConvertFrom-Json). ConvertFrom-Json emits a JSON array as ONE pipeline object, so
    # @() wraps it into a single-element array holding the array; the loop then runs once with the
    # whole array bound, and $item.name becomes a COLLECTION of every name. That produced a query
    # literally called "StudentResults BrandNewQuery". Convert first, normalize shape second.
    $payload = ConvertFrom-Json $raw
    if ($null -eq $payload) { $payload = @() }
    elseif ($payload -isnot [object[]]) { $payload = @($payload) }
} catch {
    Fail 'payload-unparseable' $_.Exception.Message
}

$existing = @{}
try {
    $count = Invoke-WithRetry { $book.Queries.Count }
    for ($i = 1; $i -le $count; $i++) {
        $q = Invoke-WithRetry { $book.Queries.Item($i) }
        $existing[$q.Name] = $q
    }
} catch {
    Fail 'queries-unavailable' $_.Exception.Message
}

$updated = @(); $added = @(); $skipped = @(); $failures = @()

foreach ($item in $payload) {
    if (-not $item.name) { continue }
    if ($item.name -isnot [string]) {
        $failures += @{ name = "$($item.name)"; message = 'Query name was not a string; payload shape is wrong.' }
        continue
    }
    if ($null -eq $item.formula -or $item.formula -isnot [string]) {
        $failures += @{ name = $item.name; message = 'Formula was missing or not a string.' }
        continue
    }

    try {
        if ($existing.ContainsKey($item.name)) {
            $current = Invoke-WithRetry { $existing[$item.name].Formula }
            if ($current -cne $item.formula) {
                Invoke-WithRetry { $existing[$item.name].Formula = $item.formula }
                $updated += $item.name
            } else {
                # Any set marks the workbook dirty. Doing that for an identical value would put an
                # unsaved-changes dot on somebody's file for no reason.
                $skipped += $item.name
            }
        } else {
            Invoke-WithRetry { [void]$book.Queries.Add($item.name, $item.formula) }
            $added += $item.name
        }
    } catch {
        $failures += @{ name = $item.name; message = $_.Exception.Message }
    }
}

Respond @{
    ok         = ($failures.Count -eq 0)
    open       = $true
    matchedHow = $matchedHow
    workbook   = (Invoke-WithRetry { $book.Name })
    updated   = $updated
    added     = $added
    unchanged = $skipped
    failures  = $failures
    dirty     = (-not [bool](Invoke-WithRetry { $book.Saved }))
    autoSaveOn = $(try { [bool](Invoke-WithRetry { $book.AutoSaveOn }) } catch { $null })
}
