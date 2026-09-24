# Read-only look inside an .xlsx without Excel/LibreOffice, even while it is open (shared read).
# Shows one sheet's rows (B C D F G H I, dates as HH:mm:ss) and the column-H notes, or profiles the
# whole package for secrets without printing them.
#
#   powershell -NoProfile -File tools\inspect_workbook.ps1 -Path scrt_test7.xlsx -From 2 -To 30
#   powershell -NoProfile -File tools\inspect_workbook.ps1 -Path scrt_test7.xlsx -Sheet Settings
#   powershell -NoProfile -File tools\inspect_workbook.ps1 -Path newbook.xlsx -Secrets
#
# -Sheet takes a name or a 1-based number (default 2 = PromptResp). -NoNotes hides the notes.
# -Secrets prints counts only (type-5/7 hashes, "secret"/"password", IPv4-like strings) per package part.
#   Known harmless hits: docProps/app.xml ipv4-like=1 is the app version (e.g. 25.8.3.2); the HostList
#   header "Password" counts once in sharedStrings. Anything in xl/comments*.xml is real captured output.
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$Sheet = '2',
    [int]$From = 1,
    [int]$To = 200,
    [switch]$NoNotes,
    [switch]$Secrets
)
Add-Type -AssemblyName System.IO.Compression
$full = (Resolve-Path $Path).Path
"file:  $full"
"saved: " + (Get-Item $full).LastWriteTime
$fs = New-Object IO.FileStream($full, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
$z = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Read)
function RT($n) { $e = $z.GetEntry($n); if (-not $e) { return $null }; $r = New-Object IO.StreamReader($e.Open()); $t = $r.ReadToEnd(); $r.Close(); $t }
try {
    if ($Secrets) {
        $any = $false
        foreach ($e in $z.Entries) {
            $r = New-Object IO.StreamReader($e.Open()); $t = $r.ReadToEnd(); $r.Close()
            $h = ([regex]::Matches($t, '\$1\$|\$9\$|\$8\$|\$6\$|(?i)\bsecret [57] ')).Count
            $w = ([regex]::Matches($t, '(?i)secret|password')).Count
            $ip = ([regex]::Matches($t, '\b\d{1,3}(\.\d{1,3}){3}\b')).Count
            if ($h + $w + $ip -gt 0) { $any = $true; "{0,-40} hashes={1} secret/password={2} ipv4-like={3}" -f $e.FullName, $h, $w, $ip }
        }
        if (-not $any) { "no hash / secret / IPv4-like strings in any part" }
        return
    }
    [xml]$wb = RT 'xl/workbook.xml'; [xml]$rels = RT 'xl/_rels/workbook.xml.rels'
    $strs = @(); $ss = RT 'xl/sharedStrings.xml'; if ($ss) { $strs = @(([xml]$ss).sst.si | ForEach-Object { $_.InnerText }) }
    $sheets = @($wb.workbook.sheets.sheet)
    "sheets: " + (($sheets | ForEach-Object { $_.name }) -join ', ')
    $sh = if ($Sheet -match '^\d+$') { $sheets[[int]$Sheet - 1] } else { $sheets | Where-Object { $_.name -eq $Sheet } }
    if (-not $sh) { "no such sheet: $Sheet"; return }
    $rid = $sh.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $target = ($rels.Relationships.Relationship | Where-Object { $_.Id -eq $rid }).Target -replace '^/?xl/', ''
    "===== sheet '$($sh.name)' rows $From-$To"
    [xml]$sx = RT ('xl/' + $target)
    foreach ($row in $sx.worksheet.sheetData.row) {
        if ([int]$row.r -lt $From -or [int]$row.r -gt $To) { continue }
        $cells = foreach ($c in $row.c) {
            $v = switch ($c.t) { 's' { $strs[[int]$c.v] } 'inlineStr' { $c.is.InnerText } default { $c.v } }
            if ("$v" -match '^4\d{4}\.\d{4,}$') { $v = [DateTime]::FromOADate([double]$v).ToString('HH:mm:ss') }
            if ("$v" -ne '') { "$($c.r)=[" + ("$v" -replace "`r?`n", ' / ') + "]" }
        }
        if ($cells) { "  " + ($cells -join '  ') }
    }
    if ($NoNotes) { return }
    # the notes part for this sheet
    $srels = RT ('xl/worksheets/_rels/' + (Split-Path $target -Leaf) + '.rels')
    if (-not $srels) { "(no notes on this sheet)"; return }
    $cref = ([xml]$srels).Relationships.Relationship | Where-Object { $_.Type -match '/comments$' }
    if (-not $cref) { "(no notes on this sheet)"; return }
    [xml]$cx = RT ('xl/' + ($cref.Target -replace '^\.\./', '' -replace '^/?xl/', ''))
    foreach ($cm in $cx.comments.commentList.comment) {
        $r = [int]($cm.ref -replace '\D', '')
        if ($r -lt $From -or $r -gt $To) { continue }
        $t = ($cm.text.InnerText -replace "`r`n", "`n") -replace "`r", "`n"
        "----- note $($cm.ref) ($($t.Length) chars)"
        ($t -split "`n") | ForEach-Object { "|" + $_ + "|" }
    }
} finally { $z.Dispose(); $fs.Dispose() }
