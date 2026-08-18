<#
.SYNOPSIS
    Converts a HardeningKitty Audit CSV report into a standalone HTML report.

.DESCRIPTION
    Reads the CSV file produced by "Invoke-HardeningKitty -Mode Audit -Report" and renders
    a self-contained HTML report: summary highlights (score, total/passed/low/medium/high),
    a per-category breakdown, and the full findings table with client-side search, severity
    filter and column sorting.

    This script does not touch HardeningKitty.psm1 and does not query the system - it only
    reads an existing CSV report file. The generated HTML has no external dependencies (no
    CDN, no internet access required) so it can be opened on an offline machine.

.PARAMETER ReportFile
    Path to the CSV report produced by Invoke-HardeningKitty (-Report / -ReportFile).

.PARAMETER OutputFile
    Path for the generated HTML file. Defaults to the report file name with a .html extension.

.PARAMETER Title
    Optional heading shown at the top of the report. Defaults to "HardeningKitty Audit Report".

.EXAMPLE
    .\Export-HardeningKittyHtmlReport.ps1 -ReportFile .\hardeningkitty_report_SRV01_finding_list_cis-20260818.csv

.EXAMPLE
    .\Export-HardeningKittyHtmlReport.ps1 -ReportFile .\report.csv -OutputFile .\report.html -Title "SRV01 - CIS Windows Server 2022"
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [String] $ReportFile,

    [String] $OutputFile,

    [String] $Title = "HardeningKitty Audit Report"
)

$ErrorActionPreference = 'Stop'

If (-not $OutputFile) {
    $OutputFile = [System.IO.Path]::ChangeExtension((Resolve-Path -LiteralPath $ReportFile).Path, 'html')
}

$Findings = Import-Csv -LiteralPath $ReportFile -Delimiter ","

If (-not $Findings -or $Findings.Count -eq 0) {
    Write-Error "No rows found in '$ReportFile'."
}

If (-not ($Findings | Get-Member -Name 'Severity' -MemberType NoteProperty)) {
    Write-Error "'$ReportFile' does not look like a HardeningKitty report (missing 'Severity' column). Generate it with -Mode Audit -Report."
}

# Only rows with a recognised Severity come from Audit mode. Config-mode rows or rows from
# checks that errored out (missing admin rights/tooling) leave Severity blank and are counted
# separately so the highlights are not silently wrong.
$KnownSeverities = 'Passed', 'Low', 'Medium', 'High'
$Audited = $Findings | Where-Object { $_.Severity -in $KnownSeverities }
$SkippedCount = $Findings.Count - $Audited.Count

$StatsPassed = ($Audited | Where-Object { $_.Severity -eq 'Passed' }).Count
$StatsLow    = ($Audited | Where-Object { $_.Severity -eq 'Low' }).Count
$StatsMedium = ($Audited | Where-Object { $_.Severity -eq 'Medium' }).Count
$StatsHigh   = ($Audited | Where-Object { $_.Severity -eq 'High' }).Count
$StatsTotal  = $Audited.Count

# Same formula HardeningKitty itself uses for the console score, recomputed from the CSV.
$Score = 1.00
If ($StatsTotal -gt 0 -and $StatsPassed -gt 0) {
    $ScoreTotal = $StatsTotal * 4
    $ScoreAchieved = ($StatsPassed * 4) + ($StatsLow * 2) + $StatsMedium
    $Score = [math]::Round((($ScoreAchieved / $ScoreTotal) * 5) + 1, 2)
}

Function Get-Pct([int] $Count) {
    If ($StatsTotal -eq 0) { return 0 }
    [math]::Round(($Count / $StatsTotal) * 100, 1)
}
$PctPassed = Get-Pct $StatsPassed
$PctLow    = Get-Pct $StatsLow
$PctMedium = Get-Pct $StatsMedium
$PctHigh   = Get-Pct $StatsHigh

Function ConvertTo-Html([String] $Text) {
    [System.Net.WebUtility]::HtmlEncode([String] $Text)
}

# Category breakdown
$CategoryRows = New-Object System.Text.StringBuilder
$Categories = $Audited | Group-Object -Property Category | Sort-Object Name
ForEach ($Cat in $Categories) {
    $CatPassed = ($Cat.Group | Where-Object { $_.Severity -eq 'Passed' }).Count
    $CatLow    = ($Cat.Group | Where-Object { $_.Severity -eq 'Low' }).Count
    $CatMedium = ($Cat.Group | Where-Object { $_.Severity -eq 'Medium' }).Count
    $CatHigh   = ($Cat.Group | Where-Object { $_.Severity -eq 'High' }).Count
    [void] $CategoryRows.Append("<tr><td>$(ConvertTo-Html $Cat.Name)</td><td>$($Cat.Count)</td><td class=`"num passed`">$CatPassed</td><td class=`"num low`">$CatLow</td><td class=`"num medium`">$CatMedium</td><td class=`"num high`">$CatHigh</td></tr>`n")
}

# Detail rows (includes any skipped/undetermined rows so nothing from the CSV is hidden)
$DetailRows = New-Object System.Text.StringBuilder
ForEach ($F in $Findings) {
    $Sev = $F.Severity
    If ([String]::IsNullOrEmpty($Sev)) { $Sev = 'Unknown' }
    $SevClass = switch ($Sev) {
        'Passed'  { 'sev-passed' }
        'Low'     { 'sev-low' }
        'Medium'  { 'sev-medium' }
        'High'    { 'sev-high' }
        default   { 'sev-unknown' }
    }
    $DefaultBadge = ''
    If ($F.DefaultValue -eq 'True') {
        $DefaultBadge = ' <span class="badge badge-default" title="Setting was not present - finding list default value was used">default</span>'
    }
    [void] $DetailRows.Append(@"
<tr data-severity="$(ConvertTo-Html $Sev)">
<td>$(ConvertTo-Html $F.ID)</td>
<td>$(ConvertTo-Html $F.Category)</td>
<td>$(ConvertTo-Html $F.Name)</td>
<td><span class="badge $SevClass">$(ConvertTo-Html $Sev)</span></td>
<td>$(ConvertTo-Html $F.Result)$DefaultBadge</td>
<td>$(ConvertTo-Html $F.Recommended)</td>
<td>$(ConvertTo-Html $F.Filter)</td>
</tr>
"@)
}

$SkippedNote = ''
If ($SkippedCount -gt 0) {
    $SkippedNote = "<p class=`"note`">$SkippedCount row(s) in the CSV had no recognised Severity (Passed/Low/Medium/High) - likely Config-mode output or checks that errored out - and are excluded from the highlights above but still listed in the details table as <span class=`"badge sev-unknown`">Unknown</span>.</p>"
}

$GeneratedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$SourceFileName = Split-Path -Leaf $ReportFile

$Html = @"
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>$(ConvertTo-Html $Title)</title>
<style>
  :root {
    --bg: #f5f6f8; --panel: #ffffff; --text: #1c1f26; --muted: #5b6270; --border: #e2e5ea;
    --passed: #2e9e5b; --low: #2f8fd1; --medium: #d99a1b; --high: #d1453d;
  }
  * { box-sizing: border-box; }
  body { margin: 0; padding: 2rem; background: var(--bg); color: var(--text); font-family: Segoe UI, Arial, sans-serif; }
  h1 { margin: 0 0 0.25rem 0; font-size: 1.5rem; }
  .meta { color: var(--muted); font-size: 0.85rem; margin-bottom: 1.5rem; }
  .panel { background: var(--panel); border: 1px solid var(--border); border-radius: 8px; padding: 1.25rem; margin-bottom: 1.5rem; }
  .cards { display: flex; flex-wrap: wrap; gap: 1rem; margin-bottom: 1rem; }
  .card { flex: 1 1 140px; background: var(--panel); border: 1px solid var(--border); border-radius: 8px; padding: 1rem; text-align: center; }
  .card .value { font-size: 1.8rem; font-weight: 600; }
  .card .label { color: var(--muted); font-size: 0.8rem; text-transform: uppercase; letter-spacing: 0.03em; }
  .card.score .value { color: var(--text); }
  .card.passed .value { color: var(--passed); }
  .card.low .value { color: var(--low); }
  .card.medium .value { color: var(--medium); }
  .card.high .value { color: var(--high); }
  .bar { display: flex; height: 22px; border-radius: 4px; overflow: hidden; border: 1px solid var(--border); margin: 0.75rem 0; }
  .bar div { height: 100%; }
  .bar .b-passed { background: var(--passed); }
  .bar .b-low { background: var(--low); }
  .bar .b-medium { background: var(--medium); }
  .bar .b-high { background: var(--high); }
  .legend { display: flex; gap: 1.25rem; font-size: 0.8rem; color: var(--muted); flex-wrap: wrap; }
  .legend span.dot { display: inline-block; width: 10px; height: 10px; border-radius: 2px; margin-right: 0.35rem; vertical-align: middle; }
  table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
  th, td { text-align: left; padding: 0.5rem 0.6rem; border-bottom: 1px solid var(--border); vertical-align: top; }
  th { color: var(--muted); text-transform: uppercase; font-size: 0.72rem; letter-spacing: 0.03em; cursor: pointer; user-select: none; white-space: nowrap; }
  th:hover { color: var(--text); }
  td.num { text-align: right; font-variant-numeric: tabular-nums; }
  .badge { display: inline-block; padding: 0.1rem 0.5rem; border-radius: 999px; font-size: 0.72rem; font-weight: 600; color: #fff; }
  .sev-passed { background: var(--passed); }
  .sev-low { background: var(--low); }
  .sev-medium { background: var(--medium); }
  .sev-high { background: var(--high); }
  .sev-unknown { background: #8a8f99; }
  .badge-default { background: #eef0f3; color: var(--muted); font-weight: 500; }
  .controls { display: flex; gap: 0.75rem; align-items: center; margin-bottom: 1rem; flex-wrap: wrap; }
  .controls input, .controls select { padding: 0.4rem 0.6rem; border: 1px solid var(--border); border-radius: 6px; font-size: 0.85rem; }
  .controls input { flex: 1 1 240px; }
  .controls .count { color: var(--muted); font-size: 0.85rem; }
  .note { color: var(--muted); font-size: 0.8rem; }
</style>
</head>
<body>

<h1>$(ConvertTo-Html $Title)</h1>
<div class="meta">Source: $(ConvertTo-Html $SourceFileName) &middot; Generated: $GeneratedAt</div>

<div class="panel">
  <div class="cards">
    <div class="card score"><div class="value">$Score</div><div class="label">HardeningKitty Score</div></div>
    <div class="card"><div class="value">$StatsTotal</div><div class="label">Total Checks</div></div>
    <div class="card passed"><div class="value">$StatsPassed</div><div class="label">Passed</div></div>
    <div class="card low"><div class="value">$StatsLow</div><div class="label">Low</div></div>
    <div class="card medium"><div class="value">$StatsMedium</div><div class="label">Medium</div></div>
    <div class="card high"><div class="value">$StatsHigh</div><div class="label">High</div></div>
  </div>
  <div class="bar">
    <div class="b-passed" style="width:${PctPassed}%"></div><div class="b-low" style="width:${PctLow}%"></div><div class="b-medium" style="width:${PctMedium}%"></div><div class="b-high" style="width:${PctHigh}%"></div>
  </div>
  <div class="legend">
    <span><span class="dot" style="background:var(--passed)"></span>Passed $PctPassed%</span>
    <span><span class="dot" style="background:var(--low)"></span>Low $PctLow%</span>
    <span><span class="dot" style="background:var(--medium)"></span>Medium $PctMedium%</span>
    <span><span class="dot" style="background:var(--high)"></span>High $PctHigh%</span>
  </div>
  $SkippedNote
</div>

<div class="panel">
  <h2 style="margin-top:0;font-size:1.1rem;">By Category</h2>
  <table>
    <thead><tr><th>Category</th><th>Total</th><th>Passed</th><th>Low</th><th>Medium</th><th>High</th></tr></thead>
    <tbody>
    $($CategoryRows.ToString())
    </tbody>
  </table>
</div>

<div class="panel">
  <h2 style="margin-top:0;font-size:1.1rem;">Findings</h2>
  <div class="controls">
    <input type="text" id="search" placeholder="Search ID, category, name, result...">
    <select id="sevFilter">
      <option value="All">All severities</option>
      <option value="High">High</option>
      <option value="Medium">Medium</option>
      <option value="Low">Low</option>
      <option value="Passed">Passed</option>
      <option value="Unknown">Unknown</option>
    </select>
    <span class="count"><span id="visibleCount">$($Findings.Count)</span> / $($Findings.Count) shown</span>
  </div>
  <table id="detailsTable">
    <thead>
      <tr>
        <th data-sort>ID</th>
        <th data-sort>Category</th>
        <th data-sort>Name</th>
        <th data-sort>Severity</th>
        <th data-sort>Result</th>
        <th data-sort>Recommended</th>
        <th data-sort>Filter</th>
      </tr>
    </thead>
    <tbody>
    $($DetailRows.ToString())
    </tbody>
  </table>
</div>

<script>
(function () {
  var searchBox = document.getElementById('search');
  var sevFilter = document.getElementById('sevFilter');
  var visibleCount = document.getElementById('visibleCount');
  var rows = Array.prototype.slice.call(document.querySelectorAll('#detailsTable tbody tr'));

  function applyFilters() {
    var q = searchBox.value.toLowerCase();
    var sev = sevFilter.value;
    var visible = 0;
    rows.forEach(function (r) {
      var matchesText = !q || r.innerText.toLowerCase().indexOf(q) !== -1;
      var matchesSev = sev === 'All' || r.getAttribute('data-severity') === sev;
      var show = matchesText && matchesSev;
      r.style.display = show ? '' : 'none';
      if (show) visible++;
    });
    visibleCount.textContent = visible;
  }
  searchBox.addEventListener('input', applyFilters);
  sevFilter.addEventListener('change', applyFilters);

  document.querySelectorAll('#detailsTable th[data-sort]').forEach(function (th, idx) {
    th.addEventListener('click', function () {
      var table = th.closest('table');
      var tbody = table.querySelector('tbody');
      var rowsArr = Array.prototype.slice.call(tbody.querySelectorAll('tr'));
      var asc = th.getAttribute('data-asc') !== 'true';
      rowsArr.sort(function (a, b) {
        var ta = a.children[idx].innerText.trim();
        var tb = b.children[idx].innerText.trim();
        return asc ? ta.localeCompare(tb, undefined, { numeric: true }) : tb.localeCompare(ta, undefined, { numeric: true });
      });
      rowsArr.forEach(function (r) { tbody.appendChild(r); });
      table.querySelectorAll('th').forEach(function (h) { h.removeAttribute('data-asc'); });
      th.setAttribute('data-asc', asc);
    });
  });
})();
</script>

</body>
</html>
"@

Set-Content -LiteralPath $OutputFile -Value $Html -Encoding UTF8

Write-Output "HTML report written to $OutputFile"
