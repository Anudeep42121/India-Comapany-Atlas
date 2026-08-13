param(
  [string]$Workbook = 'C:\Users\isra\Downloads\Hyderabad-Companies-Jan-27-2026.xlsx',
  [string]$FoundersCsv = 'C:\Users\isra\Downloads\Tracxn Hyd Cos.csv',
  [string]$Output = 'outputs\company-dashboard\dashboard-data.js'
)

$ErrorActionPreference = 'Stop'
$outDir = Split-Path -Parent $Output
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

function Clean([object]$value) {
  if ($null -eq $value) { return '' }
  return ([string]$value).Trim()
}

# Index founders by both company name and domain. The CSV is person-level, so a
# company can have more than one founder.
$founderIndex = @{}
foreach ($row in (Import-Csv -LiteralPath $FoundersCsv)) {
  $name = ((Clean $row.'First Name') + ' ' + (Clean $row.'Last name')).Trim()
  if (-not $name) { continue }
  $person = [ordered]@{ name = $name; title = Clean $row.'Job Title'; linkedin = Clean $row.'Linkenidn' }
  foreach ($key in @((Clean $row.'Company name'), (Clean $row.'Domain'))) {
    if (-not $key) { continue }
    $normal = $key.ToLowerInvariant()
    if (-not $founderIndex.ContainsKey($normal)) { $founderIndex[$normal] = New-Object System.Collections.ArrayList }
    [void]$founderIndex[$normal].Add($person)
  }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $book = $excel.Workbooks.Open((Resolve-Path $Workbook))
  $sheet = $book.Worksheets.Item('Companies Covered 1.1')
  $rows = $sheet.UsedRange.Value2
  $companies = New-Object System.Collections.Generic.List[object]
  for ($r = 7; $r -le $rows.GetLength(0); $r++) {
    $companyName = Clean $rows[$r,2]
    if (-not $companyName) { continue }
    $state = Clean $rows[$r,7]
    $city = Clean $rows[$r,8]
    if ($state -ne 'Telangana' -or $city -notmatch 'Hyderabad') { continue }
    $domain = Clean $rows[$r,3]
    $founders = @()
    foreach ($key in @($companyName.ToLowerInvariant(), $domain.ToLowerInvariant())) {
      if ($founderIndex.ContainsKey($key)) { $founders = @($founderIndex[$key]); break }
    }
    $sector = Clean $rows[$r,10]
    if ($sector.Length -gt 130) { $sector = $sector.Substring(0,130) + '…' }
    $description = Clean $rows[$r,4]
    if (-not $description) { $description = Clean $rows[$r,9] }
    if ($description.Length -gt 220) { $description = $description.Substring(0,220) + '…' }
    $companies.Add([ordered]@{
      name = $companyName; domain = $domain; state = $state; district = 'Hyderabad'; city = $city
      founded = Clean ($rows[$r,5]); sector = $sector; stage = Clean ($rows[$r,16])
      funded = Clean ($rows[$r,18]); funding = Clean ($rows[$r,19]); employees = Clean ($rows[$r,31])
      website = Clean ($rows[$r,51]); linkedin = Clean ($rows[$r,54]); description = $description; founders = $founders
    })
  }
  $book.Close($false)
} finally {
  $excel.Quit()
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
}

$data = [ordered]@{ generatedAt = (Get-Date).ToString('s'); companies = @($companies) }
$json = $data | ConvertTo-Json -Depth 6 -Compress
"window.DASHBOARD_DATA = $json;" | Set-Content -LiteralPath $Output -Encoding utf8
Write-Output ("Exported {0} Hyderabad companies to {1}" -f $companies.Count, $Output)
