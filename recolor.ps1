$f = Resolve-Path "outputs\company-dashboard\index.html"
$c = [System.IO.File]::ReadAllText($f, [System.Text.Encoding]::UTF8)

# Fix garbled characters first
$c = $c -replace [char]0xE2+[char]0x80+[char]0x94, [char]0x2014  # â€" -> —
$c = $c -replace [char]0xC2+[char]0xB7, [char]0xB7               # Â· -> ·
$c = $c -replace 'â€"', '—'
$c = $c -replace 'Â·', '·'
$c = $c -replace 'â€™', "'"
$c = $c -replace 'â€˜', "'"

# Color replacements
$c = $c -replace '#14b8a6','#3ecfb2'
$c = $c -replace '#2dd4bf','#2ec4a9'
$c = $c -replace '#0f766e','#1a9e88'
$c = $c -replace '#115e59','#0d7a66'
$c = $c -replace '#1d2534','#163028'
$c = $c -replace '#124f47','#1a4a3a'
$c = $c -replace 'rgba\(20,184,166,','rgba(62,207,178,'
$c = $c -replace 'rgba\(15,118,110,','rgba(26,158,136,'
$c = $c -replace '#0a0d12','#0d2420'
$c = $c -replace '#181f2d','#163830'
$c = $c -replace '#121723','#122820'
$c = $c -replace '#1a2232','#163028'
$c = $c -replace '#222a3c','#1c3c32'
$c = $c -replace '#f8fafc','#f5f0e8'
$c = $c -replace '#eef2ff','#ede8e0'
$c = $c -replace '#f1f5f9','#f0ebe3'
$c = $c -replace '#0f172a','#1a2e2a'
$c = $c -replace '#667085','#5a8a7e'
$c = $c -replace 'rgba\(15,20,30,','rgba(13,36,32,'
$c = $c -replace '#a8bacd','#8ab8ae'
$c = $c -replace '#c3cad4','#a8c4be'
$c = $c -replace '#d4d8de','#c8d8d4'
$c = $c -replace 'rgba\(155,169,184,','rgba(138,184,174,'
$c = $c -replace '#b3bbcc','#8ab8ae'
$c = $c -replace '#e7ecf3','#d4e8e2'

[System.IO.File]::WriteAllText($f, $c, [System.Text.Encoding]::UTF8)
Write-Host "Done"
