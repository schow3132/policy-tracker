param(
  [string]$Xlsx = (Join-Path $env:USERPROFILE "Downloads\Government  Congress Policy Tracker.xlsx"),
  [string]$Out  = (Join-Path $PSScriptRoot "data\tracker.json")
)
# Re-reads the policy tracker spreadsheet and writes data/tracker.json for the website.
# Usage (from this folder):  powershell -ExecutionPolicy Bypass -File update-data.ps1
#   or pass a different file:  ... -File update-data.ps1 -Xlsx "C:\path\to\tracker.xlsx"
# The site reads sheets by their names and header labels, so keep those the same in the workbook.

$ErrorActionPreference = "Stop"
if (-not (Test-Path -LiteralPath $Xlsx)) { throw "Spreadsheet not found: $Xlsx" }

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("tracker-" + [guid]::NewGuid())
New-Item -ItemType Directory $tmp | Out-Null
try {
  Copy-Item -LiteralPath $Xlsx (Join-Path $tmp "book.zip")
  Expand-Archive (Join-Path $tmp "book.zip") (Join-Path $tmp "x")
  $Dir = Join-Path $tmp "x"

  $ns = @{ m = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"; r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships" }
  function Read-Xml($p) { $x = New-Object System.Xml.XmlDocument; $x.PreserveWhitespace = $true; $x.Load($p); return $x }
  function ColIndex($ref) {
    $n = 0; foreach ($ch in ($ref -replace "[0-9]", "").ToCharArray()) { $n = $n * 26 + ([int][char]$ch - 64) }; return $n - 1
  }

  $ss = New-Object System.Collections.ArrayList
  $ssPath = Join-Path $Dir "xl\sharedStrings.xml"
  if (Test-Path $ssPath) {
    $sx = Read-Xml $ssPath
    foreach ($si in (Select-Xml -Xml $sx -XPath "//m:si" -Namespace $ns)) {
      [void]$ss.Add((($si.Node.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText }) -join ""))
    }
  }

  $wb = Read-Xml (Join-Path $Dir "xl\workbook.xml")
  $wbRels = Read-Xml (Join-Path $Dir "xl\_rels\workbook.xml.rels")
  $relMap = @{}; foreach ($r in $wbRels.DocumentElement.ChildNodes) { if ($r.Id) { $relMap[$r.Id] = $r.Target } }

  $sheets = [ordered]@{}
  foreach ($s in (Select-Xml -Xml $wb -XPath "//m:sheet" -Namespace $ns)) {
    $sheetPath = Join-Path $Dir ("xl\" + ($relMap[$s.Node.GetAttribute("id", $ns.r)] -replace "^/?xl/", ""))
    $sx = Read-Xml $sheetPath
    $rows = New-Object System.Collections.ArrayList
    foreach ($row in (Select-Xml -Xml $sx -XPath "//m:sheetData/m:row" -Namespace $ns)) {
      $ri = [int]$row.Node.r - 1
      while ($rows.Count -le $ri) { [void]$rows.Add((New-Object System.Collections.ArrayList)) }
      $cells = $rows[$ri]
      foreach ($c in $row.Node.ChildNodes) {
        if ($c.LocalName -ne "c") { continue }
        $ci = ColIndex $c.r
        $vn = $c.ChildNodes | Where-Object { $_.LocalName -eq "v" } | Select-Object -First 1
        $isn = $c.ChildNodes | Where-Object { $_.LocalName -eq "is" } | Select-Object -First 1
        $v = $null
        if ($c.t -eq "s" -and $vn) { $v = $ss[[int]$vn.InnerText] }
        elseif ($c.t -eq "inlineStr" -and $isn) { $v = $isn.InnerText }
        elseif ($vn) { $v = $vn.InnerText }
        if ($v -eq $null) { continue }
        while ($cells.Count -le $ci) { [void]$cells.Add($null) }
        $cells[$ci] = $v
      }
    }
    $links = [ordered]@{}
    $relsPath = Join-Path (Split-Path $sheetPath) ("_rels\" + (Split-Path $sheetPath -Leaf) + ".rels")
    if (Test-Path $relsPath) {
      $sr = Read-Xml $relsPath
      $sm = @{}; foreach ($r in $sr.DocumentElement.ChildNodes) { if ($r.Id) { $sm[$r.Id] = $r.Target } }
      foreach ($h in (Select-Xml -Xml $sx -XPath "//m:hyperlinks/m:hyperlink" -Namespace $ns)) {
        $hid = $h.Node.GetAttribute("id", $ns.r)
        if ($hid -and $sm[$hid]) { $links[$h.Node.ref] = $sm[$hid] }
      }
    }
    $sheets[$s.Node.name] = [ordered]@{ rows = $rows; links = $links }
  }

  $outObj = [ordered]@{ generated = (Get-Date).ToString("yyyy-MM-dd HH:mm"); source = (Split-Path $Xlsx -Leaf); sheets = $sheets }
  New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
  [System.IO.File]::WriteAllText($Out, ($outObj | ConvertTo-Json -Depth 8 -Compress), (New-Object System.Text.UTF8Encoding $false))
  Write-Host "Updated $Out from $(Split-Path $Xlsx -Leaf) ($($sheets.Count) sheets)"
} finally {
  Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
}
