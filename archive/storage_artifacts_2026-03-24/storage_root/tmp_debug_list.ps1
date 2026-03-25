$json = Get-Content 'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\tmp_mayor_chan.json' -Raw | ConvertFrom-Json
$mayorRows = @($json.rows)
$salesMayorByAccount = @{}
foreach ($mayorRow in $mayorRows) {
  $account = (($mayorRow.account -replace '[^0-9]',''))
  if ($account.Length -gt 0 -and $account.Length -lt 12) { $account = $account.PadLeft(12,'0') }
  if (-not $salesMayorByAccount.ContainsKey($account)) { $salesMayorByAccount[$account] = New-Object 'System.Collections.Generic.List[object]' }
  $salesMayorByAccount[$account].Add($mayorRow) | Out-Null
}
foreach($acct in '040101120001','040101120003','040101120010','040101120012') {
  $list = $salesMayorByAccount[$acct]
  Write-Output ('acct=' + $acct + ' listtype=' + $list.GetType().FullName + ' count=' + $list.Count)
  $arr = @($list)
  Write-Output ('  wrapcount=' + $arr.Count + ' firsttype=' + $arr[0].GetType().FullName)
  $arr2 = @($list.ToArray())
  Write-Output ('  arraycount=' + $arr2.Count + ' firsttype=' + $arr2[0].GetType().FullName + ' firstcredit=' + $arr2[0].credit)
}
