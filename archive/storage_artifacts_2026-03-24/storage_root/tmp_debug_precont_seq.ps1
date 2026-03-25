# inspect critical generation standalone
$excel = New-Object -ComObject Excel.Application
$excel.Visible=$false; $excel.DisplayAlerts=$false
$protos=@()
try {
  $wb=$excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios CHANGAN  2026.xls',$false,$true)
  $ws=$wb.Worksheets.Item('PrecontabilizacionVentas')
  $used=$ws.UsedRange
  $last=[int]($used.Row + $used.Rows.Count - 1)
  for($row=2;$row -le $last;$row++){
    $acctText=[string]$ws.Cells.Item($row,5).Text
    $acctDigits=$acctText -replace '[^0-9]',''
    $acct=if($acctDigits.Length -ge 10){$acctDigits.PadLeft(12,'0')}else{$acctText}
    $doc=[string]$ws.Cells.Item($row,3).Text
    if($acct -eq '' -or $doc -eq ''){continue}
    $protos += [pscustomobject]@{TemplateRow=$row;Doc=$doc.Trim();Account=$acct.Trim();Description=([string]$ws.Cells.Item($row,6).Text).Trim()}
  }
  $wb.Close($false)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($used)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($ws)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
}
finally { $excel.Quit(); [void][Runtime.Interopservices.Marshal]::ReleaseComObject($excel); [gc]::Collect(); [gc]::WaitForPendingFinalizers() }

$json = Get-Content 'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\tmp_mayor_chan.json' -Raw | ConvertFrom-Json
$mayorRows=@($json.rows)
$salesMayorByAccount=@{}
foreach($mayorRow in $mayorRows){
  $account = ($mayorRow.account -replace '[^0-9]','')
  if($account.Length -gt 0 -and $account.Length -lt 12){$account=$account.PadLeft(12,'0')}
  if(-not $salesMayorByAccount.ContainsKey($account)){ $salesMayorByAccount[$account] = New-Object 'System.Collections.Generic.List[object]' }
  $salesMayorByAccount[$account].Add($mayorRow)|Out-Null
}
function Add-GeneratedTemplateRow { param([System.Collections.Generic.List[object]]$Target,[pscustomobject]$Prototype,[double]$Debit=0,[double]$Credit=0) $Target.Add([pscustomobject]@{Row=$Prototype.TemplateRow;Doc=$Prototype.Doc;Acct=$Prototype.Account;Desc=$Prototype.Description;Debit=$Debit;Credit=$Credit})|Out-Null }
function Add-TemplateSequentialRowsFromMayor { param([System.Collections.Generic.List[object]]$Target,[object[]]$Prototypes,[object[]]$MayorRows,[switch]$EnsureAtLeastOneRow)
  $orderedPrototypes=@($Prototypes)
  $items=@($MayorRows)
  Write-Host ('itemscount=' + $items.Count)
  for($index=0;$index -lt $items.Count;$index++){
    $prototype = if($index -lt $orderedPrototypes.Count){$orderedPrototypes[$index]}else{$orderedPrototypes[$orderedPrototypes.Count-1]}
    $item=$items[$index]
    Add-GeneratedTemplateRow -Target $Target -Prototype $prototype -Debit ([double]$item.debit) -Credit ([double]$item.credit)
  }
}
$generated = New-Object 'System.Collections.Generic.List[object]'
$account='040101120001'
$accountPrototypes=@($protos | Where-Object Account -eq $account | Sort-Object TemplateRow)
$mayorGroup=@($salesMayorByAccount[$account].ToArray())
Write-Host ('protos=' + $accountPrototypes.Count + ' mayor=' + $mayorGroup.Count)
Add-TemplateSequentialRowsFromMayor -Target $generated -Prototypes $accountPrototypes -MayorRows $mayorGroup
$generated | Format-Table -AutoSize | Out-String -Width 300 | Write-Output
