$brands = @(
  @{ Name='changan'; Path='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirchan\CON_MAYORGEN2CHAN.TXT' },
  @{ Name='peug'; Path='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirpeu\CON_MAYORGEN2PEU.TXT' },
  @{ Name='szk'; Path='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirsuz\CON_MAYORGEN2SUZ.TXT' },
  @{ Name='tyt'; Path='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirtoy\CON_MAYORGEN2TOY.TXT' }
)
foreach ($brand in $brands) {
  node .\scripts\cxp\servicios_marcas\read_mayor_txt.js --input $brand.Path --output-json .\storage\tmp_mayor_audit.json | Out-Null
  $json = Get-Content .\storage\tmp_mayor_audit.json -Raw | ConvertFrom-Json
  $rows = @($json.rows)
  $matches = @($rows | Where-Object { $_.origin -eq 'AGCM' -or $_.detail -match 'REGISTRO DE PX AJUSTE DE EGRESO' })
  Write-Output ('=== ' + $brand.Name + ' matches=' + $matches.Count)
  $matches | ForEach-Object { Write-Output ('  acct=' + $_.account + ' origin=' + $_.origin + ' seat=' + $_.seat + ' detail=' + $_.detail + ' debit=' + $_.debit + ' credit=' + $_.credit) }
}
