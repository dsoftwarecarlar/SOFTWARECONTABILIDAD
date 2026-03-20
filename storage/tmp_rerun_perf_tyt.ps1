$outDir='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\perf_runs'
Get-ChildItem $outDir -File | Remove-Item -Force -ErrorAction SilentlyContinue
powershell.exe -Sta -NoProfile -ExecutionPolicy Bypass -File 'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\tmp_run_perf_tyt.ps1'
