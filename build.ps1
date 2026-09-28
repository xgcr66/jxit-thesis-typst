param([ValidateSet('all','pdf','docx')][string]$Format='all',[int]$Version=2,[switch]$UpdateFields)
$ErrorActionPreference='Stop'
Set-Location -LiteralPath $PSScriptRoot
$python=$null
if(Test-Path '.venv\Scripts\python.exe'){$python=Join-Path $PSScriptRoot '.venv\Scripts\python.exe'}
elseif(Test-Path 'local-tools.json'){$python=(Get-Content -Raw 'local-tools.json' | ConvertFrom-Json).python}
elseif(Get-Command python -ErrorAction SilentlyContinue){$python=(Get-Command python).Source}
if(-not $python -or -not(Test-Path -LiteralPath $python)){throw '请先运行 setup.ps1 安装依赖，或将 Python 加入 PATH。'}
$argsList=@('scripts/build.py','--format',$Format,'--version',"$Version")
if($UpdateFields){$argsList+='--update-fields'}
& $python @argsList
if($LASTEXITCODE -ne 0){throw '导出失败，请检查上方错误。'}
