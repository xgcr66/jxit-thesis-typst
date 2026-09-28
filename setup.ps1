param([string]$Python='python')
$ErrorActionPreference='Stop'
Set-Location -LiteralPath $PSScriptRoot
& $Python -c "import sys; assert sys.version_info >= (3,10), 'Python 3.10+ required'"
if($LASTEXITCODE -ne 0){throw '需要 Python 3.10 或更高版本。'}
if(-not(Test-Path '.venv')){& $Python -m venv .venv;if($LASTEXITCODE -ne 0){throw '创建虚拟环境失败'}}
& '.\.venv\Scripts\python.exe' -m pip install -r requirements.txt
if($LASTEXITCODE -ne 0){throw '安装 Python 依赖失败'}
if(-not(Test-Path '.tools\typst.exe')){
  New-Item -ItemType Directory -Path '.tools' -Force | Out-Null
  $archive=Join-Path $PSScriptRoot '.tools\typst.zip'
  Invoke-WebRequest -Uri 'https://github.com/typst/typst/releases/download/v0.15.1/typst-x86_64-pc-windows-msvc.zip' -OutFile $archive
  $actual=(Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant()
  if($actual -ne '19ce3551153c2fe7ee9fa2f95208310c8f4d3209fedb699e0333faf8913f6736'){throw 'Typst 下载校验失败，停止安装。'}
  Expand-Archive -LiteralPath $archive -DestinationPath '.tools\typst-0.15.1' -Force
  Copy-Item -LiteralPath '.tools\typst-0.15.1\typst-x86_64-pc-windows-msvc\typst.exe' -Destination '.tools\typst.exe'
}
& '.\.tools\typst.exe' --version
Write-Output '安装完成。双击 build.cmd，或执行 .\build.ps1 -UpdateFields。'
