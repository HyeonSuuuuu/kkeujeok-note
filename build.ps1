# 끄적노트.exe를 만든다. 윈도우에 들어 있는 .NET Framework 4 컴파일러(csc)만 있으면 된다.
#   powershell -ExecutionPolicy Bypass -File build.ps1
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$ico = Join-Path $here '끄적노트.ico'
$out = Join-Path $here 'build'
if (-not (Test-Path $ico)) { & (Join-Path $here 'make-icon.ps1') -Out $ico }
New-Item -ItemType Directory -Force $out | Out-Null
& $csc /nologo /target:winexe /codepage:65001 /optimize+ "/win32icon:$ico" "/out:$out\끄적노트.exe" `
    /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll /r:System.Core.dll `
    (Join-Path $here 'Kkeujeok.cs')
if ($LASTEXITCODE -ne 0) { throw "빌드 실패" }
"만듦: $out\끄적노트.exe"
