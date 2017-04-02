<#
.SYNOPSIS
TaskMemoで入力した実績をSRA謹製WBSに反映

.DESCRIPTION
指定された日付の実績を読み込み、指定されたディレクトリ内のEXCELファイルもしくは、
直接指定されたEXCELファイルの該当するタスクへ進捗と実績を書き込む。

デフォルトでは以下の実績ファイルを読み込みます。
C:\Users\<UserName>\Documents\taskmemo_<yyyyMM>

PowerShellコンソールから起動する場合、実行ポリシーを変更する必要があります。

管理者権限でPowerShellコンソールを起動し、以下を実行。
PS1> Set-ExecutionPolicy Unrestricted

元に戻す場合。
PS1> Set-ExecutionPolicy Restricted

一時的に実行ポリシーを変更したい、もしくはコマンドプロンプトが実行する場合は以下の通りです。
powershell -executionPolicy Unrestricted -F .\WBS2XML.ps1

.PARAMETER Path
WBSのパス、もしくは存在するディレクトリパス。
省略した場合はカレントディレクトリとなります。

.PARAMETER Date
反映対象の実績日。
省略した場合は当日。

.EXAMPLE
XML2WBS.ps1
カレントディレクトリにあるWBSに本日の実績を反映。

.EXAMPLE
WBS2XML.ps1 -Path C:\Docs
C:\Docs にあるWBSに本日の実績を反映。

.EXAMPLE
WBS2XML.ps1 -Path "C:\Docs\A.xlsx","C:\Docs\B.xlsx" -Date 2017/01/01
C:\Docs\A.xlsx と C:\Docs\B.xlsx の該当するタスクに2017/01/01の実績を反映。

.NOTES
実装詳細はShow-UpdateWBSを参照。
#>
Param([Parameter(Position=1)][String[]]$Path = ".", [Parameter(Position=2)][DateTime]$Date = (Get-Date))

$baseday = $Date.ToString("yyyyMMdd")
$basemonth = $Date.ToString("yyyyMM")
$datafile = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath "taskmemo_$basemonth"

# インクルード
$libpath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "${libpath}\..\TaskMemoClasses.ps1"
. "${libpath}\TaskMemoConfigFile.ps1"
. "${libpath}\Show-UpdateWBS.ps1"

#DEBUG
$conffile = Join-Path -Path $libpath -ChildPath "taskmemo.cfg"
$datafile = Join-Path -Path $libpath -ChildPath "taskmemo_$basemonth"

# 設定ファイル読み込み
Load-ConfigFile
Show-UpdateWBS $datafile $Date
Write-ConfigFile
