<#
.SYNOPSIS
WBSからTaskMemoのタスク一覧を生成

.DESCRIPTION
指定されたディレクトリ内のEXCELファイルもしくは直接指定されたEXCELファイルを読み込み、
内部に書かれた各タスクをTaskMemoのタスク一覧として出力する。

デフォルトでは以下のファイルを作成します。
C:\Users\<UserName>\AppData\Local\taskmemo.tree

既にファイルが存在する場合、以下の名称でバックアップを作成します。
taskmemo.tree.<yyyyMMddhhmmss>

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

.EXAMPLE
WBS2XML.ps1
カレントディレクトリにあるWBSが対象。

.EXAMPLE
WBS2XML.ps1 -Path C:\Docs
C:\Docs にあるWBSが対象。

.EXAMPLE
WBS2XML.ps1 -Path "C:\Docs\A.xlsx","C:\Docs\B.xlsx"
C:\Docs\A.xlsx と C:\Docs\B.xlsx が対象。
#>
Param([Parameter(Position=1)][String[]]$Path = ".")

$libpath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "${libpath}\..\TaskMemoClasses.ps1"
. "${libpath}\TaskMemoConfigFile.ps1"
. "${libpath}\Update-TaskList.ps1"

#DEBUG
$conffile = Join-Path -Path $libpath -ChildPath "taskmemo.cfg"
$datafile = Join-Path -Path $libpath -ChildPath "taskmemo_$basemonth"

# 設定ファイル読み込み
#セパレータと同じ文字が使われている場合の措置が必要
Load-ConfigFile

Update-TaskList -Path $Path
