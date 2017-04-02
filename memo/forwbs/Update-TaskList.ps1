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
SRA謹製WBSのパス、もしくは存在するディレクトリパス。
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

# .NET Framework クラス宣言
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# インクルード
$libpath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "${libpath}\..\TaskMemoClasses.ps1"
. "${libpath}\..\Get-ProgressBar.ps1"

Function Update-TaskList {
	Param([Parameter(Position=1)][String[]]$Path)

	if (Test-Path -Path $script:conf["TREEPATH"]) {
		Move-Item -Path $script:conf["TREEPATH"] -Destination ($script:conf["TREEPATH"] + "." + (Get-Date).ToString("yyyyMMddhhmmss"))
	}

	$xml = &$TaskTreeClass $script:conf["TREEPATH"]
	$progressform = Get-ProgressBar
	$progressform.Location = New-Object System.Drawing.Point(([int]($script:conf["LEFT"]) + 10), ([int]($script:conf["TOP"]) + 30))
	$progressform.Show()
	$progressform.Refresh()
	$parents = @{}
	$count = 1
	try {
		
		$excel = New-Object -ComObject Excel.Application
		$excel.Visible = $false
		Get-ChildItem -Path $Path | Where-Object { $_.Extension -eq ".xlsx" } | ForEach-Object {
			$fileName = $_.Name
			$progressform.Label.Text = "オープン中 ${fileName}"
			$progressform.Progress.Value = 0
			$progressform.Refresh()
			$workBooks = $excel.Workbooks.Open($_.FullName, 0, $true)
			$excel.ScreenUpdating = $false
			foreach ($sheet in $workBooks.WorkSheets) {
				if ($sheet.Name -match "^実績[_-](.*)") {
					$progressform.Label.Text = "解析中 ${fileName}"
					$progressform.Refresh()
					# テーマが記載されていない場合に使用
					$name = $Matches[1]
					# №セル
					$firstCell = $sheet.Range("A1:J10").Find("№", $sheet.Range("A1"), -4163, 1)
					if (-not $firstCell) {
						continue
					}
					# ▲セル
					$lastCell = $sheet.Cells.Find("▲", $sheet.Range("A1"), -4163, 1)
					if (-not $lastCell) {
						continue
					}
					# №セルの次の行から▲セルの手前まで各行6列(▲セルの右三つ目まで)の内容を精査
					$maxrec = $lastCell.Offset(-1, 3).Row - $firstCell.Offset(1, 0).Row + 1
					$texts = $sheet.Range($firstCell.Offset(1, 0), $lastCell.Offset(-1, 3)).Value()
					for ($r = 1; $r -le $maxrec; $r += 1) {
						for ($c = 1; $c -le 6; $c += 1) {
							$value = $texts[$r, $c]
							if (-not (($value -match "^\s*$") -or ($value -match "^\s*[0-9]+\s*$"))) {
								# 最初に見つかったのがテーマ列ではない場合はシート名を使用
								if ($count -eq 1 -and $c -ne 1) {
									$xml.AddName($name, $count)
									$parents = @{1=$count}
									$count += 1
								}
								# 親探し
								$preCount = -1
								for ($index = $c - 1; $index -ge 1; $index += -1) {
									if ($parents.ContainsKey($index)) {
										$preCount = $parents.Item($index)
										break
									}
								}
								# タスク追加
								if ($preCount -lt 0) {
									$xml.AddName($value, $count)
								} else {
									$xml.AddName($value, $count, $preCount)
								}
								# 親情報保存
								$parents.Item($c) = $count
								$count += 1
								# タスクより低いレベルの親情報削除
								for ($c += 1; $c -le 6; $c += 1) {
									$parents.Remove($c)
								}
							}
						}
						# 10%単位でプログレスバー表示
						if ($progress -ne ([Math]::Floor($r / $maxrec * 100 / 10) * 10)) {
							$progress = [Math]::Floor($r / $maxrec * 100 / 10) * 10
							#Write-Progress -Activity $fileName -Status "「${name}」読み込み中..." -PercentComplete $progress
							$progressform.Progress.PerformStep()
							$progressform.Refresh()
						}
					}
					#Write-Progress -Activity $fileName -Completed -Status "完了"
					$progressform.Label.Text = "完了 ${fileName}"
					$progressform.Refresh()
				}
			}
			$excel.ScreenUpdating = $true
			$workBooks.Close($false)
			Remove-Variable workBooks
		}
	} finally {
		if ($workBooks) {
			$workBooks.Close($false)
			Remove-Variable workBooks
		}
		$excel.Quit()
		[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
		Remove-Variable excel
		[System.GC]::Collect()
		$progressform.Close()
	}
	$xml.Commit()
}
