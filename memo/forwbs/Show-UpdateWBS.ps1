<#
.SYNOPSIS
実績反映画面表示

.DESCRIPTION
TaskMemoで入力された実績を読み込み一覧表示。
入力された進捗と実績時間をWBSに反映する。

.PARAMETER XmlPath
TaskMemoで入力された実績ファイル

.PARAMETER Date
WBSに反映させる実績日

.EXAMPLE

.NOTES
TaskMemoのコンフィグファイルが読み込まれていることが前提。
#>

# .NET Framework クラス宣言
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# インクルード
$libpath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "${libpath}\..\TaskMemoClasses.ps1"
. "${libpath}\..\Get-ProgressBar.ps1"

Function Show-UpdateWBS {
	Param([Parameter(Position=1)][String]$XmlPath, [Parameter(Position=2)][DateTime]$Date)

	$baseday = $Date.ToString("yyyyMMdd")

	$form = New-Object System.Windows.Forms.Form
	$dataGridView = New-Object System.Windows.Forms.DataGridView
	$dataSet = New-Object System.Data.DataSet
	$okButton = New-Object System.Windows.Forms.Button
	$cancelButton = New-Object System.Windows.Forms.Button

	$data = &$TaskDataClass $XmlPath
	$sta = $data.GetTasksTime($baseday, $true)
	if ($sta.Count -gt 0) {
		$reader = New-Object System.Xml.XmlNodeReader($sta.Xml)

		$dataSet.DataSetName = "root"
		[void]$dataSet.ReadXml($reader)
		$dataSet.Tables[0].Columns["name"].ColumnName = "タスク"
		$dataSet.Tables[0].Columns["time"].ColumnName = "実績(時)"
		$dataSet.Tables[0].Columns["progress"].ColumnName = "進捗率(％)"
		$dataSet.Tables[0].Columns["result"].ColumnName = "反映結果"

		$dataGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
		$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
		$dataGridView.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
		$dataGridView.Location = New-Object System.Drawing.Point(0, 0)
		$dataGridView.Size = New-Object System.Drawing.Size(([int]($script:conf["XL.WIDTH"]) - 1), ([int]($script:conf["XL.HEIGHT"]) - 33))
		$dataGridView.RowTemplate.Height = 21
		$dataGridView.TabIndex = 0
		$dataGridView.DataSource = $dataSet.Tables[0]
		$dataGridView.AllowUserToAddRows = $false
		$dataGridView.AllowUserToDeleteRows = $false
		$dataGridView.RowHeadersVisible = $false
		$dataGridView.ScrollBars = [System.Windows.Forms.ScrollBars]::None
		$openflag = $true
		$dataGridView.Add_GotFocus({
			if ($openflag) {
				# ShowDialogで描画されないとDataGridViewのColumnsが構成されない。
				# 暫定対応として初回フォーカス時に設定。
				$dataGridView.Columns[0].ReadOnly = $true
				$dataGridView.Columns[1].ReadOnly = $true
				$dataGridView.Columns[1].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
				$dataGridView.Columns[2].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
				$dataGridView.Columns[3].ReadOnly = $true
				$dataGridView.Columns[3].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
				$height = $dataGridView.ColumnHeadersHeight
				$dataGridView.Rows | ForEach-Object {
					$height += $_.Height
				}
				$form.MaximumSize =  New-Object System.Drawing.Size(1024, ([int]$height + 45 + 38))
				if ($script:conf["XL.HEIGHT"] -lt ([int]$height + 45)) {
					$form.ClientSize =  New-Object System.Drawing.Size($script:conf["XL.WIDTH"], $script:conf["XL.HEIGHT"])
					$dataGridView.Size = New-Object System.Drawing.Size(([int]($script:conf["XL.WIDTH"]) - 1), ([int]$script:conf["XL.HEIGHT"] - 45))
				} else {
					$form.ClientSize =  New-Object System.Drawing.Size($script:conf["XL.WIDTH"], ([int]$height + 45))
					$dataGridView.Size = New-Object System.Drawing.Size(([int]($script:conf["XL.WIDTH"]) - 1), ([int]$height + 2))
				}
				$openflag = $false
			}
		})

		$okButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
		$okButton.Location = New-Object System.Drawing.Point(([int]($script:conf["XL.WIDTH"]) - 152), ([int]($script:conf["XL.HEIGHT"]) - 37))
		$okButton.Size = New-Object System.Drawing.Size(75, 23)
		$okButton.TabIndex = 2
		$okButton.Text = "更新"
		$okButton.UseVisualStyleBackColor = $true
		$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

		$cancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
		$cancelButton.Location = New-Object System.Drawing.Point(([int]($script:conf["XL.WIDTH"]) - 76), ([int]($script:conf["XL.HEIGHT"]) - 37))
		$cancelButton.Size = New-Object System.Drawing.Size(75, 23)
		$cancelButton.TabIndex = 3
		$cancelButton.Text = "閉じる"
		$cancelButton.UseVisualStyleBackColor = $true
		$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

		$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
		$form.Location = New-Object System.Drawing.Point(([int]($script:conf["LEFT"]) + 10), ([int]($script:conf["TOP"]) + 30))
		$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(6, 12)
		$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
		$form.ClientSize =  New-Object System.Drawing.Size($script:conf["XL.WIDTH"], $script:conf["XL.HEIGHT"])
		$form.Controls.Add($cancelButton)
		$form.Controls.Add($okButton)
		$form.Controls.Add($dataGridView)
		$form.Text = $Date.ToString("yyyy/MM/dd")
		$form.Add_ResizeEnd({
			$script:conf["XL.WIDTH"] = $this.Width
			$script:conf["XL.HEIGHT"] = $this.Height
		})

		do {
			$result = $form.ShowDialog()
			if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
				$progressform = Get-ProgressBar
				$progressform.Location = New-Object System.Drawing.Point(([int]($script:conf["LEFT"]) + 10 + 10), ([int]($script:conf["TOP"]) + 30 + 30))
				$progressform.Show()
				$progressform.Refresh()
				$parents = @{}
				$count = 1
				try {
					$excel = New-Object -ComObject Excel.Application
					$excel.Visible = $false
					Get-ChildItem -Path $script:conf["WBSPATH"] | Where-Object { $_.Extension -eq ".xlsx" } | ForEach-Object {
						$fileName = $_.Name
						$progressform.Label.Text = "オープン中 ${fileName}"
						$progressform.Progress.Value = 0
						$progressform.Refresh()
						$workBooks = $excel.Workbooks.Open($_.FullName, 0)
						$excel.ScreenUpdating = $false
						$excel.EnableEvents = $false
						$excel.Calculation = -4135
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
								# 進捗セル
								$progCell = $sheet.Range($sheet.Rows.Item($firstCell.Row).Address()).Find("(%)", $sheet.Range($sheet.Rows.Item($firstCell.Row).Address()).Item(1), -4123, 1)
								if (-not $progCell) {
									continue
								}
								# 日付セル
								$rowarray = $sheet.Rows.Item($firstCell.Row).Value()
								for ($i = 1; $i -lt $rowarray.Count; $i += 1) {
									if ($rowarray[1,$i] -eq $Date) {
										$dateCell = $sheet.Cells.Item($firstCell.Row, $i)
										continue
									}
								}
								if (-not $dateCell) {
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
												$parents = @{1=$name}
												$count += 1
											}
											# 親探し
											$preName = ""
											for ($index = $c - 1; $index -ge 1; $index += -1) {
												if ($parents.ContainsKey($index)) {
													$preName = $parents.Item($index)
													break
												}
											}
											# タスク名確定
											if ($preName -eq "") {
												$taskName = $value
											} else {
												$taskName = @($preName, $value) -join $conf["SEP"]
											}
											# 親情報保存
											$parents.Item($c) = $taskName
											$count += 1
											# タスクより低いレベルの親情報削除
											for ($c += 1; $c -le 6; $c += 1) {
												$parents.Remove($c)
											}
											# マッチング
											foreach ($row in $dataGridView.Rows) {
												if ($row.cells[0].Value -eq $taskName) {
													# 反映処理
													if ($row.Cells[3].Value -eq "未") {
														if ($row.Cells[2].Value -ne "") {
															$progCell.Offset($r, 0).Value2 = [Int]$row.Cells[2].Value / 100
														}
														$dateCell.Offset($r, 0).Value2 = $row.Cells[1].Value
														$row.Cells[3].Value = "済(${fileName})"
														continue
													} else {
														$row.Cells[3].Value = "済(${fileName})(重複有)"
														continue
													}
												}
											}
										}
									}
									# 10%単位でプログレスバー表示
									if ($progress -ne ([Math]::Floor($r / $maxrec * 100 / 10) * 10)) {
										$progress = [Math]::Floor($r / $maxrec * 100 / 10) * 10
										#Write-Progress -Activity $fileName -Status "「${name}」マッチング中..." -PercentComplete $progress
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
						$excel.EnableEvents = $true
						$excel.Calculation = -4105
						$workBooks.Close($true)
						Remove-Variable workBooks
					}

				} finally {
					if ($workBooks) {
						$excel.ScreenUpdating = $true
						$workBooks.Close($true)
						Remove-Variable workBooks
					}
					$excel.Quit()
					[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
					Remove-Variable excel
					[System.GC]::Collect()
					$progressform.Close()
				}
				# 対象外
				foreach ($row in $dataGridView.Rows) {
					if ($row.Cells[3].Value -eq "未") {
						$row.Cells[3].Value = "対象外"
					}
				}
			}
		} while ($result -eq [System.Windows.Forms.DialogResult]::OK)
	} else {
		[void][System.Windows.Forms.MessageBox]::Show("[${baseday}]`nデータがありません。", "メッセージ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
	}

}
