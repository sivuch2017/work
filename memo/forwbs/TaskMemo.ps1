<#
.SYNOPSIS
タスクメモ

.DESCRIPTION
日々のタスク(何を何時間行ったか)を簡単に入力できます。

デフォルトでは以下のファイルを作成します。
C:\Users\<UserName>\AppData\Local\taskmemo.cfg
C:\Users\<UserName>\AppData\Local\taskmemo.lst
C:\Users\<UserName>\Documents\taskmemo_<YYYYMMDD>

PowerShellコンソールから起動する場合、実行ポリシーを変更する必要があります。

管理者権限でPowerShellコンソールを起動し、以下を実行。
PS1> Set-ExecutionPolicy Unrestricted

元に戻す場合。
PS1> Set-ExecutionPolicy Restricted

一時的に実行ポリシーを変更したい、もしくはコマンドプロンプトが実行する場合は以下の通りです。
powershell -executionPolicy Unrestricted -F .\TaskMemo.ps1

.PARAMETER TaskMode
タスク入力画面のみ起動します。
Windowsのタスクスケジューラーと合わせて使用すると、定期的に入力できるようになります。

.EXAMPLE
TaskMemo.ps1
通常起動

.EXAMPLE
TaskMemo.ps1 -TaskMode
タスク入力画面のみ起動
#>

#起動パラメーター
Param([switch]$TaskMode)
#EXE化するとParamが認識されないので独自解析
$args | ForEach-Object { if ($_ -ieq "-TaskMode") { $TaskMode = $true } }

#.NET Framework クラス宣言
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

#scriptスコープ変数宣言
$baseday = (Get-Date).ToString("yyyyMMdd")
$basemonth = (Get-Date).ToString("yyyyMM")
$datafile = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath "taskmemo_$basemonth"
$taskfile = Join-Path -Path ([Environment]::GetFolderPath('LocalApplicationData')) -ChildPath "taskmemo.lst"
$tasktreeevent = $true
$libpath = Split-Path -Parent $MyInvocation.MyCommand.Path
#DEBUG
$datafile = Join-Path -Path $libpath -ChildPath "taskmemo_$basemonth"

#インクルード
. "${libpath}\..\ContainsProcess.ps1"
. "${libpath}\..\TaskMemoClasses.ps1"
. "${libpath}\TaskMemoConfigFile.ps1"
. "${libpath}\..\Show-TaskInputForm.ps1"
. "${libpath}\Show-OptionForm.ps1"
. "${libpath}\Show-UpdateWBS.ps1"
. "${libpath}\Update-TaskList.ps1"

#生テキスト画面
Function Get-PlainText {
	Param([string]$str, [string]$title = "デバッグ")

	$form = New-Object System.Windows.Forms.Form
	$label = New-Object System.Windows.Forms.Label

	$label.Font = (New-Object System.Drawing.Font("ＭＳ ゴシック", 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point))
	$label.AutoSize = $true
	$label.Location = New-Object System.Drawing.Point(0, 0)
	$label.TabIndex = 0
	$label.Text = $str

	$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
	$form.Location = New-Object System.Drawing.Point(([int]($script:conf["LEFT"]) + 10), ([int]($script:conf["TOP"]) + 30))
	$form.ClientSize =  New-Object System.Drawing.Size($script:conf["PT.WIDTH"], $script:conf["PT.HEIGHT"])
	$form.Controls.Add($label)
	$form.Text = $title
	$form.KeyPreview = $true
	$form.Add_KeyDown({
		switch ($_.KeyCode) {
			([System.Windows.Forms.Keys]::Escape) { $this.Close() }
		}
	})
	$form.Add_ResizeEnd({
		$script:conf["PT.WIDTH"] = $this.Width
		$script:conf["PT.HEIGHT"] = $this.Height
	})

	[void]$form.ShowDialog()
}

#メッセージダイアログ
Function MsgBox {
	Param([string]$message, [string]$caption, [Enum]$button = [System.Windows.Forms.MessageBoxButtons]::YesNo)
	[System.Windows.Forms.MessageBox]::Show($message, $caption, $button, [System.Windows.Forms.MessageBoxIcon]::Information)
}

#出社時間入力画面
Function Get-StartTime {
	#各オブジェクト作成
	$form = New-Object System.Windows.Forms.Form
	$label = New-Object System.Windows.Forms.Label
	$textBox = New-Object System.Windows.Forms.TextBox
	$okButton = New-Object System.Windows.Forms.Button

	#ラベル設定
	$label.AutoSize = $true
	$label.Location = New-Object System.Drawing.Point(12, 15)
	$label.Size = New-Object System.Drawing.Size(53, 12)
	$label.Text = "出社時間"

	#テキストボックス設定
	$textBox.Location = New-Object System.Drawing.Point(71, 12)
	$textBox.Size = New-Object System.Drawing.Size(34, 19)
	$textBox.TabIndex = 0
	$textBox.Text = $script:conf["START"]

	#OKボタン設定
	$okButton.Location = New-Object System.Drawing.Point(115, 10)
	$okButton.Size = New-Object System.Drawing.Size(64, 23)
	$okButton.TabIndex = 1
	$okButton.Text = "OK"
	$okButton.UseVisualStyleBackColor = $true
	$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

	#フォーム構築
	$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
	$form.Location = New-Object System.Drawing.Point(([int]($script:conf["LEFT"]) + 10), ([int]($script:conf["TOP"]) + 30))
	$form.ClientSize = New-Object System.Drawing.Size(187, 43)
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
	$form.Controls.Add($label)
	$form.Controls.Add($textBox)
	$form.Controls.Add($okButton)
	$form.Text = "出社時間"
	$form.KeyPreview = $true
	$form.Add_KeyDown({
		switch ($_.KeyCode) {
			([System.Windows.Forms.Keys]::Escape) {
				$okButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
				$okButton.PerformClick()
			}
			([System.Windows.Forms.Keys]::Enter) { $okButton.PerformClick() }
		}
	})

	#フォーム表示
	$result = $form.ShowDialog()
	if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
		$data = &$TaskDataClass $script:datafile
		$data.AddTask($script:baseday, $textBox.Text, "出社")
		$data.Commit()
	}
}

#XML編集画面
Function Get-TaskEdit {
	Param([string]$path, [string]$title = "XML編集", [string]$datasetname = "root")

	$form = New-Object System.Windows.Forms.Form
	$dataGridView = New-Object System.Windows.Forms.DataGridView
	$dataSet = New-Object System.Data.DataSet
	$okButton = New-Object System.Windows.Forms.Button
	$cancelButton = New-Object System.Windows.Forms.Button

	$dataSet.DataSetName = $datasetname
	$dataSet.ReadXml($path)

	$dataGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
	$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
	$dataGridView.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	$dataGridView.Location = New-Object System.Drawing.Point(0, 0)
	$dataGridView.Size = New-Object System.Drawing.Size(([int]($script:conf["XL.WIDTH"]) - 1), ([int]($script:conf["XL.HEIGHT"]) - 33))
	$dataGridView.RowTemplate.Height = 21
	$dataGridView.TabIndex = 0
	if ($datasetname -eq "recs") {
		$recsBindingSource = New-Object System.Windows.Forms.BindingSource
		$recBindingSource = New-Object System.Windows.Forms.BindingSource
		$recsBindingSource.DataSource = $dataSet
		$recsBindingSource.DataMember = "recs"
		$recBindingSource.DataSource = $recsBindingSource
		$recBindingSource.DataMember = "recs_rec"
		if ($recsBindingSource.Find("date", $script:baseday) -lt 0) {
			$row = $dataSet.Tables[0].NewRow()
			$row.date = $script:baseday
			$dataSet.Tables[0].Rows.Add($row)
			$recsBindingSource.Position = $row.recs_Id
		} else {
			$recsBindingSource.Position = $recsBindingSource.Find("date", $script:baseday)
		}
		$dataGridView.DataSource = $recBindingSource
	} else {
		$dataGridView.DataSource = $dataSet.Tables[0]
	}

	$okButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
	$okButton.Location = New-Object System.Drawing.Point(([int]($script:conf["XL.WIDTH"]) - 152), ([int]($script:conf["XL.HEIGHT"]) - 27))
	$okButton.Size = New-Object System.Drawing.Size(75, 23)
	$okButton.TabIndex = 2
	$okButton.Text = "更新"
	$okButton.UseVisualStyleBackColor = $true
	$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

	$cancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
	$cancelButton.Location = New-Object System.Drawing.Point(([int]($script:conf["XL.WIDTH"]) - 76), ([int]($script:conf["XL.HEIGHT"]) - 27))
	$cancelButton.Size = New-Object System.Drawing.Size(75, 23)
	$cancelButton.TabIndex = 3
	$cancelButton.Text = "キャンセル"
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
	$form.Text = $title
	$form.Add_ResizeEnd({
		$script:conf["XL.WIDTH"] = $this.Width
		$script:conf["XL.HEIGHT"] = $this.Height
	})

	$result = $form.ShowDialog()
	if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
		$writer = New-Object System.Xml.XmlTextWriter($path, [System.Text.Encoding]::UTF8)
		$writer.Formatting = [System.Xml.Formatting]::Indented
		$writer.WriteStartDocument()
		$dataSet.WriteXml($writer)
		$writer.Close()
	}
}

#メイン画面
Function Get-MainMenu {

	#各オブジェクト作成
	$form = New-Object System.Windows.Forms.Form
	$label = New-Object System.Windows.Forms.Label
	$button = New-Object System.Windows.Forms.Button
	$menuStrip = New-Object System.Windows.Forms.MenuStrip
	$fileToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$optionToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$listToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$editToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$timeListToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$taskListToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$histEditToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$exitToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$openToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$wbsToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$wbs2xmlToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$xml2wbsToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$dateTimePicker = New-Object System.Windows.Forms.DateTimePicker

	#TOP画面再描画用関数定義
	$redraw = {
		if ($script:conf["TOPLIST"] -ne 0) {
			$data = &$TaskDataClass $script:datafile
			$label.Text = $data.GetString($script:baseday, "time")
			$width = 219
			if ($width -lt $label.Width) { $width = $label.Width }
			$form.ClientSize = New-Object System.Drawing.Size($width, ($label.Height + 64))
		} else {
			$form.ClientSize = New-Object System.Drawing.Size(219, 64)
		}
	}

	#メニューバー設定
	$menuStrip.Items.AddRange(@($fileToolStripMenuItem, $wbsToolStripMenuItem))
	$menuStrip.Location = New-Object System.Drawing.Point(0, 0)
	$menuStrip.Size = New-Object System.Drawing.Size(189, 26)
	$menuStrip.Text = "menuStrip"

	#メニュー(ファイル)設定
	$fileToolStripMenuItem.DropDownItems.AddRange(@($openToolStripMenuItem, $optionToolStripMenuItem, $listToolStripMenuItem, $histEditToolStripMenuItem, $exitToolStripMenuItem))
	$fileToolStripMenuItem.Text = "ファイル"

	#メニュー(ファイル-設定)設定
	$optionToolStripMenuItem.Text = "設定"
	$optionToolStripMenuItem.Add_Click({
		Show-OptionForm
		&$redraw
	})

	#メニュー(ファイル-入力済一覧)設定
	$listToolStripMenuItem.DropDownItems.AddRange(@($editToolStripMenuItem, $timeListToolStripMenuItem, $taskListToolStripMenuItem))
	$listToolStripMenuItem.Text = "入力済一覧"

	#メニュー(ファイル-タスクツリー・履歴編集)設定
	$histEditToolStripMenuItem.Text = "タスクツリー・履歴編集"
	$histEditToolStripMenuItem.Add_Click({
		if ($script:conf["TREE"] -ne 0) {
			if (-not (Test-Path -Path $script:conf["TREEPATH"])) {
				$dmy = &$TaskTreeClass $script:conf["TREEPATH"]
				$dmy.AddName("dummy", "1", "")
				$dmy.Commit()
			}
			Get-TaskEdit $script:conf["TREEPATH"] "タスクツリー編集"
		} else {
			if (-not (Test-Path -Path $script:taskfile)) {
				MsgBox "履歴がありません。" "タスク履歴編集" ([System.Windows.Forms.MessageBoxButtons]::OK)
			} else {
				Get-TaskEdit $script:taskfile "タスク履歴編集"
			}
		}
	})

	#メニュー(ファイル-終了)設定
	$exitToolStripMenuItem.Text = "終了"
	$exitToolStripMenuItem.Add_Click({ $form.Dispose() })

	#メニュー(ファイル-入力済一覧-タスク編集)設定
	$editToolStripMenuItem.Text = "タスク編集"
	$editToolStripMenuItem.Add_Click({
		Get-TaskEdit $script:datafile "タスク編集" "recs"
		&$redraw
	})

	#メニュー(ファイル-入力済一覧-時系列一覧)設定
	$timeListToolStripMenuItem.Text = "時系列一覧"
	$timeListToolStripMenuItem.Add_Click({
		$data = &$TaskDataClass $script:datafile
		Get-PlainText $data.GetString($script:baseday, "time") "時系列一覧"
	})

	#メニュー(ファイル-入力済一覧-タスク別一覧)設定
	$taskListToolStripMenuItem.Text = "タスク別一覧"
	$taskListToolStripMenuItem.Add_Click({
		$data = &$TaskDataClass $script:datafile
		Get-PlainText $data.GetString($script:baseday, "task") "タスク別一覧"
	})

	#メニュー(ファイル-開く)設定
	$openToolStripMenuItem.Text = "開く"
	$openToolStripMenuItem.Add_Click({
		$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
		$fileDialog.ShowHelp = $true
		$fileDialog.CheckFileExists = $false
		$fileDialog.CheckPathExists = $false
		$fileDialog.Filter = "データファイル(taskmemo_*)|taskmemo_*"
		$fileDialog.InitialDirectory = (Split-Path -Parent $script:datafile)
		$fileDialog.Title = "ファイルを選択してください"
		if ($fileDialog.ShowDialog() -eq "OK") {
			$script:datafile = $fileDialog.FileName
			$data = &$TaskDataClass $script:datafile
			$script:baseday = $data.GetLastDay().ToString("yyyyMMdd")
			$dateTimePicker.MinDate = Get-Date -Date "1/1/1753 00:00:00"
			$dateTimePicker.MaxDate = Get-Date -Date "12/31/9998 00:00:00"
			$dateTimePicker.Value = [System.DateTime]::ParseExact($script:baseday, "yyyyMMdd", $null)
			$dateTimePicker.MinDate = New-Object System.DateTime($data.GetFirstDay().Year, $data.GetFirstDay().Month, 1)
			if ($data.GetLastDay() -lt (New-Object System.DateTime((Get-Date).Year, (Get-Date).Month, 1))) {
				$dateTimePicker.MaxDate = ((New-Object System.DateTime($data.GetLastDay().Year, $data.GetLastDay().Month, 1)).AddMonths(1).AddSeconds(-1))
			} else {
				$dateTimePicker.MaxDate = ((New-Object System.DateTime($data.GetLastDay().Year, $data.GetLastDay().Month, $data.GetLastDay().Day)).AddDays(1).AddSeconds(-1))
			}
			&$redraw
		}
	})

	#メニュー(WBS)設定
	$wbsToolStripMenuItem.DropDownItems.AddRange(@($wbs2xmlToolStripMenuItem, $xml2wbsToolStripMenuItem))
	$wbsToolStripMenuItem.Text = "WBS"

	#メニュー(WBS-タスクリスト読込)設定
	$wbs2xmlToolStripMenuItem.Text = "タスクリスト読込"
	$wbs2xmlToolStripMenuItem.Add_Click({
		Update-TaskList $script:conf["WBSPATH"]
	})

	#メニュー(WBS-実績書込)設定
	$xml2wbsToolStripMenuItem.Text = "実績書込"
	$xml2wbsToolStripMenuItem.Add_Click({
		Show-UpdateWBS $script:datafile ([System.DateTime]::ParseExact($script:baseday, "yyyyMMdd", $null))
	})

	#カレンダー設定
	$dateTimePicker.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
	$dateTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
	$dateTimePicker.Location = New-Object System.Drawing.Point(100, 2)
	$dateTimePicker.Size = New-Object System.Drawing.Size(90, 19)
	$dateTimePicker.TabIndex = 1
	$data = &$TaskDataClass $script:datafile
	$dateTimePicker.MinDate = New-Object System.DateTime($data.GetFirstDay().Year, $data.GetFirstDay().Month, 1)
	if ($data.GetLastDay() -lt (New-Object System.DateTime((Get-Date).Year, (Get-Date).Month, 1))) {
		$dateTimePicker.MaxDate = ((New-Object System.DateTime($data.GetLastDay().Year, $data.GetLastDay().Month, 1)).AddMonths(1).AddSeconds(-1))
	} else {
		$dateTimePicker.MaxDate = ((New-Object System.DateTime($data.GetLastDay().Year, $data.GetLastDay().Month, $data.GetLastDay().Day)).AddDays(1).AddSeconds(-1))
	}
	$dateTimePicker.Value = [System.DateTime]::ParseExact($script:baseday, "yyyyMMdd", $null)
	$dateTimePicker.Add_ValueChanged({
		$script:baseday = ([System.DateTime]$dateTimePicker.Value).ToString("yyyyMMdd")
		$data = &$TaskDataClass $script:datafile
		if ($data.GetDayNode($script:baseday) -eq $null) {
			Get-StartTime
		}
		&$redraw
	})

	#入力履歴ラベル設定
	$label.Font = (New-Object System.Drawing.Font("ＭＳ ゴシック", 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point))
	$label.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
	$label.Location = New-Object System.Drawing.Point(0, 26)
	$label.Size = New-Object System.Drawing.Size(189, 0)
	$label.AutoSize = $true

	#タスク入力ボタン設定
	$button.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
	$button.Location = New-Object System.Drawing.Point(0, 26)
	$button.Size = New-Object System.Drawing.Size(189, 38)
	$button.TabIndex = 0
	$button.Text = "入力"
	$button.UseVisualStyleBackColor = $true
	$button.DialogResult = [System.Windows.Forms.DialogResult]::OK

	#フォーム構築
	$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
	$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(6, 12)
	$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
	$form.ClientSize = New-Object System.Drawing.Size(189, 64)
	$form.Controls.Add($dateTimePicker)
	$form.Controls.Add($button)
	$form.Controls.Add($label)
	$form.Controls.Add($menuStrip)
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
	$form.Text = "タスクメモ"
	$form.KeyPreview = $true
	$form.Add_KeyDown({
		switch ($_.KeyCode) {
			([System.Windows.Forms.Keys]::Escape) { $this.Close() }
			([System.Windows.Forms.Keys]::Enter) { $button.PerformClick() }
		}
	})
	$form.Add_ResizeEnd({
		$script:conf["TOP"] = $this.Top
		$script:conf["LEFT"] = $this.Left
	})

	#フォーム表示
	do {
		&$redraw
		$form.Location = New-Object System.Drawing.Point($script:conf["LEFT"], $script:conf["TOP"])
		$result = $form.ShowDialog()
		if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
			Show-TaskInputForm
		}
	} while($result -eq [System.Windows.Forms.DialogResult]::OK)
}

#主処理

# 多重起動チェック
$powershellTitleName = ""
if ((Get-Process -Id $PID).ProcessName -eq "powershell") {
	$powershellTitleName = $Host.UI.RawUI.WindowTitle
	$Host.UI.RawUI.WindowTitle = "タスクメモ"
}
if (ContainsProcess) {
	exit
} elseif (ContainsProcess -id $PID -name "powershell" -title "タスクメモ") {
	exit
}

#DEBUG
$conffile = Join-Path -Path $libpath -ChildPath "taskmemo.cfg"
#設定ファイル読み込み
Load-ConfigFile

#メイン画面表示
try {
	if ($TaskMode) {
		$result = Show-TaskInputForm
	} else {
		if (Test-Path -Path $script:datafile) {
			$data = &$TaskDataClass $script:datafile
			if ($data.GetDayNode($script:baseday) -eq $null) {
				Get-StartTime
			}
		} else {
			Get-StartTime
		}
		$result = Get-MainMenu
	}
} finally {
	Write-ConfigFile
	if ((Get-Process -Id $PID).ProcessName -eq "powershell") {
		$Host.UI.RawUI.WindowTitle = $powershellTitleName
	}
}
