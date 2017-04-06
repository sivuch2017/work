# .NET Framework クラス宣言
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# 設定画面
Function Show-OptionForm {

	#各オブジェクト作成
	$form = New-Object System.Windows.Forms.Form
	$histCheckBox = New-Object System.Windows.Forms.CheckBox
	$treeCheckBox = New-Object System.Windows.Forms.CheckBox
	$treeLabel = New-Object System.Windows.Forms.Label
	$treeTextBox = New-Object System.Windows.Forms.TextBox
	$treeButton = New-Object System.Windows.Forms.Button
	$wbsLabel = New-Object System.Windows.Forms.Label
	$wbsTextBox = New-Object System.Windows.Forms.TextBox
	$wbsButton = New-Object System.Windows.Forms.Button
	$unitLabel = New-Object System.Windows.Forms.Label
	$unitComboBox = New-Object System.Windows.Forms.ComboBox
	$okButton = New-Object System.Windows.Forms.Button
	$cancelButton = New-Object System.Windows.Forms.Button

	#履歴表示チェックボックス設定
	$histCheckBox.AutoSize = $true
	$histCheckBox.Location = New-Object System.Drawing.Point(12, 12)
	$histCheckBox.Size = New-Object System.Drawing.Size(140, 16)
	$histCheckBox.TabIndex = 0
	$histCheckBox.Text = "メイン画面に履歴を表示"
	$histCheckBox.UseVisualStyleBackColor = $true
	if ($script:conf["TOPLIST"] -ne 0) {
		$histCheckBox.Checked = $true
	}

	#タスクツリーチェックボックス設定
	$treeCheckBox.AutoSize = $true
	$treeCheckBox.Location = New-Object System.Drawing.Point(12, 34)
	$treeCheckBox.Size = New-Object System.Drawing.Size(118, 16)
	$treeCheckBox.TabIndex = 1
	$treeCheckBox.Text = "タスク選択を固定化"
	$treeCheckBox.UseVisualStyleBackColor = $true
	if ($script:conf["TREE"] -ne 0) {
		$treeCheckBox.Checked = $true
	}
	$treeCheckBox.Add_CheckedChanged({
		if ($this.Checked) {
			$treeTextBox.ReadOnly = $false
		} else {
			$treeTextBox.ReadOnly = $true
		}
	})

	#タスクツリーラベル設定
	$treeLabel.AutoSize = $true
	$treeLabel.Location = New-Object System.Drawing.Point(12, 61)
	$treeLabel.Size = New-Object System.Drawing.Size(64, 12)
	$treeLabel.TabIndex = 2
	$treeLabel.Text = "タスクファイル"

	#タスクツリーテキストボックス設定
	$treeTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
	$treeTextBox.Location = New-Object System.Drawing.Point(12, 78)
	$treeTextBox.Size = New-Object System.Drawing.Size(240, 19)
	$treeTextBox.TabIndex = 3
	$treeTextBox.Text = $script:conf["TREEPATH"]
	$treeTextBox.ReadOnly = $true
	if ($script:conf["TREE"] -ne 0) {
		$treeTextBox.ReadOnly = $false
	}

	#タスクツリーボタン設定
	$treeButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
	$treeButton.Location = New-Object System.Drawing.Point(252, 78)
	$treeButton.Size = New-Object System.Drawing.Size(20, 19)
	$treeButton.TabIndex = 4
	$treeButton.Text = "..."
	$treeButton.UseVisualStyleBackColor = $true
	$treeButton.Add_Click({
		$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
		$fileDialog.ShowHelp = $true
		$fileDialog.CheckFileExists = $false
		$fileDialog.CheckPathExists = $false
		$fileDialog.Filter = "タスクファイル(*.tree)|*.tree"
		$fileDialog.InitialDirectory = (Split-Path -Parent $script:conf["TREEPATH"])
		$fileDialog.Title = "ファイルを選択してください"
		if ($fileDialog.ShowDialog() -eq "OK") {
			 $treeTextBox.Text = $fileDialog.FileName
		}
    Remove-Variable fileDialog
	})

  #WBSパスラベル設定
	$wbsLabel.AutoSize = $true
	$wbsLabel.Location = New-Object System.Drawing.Point(12, 100)
	$wbsLabel.Size = New-Object System.Drawing.Size(64, 12)
	$wbsLabel.TabIndex = 5
	$wbsLabel.Text = "WBSパス"

	#WBSパステキストボックス設定
	$wbsTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
	$wbsTextBox.Location = New-Object System.Drawing.Point(12, 117)
	$wbsTextBox.Size = New-Object System.Drawing.Size(240, 19)
	$wbsTextBox.TabIndex = 6
	$wbsTextBox.Text = $script:conf["WBSPATH"]

	#WBSパスボタン設定
	$wbsButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
	$wbsButton.Location = New-Object System.Drawing.Point(252, 117)
	$wbsButton.Size = New-Object System.Drawing.Size(20, 19)
	$wbsButton.TabIndex = 7
	$wbsButton.Text = "..."
	$wbsButton.UseVisualStyleBackColor = $true
	$wbsButton.Add_Click({
		$dirDialog = New-Object System.Windows.Forms.FolderBrowserDialog
		$dirDialog.SelectedPath = $script:conf["WBSPATH"]
		$dirDialog.Description = "ディレクトリを選択してください"
		$dirDialog.ShowNewFolderButton = $false
		if ($dirDialog.ShowDialog() -eq "OK") {
			 $wbsTextBox.Text = $dirDialog.SelectedPath
		}
    Remove-Variable dirDialog
	})

	#集計単位ラベル設定
	$unitLabel.AutoSize = $true
	$unitLabel.Location = New-Object System.Drawing.Point(12, 139)
	$unitLabel.Size = New-Object System.Drawing.Size(53, 12)
	$unitLabel.TabIndex = 8
	$unitLabel.Text = "集計単位"

	#集計単位コンボボックス設定
	$unitComboBox.FormattingEnabled = $true
	$unitComboBox.Items.AddRange(@("1分毎", "15分毎", "30分毎"))
	$unitComboBox.Location = New-Object System.Drawing.Point(12, 156)
	$unitComboBox.Size = New-Object System.Drawing.Size(121, 20)
	$unitComboBox.TabIndex = 9
	switch ($script:conf["UNIT"]) {
		60 { $unitComboBox.Text = "1分毎" }
		2  { $unitComboBox.Text = "30分毎" }
		4  { $unitComboBox.Text = "15分毎" }
	}

	#OKボタン設定
	$okButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
	$okButton.Location = New-Object System.Drawing.Point(116, 191)
	$okButton.Size = New-Object System.Drawing.Size(75, 23)
	$okButton.TabIndex = 10
	$okButton.Text = "保存"
	$okButton.UseVisualStyleBackColor = $true
	$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

	#CANCELボタン設定
	$cancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
	$cancelButton.Location = New-Object System.Drawing.Point(197, 191)
	$cancelButton.Size = New-Object System.Drawing.Size(75, 23)
	$cancelButton.TabIndex = 11
	$cancelButton.Text = "キャンセル"
	$cancelButton.UseVisualStyleBackColor = $true
	$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

	#フォーム構築
	$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
	$form.Location = New-Object System.Drawing.Point(([int]($script:conf["LEFT"]) + 10), ([int]($script:conf["TOP"]) + 30))
	$form.ClientSize = New-Object System.Drawing.Size(284, 262)
	$form.MaximumSize = New-Object System.Drawing.Size(1024, 262)
	$form.MinimumSize = New-Object System.Drawing.Size(300, 262)
	$form.Controls.Add($histCheckBox)
	$form.Controls.Add($treeCheckBox)
	$form.Controls.Add($treeLabel)
	$form.Controls.Add($treeTextBox)
	$form.Controls.Add($treeButton)
	$form.Controls.Add($wbsLabel)
	$form.Controls.Add($wbsTextBox)
	$form.Controls.Add($wbsButton)
	$form.Controls.Add($unitLabel)
	$form.Controls.Add($unitComboBox)
	$form.Controls.Add($okButton)
	$form.Controls.Add($cancelButton)
	$form.Text = "設定"
	$form.KeyPreview = $true
	$form.Add_KeyDown({
		switch ($_.KeyCode) {
			([System.Windows.Forms.Keys]::Escape) { $cancelButton.PerformClick() }
			([System.Windows.Forms.Keys]::Enter) { $okButton.PerformClick() }
		}
	})

	#フォーム表示
	$result = $form.ShowDialog()
	if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
		if ($histCheckBox.Checked) {
			$script:conf["TOPLIST"] = 1
		} else {
			$script:conf["TOPLIST"] = 0
		}
		if ($treeCheckBox.Checked) {
			$script:conf["TREE"] = 1
		} else {
			$script:conf["TREE"] = 0
		}
		$script:conf["TREEPATH"] = $treeTextBox.Text
		$script:conf["WBSPATH"] = $wbsTextBox.Text
		switch ($unitComboBox.Text) {
			"1分毎" { $script:conf["UNIT"] = 60 }
			"15分毎" { $script:conf["UNIT"] = 4 }
			"30分毎" { $script:conf["UNIT"] = 2 }
		}
	}
}
