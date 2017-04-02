# .NET Framework クラス宣言
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# タスク入力画面
Function Show-TaskInputForm {

	#各変数設定
	$data = &$TaskDataClass $script:datafile
	$taskname = &$TaskNameClass $script:taskfile
	$now = (Get-Date).ToString("HH:mm")

	#各オブジェクト作成
	$form = New-Object System.Windows.Forms.Form
	$timeTextBox = New-Object System.Windows.Forms.TextBox
	$taskComboBox = New-Object System.Windows.Forms.ComboBox
	$okButton = New-Object System.Windows.Forms.Button
	$cancelButton = New-Object System.Windows.Forms.Button

	#時間入力用テキストボックス設定
	$timeTextBox.Location = New-Object System.Drawing.Point(1, 2)
	$timeTextBox.Size = New-Object System.Drawing.Size(34, 19)
	$timeTextBox.TabIndex = 3
	$timeTextBox.Text = $now

	#タスク入力用コンボボックス設定
	$taskComboBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
	$taskComboBox.FormattingEnabled = $true
	$taskComboBox.Location = New-Object System.Drawing.Point(36, 1)
	$taskComboBox.Size = New-Object System.Drawing.Size(([int]($script:conf["TI.WIDTH"]) - 37), 20)
	$taskComboBox.TabIndex = 0
	$taskComboBox.Items.AddRange($taskname.GetNames())
	$taskComboBox.Text = $data.LastTask($script:baseday, $now).InnerText
	if ($script:conf["TREE"] -ne 0) {
		#タスクツリー形式が選択されている場合はTreeViewでドロップダウンを隠蔽
		$taskComboBox.DropDownHeight = 1
		$taskComboBox.Add_GotFocus({
			if ($taskComboBox.DroppedDown) {
				#ALT+↓でTreeViewを開いた場合、隠蔽したDropDownも表示されているので、フォーカスが戻ったらDropDownも閉じる
				[System.Windows.Forms.SendKeys]::Send("{ESC}")
			}
		})
		$taskComboBox.Add_DropDown({
			#変数設定
			$tasktree = &$TaskTreeClass $script:conf["TREEPATH"]

			#各オブジェクト作成
			$subForm = New-Object System.Windows.Forms.Form
			$taskTreeView = New-Object System.Windows.Forms.TreeView

			#タスク選択用ツリービュー設定
			$tasktree.GetNode() | ForEach-Object {
				#何故か配列の頭に[Int]0が追加されてしまうので除外
				if ($_ -is [System.Windows.Forms.TreeNode]) {
					$taskTreeView.Nodes.Add($_)
				}
			}
			$taskTreeView.Dock = [System.Windows.Forms.DockStyle]::Fill
			$taskTreeView.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
			$taskTreeView.Add_KeyDown({
				switch ($_.KeyCode) {
					([System.Windows.Forms.Keys]::Escape) {
						$this.Parent.Close()
					}
					([System.Windows.Forms.Keys]::Enter) {
						$node = $this.SelectedNode
						$text = $node.Text
						while ($node.Parent -ne $null) {
							$text = $node.Parent.Text + $script:conf["SEP"] + $text
							$node = $node.Parent
						}
						$taskComboBox.Text = $text
						$this.Parent.Close()
					}
				}
			})
			$taskTreeView.Add_KeyUp({
				$script:tasktreeevent = $true
			})
			$taskTreeView.Add_NodeMouseClick({
				if ($script:tasktreeevent) {
					$node = $_.Node
					$text = $node.Text
					while ($node.Parent -ne $null) {
						$text = $node.Parent.Text + $script:conf["SEP"] + $text
						$node = $node.Parent
					}
					$taskComboBox.Text = $text
					$this.Parent.Close()
				} else {
					$script:tasktreeevent = $true
				}
			})
			$taskTreeView.Add_AfterExpand({
				$this.Parent.ClientSize = New-Object System.Drawing.Size($this.Parent.ClientSize.Width, ($_.Node.Nodes.Count * 16 + $this.Parent.ClientSize.Height))
				$script:tasktreeevent = $false
			})
			$taskTreeView.Add_AfterCollapse({
				$this.Parent.ClientSize = New-Object System.Drawing.Size($this.Parent.ClientSize.Width, ($this.Parent.ClientSize.Height - ($_.Node.Nodes.Count * 16)))
				$script:tasktreeevent = $false
			})

			#フォーム構築
			$subForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
			$subForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
			$pointX = ($taskComboBox.PointToScreen($taskComboBox.Location)).X - 36
			$pointY = ($taskComboBox.PointToScreen($taskComboBox.Location)).Y + $taskComboBox.Height
			$subForm.Location = New-Object System.Drawing.Point($pointX, $pointY)
			$subForm.ClientSize = New-Object System.Drawing.Size($taskComboBox.Width, ($taskTreeView.Nodes.Count * 16))
			$subForm.Controls.Add($taskTreeView)
			$subForm.Add_Deactivate({
				$this.Close()
			})

			#フォーム表示
			$subForm.Show()
		})
	}

	#OKボタン設定
	$okButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
	$okButton.Location = New-Object System.Drawing.Point(([int]($script:conf["TI.WIDTH"]) - 152), 22)
	$okButton.Size = New-Object System.Drawing.Size(75, 23)
	$okButton.TabIndex = 1
	$okButton.Text = "登録"
	$okButton.UseVisualStyleBackColor = $true
	$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

	#CANCELボタン設定
	$cancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
	$cancelButton.Location = New-Object System.Drawing.Point(([int]($script:conf["TI.WIDTH"]) - 76), 22)
	$cancelButton.Size = New-Object System.Drawing.Size(75, 23)
	$cancelButton.TabIndex = 2
	$cancelButton.Text = "キャンセル"
	$cancelButton.UseVisualStyleBackColor = $true
	$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

	#フォーム構築
	$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
	$form.Location = New-Object System.Drawing.Point(([int]($script:conf["LEFT"]) + 10), ([int]($script:conf["TOP"]) + 30))
	$form.ClientSize =  New-Object System.Drawing.Size($script:conf["TI.WIDTH"], $script:conf["TI.HEIGHT"])
	$form.MaximumSize =  New-Object System.Drawing.Size(1024, $form.Height)
	$form.MinimumSize =  New-Object System.Drawing.Size(291, $form.Height)
	$form.Controls.Add($cancelButton)
	$form.Controls.Add($okButton)
	$form.Controls.Add($taskComboBox)
	$form.Controls.Add($timeTextBox)
	$form.Text = "タスク入力"
	$form.KeyPreview = $true
	$form.Add_KeyDown({
		switch ($_.KeyCode) {
			([System.Windows.Forms.Keys]::Escape) {
				if (-not $taskComboBox.DroppedDown) {
					$cancelButton.PerformClick()
				}
			}
			([System.Windows.Forms.Keys]::Enter) {
				if (-not $taskComboBox.DroppedDown) {
					$okButton.PerformClick()
				}
			}
		}
	})
	$form.Add_ResizeEnd({
		$script:conf["TI.WIDTH"] = $this.Width
	})

	#フォーム表示
	$result = $form.ShowDialog()
	if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
		$data.AddTask($script:baseday, $timeTextBox.Text, $taskComboBox.Text)
		$data.Commit()
		$taskname.AddName($taskComboBox.Text)
		$taskname.Commit()
	}
}
