<#
.SYNOPSIS
.DESCRIPTION
.EXAMPLE
.NOTES
#>

#.NET Framework クラス宣言
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

Function Get-ProgressBar {
	#各オブジェクト作成
	$form = New-Object System.Windows.Forms.Form
	$label = New-Object System.Windows.Forms.Label
	$statusStrip = New-Object System.Windows.Forms.StatusStrip
	$toolStripProgressBar = New-Object System.Windows.Forms.ToolStripProgressBar
	#ラベル設定
	$label.Font = (New-Object System.Drawing.Font("ＭＳ ゴシック", 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point))
	$label.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
	$label.Location = New-Object System.Drawing.Point(12, 9)
	$label.Size = New-Object System.Drawing.Size(163, 16)
	$label.AutoSize = $true
	$label.Add_ReSize({
		if ($this.Parent.ClientSize.Width -lt ($this.Width + 24)) {
			$this.Parent.Width = $this.Width + 24 + 6
		}
	})
	#ステータスバー設定
	$statusStrip.Items.AddRange(@($toolStripProgressBar))
	$statusStrip.Location = New-Object System.Drawing.Point(0, 33)
	$statusStrip.Size = New-Object System.Drawing.Size(284, 22)
	$statusStrip.SizingGrip = $false
	#プログレスバー設定
	$toolStripProgressBar.Size = New-Object System.Drawing.Size(100, 16)
	#フォーム設定
	$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
	$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(6, 12)
	$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
	$form.ClientSize = New-Object System.Drawing.Size(284, 55)
	$form.Controls.Add($label)
	$form.Controls.Add($statusStrip)
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.Text = "進捗"
	#追加プロパティ
	Add-Member -InputObject $form -MemberType NoteProperty -Name Label -value $label
	Add-Member -InputObject $form -MemberType NoteProperty -Name Progress -value $toolStripProgressBar
	#戻り
	return $form
}
