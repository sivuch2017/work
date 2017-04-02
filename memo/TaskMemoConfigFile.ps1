<#
.SYNOPSIS

.DESCRIPTION

.NOTES
#>

# scriptスコープ変数宣言
$script:conf = @{}
$script:conffile = Join-Path -Path ([Environment]::GetFolderPath('LocalApplicationData')) -ChildPath "taskmemo.cfg"

# 設定ファイル読み込み
Function Load-ConfigFile {
	if (Test-Path -Path $script:conffile) {
		Get-Content $script:conffile | ForEach-Object { $script:conf.Add(($_ -split "`v")[0], ($_ -split "`v")[1]) }
	}
	if ($script:conf.ContainsKey("TOP") -eq $false) { $script:conf.Add("TOP", 100) }
	if ($script:conf.ContainsKey("LEFT") -eq $false) { $script:conf.Add("LEFT", 100) }
	if ($script:conf.ContainsKey("TREE") -eq $false) { $script:conf.Add("TREE", 0) }
	if ($script:conf.ContainsKey("TREEPATH") -eq $false) { $script:conf.Add("TREEPATH", (Join-Path -Path ([Environment]::GetFolderPath('LocalApplicationData')) -ChildPath "taskmemo.tree")) }
	if ($script:conf.ContainsKey("TOPLIST") -eq $false) { $script:conf.Add("TOPLIST", 0) }
	if ($script:conf.ContainsKey("PT.WIDTH") -eq $false) { $script:conf.Add("PT.WIDTH", 284) }
	if ($script:conf.ContainsKey("PT.HEIGHT") -eq $false) { $script:conf.Add("PT.HEIGHT", 262) }
	if ($script:conf.ContainsKey("TI.WIDTH") -eq $false) { $script:conf.Add("TI.WIDTH", 291) }
	if ($script:conf.ContainsKey("TI.HEIGHT") -eq $false) { $script:conf.Add("TI.HEIGHT", 46) }
	if ($script:conf.ContainsKey("XL.WIDTH") -eq $false) { $script:conf.Add("XL.WIDTH", 284) }
	if ($script:conf.ContainsKey("XL.HEIGHT") -eq $false) { $script:conf.Add("XL.HEIGHT", 262) }
	if ($script:conf.ContainsKey("UNIT") -eq $false) { $script:conf.Add("UNIT", 60) }
	if ($script:conf.ContainsKey("SEP") -eq $false) { $script:conf.Add("SEP", "-") }
	if ($script:conf.ContainsKey("START") -eq $false) { $script:conf.Add("START", "09:00") }
}

# 設定ファイル書き込み
Function Write-ConfigFile {
	$script:conf.Keys | ForEach-Object { ($_, $script:conf[$_]) -join "`v" } | Out-File -FilePath $script:conffile
}
