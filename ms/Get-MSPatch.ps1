<#
.SYNOPSIS
インターネットからMicrosoftUpdateの情報を回収

.DESCRIPTION
KB番号等を元にMicrosoftダウンロードセンターやMicrosoftUpdateカタログから、パッチ情報を回収し出力する。
ダウンロードパスをしてしている場合はパッチのダウンロードも行う。

.PARAMETER MBSALogPath
MBSAが出力したXMLファイルを指定。

.PARAMETER $CSVPath
MicrosoftUpdate情報を出力したCSVファイルを指定。

.PARAMETER $PatchPath
ダウンロード済みファイルがあるパス。
デフォルトは"\\172.127.56.58\patch\各基盤標準"。

.PARAMETER $ProxyUrl
.PARAMETER $ProxyUser
.PARAMETER $ProxyPass
.PARAMETER $DownloadPath
ダウンロードしたファイルの保存するパス。
指定しない場合はダウンロードを実施しない。

.PARAMETER $OldCSVPath
前回出力したCSVファイル。
#>
[CmdletBinding(DefaultParameterSetName="MBSA")]
Param(
	[Parameter(Mandatory=$true,Position=1,ParameterSetName="MBSA")][ValidateScript({ Test-Path $_ })][String]$MBSALogPath,
	[Parameter(Mandatory=$true,Position=1,ParameterSetName="CSV")][ValidateScript({ Test-Path $_ })][String]$CSVPath,
	[Parameter(Mandatory=$true,Position=1,ParameterSetName="SECG")][ValidateScript({ Test-Path $_ })][String]$SecurityGuidancePath,
	[String[]]$PatchPath = "\\172.127.56.58\patch\各基盤標準",
	[String]$ProxyUrl,
	[String]$ProxyUser,
	[String]$ProxyPass,
	[String]$DownloadPath,
	[ValidateScript({ Test-Path $_ })][String]$OldCSVPath
)

# インクルード
Add-Type -AssemblyName System.Web
$libpath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "${libpath}\SharedLib.ps1"

# 内部関数
# 入力情報オブジェクト生成
Function New-BaseObject {
	Param([String]$kb, [String]$title, [String]$category, [String]$url)
	$obj = New-Object PSObject
	$obj | Add-Member -MemberType NoteProperty -Name kb -Value $kb
	$obj | Add-Member -MemberType NoteProperty -Name title -Value $title
	$obj | Add-Member -MemberType NoteProperty -Name category -Value $category
	$obj | Add-Member -MemberType NoteProperty -Name url -Value $url
	return $obj
}

# 内部関数
# 出力オブジェクト生成
Function New-DetailObject {
	Param([String]$kb, [String]$desc, [String]$file, [String]$url)
	$obj = New-Object PSObject
	$obj | Add-Member -MemberType NoteProperty -Name kb -Value $kb
	$obj | Add-Member -MemberType NoteProperty -Name desc -Value $desc
	$obj | Add-Member -MemberType NoteProperty -Name detail -Value ""
	$obj | Add-Member -MemberType NoteProperty -Name update -Value ""
	$obj | Add-Member -MemberType NoteProperty -Name dummy -Value ""
	$obj | Add-Member -MemberType NoteProperty -Name build -Value ""
	$obj | Add-Member -MemberType NoteProperty -Name file -Value $file
	$obj | Add-Member -MemberType NoteProperty -Name path -Value ""
	$obj | Add-Member -MemberType NoteProperty -Name url -Value $url
	return $obj
}

# 内部関数
Function Get-InfoFromDownloadCenter {
	Param([ref]$obj)
	$script:ie.Navigate($obj.value.url)
	While ($script:ie.Busy) {
		Start-Sleep -s 1
	}
	[System.__ComObject].InvokeMember("getElementsByTagName", [System.Reflection.BindingFlags]::InvokeMethod, $null, $script:ie.Document, "noscript") | ForEach-Object {
		if ($_.innerText -match "data-bi-dlnm=`"([^`"]+)`".*href=`"confirmation\.aspx\?id=([0-9]+)`"") {
			$obj.value.title = $Matches[1]
			$id = $Matches[2]
			$obj.value.url = "https://www.microsoft.com/ja-jp/download/confirmation.aspx?id=${id}"
		}
	}
}

# 内部関数
Function Download-FileFromDownloadCenter {
	Param([ref]$obj, [ref]$outobj)
	$script:ie.Navigate($obj.value.url)
	While ($script:ie.Busy) {
		Start-Sleep -s 1
	}
	[System.__ComObject].InvokeMember("getElementsByTagName", [System.Reflection.BindingFlags]::InvokeMethod, $null, $script:ie.Document, "a") | ForEach-Object {
		if ($_.href -match "^https://download.microsoft.com/download/") {
			$uri = [uri]$_.href
			if (-not $outobj.value) {
				$outobj.value = New-DetailObject -kb $obj.value.kb -url $obj.value.url -desc $obj.value.title
			}
			$outobj.value.file = Split-Path -Leaf -Path $uri.LocalPath
			if ($script:files.Contains($outobj.value.file)) {
				Write-Debug "Downloded"
				$outobj.value.path = $script:files[$outobj.file]
			} else {
				$outobj.value.path = $uri.OriginalString
				$target = Join-Path $script:DownloadPath $outobj.value.file
				$script:webcient.DownloadFile($uri.OriginalString, $target)
			}
		}
	}
}

# 内部関数
Function Get-LinkByMicrosoftUpdateCatalog {
	Param([String]$kb, [String]$title)
	$linkelements = @()
	if ($title) {
		# タイトルで検索
		$encodetitle = [System.Web.HttpUtility]::UrlEncode($title)
		$url = "http://www.catalog.update.microsoft.com/Search.aspx?q=${encodetitle}"
		$script:ie.Navigate($url)
		While ($script:ie.Busy) {
			Start-Sleep -s 1
		}
		[System.__ComObject].InvokeMember("getElementsByTagName", [System.Reflection.BindingFlags]::InvokeMethod, $null, $script:ie.Document, "a") | ForEach-Object {
			if ($_.outerHTML -match "<A[^>]*goToDetails\(.([^>)]+).\);'") {
				$linkelements += @{id = $Matches[1]; name = $_.innerText.Trim(); element = $null;}
			}
		}
	}
	if (-not $linkelements) {
		# KBで検索
		$url = "http://www.catalog.update.microsoft.com/Search.aspx?q=${kb}"
		$script:ie.Navigate($url)
		While ($script:ie.Busy) {
			Start-Sleep -s 1
		}
		[System.__ComObject].InvokeMember("getElementsByTagName", [System.Reflection.BindingFlags]::InvokeMethod, $null, $script:ie.Document, "a") | ForEach-Object {
			if ($_.outerHTML -match "<A[^>]*goToDetails\(.([^>)]+).\);'") {
				$linkelements += @{id = $Matches[1]; name = $_.innerText.Trim(); element = $null;}
			}
		}
	}
	# ダウンロードボタン取得
	$inputelements = @([System.__ComObject].InvokeMember("getElementsByTagName", [System.Reflection.BindingFlags]::InvokeMethod, $null, $script:ie.Document, "input") | ? { $_.Value -eq "ダウンロード" })
	if ($linkelements.Count -ne $inputelements.Count) {
		Write-Error "リンクとボタンの数が合わない。$kb"
		return @()
	}
	for ($index = 0; $index -lt $linkelements.Count; $index++) {
		$linkelements[$index].element = $inputelements[$index]
	}
	return $linkelements
}

# 内部関数
Function Get-InfoByMicrosoftUpdateCatalog {
	Param([String]$id, [String]$kb, [ref]$refobj)
	# KB情報を取得
	$url = "http://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid=" + $id
	if ($refobj.value) {
		$obj = $refobj.value
		$obj.url = $url
	} else {
		$obj = New-DetailObject -kb $kb -url $url
		$refobj.value = $obj
	}
	try {
		$content = Get-HttpContents $url $script:ProxyUrl $script:ProxyUser $script:ProxyPass
	} catch [Exception] {
		$obj.desc = "ダウンロード情報取得エラー"
		$obj.detail = ($error[0] -replace "\n","")
	}
	if ($content -ne "") {
		$match = [RegEx]::Matches($content, '<span id="ScopedViewHandler_titleText">(.+)?</span>')
		if ($match.Count -gt 0) {
			$update = [RegEx]::Matches($content, '<span id="ScopedViewHandler_date">(.+)?</span>')
			$detail = [RegEx]::Matches($content, '<span id="ScopedViewHandler_desc">(.+)?</span>')
			$build = [RegEx]::Matches($content, '<div id="securityBullitenDiv">((?!</div>)[\s\S])*</span>(((?!</div>)[\s\S])*)</div>')
			$obj.desc = $match[0].Groups[1].Value
			$obj.detail = $detail[0].Groups[1].Value
			$obj.update = $update[0].Groups[1].Value
			$obj.build = $build[0].Groups[2].Value.Trim()
		} else {
			$obj.desc = "ダウンロード情報取得エラー"
			$obj.detail = "情報なし"
		}
	}
}

# 内部関数
Function Download-FileFromMicrosoftUpdateCatalog {
	Param([ref]$inputelement, [ref]$obj)
	# MicrosoftUpdateCatalog
	$inputelement.value.click()
	Start-Sleep -s 10
	$windows = $script:shell.windows()
	foreach ($window in $windows) {
		if ($window.locationURL.contains("DownloadDialog.aspx")) {
			$window.document.GetElementsByTagName("A") | ForEach-Object {
				$uri = [uri]$_.href
				$obj.value.file = Split-Path -Leaf -Path $uri.LocalPath
				if ($script:files.Contains($obj.value.file)) {
					$obj.value.path = $script:files[$obj.value.file]
				} else {
					$obj.value.path = $uri.OriginalString
					$target = Join-Path $DownloadPath $obj.value.file
					$script:webcient.DownloadFile($uri.OriginalString, $target)
				}
			}
			$window.Quit()
		}
	}
}

# ダウンロード済みリスト ファイル名：パス パスは複数の場合カンマ区切り
$files = @{}
Get-ChildItem -Path $PatchPath -Recurse | Where-Object { (-not $_.PSIsContainer) -and ($_.Name -notmatch "\.lnk$") } | ForEach-Object {
	if ($files.Contains($_.Name)) {
		$files[$_.Name] += ("," + $_.FullName)
	} else {
		$files.Add($_.Name, $_.FullName)
	}
}

# 前回作成リスト KB番号：オブジェクトのリスト
$olddata = @{}
if ($OldCSVPath) {
	Get-Content $OldCSVPath | ConvertFrom-Csv | ForEach-Object {
		if ($_.update) {
			if ($olddata.Contains($_.kb)) {
				$olddata[$_.kb] += $_
			} else {
				$olddata.Add($_.kb, @($_))
			}
		}
	}
}

# 入力情報作成
$inputinfoarray = @()
if ($MBSALogPath) {
	$xml = [xml](Get-Content $MBSALogPath)
	$xml.GetElementsByTagName("Check") | ForEach-Object {
		$category = $_.GroupName
		$_.GetElementsByTagName("UpdateData") | ForEach-Object {
			$kb = "KB" + $_.KBID
			if ($_.IsInstalled -ne "true") {
				$_.GetElementsByTagName("Title") | ForEach-Object {
					$inputinfoarray += New-BaseObject -kb $kb -title $_.InnerText -category $category
				}
			}
		}
	}
} elseif ($CSVPath) {
	Get-Content $CSVPath | ConvertFrom-Csv | ForEach-Object {
		$inputinfoarray += New-BaseObject -kb $_.KB -title $_.Title -category $_.Categories
	}
} elseif ($SecurityGuidancePath) {
	Get-Content $SecurityGuidancePath | ConvertFrom-Csv | ForEach-Object {
		$kb = "KB" + $_.記事
		$inputinfoarray += New-BaseObject -kb $kb -category $_.製品ファミリ -url $_.ダウンロード
	}
}

# 制御用IEオブジェクト取得
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $true

# 制御用シェルオブジェクト取得
$shell = New-Object -ComObject "Shell.Application"

# ダウンロード用オブジェクト取得
$webcient = Get-WebClient $ProxyUrl $ProxyUser $ProxyPass

# MSからパッチ情報・ファイルを取得
$inputinfoarray | ForEach-Object {
	$obj = $_
	Write-Debug $obj.category
	if ($olddata.Contains($obj.kb)) {
		# 過去に取得済みの場合は過去データを出力
		Write-Debug ("Matched" + $obj.kb)
		$olddata[$obj.kb]
	} else {
		if (-not $obj.title -and $obj.url) {
			# SECGパターン (タイトルなし URLあり)
			$outobj = $null
			Get-InfoFromDownloadCenter([ref]$obj)
			if ($DownloadPath) {
				Write-Debug $obj.category
				if ($obj.category -match "Office") {
					Write-Debug "bp2"
					Download-FileFromDownloadCenter -obj ([ref]$obj) -ouyobj ([ref]$outobj)
				}
			}
		}
		Get-LinkByMicrosoftUpdateCatalog -kb $kb -title $title | ForEach-Object {
			if ($title -eq "" -or $_.name -eq $title) {
				Get-InfoByMicrosoftUpdateCatalog -kb $obj.kb -id $_.id -refobj ([ref]$outobj)
				if ($DownloadPath) {
					if ($obj.category -match "Office") {
						if (-not $outobj.file) {
							Download-FileFromDownloadCenter -obj ([ref]$obj) -ouyobj ([ref]$outobj)
						}
					} else {
						Download-FileFromMicrosoftUpdateCatalog -inputelement ([ref]$_.element) -obj ([ref]$outobj)
					}
				}
				$outobj
			}
		}
	}
}
