<#
.SYNOPSIS
�C���^�[�l�b�g����MicrosoftUpdate�̏������

.DESCRIPTION
KB�ԍ���������Microsoft�_�E�����[�h�Z���^�[��MicrosoftUpdate�J�^���O����A�p�b�`����������o�͂���B
�_�E�����[�h�p�X�����Ă��Ă���ꍇ�̓p�b�`�̃_�E�����[�h���s���B

.PARAMETER MBSALogPath
MBSA���o�͂���XML�t�@�C�����w��B

.PARAMETER $CSVPath
MicrosoftUpdate�����o�͂���CSV�t�@�C�����w��B

.PARAMETER $PatchPath
�_�E�����[�h�ς݃t�@�C��������p�X�B
�f�t�H���g��"\\172.127.56.58\patch\�e��ՕW��"�B

.PARAMETER $ProxyUrl
.PARAMETER $ProxyUser
.PARAMETER $ProxyPass
.PARAMETER $DownloadPath
�_�E�����[�h�����t�@�C���̕ۑ�����p�X�B
�w�肵�Ȃ��ꍇ�̓_�E�����[�h�����{���Ȃ��B

.PARAMETER $OldCSVPath
�O��o�͂���CSV�t�@�C���B
#>
[CmdletBinding(DefaultParameterSetName="MBSA")]
Param(
	[Parameter(Mandatory=$true,Position=1,ParameterSetName="MBSA")][ValidateScript({ Test-Path $_ })][String]$MBSALogPath,
	[Parameter(Mandatory=$true,Position=1,ParameterSetName="CSV")][ValidateScript({ Test-Path $_ })][String]$CSVPath,
	[Parameter(Mandatory=$true,Position=1,ParameterSetName="SECG")][ValidateScript({ Test-Path $_ })][String]$SecurityGuidancePath,
	[String[]]$PatchPath = "\\172.127.56.58\patch\�e��ՕW��",
	[String]$ProxyUrl,
	[String]$ProxyUser,
	[String]$ProxyPass,
	[String]$DownloadPath,
	[ValidateScript({ Test-Path $_ })][String]$OldCSVPath
)

# �C���N���[�h
Add-Type -AssemblyName System.Web
$libpath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "${libpath}\SharedLib.ps1"

# �����֐�
# ���͏��I�u�W�F�N�g����
Function New-BaseObject {
	Param([String]$kb, [String]$title, [String]$category, [String]$url)
	$obj = New-Object PSObject
	$obj | Add-Member -MemberType NoteProperty -Name kb -Value $kb
	$obj | Add-Member -MemberType NoteProperty -Name title -Value $title
	$obj | Add-Member -MemberType NoteProperty -Name category -Value $category
	$obj | Add-Member -MemberType NoteProperty -Name url -Value $url
	return $obj
}

# �����֐�
# �o�̓I�u�W�F�N�g����
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

# �����֐�
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

# �����֐�
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

# �����֐�
Function Get-LinkByMicrosoftUpdateCatalog {
	Param([String]$kb, [String]$title)
	$linkelements = @()
	if ($title) {
		# �^�C�g���Ō���
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
		# KB�Ō���
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
	# �_�E�����[�h�{�^���擾
	$inputelements = @([System.__ComObject].InvokeMember("getElementsByTagName", [System.Reflection.BindingFlags]::InvokeMethod, $null, $script:ie.Document, "input") | ? { $_.Value -eq "�_�E�����[�h" })
	if ($linkelements.Count -ne $inputelements.Count) {
		Write-Error "�����N�ƃ{�^���̐�������Ȃ��B$kb"
		return @()
	}
	for ($index = 0; $index -lt $linkelements.Count; $index++) {
		$linkelements[$index].element = $inputelements[$index]
	}
	return $linkelements
}

# �����֐�
Function Get-InfoByMicrosoftUpdateCatalog {
	Param([String]$id, [String]$kb, [ref]$refobj)
	# KB�����擾
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
		$obj.desc = "�_�E�����[�h���擾�G���["
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
			$obj.desc = "�_�E�����[�h���擾�G���["
			$obj.detail = "���Ȃ�"
		}
	}
}

# �����֐�
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

# �_�E�����[�h�ς݃��X�g �t�@�C�����F�p�X �p�X�͕����̏ꍇ�J���}��؂�
$files = @{}
Get-ChildItem -Path $PatchPath -Recurse | Where-Object { (-not $_.PSIsContainer) -and ($_.Name -notmatch "\.lnk$") } | ForEach-Object {
	if ($files.Contains($_.Name)) {
		$files[$_.Name] += ("," + $_.FullName)
	} else {
		$files.Add($_.Name, $_.FullName)
	}
}

# �O��쐬���X�g KB�ԍ��F�I�u�W�F�N�g�̃��X�g
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

# ���͏��쐬
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
		$kb = "KB" + $_.�L��
		$inputinfoarray += New-BaseObject -kb $kb -category $_.���i�t�@�~�� -url $_.�_�E�����[�h
	}
}

# ����pIE�I�u�W�F�N�g�擾
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $true

# ����p�V�F���I�u�W�F�N�g�擾
$shell = New-Object -ComObject "Shell.Application"

# �_�E�����[�h�p�I�u�W�F�N�g�擾
$webcient = Get-WebClient $ProxyUrl $ProxyUser $ProxyPass

# MS����p�b�`���E�t�@�C�����擾
$inputinfoarray | ForEach-Object {
	$obj = $_
	Write-Debug $obj.category
	if ($olddata.Contains($obj.kb)) {
		# �ߋ��Ɏ擾�ς݂̏ꍇ�͉ߋ��f�[�^���o��
		Write-Debug ("Matched" + $obj.kb)
		$olddata[$obj.kb]
	} else {
		if (-not $obj.title -and $obj.url) {
			# SECG�p�^�[�� (�^�C�g���Ȃ� URL����)
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
