<#
.SYNOPSIS
スクリプトや関数の役割の概要。

.DESCRIPTION
スクリプトや関数の役割の詳細な説明。

.PARAMETER Edit
特定のパラメーターの説明。<名前> をパラメーター名に置き換えます。スクリプトや関数で使用する各パラメーターにつき、1 つの .PARAMETER セクションを使用できます。

.EXAMPLE
 スクリプトや関数の使用例。複数の例を提示する場合は、複数の .EXAMPLE セクションを使用できます。

.NOTES
スクリプトや関数を使用するにあたってのさまざまな注釈。

.LINK
他のヘルプ トピックとの相互参照。このセクションは複数配置できます。http:// や https:// で始まる URL を含めると、Windows PowerShell では、Help コマンドの -online パラメーターが使用されたときに、指定の URL を開きます。
#>
[CmdletBinding(DefaultParameterSetName="Normal")]
Param(
	[Parameter(Mandatory=$true,ParameterSetName="Edit")][switch]$Edit,
	[Parameter(ParameterSetName="Normal")][switch]$List,
	[Parameter(ParameterSetName="Normal")][string[]]$Keys = "",
	[Parameter(ParameterSetName="Normal")][string[]]$Find = "",
	[ValidateScript({ Test-Path $_ })][string]$FilePath = "D:\UserData\Documents\command.fip"
)

#エディター起動(終了)
if ($Edit -eq $true) {
	& 'C:\Program Files\feditor\editor.exe' $FilePath
	exit
}

#セクション値宣言
[Hashtable]$section = @{
	"None" = 0;
	"Label" = 1;
	"CardData" = 2;
}

#行内容からセクション値取得
Function GetSection {
	Param([string]$rec)
	switch -regex ($rec) {
		"^\[Label\]$" { $section.Label }
		"^\[CardData\]$" { $section.CardData }
		default { $section.None }
	}
}

#キーワードマッチング
Function CheckKey {
	Param([string[]]$words)
	if ($Keys.Count -eq 1 -and $Keys[0] -eq "") {
		return $true
	} else {
		foreach ($word in $words) {
			foreach ($key in $Keys) {
				if ($word -ieq $key) {
					return $true
				}
			}
		}
		return $false
	}
}

#文字列検索
Function FindString {
	Param([string[]]$words)
	if ($Find.Count -eq 1 -and $Find[0] -eq "") {
		return $true
	} else {
		foreach ($word in $words) {
			foreach ($str in $Find) {
				if ($word -match $str) {
					return $true
				}
			}
		}
		return $false
	}
}

#変数宣言
[Hashtable]$taghash = @{}
[Array]$namelist = @()
[Boolean]$dataflag = $false
[int]$nowSection = $section.None
[int]$nowLine = 0
[string]$title = ""
[Array]$words = @()
[Array]$data = @()
[string]$newline = ""

#構文解析
Get-Content -Path $FilePath | ForEach-Object {
	$rec = $_
	switch ($nowSection) {
		$section.Label {
			#ラベル行処理
			if ($rec -match "^([0-9]+)=.+Na([^,]+)$") {
				$taghash.Add($matches[1], $matches[2])
			}
			if ($rec -match "^\[") {
				$nowSection = GetSection($rec)
			}
		}
		$section.CardData {
			#カード行処理
			if ($nowLine -eq 0) {
				if ($rec -match "^([0-9]+)$") {
					#カード開始行処理
					$nowLine = [int]$matches[1]
					$title = ""
					$words = @()
					$data = @()
					$dataflag = $false
				}
			} else {
				if ($dataflag) {
					#カード内容
					$data += $rec
				} elseif ($rec -match "^Title:(.+)$") {
					#カードタイトル
					$title = $matches[1]
					$namelist += $title
				} elseif ($rec -match "^Label:(.+)$") {
					#カードに付与されているラベル
					$matches[1] -split "," | ForEach-Object {
						[int]$index = [int]$_ - 1
						$words += $taghash[[string]$index]
					}
				} elseif ($rec -eq "-") {
					#カード情報行の終わり
					$dataflag = $true
				}
				$nowLine -= 1
				if ($nowLine -eq 0 -and $List -eq $false) {
					#カード最終行処理
					if ((CheckKey($title, $words)) -and (FindString($title, $words, $data))) {
						if ($words.Count -eq 0) {
							"${newline}■ $title"
						} else {
							"${newline}■ $title (" + ($words -join ",") + ")"
						}
						$newline = "`n"
						$data -join "`n"
					}
				}
			}
		}
		default {
			$nowSection = GetSection($rec)
		}
	}
}

#リスト出力の場合
if ($List) {
	"■ラベル"
	$taghash.Values | Where-Object { (CheckKey($_)) -and (FindString($_)) } | Sort-Object
	"`n■タイトル"
	$namelist | Where-Object { (CheckKey($_)) -and (FindString($_)) } | Sort-Object
}

### 次世代作成中
Function Load-FrieveData {
	Param([string]$filepath)
	$result = @{}
	if (Test-Path -Path $filepath) {
		# Frieveデータ読み込み用ベースクラス
		$FrieveBaseClass = {
			$frievetext = ""
			$MyInvocation.MyCommand.ScriptBlock `
			| Add-Member -MemberType NoteProperty -Name Name -Force -Value "Base" -PassThru `
			| Add-Member -MemberType NoteProperty -Name Text -Force -Value $frievetext -PassThru `
			| Add-Member -MemberType ScriptMethod -Name AddText -Value {
				Param($addtext)
				$frievetext += "${addtext}`r`n"
			} -PassThru
		}
		# Globalクラス
		$FrieveGlobalClass = {
			&$FrieveBaseClass `
			| Add-Member -MemberType NoteProperty -Name Name -Force -Value "Global" -PassThru
		}
		# Cardクラス
		$FrieveCardClass = {
			&$FrieveBaseClass `
			| Add-Member -MemberType NoteProperty -Name Name -Force -Value "Card" -PassThru
		}
		# Linkクラス
		$FrieveLinkClass = {
			&$FrieveBaseClass `
			| Add-Member -MemberType NoteProperty -Name Name -Force -Value "Link" -PassThru
		}
		# Labelクラス
		$FrieveLabelClass = {
			$labelhash = @{}
			&$FrieveBaseClass `
			| Add-Member -MemberType NoteProperty -Name Name -Force -Value "Label" -PassThru `
			| Add-Member -MemberType NoteProperty -Name LabelHash -Force -Value $labelhash -PassThru `
			| Add-Member -MemberType ScriptMethod -Name AddText -Force -Value {
				Param($addtext)
				$this.Text += "${addtext}`r`n"
				if ($addtext -match "^([0-9]+)=.*,Na(.*)$") {
					$this.LabelHash.Add([string]([int]($matches[1]) + 1), $matches[2])
				}
			} -PassThru `
			| Add-Member -MemberType ScriptMethod -Name LabelNames -Force -Value {
				Param([array]$indexlist)
				$ary = @()
				foreach ($index in $indexlist) {
					if ($this.LabelHash.ContainsKey($index)) {
						$ary += $this.LabelHash[$index]
					}
				}
				if ($ary.Count -gt 0) {
					return "(" + ($ary -join ",") + ")"
				} else {
					return ""
				}
			} -PassThru
		}
		# LinkLabelクラス
		$FrieveLinkLabelClass = {
			&$FrieveBaseClass `
			| Add-Member -MemberType NoteProperty -Name Name -Force -Value "LinkLabel" -PassThru
		}
		# CardDataクラス
		$FrieveCardDataClass = {
			$array = @()
			$lineno = 0
			$frievetext = ""
			$MyInvocation.MyCommand.ScriptBlock `
			| Add-Member -MemberType NoteProperty -Name Text -Force -Value $frievetext -PassThru `
			| Add-Member -MemberType NoteProperty -Name Name -Force -Value "CardData" -PassThru `
			| Add-Member -MemberType ScriptMethod -Name AddText -Force -Value {
				Param($addtext)
				$frievetext += "${addtext}`r`n"
				if (($this.Line -le 0) -and ($addtext -match "^[0-9]+$")) {
					$this.Line = [int]$addtext
					$o = $this.CreateItem()
					$this.Values += $o
				} else {
					$o = $this.Values
					$o[-1].AddText($addtext)
					$this.Line = $this.Line - 1
				}
			} -PassThru `
			| Add-Member -MemberType NoteProperty -Name Line -Force -Value $lineno -PassThru `
			| Add-Member -MemberType NoteProperty -Name Values -Force -Value $array -PassThru `
			| Add-Member -MemberType ScriptMethod -Name CreateItem -Force -Value {
				return &{
					$alltext = ""
					$cardinfoflag = $true
					$cardtitle = ""
					$cardlabels = @()
					$cardtext = ""
					$MyInvocation.MyCommand.ScriptBlock `
					| Add-Member -MemberType NoteProperty -Name Name -Force -Value "CardDataItem" -PassThru `
					| Add-Member -MemberType NoteProperty -Name Text -Force -Value $alltext -PassThru `
					| Add-Member -MemberType NoteProperty -Name CardInfoFlag -Force -Value $cardinfoflag -PassThru `
					| Add-Member -MemberType NoteProperty -Name Title -Force -Value $cardtitle -PassThru `
					| Add-Member -MemberType NoteProperty -Name Labels -Force -Value $cardlabels -PassThru `
					| Add-Member -MemberType NoteProperty -Name CardText -Force -Value $cardtext -PassThru `
					| Add-Member -MemberType ScriptMethod -Name AddText -Force -Value {
						Param($addtext)
						$this.Text += "${addtext}`r`n"
						if ($this.CardInfoFlag -and $addtext -match "^Title:(.+)$") {
							$this.Title = $matches[1]
						} elseif ($this.CardInfoFlag -and $addtext -match "^Label:(.+)$") {
							$this.Labels = $matches[1] -split ","
						} elseif ($this.CardInfoFlag -and $addtext -eq "-") {
							$this.CardInfoFlag = $false
						} elseif ($this.CardInfoFlag -eq $false) {
							$this.CardText += "${addtext}`r`n"
						}
					} -PassThru
				}
			} -PassThru
		}

		$global = &$FrieveGlobalClass
		$card = &$FrieveCardClass
		$link = &$FrieveLinkClass
		$label = &$FrieveLabelClass
		$linklabel = &$FrieveLinkLabelClass
		$carddata = &$FrieveCardDataClass

		Get-Content $filepath | ForEach-Object -Begin {$infoflag = $true} -Process {
			if ($infoflag -and $_ -eq "[Global]") {
				$obj = $global
			} elseif ($infoflag -and $_ -eq "[Card]") {
				$obj = $card
			} elseif ($infoflag -and $_ -eq "[Link]") {
				$obj = $link
			} elseif ($infoflag -and $_ -eq "[Label]") {
				$obj = $label
			} elseif ($infoflag -and $_ -eq "[LinkLabel]") {
				$obj = $linklabel
			} elseif ($infoflag -and $_ -eq "[CardData]") {
				$obj = $carddata
				$infoflag = $false
			} else {
				$obj.AddText($_)
			}
		}
		$result = @{GLOBAL = $global; CARD = $card; LINK = $link; LABEL = $label; LINKLABEL = $linklabel; CARDDATA = $carddata}
	}
	return $result
}

Function Test-Load-FrieveData {
	$hash = Load-FrieveData "D:\UserData\Documents\net.fip"
	foreach ($a in $hash["CARDDATA"].Values) {
		Write-Host "■" $a.Title $hash["LABEL"].LabelNames($a.Labels)
	}
}
