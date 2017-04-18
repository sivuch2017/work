Function Load-FrieveData {
	Param([String]$filepath)
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
	} else {
		throw New-Object System.IO.FileNotFoundException
	}
	return $result
}
