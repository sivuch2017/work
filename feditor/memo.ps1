<#
.SYNOPSIS
�X�N���v�g��֐��̖����̊T�v�B

.DESCRIPTION
�X�N���v�g��֐��̖����̏ڍׂȐ����B

.PARAMETER Edit
����̃p�����[�^�[�̐����B<���O> ���p�����[�^�[���ɒu�������܂��B�X�N���v�g��֐��Ŏg�p����e�p�����[�^�[�ɂ��A1 �� .PARAMETER �Z�N�V�������g�p�ł��܂��B

.EXAMPLE
 �X�N���v�g��֐��̎g�p��B�����̗��񎦂���ꍇ�́A������ .EXAMPLE �Z�N�V�������g�p�ł��܂��B

.NOTES
�X�N���v�g��֐����g�p����ɂ������Ă̂��܂��܂Ȓ��߁B

.LINK
���̃w���v �g�s�b�N�Ƃ̑��ݎQ�ƁB���̃Z�N�V�����͕����z�u�ł��܂��Bhttp:// �� https:// �Ŏn�܂� URL ���܂߂�ƁAWindows PowerShell �ł́AHelp �R�}���h�� -online �p�����[�^�[���g�p���ꂽ�Ƃ��ɁA�w��� URL ���J���܂��B
#>
[CmdletBinding(DefaultParameterSetName="Normal")]
Param(
	[Parameter(Mandatory=$true,ParameterSetName="Edit")][switch]$Edit,
	[Parameter(ParameterSetName="Normal")][switch]$List,
	[Parameter(ParameterSetName="Normal")][string[]]$Keys = "",
	[Parameter(ParameterSetName="Normal")][string[]]$Find = "",
	[ValidateScript({ Test-Path $_ })][string]$FilePath = "D:\UserData\Documents\command.fip"
)

#�G�f�B�^�[�N��(�I��)
if ($Edit -eq $true) {
	& 'C:\Program Files\feditor\editor.exe' $FilePath
	exit
}

#�Z�N�V�����l�錾
[Hashtable]$section = @{
	"None" = 0;
	"Label" = 1;
	"CardData" = 2;
}

#�s���e����Z�N�V�����l�擾
Function GetSection {
	Param([string]$rec)
	switch -regex ($rec) {
		"^\[Label\]$" { $section.Label }
		"^\[CardData\]$" { $section.CardData }
		default { $section.None }
	}
}

#�L�[���[�h�}�b�`���O
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

#�����񌟍�
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

#�ϐ��錾
[Hashtable]$taghash = @{}
[Array]$namelist = @()
[Boolean]$dataflag = $false
[int]$nowSection = $section.None
[int]$nowLine = 0
[string]$title = ""
[Array]$words = @()
[Array]$data = @()
[string]$newline = ""

#�\�����
Get-Content -Path $FilePath | ForEach-Object {
	$rec = $_
	switch ($nowSection) {
		$section.Label {
			#���x���s����
			if ($rec -match "^([0-9]+)=.+Na([^,]+)$") {
				$taghash.Add($matches[1], $matches[2])
			}
			if ($rec -match "^\[") {
				$nowSection = GetSection($rec)
			}
		}
		$section.CardData {
			#�J�[�h�s����
			if ($nowLine -eq 0) {
				if ($rec -match "^([0-9]+)$") {
					#�J�[�h�J�n�s����
					$nowLine = [int]$matches[1]
					$title = ""
					$words = @()
					$data = @()
					$dataflag = $false
				}
			} else {
				if ($dataflag) {
					#�J�[�h���e
					$data += $rec
				} elseif ($rec -match "^Title:(.+)$") {
					#�J�[�h�^�C�g��
					$title = $matches[1]
					$namelist += $title
				} elseif ($rec -match "^Label:(.+)$") {
					#�J�[�h�ɕt�^����Ă��郉�x��
					$matches[1] -split "," | ForEach-Object {
						[int]$index = [int]$_ - 1
						$words += $taghash[[string]$index]
					}
				} elseif ($rec -eq "-") {
					#�J�[�h���s�̏I���
					$dataflag = $true
				}
				$nowLine -= 1
				if ($nowLine -eq 0 -and $List -eq $false) {
					#�J�[�h�ŏI�s����
					if ((CheckKey($title, $words)) -and (FindString($title, $words, $data))) {
						if ($words.Count -eq 0) {
							"${newline}�� $title"
						} else {
							"${newline}�� $title (" + ($words -join ",") + ")"
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

#���X�g�o�͂̏ꍇ
if ($List) {
	"�����x��"
	$taghash.Values | Where-Object { (CheckKey($_)) -and (FindString($_)) } | Sort-Object
	"`n���^�C�g��"
	$namelist | Where-Object { (CheckKey($_)) -and (FindString($_)) } | Sort-Object
}

### ������쐬��
Function Load-FrieveData {
	Param([string]$filepath)
	$result = @{}
	if (Test-Path -Path $filepath) {
		# Frieve�f�[�^�ǂݍ��ݗp�x�[�X�N���X
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
		# Global�N���X
		$FrieveGlobalClass = {
			&$FrieveBaseClass `
			| Add-Member -MemberType NoteProperty -Name Name -Force -Value "Global" -PassThru
		}
		# Card�N���X
		$FrieveCardClass = {
			&$FrieveBaseClass `
			| Add-Member -MemberType NoteProperty -Name Name -Force -Value "Card" -PassThru
		}
		# Link�N���X
		$FrieveLinkClass = {
			&$FrieveBaseClass `
			| Add-Member -MemberType NoteProperty -Name Name -Force -Value "Link" -PassThru
		}
		# Label�N���X
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
		# LinkLabel�N���X
		$FrieveLinkLabelClass = {
			&$FrieveBaseClass `
			| Add-Member -MemberType NoteProperty -Name Name -Force -Value "LinkLabel" -PassThru
		}
		# CardData�N���X
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
		Write-Host "��" $a.Title $hash["LABEL"].LabelNames($a.Labels)
	}
}
