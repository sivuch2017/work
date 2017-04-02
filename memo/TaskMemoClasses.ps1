<#
.SYNOPSIS
TaskMemoで使用するデータクラス

.DESCRIPTION
XML簡易制御用ベースクラス[$XmlBaseClass]
├データクラス[$TaskDataClass]
├統計データクラス[$StatisticalTaskDataClass]
├タスク名クラス[$TaskNameClass]
└タスクツリークラス[$TaskTreeClass]

.EXAMPLE
$obj = &$XmlBaseClass "C:\Work\Data.xml"
#>

# XML簡易制御用(疑似)ベースクラス
$XmlBaseClass = {
	Param([String]$filepath)
	if ([String]::IsNullOrEmpty($filepath)) {
		$xml = New-Object Xml
		[void]$xml.AppendChild($xml.CreateXmlDeclaration("1.0","UTF-8",$null))
		$root = $xml.CreateElement("root")
		[void]$xml.AppendChild($root)
	} else {
		if (Test-Path -Path $filepath) {
			$filepath = (Resolve-Path $filepath).Path
			$xml = [xml](Get-Content $filepath)
		} else {
			$parent = Split-Path $filepath -Parent
			if (Test-Path -Path $parent) {
				$filepath = Join-Path -Path (Resolve-Path $parent).Path -ChildPath (Split-Path $filepath -Leaf)
				$xml = New-Object Xml
				[void]$xml.AppendChild($xml.CreateXmlDeclaration("1.0","UTF-8",$null))
				$root = $xml.CreateElement("root")
				[void]$xml.AppendChild($root)
			} else {
				throw "指定されたパスの親ディレクトリが存在しません。"
			}
		}
	}
	$MyInvocation.MyCommand.ScriptBlock `
	| Add-Member -MemberType NoteProperty -Name FilePath -Force -Value $filepath -PassThru `
	| Add-Member -MemberType NoteProperty -Name Xml -Force -Value $xml -PassThru `
	| Add-Member -MemberType ScriptMethod -Name Commit -Value { $this.Xml.Save($this.FilePath) } -PassThru `
	| Add-Member -MemberType ScriptMethod -Name GetString -Value {} -PassThru `
	| Add-Member -MemberType ScriptMethod -Name AddRecord -Value {
		Param($child)
		$root = $this.Xml.DocumentElement
		[void]$root.AppendChild($child)
	} -PassThru
}

# データ(疑似)クラス
$TaskDataClass = {
	Param([string]$filepath)
	&$XmlBaseClass $filepath `
	| Add-Member -MemberType ScriptMethod -Name GetDayNode -Force -Value {
		Param($date)
		$result = $this.Xml.GetElementsByTagName("recs") | Where-Object {$_.date -eq $date} | Select-Object -First 1
		if ($result -isnot [System.Xml.XmlLinkedNode]) {
			$result = $null
		}
		return $result
	} -PassThru `
	| Add-Member -MemberType ScriptMethod -Name GetFirstDay -Force -Value {
		$node = $this.Xml.GetElementsByTagName("recs") | Sort-Object -Property date | Select-Object -First 1
		if ($node -is [System.Xml.XmlLinkedNode]) {
			$result = [System.DateTime]::ParseExact($node.date, "yyyyMMdd", $null)
		} else {
			$result = $null
		}
		return $result
	} -PassThru `
	| Add-Member -MemberType ScriptMethod -Name GetLastDay -Force -Value {
		$node = $this.Xml.GetElementsByTagName("recs") | Sort-Object -Property date -Descending | Select-Object -First 1
		if ($node -is [System.Xml.XmlLinkedNode]) {
			$result = [System.DateTime]::ParseExact($node.date, "yyyyMMdd", $null)
		} else {
			$result = $null
		}
		return $result
	} -PassThru `
	| Add-Member -MemberType ScriptMethod -Name GetString -Force -Value {
		Param($date, $mode)
		$node = $this.GetDayNode($date)
		if ($node -eq $null) {
			return ""
		}
		$result = @()
		if ($mode -eq "time") {
			$node.GetElementsByTagName("rec") | ForEach-Object -Begin { $seq = 0 } -Process {
				#同時間タスクがあった場合「読み込み順」に並び替える
				#本来はXML化する際に情報を入れることが正しいのだがXML構成の変更が必要なので暫定対応
				$obj = New-Object PSObject
				$obj | Add-Member -MemberType NoteProperty -Name time -Value $_.time
				$obj | Add-Member -MemberType NoteProperty -Name seq -Value $seq
				$obj | Add-Member -MemberType NoteProperty -Name InnerText -Value $_.InnerText
				$obj
				$seq += 1
			} | Sort-Object -Property @{Expression="time";Ascending=$true},@{Expression="seq";Ascending=$true} | ForEach-Object -Begin { [System.DateTime]$pretime = 0 } -Process {
				if ($pretime -eq 0) {
					$result += [string]::Format("{0}    - {1}", $_.time, $_.InnerText)
					$pretime = [System.DateTime]($_.time)
				} else {
					$time = [System.DateTime]($_.time)
					$result += [string]::Format("{0} {1,4} {2}", $_.time, [string]($time - $pretime).TotalMinutes, $_.InnerText)
					$pretime = $time
				}
			}
		} elseif ($mode -eq "task") {
			$node.GetElementsByTagName("rec") | Sort-Object time | ForEach-Object -Begin { [System.DateTime]$pretime = 0; $hash = @{}; $total = 0; $difftotal = 0} -Process {
				if ($pretime -eq 0) {
					$pretime = [System.DateTime]($_.time)
				} else {
					$time = [System.DateTime]($_.time)
					if ($hash.ContainsKey($_.InnerText)) {
						$hash[$_.InnerText] += ($time - $pretime).TotalMinutes
					} else {
						$hash.Add($_.InnerText, ($time - $pretime).TotalMinutes)
					}
					$pretime = $time
				}
			}
			$first = ($node.GetElementsByTagName("rec") | Sort-Object time | Select-Object -First 1).time
			$last = ($node.GetElementsByTagName("rec") | Sort-Object time | Select-Object -Last 1).time
			$result += [string]::Format("開始{0} 終了{1} 時間{2,4:F2}h", $first, $last, (([System.DateTime]$last - [System.DateTime]$first).TotalMinutes / 60))
			$result += ""
			$result += "分    時     差     タスク"
			$result += "----- ------ ------ ------"
			foreach ($key in $hash.Keys) {
				$unit = [int]$script:conf["UNIT"]
				$total += [int]($hash[$key] / 60 * $unit) / $unit
				$difftotal += $hash[$key] / 60 - [int]($hash[$key] / 60 * $unit) / $unit
				$result += [string]::Format("{0,4}m {1,5:F2}h {2,5:F2}h {3}", $hash[$key], ([int]($hash[$key] / 60 * $unit) / $unit), ($hash[$key] / 60 - [int]($hash[$key] / 60 * $unit) / $unit), $key)
			}
			$result += "----- ------ ------ ------"
			$result += [string]::Format("計    {0,5:F2}h {1,5:F2}h", $total, $difftotal)
		} else {
			$node.GetElementsByTagName("rec") | Sort-Object time | ForEach-Object { $result += $_.time + " " + $_.InnerText }
		}
		return ($result -join "`n")
	} -PassThru `
	| Add-Member -MemberType ScriptMethod -Name AddDay -Force -Value {
		Param($date)
		$node = $this.GetDayNode($date)
		if ($node -eq $null) {
			$node = $this.Xml.CreateElement("recs")
			$atr = $this.Xml.CreateAttribute("date")
			$atr.Value = $date
			[void]$node.Attributes.Append($atr)
			$this.AddRecord($node)
		}
		return $node
	} -PassThru `
	| Add-Member -MemberType ScriptMethod -Name AddTask -Force -Value {
		Param($date, $time, $task)
		$node = $this.AddDay($date)
		$rec = $this.Xml.CreateElement("rec")
		$atr = $this.Xml.CreateAttribute("time")
		$atr.Value = $time
		[void]$rec.Attributes.Append($atr)
		$rec.InnerText = $task
		[void]$node.AppendChild($rec)
	} -PassThru `
	| Add-Member -MemberType ScriptMethod -Name LastTask -Force -Value {
		Param($date, $time)
		$node = $this.GetDayNode($date)
		if ($node -eq $null) {
			return $null
		}
		$result = $node.GetElementsByTagName("rec") | ForEach-Object -Begin { $seq = 0 } -Process {
			#同時間タスクがあった場合「読み込み順」に並び替える
			#本来はXML化する際に情報を入れることが正しいのだがXML構成の変更が必要なので暫定対応
			$obj = New-Object PSObject
			$obj | Add-Member -MemberType NoteProperty -Name time -Value $_.time
			$obj | Add-Member -MemberType NoteProperty -Name seq -Value $seq
			$obj | Add-Member -MemberType NoteProperty -Name InnerText -Value $_.InnerText
			$obj
			$seq += 1
		} | Sort-Object -Property @{Expression="time";Ascending=$true},@{Expression="seq";Ascending=$true} | Select-Object -Last 1
		return $result
	} -PassThru `
	| Add-Member -MemberType ScriptMethod -Name GetTasksTime -Force -Value {
		Param($date, $toxml = $false)
		$result = @()
		if ($toxml) {
			$result = &$StatisticalTaskDataClass
		}
		$node = $this.GetDayNode($date)
		if ($node -eq $null) {
			return $result
		}
		$node.GetElementsByTagName("rec") | Sort-Object time | ForEach-Object -Begin { [System.DateTime]$pretime = 0; $hash = @{}; $total = 0; $difftotal = 0} -Process {
			if ($pretime -eq 0) {
				$pretime = [System.DateTime]($_.time)
			} else {
				$time = [System.DateTime]($_.time)
				if ($hash.ContainsKey($_.InnerText)) {
					$hash[$_.InnerText] += ($time - $pretime).TotalMinutes
				} else {
					$hash.Add($_.InnerText, ($time - $pretime).TotalMinutes)
				}
				$pretime = $time
			}
		}
		$first = ($node.GetElementsByTagName("rec") | Sort-Object time | Select-Object -First 1).time
		$last = ($node.GetElementsByTagName("rec") | Sort-Object time | Select-Object -Last 1).time
		foreach ($key in $hash.Keys) {
			$unit = [int]$script:conf["UNIT"]
			$total += [int]($hash[$key] / 60 * $unit) / $unit
			$difftotal += $hash[$key] / 60 - [int]($hash[$key] / 60 * $unit) / $unit
			if ($toxml) {
				$result.AddTask($key, ([int]($hash[$key] / 60 * $unit) / $unit))
			} else {
				$obj = New-Object PSObject
				$obj | Add-Member -MemberType NoteProperty -Name name -Value $key
				$obj | Add-Member -MemberType NoteProperty -Name minuts -Value $hash[$key]
				$obj | Add-Member -MemberType NoteProperty -Name diff -Value ($hash[$key] / 60 - [int]($hash[$key] / 60 * $unit) / $unit)
				$result += $obj
			}
		}
		return $result
	} -PassThru
}

# 統計データクラス
$StatisticalTaskDataClass = {
	Param([string]$filepath)
	&$XmlBaseClass $filepath `
	| Add-Member -MemberType ScriptProperty -Name Count -Force -Value {
		$this.Xml.GetElementsByTagName("rec").Count
	} -PassThru `
	| Add-Member -MemberType ScriptMethod -Name AddTask -Force -Value {
		Param($name, $time = "", $progress = "", $result = "未")
		$rec = $this.Xml.CreateElement("rec")
		$atrA = $this.Xml.CreateAttribute("name")
		$atrA.Value = $name
		[void]$rec.Attributes.Append($atrA)
		$atrB = $this.Xml.CreateAttribute("time")
		$atrB.Value = $time
		[void]$rec.Attributes.Append($atrB)
		$atrC = $this.Xml.CreateAttribute("progress")
		$atrC.Value = $progress
		[void]$rec.Attributes.Append($atrC)
		$atrD = $this.Xml.CreateAttribute("result")
		$atrD.Value = $result
		[void]$rec.Attributes.Append($atrD)
		$this.AddRecord($rec)
	} -PassThru
}

# タスク名(疑似)クラス
$TaskNameClass = {
	Param([string]$filepath)
	&$XmlBaseClass $filepath `
	| Add-Member -MemberType ScriptMethod -Name AddName -Force -Value {
		Param($name)
		$exist = $this.Xml.GetElementsByTagName("task") | ForEach-Object -Begin {$result = $false} -Process {if ($_.InnerText -eq $name) {$result = $true}} -End {return $result}
		if ($exist -eq $false) {
			$rec = $this.Xml.CreateElement("task")
			$rec.InnerText = $name
			$this.AddRecord($rec)
		}
	} -PassThru `
	| Add-Member -MemberType ScriptMethod -Name GetNames -Force -Value {
		$ary = @()
		$this.Xml.GetElementsByTagName("task") | ForEach-Object { $ary += $_.InnerText }
		if ($ary.Count -gt 0) {
			$ary = $ary | Sort-Object | Get-Unique
		}
		return $ary
	} -PassThru
}

# タスクツリー(疑似)クラス
$TaskTreeClass = {
	Param([string]$filepath)
	&$XmlBaseClass $filepath `
	| Add-Member -MemberType ScriptMethod -Name GetNode -Force -Value {
		$top = @()
		$hash = @{}
		$this.Xml.GetElementsByTagName("task") | ForEach-Object {
			if ($_.id) {
				if ($hash.ContainsKey($_.id)) {
					$hash[$_.id].Name($_.InnerText)
				} else {
					$hash.Add($_.id, (New-Object System.Windows.Forms.TreeNode($_.InnerText)))
				}
				if ($_.parent) {
					if (-not $hash.ContainsKey($_.parent)) {
						$hash.Add($_.parent, (New-Object System.Windows.Forms.TreeNode("仮")))
					}
					$hash[$_.parent].Nodes.Add($hash[$_.id])
				} else{
					$top += $hash[$_.id]
				}
			} else {
				#idが設定されていない場合は無視
			}
		}
		if ($top.Count -eq 0) {
			$top += New-Object System.Windows.Forms.TreeNode("空")
		}
		return $top
	} -PassThru `
	| Add-Member -MemberType ScriptMethod -Name AddName -Force -Value {
		Param([String]$name, [String]$id, [String]$parent)
		$rec = $this.Xml.CreateElement("task")
		if (-not [String]::IsNullOrEmpty($id)) {
			$idatr = $this.Xml.CreateAttribute("id")
			$idatr.Value = $id
			[void]$rec.Attributes.Append($idatr)
		}
		if (-not [String]::IsNullOrEmpty($parent)) {
			$patr = $this.Xml.CreateAttribute("parent")
			$patr.Value = $parent
			[void]$rec.Attributes.Append($patr)
		}
		$rec.InnerText = $name
		$this.AddRecord($rec)
	} -PassThru
}
