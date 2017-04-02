Param([string]$o = "dmy", [Parameter(Mandatory=$true,Position=2)][ValidateScript({Test-Path (".\" + (Get-Item $_).Name)})][string]$source)

Function MainProcess {
	Param([string]$o, [string]$source)
	if (-not (Test-Path ".\ps2exe.ps1")) {
		Write-Host "ps2exe.ps1 がありません。"
		return
	}
	if ($o -eq "dmy") {
		$o = (Get-Item $source).BaseName + ".exe"
	} else {
		$o = (Split-Path -Path $o -Leaf)
	}
	$f = (Get-Item $source).Name
	$work = (Get-Date).ToString("_yyyyMMddhhmmss.p\s1")
	SubProcess ".\$f" | Out-File -FilePath ".\$work"
	$i = (Get-Item $source).BaseName + ".ico"
	if (Test-Path ".\$i") {
		powershell -command "&'.\ps2exe.ps1' -inputFile $work -outputFile $o -x86 -runtime20 -noConsole -iconFile $i -verbose"
	} else {
		powershell -command "&'.\ps2exe.ps1' -inputFile $work -outputFile $o -x86 -runtime20 -noConsole -verbose"
	}
	Remove-Item $work
}

Function SubProcess {
	Param([string]$source)
	Get-Content $source | ForEach-Object {
		if ($_ -match "^\. +`"*([^`"]+)`"*$") {
			$f = (Split-Path -Leaf $Matches[1])
			if (($f -ne (Get-Item $source).Name) -and (Test-Path ".\$f")) {
				SubProcess ".\$f"
			} else {
				$_
			}
		} else {
			$_
		}
	}
}

MainProcess -o $o -source $source
