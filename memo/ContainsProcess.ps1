<#
.SYNOPSIS
指定アプリケーションが起動しているかを返す。

.DESCRIPTION
指定されたプロセスID「以外」で、同一のプロセス名のものが存在するか確認する。
存在した場合、そのプロセスを前面表示し、真を返す。
ただしプロセス名がPowerShellだった場合は、ウィンドウタイトルで判断し存在した場合は、確認ダイアログを開く。
「無視しない」を選択した場合のみ真を返す。

.PARAMETER id
除外プロセスID
省略した場合は自分自身のプロセスID

.PARAMETER name
プロセス名
省略した場合は除外プロセスIDのプロセス名

.PARAMETER title
ウィンドウタイトル
PowerShellの場合だけ比較に使用される
省略した場合は除外プロセスIDのウィンドウタイトル
#>

# .NET Framework クラス宣言
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

Function ContainsProcess {
	Param([Parameter(Position=1)][Int32]$id = $PID, [Parameter(Position=2)][String]$name, [Parameter(Position=3)][String]$title)

	# WIN32API制御関数
	# http://d.hatena.ne.jp/ps1/20080312/1205339720
	Function New-PMethod([string]$dll_name, [Type]$return_type, [string]$method_name, [Type[]]$parameter_types) {
		$ptype = [AppDomain]::CurrentDomain.DefineDynamicAssembly((New-Object Reflection.AssemblyName 'PInvokeAssembly'), 'Run').DefineDynamicModule('PInvokeModule').DefineType('PInvokeType', "Public,BeforeFieldInit")
		$ptype.DefineMethod($method_name, 'Public,HideBySig,Static,PinvokeImpl', $return_type, $parameter_types).SetCustomAttribute((New-Object Reflection.Emit.CustomAttributeBuilder ([Runtime.InteropServices.DllImportAttribute].GetConstructor([string])), $dll_name))
		$invoke  = {
			for ($i = 0; $i -lt $this.ParameterTypes.Length; $i++) { $args[$i] = $args[$i] -as $this.ParameterTypes[$i] }
			return $this.InvokeMember($this.MethodName, 'Public,Static,InvokeMethod', $null, $null, $args)
		}
		$api = $ptype.CreateType() 
		$api =
			Add-Member -InputObject $api -MemberType NoteProperty -Name MethodName     -Value $method_name -passthru
			Add-Member -InputObject $api -MemberType NoteProperty -Name ParameterTypes -Value $parameter_types
			Add-Member -InputObject $api -MemberType ScriptMethod -Name Invoke         -Value $invoke
		return $api
	}

	# 多重起動チェック
	$result = $false
	if ([String]::IsNullOrEmpty($name)) {
		$name = (Get-Process -Id $id).ProcessName
	}
	if ($name -ne "powershell") {
		$proc = Get-Process | Where-Object {$_.Id -ne $id -and $_.ProcessName -eq $name} | Select-Object -First 1
		if ($proc) {
			# 起動済みプロセスを前面に表示
			$ShowWindow = New-PMethod "user32.dll" ([Bool]) "ShowWindow" @([IntPtr], [Int32])
			$SetForegroundWindows = New-PMethod "user32.dll" ([Bool]) "SetForegroundWindow" @([IntPtr])
			[void]$ShowWindow.Invoke($proc.MainWindowHandle, 5)
			[void]$SetForegroundWindows.Invoke($proc.MainWindowHandle)
			$result = $true
		}
	} else {
		if ([String]::IsNullOrEmpty($title)) {
			$title = (Get-Process -Id $id).MainWindowTitle
		}
		$proc = Get-Process | Where-Object {$_.Id -ne $id -and $_.ProcessName -eq $name -and $_.MainWindowTitle -eq $title} | Select-Object -First 1
		if ($proc) {
			if ([System.Windows.Forms.MessageBox]::Show("多重起動の可能性があります。`n無視しますか？", "多重起動チェック", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information) -eq [System.Windows.Forms.DialogResult]::No) {
				$result = $true
			}
		}
	}
	return $result

}
