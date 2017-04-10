<#
.SYNOPSIS
カスタムクラス、Microsoft .NET Framework または COM オブジェクトのインスタンスを作成します。

.DESCRIPTION
New-LocalObject は、カスタムクラス、.NET Framework または COM オブジェクトのインスタンスを作成します。
カスタムクラスの型、.NET Framework クラスの型または COM オブジェクトの ProgID を指定できます。
既定では、完全修飾されたカスタムクラスの型、.NET Framework クラスの名前を入力とし、そのクラスのインスタンスへの参照を返します。
COM オブジェクトのインスタンスを作成するには、ComObject パラメーターの値にオブジェクトの ProgID を指定します。

.PARAMETER ArgumentList
カスタムクラス、.NET Framework クラスのコンストラクターに渡す引数の一覧を指定します。引数の一覧の各要素はコンマ (,) で区切ります。

.PARAMETER ComObject
COM オブジェクトのプログラム識別子 (ProgID) を指定します。

.PARAMETER Property
プロパティ値を設定し、新しいオブジェクトのメソッドを呼び出します。
キーがプロパティ名またはメソッド名で、値がプロパティ値またはメソッドの引数であるハッシュ テーブルを入力します。
New-LocalObject は、オブジェクトを作成し、ハッシュ テーブルで出現する順に各プロパティ値を設定し、各メソッドを呼び出します。
新しいオブジェクトが PSObject クラスから派生しているときに、そのオブジェクトに存在しないプロパティを指定した場合、
New-LocalObject は指定されたプロパティを NoteProperty としてオブジェクトに追加します。
オブジェクトが PSObject ではない場合、コマンドによって未終了エラーが生成されます。

.PARAMETER Strict
作成しようとしている COM オブジェクトが相互運用機能アセンブリを使用すると、エラーになるよう指定します。
これにより、COM 呼び出しが可能なラッパーを備えた .NET Framework オブジェクトと実際の COM オブジェクトとを区別できます。

.PARAMETER TypeName
カスタムクラス、.NET Framework クラスの完全修飾名を指定します。TypeName パラメーターと ComObject パラメーターの両方を指定することはできません。
#>
Function New-LocalObject {
	[CmdletBinding(DefaultParameterSetName="_TypeName")]
	[OutputType([Object])]
	Param(
		[Parameter(Mandatory=$true,Position=0,ParameterSetName="_TypeName")][String]$TypeName,
		[Parameter(Position=1,ParameterSetName="_TypeName")][Object[]]$ArgumentList,
		[Parameter(Mandatory=$true,Position=0,ParameterSetName="_ComObject")][String]$ComObject,
		[Parameter(ParameterSetName="_ComObject")][Switch]$Strict,
		[System.Collections.IDictionary]$Property
	)
	if ([String]::IsNullOrEmpty($TypeName)) {
		if ($Property) {
			return New-Object -ComObject $ComObject -Strict:$Strict -Property $Property
		} else {
			return New-Object -ComObject $ComObject -Strict:$Strict
		}
	} else {
		switch -regex ($TypeName) {
			"^()$" {
			}
			default {
				if ($Property) {
					return New-Object -TypeName $TypeName -ArgumentList $ArgumentList -Property $Property
				} else {
					return New-Object -TypeName $TypeName -ArgumentList $ArgumentList
				}
			}
		}
	}
}
