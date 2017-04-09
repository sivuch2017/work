<#
#>
Function New-LocalObject {
    [CmdletBinding(DefaultParameterSetName="TypeName")]
    [OutputType([Object])]
    Param(
        [Parameter(Mandatory=$true,Position=0,ParameterSetName="TypeName")][String]$TypeName,
        [Parameter(Position=1,ParameterSetName="TypeName")][Object[]]$ArgumentList,
        [Parameter(Mandatory=$true,Position=0,ParameterSetName="ComObject")][String]$ComObject,
        [Parameter(ParameterSetName="ComObject")][Switch]$Strict,
        [System.Collections.IDictionary]$Property
    )
    if ([String]::IsNullOrEmpty($TypeName)) {
        return New-Object -ComObject $ComObject -Strict:$Strict -Property $Property
    } else {
        switch ($TypeName) {
            default {
                return New-Object -TypeName $TypeName -ArgumentList $ArgumentList -Property $Property
            }
        }
    }
}