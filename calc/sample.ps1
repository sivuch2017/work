$BaseDir = Split-Path $MyInvocation.MyCommand.path
$EnvFile = Join-Path $BaseDir "env.ps1"
. $EnvFile

$script:eigyo = 0
$script:shell = New-Object -ComObject Shell.Application

Function execMain {
    Param([String]$url, $user, $pass, $outfile)
    
    $ie = New-Object -ComObject InternetExplorer.Application
    $script:hwnd = $ie.HWND

    $ie.Visible = $true
    $ie.Navigate($url, 4)
    waitBusy([ref]$ie)

    @($ie.Document.getElementsByName("loginId"))[0].value = $user
    @($ie.Document.getElementsByName("passwd"))[0].value = $pass
    $ie.Document.getElementById("login_button").click()
    waitBusy([ref]$ie)

    @(@($ie.Document.getElementsByTagName("table"))[3].getElementsByTagName("a"))[2].click()
    waitBusy([ref]$ie)

    @($ie.Document.getElementsByTagName("table"))[7].getElementsByTagName("tr") | ForEach-Object {
        $td = $_.getElementsByTagName("td") | %{$_.innerText.ToString()}
        if ($td) {
            if ($td[0] -match ".+\(([^)]+)\)") {
                $td[0] = $Matches[1] + "`t"
            }
            $td -join "," | Out-File -FilePath $outfile -Encoding default -Append
        }
    }

    if ($eigyo -eq 0) {
        $a = @(@($ie.Document.getElementsByTagName("table"))[7].getElementsByTagName("a"))[0]
        $a.target = "_self"
        $a.click()
        waitBusy([ref]$ie)

        @($ie.Document.getElementsByTagName("table"))[4].getElementsByTagName("tr") | foreach-object -begin {$c = 0; $p = 0} -process {if ($_.outerHTML -match "workhourtable") {$c++; if ($_.outerHTML -match "pink") {$p++};}} -end {$eigyo = $c - $p}
    }
    write-host "ç°åéÇÃâcã∆ì˙ÇÕ${eigyo}ì˙"
    $ie.Quit()
}

Function waitBusy {
    Param([ref]$ie)
    $ie.value = @($shell.Windows() | ? {$_.HWND -eq $hwnd})[-1]
    while ($ie.value.busy -or $ie.value.readystate -ne 4) {
        Start-Sleep -Milliseconds 100
    }
}

# Main
foreach ($key in $users.Keys) {
    execMain $topurl $users[$key][0] $users[$key][1] $csvfile
}
