#[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$url = "http://172.127.56.72:88/sra00346/4.html"
$shell = New-Object -ComObject Shell.Application
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $true
$hwnd = $ie.HWND
$ie.Navigate($url,4)
while ($ie.Document -isnot [mshtml.HTMLDocumentClass]) {
    $ie = $shell.Windows() | ? {$_.HWND -eq $hwnd}
}
while ($ie.busy -or $ie.readystate -ne 4) {
    Start-Sleep -Milliseconds 100
}
$ie.Document.getElementsByName("workDate[]") | ForEach-Object {
    # Write-Host $_.Value
    # $_.parentElement.childNodes | foreach-object -begin {$bln = $true} -process {
    #     if ($bln -and $_.TagName -eq "TD") {
    #         if ($_.innerText -match "[0-9.]+\(([0-9.]+)\)") {
    #             Write-Host $Matches[1]
    #             $bln = $false
    #         }
    #     }
    # }
}
