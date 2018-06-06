[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.mshtml")

$url = "http://172.127.56.72:88/sra00346/1.html"
$user = "2817"
$pass = "2817"

$script:shell = New-Object -ComObject Shell.Application
$ie = New-Object -ComObject InternetExplorer.Application
$script:hwnd = $ie.HWND

Function waitBusy {
    Param([ref]$ie)
    while ($ie.value.Document -isnot [mshtml.HTMLDocumentClass]) {
        $ie.value = @($shell.Windows() | ? {$_.HWND -eq $hwnd})[-1]
    }
    while ($ie.value.busy -or $ie.value.readystate -ne 4) {
        Start-Sleep -Milliseconds 100
    }
}

$ie.Visible = $true
$ie.Navigate($url, 4)
waitBusy([ref]$ie)

@($ie.Document.getElementsByName("loginId"))[0].value = $user
@($ie.Document.getElementsByName("passwd"))[0].value = $pass
$ie.Document.getElementById("login_button").click()
waitBusy([ref]$ie)

# getElementsByClassNamemesoddo 
# 漢字コードによってマッチングできないためAタグ3番目決め打ち
# 環境によっては getElementsByClassName が使えない
# @(@([System.__ComObject].InvokeMember("getElementsByClassName", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "box"))[0].getElementsByTagName("a"))[2].click()
# @($ie.Document.getElementsByTagName("table") | Where-Object { $_.class -eq "box" })[0].@($_.getElementsByTagName("a"))[2].click()
@(@($ie.Document.getElementsByTagName("table") | Where-Object { $_.getAttributeNode("class").textContent -eq "box" })[0].getElementsByTagName("a"))[2].click()
waitBusy([ref]$ie)

@($ie.Document.getElementsByTagName("table") | Where-Object { $_.getAttributeNode("class").textContent -eq "box" })[2].getElementsByTagName("a") | ForEach-Object {
    $_.click()
    $ie2 = @($shell.Windows() | ? {$_.HWND -eq $hwnd})[-1]
    waitBusy([ref]$ie2)
    $ie2.Document.getElementsByName("workDate[]") | ForEach-Object {
        Write-Host $_.Value $_.parentElement.childNodes.Item(4).innerText
    }
}

