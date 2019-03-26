<#
.SYNOPSIS
.DESCRIPTION
.PARAMETER
.EXAMPLE
#>
Param([String]$Address, [String]$UserName, [String]$Password, [Switch]$AddService2 = $false)
$session = New-Object -ComObject Microsoft.Update.Session -Strict
if ($session) {
    if ($Address) { $session.WebProxy.Address = $Address }
    if ($UserName) { $session.WebProxy.UserName = $UserName }
    if ($Password) { $session.WebProxy.SetPassword($Password) }
} else {
    Write-Error "Microsoft.Update.Session使用不能"
}
if ($AddService2) {
    $manager = $session.CreateUpdateServiceManager()
    $manager.AddService2("7971F918-A847-4430-9279-4A52D1EFE18D",7, "")
} else {
    $searcher = $session.CreateUpdateSearcher()
    $search = $searcher.Search("Type='Software'")
    $search.Updates | Select-Object `
        @{Name = "KB"; Expression={"KB" + $_.KBArticleIDs}},
        AutoSelectOnWebSites,
        CanRequireSource,
        @{Name="Categories"; Expression={$_.Categories|ForEach-Object -Begin {$w = @()} -Process {$w += $_.Name} -End {$w -join ", "}}},
        DeltaCompressedContentAvailable,
        DeltaCompressedContentPreferred,
        DeploymentAction,Description,
        DownloadPriority,EulaAccepted,
        IsBeta,
        IsDownloaded,
        IsHidden,
        IsInstalled,
        IsMandatory,
        IsUninstallable,
        @{Name = "KBArticleIDs"; Expression={$_.KBArticleIDs}},
        @{Name = "LastDeploymentChangeTime"; Expression={$_.LastDeploymentChangeTime.ToString("yyyy/MM/dd")}},
        MaxDownloadSize,
        MinDownloadSize,
        @{Name = "MoreInfoUrls"; Expression={$_.MoreInfoUrls}},
        MsrcSeverity,
        RecommendedCPUSpeed,
        RecommendedHardDiskSpace,
        RecommendedMemory,
        @{Name = "SecurityBulletinIDs"; Expression={$_.SecurityBulletinIDs}},
        SupportUrl,
        Title,
        Type,
        @{Name = "UninstallationNotes"; Expression={$_.UninstallationNotes}}
}