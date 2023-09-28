[CmdletBinding(DefaultParameterSetName = 'Credential+FileExtension')]
param (
    [Parameter(Mandatory, Position = 0)]
    [String[]]
    $SiteURL,

    [Parameter( Mandatory, ParameterSetName = 'Credential+FileExtension' )]
    [pscredential]
    $Credential,

    [Parameter( Mandatory, ParameterSetName = 'Certificate+FileExtension' )]
    [String]
    $ClientId,

    [Parameter( Mandatory, ParameterSetName = 'Certificate+FileExtension' )]
    [String]
    $Tenant,

    [Parameter( Mandatory, ParameterSetName = 'Certificate+FileExtension' )]
    [String]
    $Thumbprint,

    [Parameter( Mandatory, ParameterSetName = 'Credential+FileExtension' )]
    [Parameter( Mandatory, ParameterSetName = 'Certificate+FileExtension' )]
    [String[]]
    $FileExtension
)

#Region Functions
Function Say {
    param(
        [Parameter(Mandatory)]
        $Text,
        [Parameter()]
        $Color = 'Cyan'
    )

    if ($Color) {
        $Host.UI.RawUI.ForegroundColor = $Color
    }
    $Text | Out-Default
    [Console]::ResetColor()
}
#EndRegion Functions

# Filter out the portal, admin portal, and search site URLs.
$urlPatternToExclude = ".*-my\.sharepoint\.com/$|.*\.sharepoint\.com/$|.*\.sharepoint\.com/search$|.*\.sharepoint\.com/portals/hub$|.*\.sharepoint\.com/sites/appcatalog$"
$SiteURL = $SiteURL | Where-Object { $_ -notmatch $urlPatternToExclude }

# Build the KQL Query
$fileExtensionQuery = ''
$FileExtension | ForEach-Object {
    $fileExtensionQuery = $fileExtensionQuery + "filetype:$($_),"
}
$fileExtensionQuery = "(" + ($fileExtensionQuery -replace ".$" -replace ',', ' OR ') + ")"

[System.Collections.Generic.List[System.Object]]$finalResult = @()
for ($urlIndex = 0 ; $urlIndex -lt $SiteURL.Count ; $urlIndex++) {
    $url = $SiteURL[$urlIndex]
    $SearchQuery = "$($fileExtensionQuery) AND path:$($url)"
    # Say $SearchQuery
    try {
        Say "Site $($urlIndex+1) of $($SiteURL.Count): [$($url)]" Cyan

        ## If using certificate authentication
        if ($PSCmdlet.ParameterSetName -like "Certificate*") {
            Connect-PnPOnline -Tenant $Tenant -Url $url -ClientId $ClientId -Thumbprint $Thumbprint -ErrorAction Stop
        }

        ## If using credential authentication (non-MFA)
        if ($PSCmdlet.ParameterSetName -like 'Credential*') {
            Connect-PnPOnline -Url $url -Credentials $Credential -ErrorAction Stop
        }

        $site = Get-PnPTenantSite -Identity $url
        if ($url -like "*-my.sharepoint.com*") {
            Say "  -> OneDrive Name: $($site.Title)" Yellow
            $siteType = 'OneDrive'
        }
        else {
            Say "  -> Site Name: $($site.Title)" Yellow
            $siteType = 'SharePoint'
        }

        $SearchResults = Submit-PnPSearchQuery -Query $SearchQuery -All

        Say "    -> Items: $($SearchResults.RowCount)" Green
        foreach ($ResultRow in $SearchResults.ResultRows) {
            $fileName = $(
                if ($ResultRow['Title'] -notmatch $ResultRow['FileType']) {
                    "$($ResultRow['Title']).$($ResultRow['FileType'])"
                }
                else {
                    $ResultRow['Title']
                }
            )

            $fullPath = $(
                if ($siteType -eq 'OneDrive') {
                    $ResultRow['Path']
                }
                else {
                    "$($ResultRow['ParentLink'])/$($fileName)"
                }
            )

            $finalResult.Add(
                $(
                    [PSCustomObject](
                        [ordered]@{
                            Filename       = $fileName
                            FullPath       = $fullPath
                            'Size(KB)'     = $([math]::Round(($ResultRow['Size'] / 1KB), 2))
                            SiteName       = $site.Title
                            SiteType       = $siteType
                            SiteOwnerName  = $site.OwnerName
                            SiteOwnerEmail = $site.OwnerEmail
                            SiteOwnerType  = $(
                                if ($site.GroupId -ne '00000000-0000-0000-0000-000000000000') {
                                    'Group'
                                }
                                else {
                                    'User'
                                }
                            )
                        }
                    )
                )
            )
        }
    }
    catch {
        Say "[ERROR] - [$($url)]: $($_.Exception.Message)" Red
    }
}

return $finalResult