
#Requires -Modules @{ ModuleName="PnP.PowerShell"; RequiredVersion="2.3.0" }
#Requires -Version 7.2

<#PSScriptInfo

.VERSION 0.1

.GUID 8197ffd7-8561-4c67-9602-32a69d59b337

.AUTHOR June Castillote

.COMPANYNAME

.COPYRIGHT june.castillote@gmail.com

.TAGS

.LICENSEURI https://raw.githubusercontent.com/junecastillote/SPO-File-Search/main/LICENSE

.PROJECTURI https://github.com/junecastillote/SPO-File-Search

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

* 0.1
      - Initial
* 0.2
      - Added the Interactive login switch
      - Changed default parameter set to Interactive+SearchString

.PRIVATEDATA

#>

<#

.PARAMETER SiteURL

The site URL you want to search.
You can enter a single URL string:
    - "https://contoso-my.sharepoint.com/personal/someone".
Or a collection:
    - @("https://contoso-my.sharepoint.com/personal/someone","https://contoso.sharepoint.com/sites/sitename")

.PARAMETER Credential

The PSCredential object of the account used to connect to the SharePoint or OneDrive site.
This credential must be non-MFA enabled and has Site Administrator or Owner access to the site.

.PARAMETER ClientId

The client ID of application ID of the Azure AD app registration, if using app-only authentication instead of a credential.

.PARAMETER Tenant

The SharePoint Online tenant ID, such as contoso.sharepoint.com, if using app-only authentication instead of a credential.

.PARAMETER Thumbprint

The public key certificate thumbprint associated with the Azure AD app registration, if using app-only authentication instead of a credential.
The corresponding private certificate must be present in your personal certificate store [cert:\CurrentUser\My\<thumbprint>] for this to work.

.PARAMETER Interactive

Switch to trigger interactive login. Note that this switch will prompt you to log in to each site to search.

.PARAMETER SearchString

One of more specific file name or pattern to search. For example, "*.pdf","filename.ext","file*.00*"

.PARAMETER ReturnResult

Indicates whether the search results will be returned.

.PARAMETER OutputFile

The custom CSV file path to write the search results, if any. If not specified, the default output file path will be: ".\search\SPO_File_Search_yyyy-MM-dd_hh-mm-ss_tt_username.csv"

.PARAMETER Quiet

Suppresses the informational output on the screen. The output will still be written to a log file with the same filename as the output filename with a LOG extension.

.SYNOPSIS
    PowerShell script to find files in a SharePoint Online and OneDrive sites.
.DESCRIPTION
    * A PowerShell script that uses the PnP.PowerShell module to search for files in one, many, or all SharePoint and OneDrive sites.
    * The search string can be a filename (filename.ext) or file extension (*.ext), for example and also accepts wildcard (f*le*) patterns.
    * The output is automatically written to a default or custom CSV file path.
.NOTES
    * This script only works in PowerShell Core (7.2+), which is a requirement of the PnP.PowerShell module.
    * This script has only been tested on Windows.

.LINK
https://github.com/junecastillote/SPO-File-Search/tree/main/README.md

.OUTPUTS

.EXAMPLE

# ??

This example connects to the SharePoint online site using a non-MFA credential stored as a credential object and searches for files matching *.txt, *.docx, and a specific filename called "my daily tracker.log"
The search results is written to the default CSV file.

```
.\Find-FileInSite.ps1 -Credential <pscredential> -SiteUrl https://contoso.sharepoint.com/sites/dummysite -SearchString *.txt, *.docx, "my daily tracker.log"
```

.EXAMPLE

# ??

This example connects to the SharePoint online site using app-only authentication and searches for files matching *.txt, *.docx, and a specific filename called "my daily tracker.log"
The search results is written to a custom CSV file path.

```
$splat = @{
    Tenant       = 'contoso.sharepoint.com'
    ClientId     = '54630388-7f0f-4a2e-b2d6-d99d95fd7f4f'
    Thumbprint   = '41C82D1E0EE6B428559387B0623FD632E2CD70C6'
    SiteUrl      = @('https://contoso.sharepoint.com/sites/dummysite1', 'https://contoso-my.sharepoint.com/personal/onedrive1')
    SearchString = @("*.txt", "*.docx", "my daily tracker.log")
    OutputFile   = "C:\Path\To\File.CSV"
}

.\Find-FileInSite.ps1 @splat
```

#>

[CmdletBinding(DefaultParameterSetName = 'Interactive+SearchString')]
param (
    [Parameter(
        Mandatory,
        Position = 0
    )]
    [String[]]
    $SiteURL,

    [Parameter( Mandatory, ParameterSetName = 'Credential+SearchString' )]
    [pscredential]
    $Credential,

    [Parameter( Mandatory, ParameterSetName = 'Certificate+SearchString' )]
    [String]
    $ClientId,

    [Parameter( Mandatory, ParameterSetName = 'Certificate+SearchString' )]
    [String]
    $Tenant,

    [Parameter( Mandatory, ParameterSetName = 'Certificate+SearchString' )]
    [String]
    $Thumbprint,

    [Parameter( ParameterSetName = 'Interactive+SearchString' )]
    [Switch]
    $Interactive,

    # [Parameter( Mandatory, ParameterSetName = 'Credential+SearchString' )]
    # [Parameter( Mandatory, ParameterSetName = 'Certificate+SearchString' )]
    [Parameter( Mandatory )]
    [String[]]
    $SearchString,

    [Parameter()]
    [Switch]
    $ReturnResult,

    [Parameter()]
    [String]
    $OutputFile,

    [Parameter()]
    [Switch]
    $Quiet
)

begin {

    #Region Functions
    Function Say {
        param(
            [Parameter(Mandatory)]
            $Text,
            [Parameter()]
            [ValidateSet(
                "Black", "DarkBlue", "DarkGreen", "DarkCyan", "DarkRed", "DarkMagenta", "DarkYellow", "Gray", "DarkGray", "Blue", "Green", "Cyan", "Red", "Magenta", "Yellow", "White"
            )]
            $Color = $([Console]::ResetColor(); $Host.UI.RawUI.ForegroundColor)
        )

        if (!$Quiet) {
            $Host.UI.RawUI.ForegroundColor = $Color
            $Text | Out-Default
            [Console]::ResetColor()
        }

        $Text | Out-File -Append $LogFile

    }

    # Function to replace the internal column names
    # e.g "[First_x0020_Name] to [First Name]"
    Function ReplaceInternalName {
        param(
            [parameter(Mandatory)]
            [string]$String
        )
        #Patterns
        $internalNames = @{
            _x0020_	=	' '
            _x007e_	=	'~'
            _x0021_	=	'!'
            _x0040_	=	'@'
            _x0023_	=	'#'
            _x0024_	=	'$'
            _x0025_	=	'%'
            _x005E_	=	'^'
            _x0026_	=	'&'
            _x002a_	=	'*'
            _x0028_	=	'('
            _x0029_	=	')'
            _x002B_	=	'+'
            _x002D_	=	'-'
            _x003D_	=	'='
            _x007B_	=	'{'
            _x007D_	=	'}'
            _x003A_	=	':'
            _x0022_	=	"'"
            _x007C_	=	'|'
            _x003B_	=	';'
            _x0027_	=	"'"
            _x005C_	=	'\'
            _x003C_	=	'<'
            _x003E_	=	'>'
            _x003F_	=	'?'
            _x002C_	=	','
            _x002E_	=	'.'
            _x002F_	=	'/'
            _x0060_	=	'`'
        }

        $internalNames.Keys | ForEach-Object {
            $stringToReplace = $_
            $String = $String -replace $stringToReplace, $internalNames[$stringToReplace]
        }
        return $String
    }

    #EndRegion Functions

    $now = Get-Date
    $nowString = $now.ToString("yyyy-MM-dd_hh-mm-ss_tt")

    if (!$OutputFile) {
        $OutputFile = ".\search\SPO_File_Search_$($nowString)_$($env:USERNAME).csv"
    }

    if ($OutputFile) {
        try {
            $LogFile = $($OutputFile).Replace('.csv', '.log')
            $null = New-Item -ItemType File -Path $OutputFile -Force -ErrorAction Stop
            $null = New-Item -ItemType File -Path $LogFile -Force -ErrorAction Stop
            Start-Sleep -Seconds 3
            Say "Results file: $((Resolve-Path $OutputFile).Path)" Yellow
            Say "Log file: $((Resolve-Path $LogFile).Path)" Yellow
        }
        catch {
            Say "[ERROR] : $($_.Exception.Message)" Red
            Continue
        }
    }

    Say "Start @ $((Get-Date).ToString("yyyy-MM-dd hh:mm:ss tt"))" Yellow
    # Filter out the portal, admin portal, and search site URLs.
    $urlPatternToExclude = ".*-my\.sharepoint\.com/$|.*\.sharepoint\.com/$|.*\.sharepoint\.com/search$|.*\.sharepoint\.com/portals/hub$|.*\.sharepoint\.com/sites/appcatalog$"
    $SiteURL = $SiteURL | Where-Object { $_ -notmatch $urlPatternToExclude }

    # System libraries to ignore
    $SystemLibraries = @('Form Templates', 'Pages', 'Preservation Hold Library', 'Site Assets', 'Site Pages', 'Images',
        'Site Collection Documents', 'Site Collection Images', 'Style Library')

    # $filterPattern = ($SearchString | ForEach-Object { [regex]::Escape($_) -replace '\\\*', '.*' }) -join '|'
    $filterPattern = (
        $SearchString | ForEach-Object {
            if ($_ -match '\*\.\w+') {
                # if the search string is *.<word>
                "$([regex]::Escape($_) -replace '\\\*', '.*')$"
            }
            elseif ($_ -match '\*') {
                # if the search string contains *
                "$([regex]::Escape($_) -replace '\\\*', '.*')$"
            }
            elseif ($_ -notmatch '\*') {
                # if the search string does not contain *
                "$([regex]::Escape($_))$"
            }
            else {
                $_
            }
        }
    ) -join '|'

    Say "SPO / ODB Sites to search: $($SiteURL.Count)" Yellow
    Say "Search Filter: $($filterPattern)" Yellow

    # Tenant URLs
    $tenantUrls = $null
}

process {
    # [System.Collections.Generic.List[System.Object]]$finalResult = @()
    for ($urlIndex = 0 ; $urlIndex -lt $SiteURL.Count ; $urlIndex++) {
        $url = $SiteURL[$urlIndex]
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

            ## If using interactive authentication (non-MFA/MFA)
            if ($PSCmdlet.ParameterSetName -like 'Interactive*') {
                Say "  -> Start interactive login to site." Yellow
                Connect-PnPOnline -Url $url -Interactive -ErrorAction Stop
            }

            ## Get tenant URLs (once)
            if (!$tenantUrls) {
                # Say "  -> Getting tenant URLs" Yellow
                $tenantUrls = Get-PnPTenantInstance
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

            # Get all document libraries that are:
            # * BaseType is DocumentLibrary
            # * Not hidden
            # * Title is not in the $SystemLibraries
            $DocumentLibraries = @(
                Get-PnPList |
                Where-Object {
                    $_.BaseType -eq 'DocumentLibrary' -and
                    $_.Hidden -eq $False -and
                    $_.Title -notin $SystemLibraries
                }
            )
            $DocumentLibraries | Add-Member -MemberType NoteProperty -Value $null -Name Leaf

            foreach ($item in $DocumentLibraries) {
                $item.Leaf = ReplaceInternalName -String $item.EntityTypeName
            }

            if ($DocumentLibraries.Count -gt 0) {
                # Process each document library
                for ($libraryIndex = 0; $libraryIndex -lt $DocumentLibraries.Count; $libraryIndex++) {
                    $library = $DocumentLibraries[$libraryIndex]
                    Say "    -> Library name: [$(($library.Title))]" Magenta
                    [System.Collections.Generic.List[System.Object]]$searchResult = @()
                    $files = Get-PnPFolderItem -FolderSiteRelativeUrl $($library.Leaf) -ItemType File -Recursive
                    $searchResult.AddRange(
                        @(
                            $files | Where-Object { $_.Name -match $filterPattern }
                        )
                    )
                    Say "      -> Items: $($searchResult.Count)" Green
                }
            }

            if ($searchResult) {
                $searchResult | Add-Member -Name SiteUrl -MemberType NoteProperty -Value $site.Url
                $searchResult | Add-Member -Name SiteName -MemberType NoteProperty -Value $site.Title
                $searchResult | Add-Member -Name SiteType -MemberType NoteProperty -Value $siteType
                $searchResult | Add-Member -Name OwnerName -MemberType NoteProperty -Value $site.OwnerName
                $searchResult | Add-Member -Name OwnerEmail -MemberType NoteProperty -Value $site.OwnerEmail
                $searchResult | Add-Member -Name OwnerType -MemberType NoteProperty -Value $(
                    if ($site.GroupId -ne '00000000-0000-0000-0000-000000000000') {
                        'Group'
                    }
                    else {
                        'User'
                    }
                )

                $output = $searchResult | Select-Object SiteUrl, SiteName, SiteType, OwnerName, OwnerEmail, OwnerType, @{
                    n = "ParentPath"; e = {
                        "/$((($_.ServerRelativeUrl -split '/') | Select-Object -Skip 3 | Select-Object -SkipLast 1) -join "/")/"
                    }
                },
                @{n = "FileName"; e = { $_.Name } },
                @{n = "FileType"; e = { ($_.Name.ToString().Split('.'))[-1] } },
                @{n = "SizeKB"; e = { $([math]::Round(($_.Length / 1KB), 2)) } },
                TimeCreated, TimeLastModified

                if ($ReturnResult) {
                    $output
                }

                if ($OutputFile) {
                    $output | Export-Csv -NoTypeInformation -Append $OutputFile
                }
            }
        }
        catch {
            Say "[ERROR] - [$($site.Title)]: $($_.Exception.Message)" Red
        }
    }
}

end {
    Say "End @ $((Get-Date).ToString("yyyy-MM-dd hh:mm:ss tt"))" Yellow
}
