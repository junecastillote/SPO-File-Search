# SharePoint and OneDrive File Search PowerShell Script

Administrators may be asked to search SharePoint and OneDrive sites for various reasons. It may be for reporting or analysis. This script utilizes the [PnP.PowerShell](https://pnp.github.io/powershell/) module to find files by name, extension, or wildcard pattern in one or more sites.

- [Other alternatives exists?](#other-alternatives-exists)
- [Why use this script (and when)?](#why-use-this-script-and-when)
- [Requirements](#requirements)
  - [Environment](#environment)
  - [Authentication](#authentication)
    - [OPTION 1: Administrator Credential](#option-1-administrator-credential)
    - [OPTION 2: App-Only Credential with Certificate (Recommended)](#option-2-app-only-credential-with-certificate-recommended)
- [Syntax](#syntax)
  - [Using an Administrator Credential](#using-an-administrator-credential)
  - [Using App-Only Authentication](#using-app-only-authentication)
- [Parameters](#parameters)
  - [`-SiteURL <String[]>`](#-siteurl-string)
  - [`-Credential <PSCredential>`](#-credential-pscredential)
  - [`-ClientId <String>`](#-clientid-string)
  - [`-Tenant <String>`](#-tenant-string)
  - [`-Thumbprint <String>`](#-thumbprint-string)
  - [`-SearchString <String[]>`](#-searchstring-string)
  - [`-ReturnResult [<SwitchParameter>]`](#-returnresult-switchparameter)
  - [`-OutputFile <String>`](#-outputfile-string)
  - [`-Quiet [<SwitchParameter>]`](#-quiet-switchparameter)
  - [`<CommonParameters>`](#commonparameters)

## Other alternatives exists?

Yes, there are alternatives to use to search for files, such as compiance search in Exchange Online PowerShell using the [`New-ComplianceSearch`](https://learn.microsoft.com/en-us/powershell/module/exchange/new-compliancesearch?view=exchange-ps) cmdlet and in the [Microsoft Purview Content Search](https://learn.microsoft.com/en-us/training/modules/search-for-content-security-compliance-center/).

The [`Submit-PnPSearchQuery`](https://pnp.github.io/powershell/cmdlets/Submit-PnPSearchQuery.html) cmdlet is another alternative that searches the SharePoint Search Index.

## Why use this script (and when)?

While the native tools provide a convenient way to search files on all sites, they don't always show accurate and real-time results. For example, if the content or file is not indexed somehow, it will not show in the search index query results.

This script connects to each specified site, gets the existing files recursively, and filter the result based on your search string. But since the script reads ALL file properties in any given document library, the speed will depend on the volume of files.

> During testing, the script can retreive approximately 14,000 files in one document library recursively in around 10 minutes.

## Requirements

### Environment

- A Windows Computer
- [PowerShell 7.2](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows)
- [PnP.PowerShell 2.3.0](https://www.powershellgallery.com/packages/PnP.PowerShell)

### Authentication

#### OPTION 1: Administrator Credential

- A non-MFA credential with a SharePoint Admin role and has full access to all SharePoint and OneDrive sites.

#### OPTION 2: App-Only Credential with Certificate (Recommended)

- Register an Azure AD App using the [Register-PnPAzureADApp](https://pnp.github.io/powershell/cmdlets/Register-PnPAzureADApp.html) cmdlet.

## Syntax

### Using an Administrator Credential

```powershell
.\Find-FileInSite.ps1
    [-SiteURL] <String[]>
    -Credential <PSCredential>
    -SearchString <String[]>
    [-ReturnResult]
    [-OutputFile <String>]
    [-Quiet]
    [<CommonParameters>]
```

### Using App-Only Authentication

```powershell
.\Find-FileInSite.ps1
    [-SiteURL] <String[]>
    -ClientId <String>
    -Tenant <String>
    -Thumbprint <String>
    -SearchString <String[]>
    [-ReturnResult]
    [-OutputFile <String>]
    [-Quiet]
    [<CommonParameters>]
```

## Parameters

### `-SiteURL <String[]>`

The site URL you want to search.

You can enter a single URL string:

```PowerShell
"https://contoso-my.sharepoint.com/personal/someone"
```

Or a collection:

```PowerShell
@(
    "https://contoso-my.sharepoint.com/personal/someone",
    "https://contoso.sharepoint.com/sites/sitename"
)
```

### `-Credential <PSCredential>`

The PSCredential object of the account used to connect to the SharePoint or OneDrive site.

This credential must be non-MFA enabled and has Site Administrator or Owner access to the site.

### `-ClientId <String>`

The client ID of application ID of the Azure AD app registration, if using app-only authentication instead of a credential.

### `-Tenant <String>`

The SharePoint Online tenant ID, such as `contoso.sharepoint.com`, if using app-only authentication instead of a credential.

### `-Thumbprint <String>`

The public key certificate thumbprint associated with the Azure AD app registration, if using app-only authentication instead of a credential.

The corresponding private certificate must be present in your personal certificate store `[cert:\CurrentUser\My\<thumbprint>]` for this to work.

### `-SearchString <String[]>`

One of more specific file name or pattern to search. For example, `"*.pdf","filename.ext","file*.00*"`

### `-ReturnResult [<SwitchParameter>]`

Indicates whether the search results will be returned.

### `-OutputFile <String>`

The custom CSV file path to write the search results, if any. If not specified, the default output file path will be: `.\search\SPO_File_Search_yyyy-MM-dd_hh-mm-ss_tt_username.csv`

### `-Quiet [<SwitchParameter>]`

Suppresses the informational output on the screen. The output will still be written to a log file with the same filename as the output filename with a `.LOG` extension.

### `<CommonParameters>`

This cmdlet supports the common parameters: `Verbose`, `Debug`, `ErrorAction`, `ErrorVariable`, `WarningAction`, `WarningVariable`, `OutBuffer`, `PipelineVariable`, and `OutVariable`.

For more information, see [about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).
