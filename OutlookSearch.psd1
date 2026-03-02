@{
    RootModule = 'OutlookSearch.psm1'
    ModuleVersion = '1.0.0'
    GUID = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890'
    Author = 'Foad Sojoodi Farimani'
    CompanyName = 'Community'
    Copyright = '(c) 2026 Foad Sojoodi Farimani. All rights reserved.'
    Description = 'Advanced CLI/TUI for Outlook email search with boolean logic and export capabilities'
    PowerShellVersion = '7.4'
    RequiredModules = @()
    FunctionsToExport = @('Search-Outlook', 'Export-OutlookEmail', 'Show-OutlookSearchTUI')
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @('oso')
    PrivateData = @{
        PSData = @{
            Tags = @('Outlook', 'Email', 'Search', 'TUI', 'CLI', 'Automation')
            LicenseUri = 'https://www.gnu.org/licenses/gpl-3.0.txt'
            ProjectUri = 'https://github.com/Foadsf/OutlookSearch'
            ReleaseNotes = 'Initial release with full CLI/TUI support'
        }
    }
}
