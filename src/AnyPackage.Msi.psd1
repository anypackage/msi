@{
    RootModule = 'MsiProvider.dll'
    ModuleVersion = '0.1.0'
    CompatiblePSEditions = @('Desktop', 'Core')
    GUID = '56d624ff-e0f4-43e0-a0ca-0af22e41f9f5'
    Author = 'Thomas Nieto'
    Copyright = '(c) 2024 Thomas Nieto. All rights reserved.'
    Description = 'Msi provider for AnyPackage.'
    PowerShellVersion = '5.1'
    RequiredModules = @(
        @{ ModuleName = 'AnyPackage'; ModuleVersion = '0.8.0' }
        'AnyPackage.Programs')
    FunctionsToExport = @()
    CmdletsToExport = @()
    AliasesToExport = @()
    PrivateData = @{
        AnyPackage = @{
            Providers = 'Msi'
        }
        PSData = @{
            Tags = @('AnyPackage', 'Provider', 'Msi', 'Msp', 'Windows')
            LicenseUri = 'https://github.com/anypackage/msi/blob/main/LICENSE'
            ProjectUri = 'https://github.com/anypackage/msi'
        }
    }
    HelpInfoURI = 'https://go.anypackage.dev/help'
}
