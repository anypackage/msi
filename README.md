# AnyPackage.Msi

[![gallery-image]][gallery-site]
[![build-image]][build-site]
[![cf-image]][cf-site]

[gallery-image]: https://img.shields.io/powershellgallery/dt/AnyPackage.Msi
[build-image]: https://img.shields.io/github/actions/workflow/status/anypackage/msi/ci.yml
[cf-image]: https://img.shields.io/codefactor/grade/github/anypackage/msi
[gallery-site]: https://www.powershellgallery.com/packages/AnyPackage.Msi
[build-site]: https://github.com/anypackage/msi/actions/workflows/ci.yml
[cf-site]: https://www.codefactor.io/repository/github/anypackage/msi

Msi provider for AnyPackage.

## Install AnyPackage.Msi

```powershell
Install-PSResource AnyPackage.Msi
```

## Import AnyPackage.Msi

```powershell
Import-Module AnyPackage.Msi
```

## Sample usages

### Find available packages

```powershell
Find-Package -Path C:\Temp\installer.msi
```

### Install packages

```powershell
Install-Package -Path C:\Temp\installer.msi
```

### Install packages with custom properties

```powershell
Install-Package -Path C:\Temp\installer.msi -Provider Msi -Properties "CustomProp=1"
```

### Get installed packages

```powershell
Get-Package
```

### Get installed packages and specify type

```powershell
# MSI Files
Get-Package -Provider Msi -InstallerType Product

# MSP Files
Get-Package -Provider Msi -InstallerType Patch
```

### Uninstall packages

```powershell
Uninstall-Package -Name neovim
```
