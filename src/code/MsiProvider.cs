// Copyright (c) Thomas Nieto - All Rights Reserved
// You may use, distribute and modify this code under the
// terms of the MIT license.

using System.Management.Automation;

using WixToolset.Dtf.WindowsInstaller;

namespace AnyPackage.Provider.Msi;

[PackageProvider("Msi", PackageByName = false, FileExtensions = [".msi", ".msp"])]
public class MsiProvider : PackageProvider, IFindPackage, IGetPackage, IInstallPackage, IUninstallPackage
{
    public void FindPackage(PackageRequest request)
    {
        PackageInfo package;

        if (Path.GetExtension(request.Path) == ".msi")
        {
            package = FindPackageMsi(request);
        }
        else
        {
            package = FindPackageMsp(request);
        }

        request.WritePackage(package);
    }

    public void GetPackage(PackageRequest request)
    {
        var installType = InstallType.All;
        if (request.DynamicParameters is GetPackageDynamicParameters dynamicParameters)
        {
            installType = dynamicParameters.InstallType;
        }

        if (installType.HasFlag(InstallType.Product))
        {
            foreach (var package in GetProductPackages())
            {
                if (request.IsMatch(package.Name, package.Version!))
                {
                    request.WritePackage(package);
                }
            }
        }

        if (installType.HasFlag(InstallType.Patch))
        {
            foreach (var package in GetPatchPackages())
            {
                if (!request.IsVersionFiltered && request.IsMatch(package.Name))
                {
                    request.WritePackage(package);
                }
            }
        }
    }

    public void InstallPackage(PackageRequest request)
    {
        var pathExtension = Path.GetExtension(request.Path);
        var packageExtension = Path.GetExtension(request.Package?.Source?.Location);
        PackageInfo package;

        if (pathExtension == ".msi" || packageExtension == ".msi")
        {
            package = InstallPackageMsi(request);
        }
        else
        {
            package = InstallPackageMsp(request);
        }

        request.WritePackage(package);
    }

    public void UninstallPackage(PackageRequest request)
    {
        IEnumerable<PackageInfo> packages = [];

        if (request.Package is null)
        {
            using var powershell = PowerShell.Create(RunspaceMode.CurrentRunspace);
            powershell.AddCommand("Get-Package")
                      .AddParameter("Name", request.Name)
                      .AddParameter("Provider", ProviderInfo.FullName);

            if (request.Version is not null)
            {
                powershell.AddParameter("Version", request.Version);
            }

            packages = powershell.Invoke<PackageInfo>();
        }
        else
        {
            packages = new [] { request.Package };
        }

        foreach (var package in packages)
        {
            if (package.Source?.Location is null)
            {
                var ex = new InvalidOperationException("Package does not contain path to MSI file.");
                var err = new ErrorRecord(ex, "MissingPackagePath", ErrorCategory.ResourceUnavailable, package);
                request.WriteError(err);
                continue;
            }

            var extension = Path.GetExtension(package.Source.Location);

            if (extension == ".msi")
            {
                try
                {
                    Installer.InstallProduct(package.Source.Location, "REMOVE=ALL");
                    request.WritePackage(package);
                }
                catch (InstallerException e)
                {
                    var err = new ErrorRecord(e, "UninstallFailed", ErrorCategory.InvalidResult, package);
                    request.WriteError(err);
                }
            }
            else
            {
                if (package.Metadata.TryGetValue("ProductCode", out var productCode))
                {
                    if (productCode is null)
                    {
                        var ex = new InvalidOperationException("Patch does not contain product code.");
                        var err = new ErrorRecord(ex, "MissingProductCode", ErrorCategory.InvalidOperation, package);
                        request.WriteError(err);
                        continue;
                    }

                    try
                    {
                        Installer.RemovePatches(new[] { package.Source.Location }, productCode.ToString(), "");
                        request.WritePackage(package);
                    }
                    catch (InstallerException e)
                    {
                        var err = new ErrorRecord(e, "UninstallFailed", ErrorCategory.InvalidResult, package);
                        request.WriteError(err);
                    }
                }
            }
        }
    }

    private PackageInfo InstallPackageMsi(PackageRequest request)
    {
        var package = request.Package ?? FindPackageMsi(request);

        if (package.Source?.Location is null)
        {
            throw new InvalidOperationException("Package does not contain path to MSI file.");
        }

        var logPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        Installer.EnableLog(InstallLogModes.LogOnlyOnError, logPath);
        request.WriteVerbose($"Error logging path: '{logPath}'");

        var properties = "REBOOT=REALLYSUPPRESS ";

        if (request.DynamicParameters is InstallPackageDynamicParameters dynamicParameters)
        {
            properties += string.Join(" ", dynamicParameters.Properties);
        }

        request.WriteVerbose($"Properties: {properties}");

        Installer.InstallProduct(package.Source.Location, properties);

        // Disable logging
        Installer.EnableLog(InstallLogModes.None, null);

        if (Installer.RebootRequired)
        {
            request.WriteWarning("Restart computer to complete installation.");
        }

        return package;
    }

    private PackageInfo InstallPackageMsp(PackageRequest request)
    {
        var package = request.Package ?? FindPackageMsp(request);

        if (package?.Source?.Location is null)
        {
            throw new InvalidOperationException("Package does not contain path to MSI file.");
        }

        var logPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        Installer.EnableLog(InstallLogModes.LogOnlyOnError, logPath);
        request.WriteVerbose($"Error logging path: '{logPath}'");

        var properties = "REBOOT=REALLYSUPPRESS ";

        if (request.DynamicParameters is InstallPackageDynamicParameters dynamicParameters)
        {
            properties += string.Join(" ", dynamicParameters.Properties);
        }

        request.WriteVerbose($"Properties: {properties}");

        Installer.ApplyPatch(package.Source.Location, properties);

        // Disable logging
        Installer.EnableLog(InstallLogModes.None, null);

        if (Installer.RebootRequired)
        {
            request.WriteWarning("Restart computer to complete installation.");
        }

        return package;
    }

    private IEnumerable<PackageInfo> GetProductPackages()
    {
        foreach (var product in ProductInstallation.AllProducts)
        {
            if (string.IsNullOrWhiteSpace(product.ProductName))
            {
                continue;
            }

            var metadata = new Dictionary<string, object?>
            {
                { "Features", product.Features },
                { "ProductCode", product.ProductCode },
                { "IsInstalled", product.IsInstalled },
                { "IsAdvertised", product.IsAdvertised },
                { "IsElevated", product.IsElevated },
                { "SourceList", product.SourceList },
                { "HelpLink", product.HelpLink },
                { "HelpTelephone", product.HelpTelephone },
                { "InstallDate", product.InstallDate },
                { "ProductName", product.ProductName },
                { "InstallLocation", product.InstallLocation },
                { "InstallSource", product.InstallSource },
                { "LocalPackage", product.LocalPackage },
                { "Publisher", product.Publisher },
                { "UrlInfoAbout", product.UrlInfoAbout },
                { "UrlUpdateInfo", product.UrlUpdateInfo },
                { "ProductVersion", product.ProductVersion },
                { "ProductId", product.ProductId },
                { "RegCompany", product.RegCompany },
                { "RegOwner", product.RegOwner },
                { "AdvertisedTransforms", product.AdvertisedTransforms },
                { "AdvertisedLanguage", product.AdvertisedLanguage },
                { "AdvertisedProductName", product.AdvertisedProductName },
                { "AdvertisedPerMachine", product.AdvertisedPerMachine },
                { "AdvertisedPackageCode", product.AdvertisedPackageCode },
                { "AdvertisedVersion", product.AdvertisedVersion },
                { "AdvertisedProductIcon", product.AdvertisedProductIcon },
                { "AdvertisedPackageName", product.AdvertisedPackageName },
                { "PrivilegedPatchingAuthorized", product.PrivilegedPatchingAuthorized },
                { "UserSid", product.UserSid },
                { "Context", product.Context },
                { "InstallType", InstallType.Product }
            };

            var source = new PackageSourceInfo(product.LocalPackage,
                                               product.LocalPackage,
                                               ProviderInfo);

            yield return new PackageInfo(product.ProductName,
                                         product.ProductVersion,
                                         source,
                                         description: "",
                                         dependencies: null,
                                         metadata,
                                         ProviderInfo);
        }
    }

    private IEnumerable<PackageInfo> GetPatchPackages()
    {
        foreach (var patch in PatchInstallation.AllPatches)
        {
            var metadata = new Dictionary<string, object?>
            {
                { "Context", patch.Context },
                { "InstallDate", patch.InstallDate },
                { "IsInstalled", patch.IsInstalled },
                { "IsObsoleted", patch.IsObsoleted },
                { "IsSuperseded", patch.IsSuperseded },
                { "LocalPackage", patch.LocalPackage },
                { "MoreInfoUrl", patch.MoreInfoUrl },
                { "PatchCode", patch.PatchCode },
                { "ProductCode", patch.ProductCode },
                { "SourceList", patch.SourceList },
                { "State", patch.State },
                { "Transforms", patch.Transforms },
                { "Uninstallable", patch.Uninstallable },
                { "UserSid", patch.UserSid },
                { "InstallType", InstallType.Patch }
            };

            var source = new PackageSourceInfo(patch.LocalPackage,
                                               patch.LocalPackage,
                                               ProviderInfo);

            yield return new PackageInfo(patch.DisplayName,
                                         version: null,
                                         source,
                                         description: "",
                                         dependencies: null,
                                         metadata,
                                         ProviderInfo);
        }
    }

    private PackageInfo FindPackageMsi(PackageRequest request)
    {
        using var database = new Database(request.Path, DatabaseOpenMode.ReadOnly);
        var metadata = GetMetadataMsi(database);

        var source = new PackageSourceInfo(request.Path, request.Path, ProviderInfo);
        var package = new PackageInfo(metadata["ProductName"]!.ToString(),
                                      metadata["ProductVersion"]!.ToString(),
                                      source,
                                      description: "",
                                      dependencies: null,
                                      metadata,
                                      ProviderInfo);

        return package;
    }

    private PackageInfo FindPackageMsp(PackageRequest request)
    {
        using var database = new Database(request.Path, DatabaseOpenMode.ReadOnly);
        var metadata = GetMetadataMsp(database);

        var source = new PackageSourceInfo(request.Path, request.Path, ProviderInfo);
        var package = new PackageInfo(metadata["DisplayName"]!.ToString(),
                                      version: null,
                                      source,
                                      description: metadata["Description"]!.ToString(),
                                      dependencies: null,
                                      metadata,
                                      ProviderInfo);

        return package;
    }

    protected override object? GetDynamicParameters(string commandName)
    {
        return commandName switch
        {
            "Get-Package" => new GetPackageDynamicParameters(),
            "Install-Package" => new InstallPackageDynamicParameters(),
            _ => null,
        };
    }

    private static Dictionary<string, object?> GetMetadataMsi(Database database)
    {
        var props = database.ExecuteStringQuery("SELECT Property FROM Property");
        var metadata = new Dictionary<string, object?>();

        foreach (var prop in props)
        {
            var value = database.ExecutePropertyQuery(prop);
            metadata.Add(prop, value);
        }

        return metadata;
    }

    private static Dictionary<string, object?> GetMetadataMsp(Database database)
    {
        var props = database.ExecuteStringQuery("SELECT Property FROM MsiPatchMetadata");
        var metadata = new Dictionary<string, object?>();

        foreach (var prop in props)
        {
            var query = $"SELECT Value FROM MsiPatchMetadata WHERE Property = '{prop}'";
            var value = database.ExecuteStringQuery(query);
            metadata.Add(prop, string.Join(Environment.NewLine, value));
        }

        return metadata;
    }
}
