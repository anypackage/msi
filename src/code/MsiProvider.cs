// Copyright (c) Thomas Nieto - All Rights Reserved
// You may use, distribute and modify this code under the
// terms of the MIT license.

using WixToolset.Dtf.WindowsInstaller;

namespace AnyPackage.Provider.Msi;

[PackageProvider("Msi", PackageByName = false, FileExtensions = [".msi", ".msp"])]
public class MsiProvider : PackageProvider, IFindPackage, IGetPackage
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
