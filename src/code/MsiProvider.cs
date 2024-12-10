// Copyright (c) Thomas Nieto - All Rights Reserved
// You may use, distribute and modify this code under the
// terms of the MIT license.

using System.Management.Automation;

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
        using var powershell = PowerShell.Create(RunspaceMode.CurrentRunspace);
        powershell.AddCommand("Get-Package")
                  .AddParameter("Name", request.Name)
                  .AddParameter("Provider", @"AnyPackage.Programs\Programs");

        if (request.Version is not null)
        {
            powershell.AddParameter("Version", request.Version);
        }

        if (request.DynamicParameters is GetPackageDynamicParameters dynamicParameters
            && dynamicParameters.SystemComponent)
        {
            powershell.AddParameter("SystemComponent");
        }

        var scriptBlock = ScriptBlock.Create("$_.Metadata['UninstallString'] -like 'MsiExec.exe*' ");

        powershell.AddCommand("Where-Object")
                  .AddParameter("FilterScript", scriptBlock);

        foreach (var result in powershell.Invoke<PackageInfo>())
        {
            PackageSourceInfo? source;
            if (result.Source is not null)
            {
                source = new PackageSourceInfo(result.Source.Name, result.Source.Location, ProviderInfo);
            }
            else
            {
                source = null;
            }

            var package = new PackageInfo(result.Name,
                                          result.Version,
                                          source,
                                          result.Description,
                                          dependencies: null,
                                          result.Metadata.ToDictionary(x => x.Key, x => x.Value),
                                          ProviderInfo);

            request.WritePackage(package);
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
