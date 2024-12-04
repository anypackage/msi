// Copyright (c) Thomas Nieto - All Rights Reserved
// You may use, distribute and modify this code under the
// terms of the MIT license.

using System.Management.Automation;

using static AnyPackage.Provider.PackageProviderManager;

namespace AnyPackage.Provider.Msi;

public sealed class Init : IModuleAssemblyInitializer, IModuleAssemblyCleanup
{
    private readonly Guid _id = new("327bc87e-5949-4f87-802b-c68cecea8c15");

    public void OnImport()
    {
        RegisterProvider(_id, typeof(MsiProvider), "AnyPackage.Msi");
    }

    public void OnRemove(PSModuleInfo psModuleInfo)
    {
        UnregisterProvider(_id);
    }
}
