// Copyright (c) Thomas Nieto - All Rights Reserved
// You may use, distribute and modify this code under the
// terms of the MIT license.

using System.Management.Automation;

namespace AnyPackage.Provider.Msi;

public class InstallPackageDynamicParameters
{
    [Parameter]
    public string[] Properties { get; set; } = [];
}
