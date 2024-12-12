namespace AnyPackage.Provider.Msi;

[Flags]
public enum InstallType : int
{
    Product = 1,
    Patch = 2,
    All = int.MaxValue
}
