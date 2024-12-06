@{
    Path = @(
        './src/code/bin/Release/netstandard2.0/publish/*'
        './src/AnyPackage.NuGet.psd1'
    )
    Destination = './module'
    Exclude = @(
        'NuGetProvider.deps.json',
        '*.pdb'
    )
}
