@{
    Path = @(
        './src/code/bin/Release/netstandard2.0/publish/*'
        './src/AnyPackage.Msi.psd1'
    )
    Destination = './module'
    Exclude = @(
        '*.deps.json',
        '*.pdb'
    )
}
