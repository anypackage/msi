#Requires -Modules AnyPackage.Msi

Describe Get-PackageProvider {
    Context 'without parameters' {
        It 'should return package provider' {
            Get-PackageProvider |
            Should -Not -BeNullOrEmpty
        }
    }
}
