#Import-Module -Force $PSScriptRoot\..\Read-ExcelVBAComponents.ps1

Describe 'Read-ExcelVBAComponents' {
    Context 'Strict mode' {
        
        It 'should be true when true' {
            $true | Should be $true
        }

    }
}