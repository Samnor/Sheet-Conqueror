Import-Module $PSScriptRoot\..\Read-ExcelVBAComponents.psm1 -Force

# Prepartions for pester tests below
$correctPath = "$PSScriptRoot\..\Test Dir\Equity Research\Models\Ericsson.xlsm"
$readComponents = Read-ExcelVBAComponents $correctPath
$readComponents | ForEach-Object {
    if($_.name -eq "Main"){
        $codeTest = $_.Code -match "sub subname"
    }
}

Describe 'Read-ExcelVBAComponents' {
    Context 'Strict mode' {
        
        It "should throw an error when the path is empty" {
            {Read-ExcelVBAComponents ""} | Should Throw
        }
        It "should throw an error when the path is random letters" {
            {Read-ExcelVBAComponents "asC:\dgfdfgfd"} | Should Throw
        }
        It "should not throw an error when the path is valid" {
            {Read-ExcelVBAComponents $correctPath} | Should Not Throw
        }
        It "should find sub within main module in Ericsson.xlsm" {
            $codeTest | Should Be $True
        }
    }
}
