Import-Module $PSScriptRoot\..\Read-ExcelVBAComponents.psm1 -Force

$correctPath = "$PSScriptRoot\..\Test Dir\Equity Research\Models\Ericsson.xlsm"

$readComponents = Read-ExcelVBAComponents $testPath
$readComponents | % { $_.Code}


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



    }
}
