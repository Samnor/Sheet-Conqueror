﻿
function Read-ExcelVBAComponents(){
    <#
    .SYNOPSIS
    Extract all VBA components from an excel workbook.

    .DESCRIPTION
    Uses commands from PSExcel to extract VBA components from an excel workbook.

    .PARAMETER workbookPath
    The full path to the excel workbook that you wish to use.

    .EXAMPLE
    Coming soon

    .INPUTS
    Path strings only.

    .OUTPUTS
    Returns an ArrayList with VBA modules found. (System.Collections.ArrayList)

    .NOTES
    Dependencies: PSExcel 
    Started writing this function 8th may 2016.
    #>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$true)]$workbookPath
    )
    try{
        if(!(Test-Path $workbookPath)){
            Throw "workbookPath not found."
        }
    }
    catch
    {
        Throw "Bad workbookPath input"
    }
    
    if (!(Get-Module -ListAvailable -Name PSExcel)) {
        Import-Module "$PSScriptRoot\Modules\PSExcel\1.0\PSExcel.psm1"  
    }

    Get-Command -Module PSExcel | Out-Null
    $returnArray = [System.Collections.ArrayList]@()
    $excelObj = New-Excel -Path $workbookPath
    $excelObj.Workbook.VbaProject.Modules | ForEach-Object {
        $returnArray.Add($_)
    }
    $returnArray
}

#$testPath = "C:\Users\Samuel\Desktop\CURRENT\Posh VBA\Test Dir\Equity Research\Models\Ericsson.xlsm"

#$readComponents = Read-ExcelVBAComponents $testPath
#$readComponents | % { $_.Code} 
