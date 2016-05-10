<#
SJUKT BRA
https://github.com/RamblingCookieMonster/PSExcel
#>

#Install-Module PSExcel
if (!(Get-Module -ListAvailable -Name PSExcel)) {
        Import-Module "$PSScriptRoot\Modules\PSExcel\1.0\PSExcel.psm1"  
}
Get-Command -Module PSExcel | Out-Null
Import-Module $PSScriptRoot\Read-ExcelVBAComponents.psm1

$reportFolderName = "macroreport_" + ((Get-Date).ToShortDateString() -replace "/","-")

$startDir = "$PSScriptRoot\Test Dir\"
$outDir = "$PSScriptRoot\" + $reportFolderName + "\"
$outDir

if(!(Test-Path $outDir)){
    new-item $outDir -ItemType Directory
}


Get-ChildItem $startDir -Directory -Recurse | ForEach-Object {
    $outputDir = $_.FullName -replace "Test Dir", $reportFolderName
    if(!(Test-Path($outputDir))){
        New-Item $outputDir -ItemType Directory
    }
}


Get-ChildItem $startDir -Include *.xlsm -Recurse | ForEach-Object {
    $outputDir = Split-Path ($_.FullName -replace "Test Dir", $reportFolderName) -Parent
    $fileToFolderDir = $outputDir + "\" + $_.BaseName

    if(!(Test-Path $fileToFolderDir)){
        New-Item $fileToFolderDir -ItemType Directory
    }

    Read-ExcelVBAComponents $_.FullName | ForEach-Object {   
        Export-Clixml -Path ($fileToFolderDir + "\" + $_.Name + ".xml") -InputObject $_.code
    }
}

