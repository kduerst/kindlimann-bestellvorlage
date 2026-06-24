param(
    [string]$OutputFileName = "Kindlimann Bestellliste.xlsm"
)

$ErrorActionPreference = "Stop"

$securityPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$propertyName = "AccessVBOM"
$generator = Join-Path $PSScriptRoot "create_kindlimann_workbook.ps1"

$hadPath = Test-Path -Path $securityPath
$hadProperty = $false
$oldValue = $null

if ($hadPath) {
    $existing = Get-ItemProperty -Path $securityPath -Name $propertyName -ErrorAction SilentlyContinue
    if ($null -ne $existing) {
        $hadProperty = $true
        $oldValue = $existing.$propertyName
    }
}
else {
    New-Item -Path $securityPath -Force | Out-Null
}

try {
    Set-ItemProperty -Path $securityPath -Name $propertyName -Type DWord -Value 1
    & $generator -OutputFileName $OutputFileName
}
finally {
    if ($hadProperty) {
        Set-ItemProperty -Path $securityPath -Name $propertyName -Type DWord -Value $oldValue
    }
    else {
        Remove-ItemProperty -Path $securityPath -Name $propertyName -ErrorAction SilentlyContinue
    }
}
