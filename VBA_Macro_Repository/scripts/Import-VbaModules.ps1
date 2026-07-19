param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $false)]
    [string]$SourceRoot = ".\src\excel"
)

$ErrorActionPreference = "Stop"

$WorkbookPath = (Resolve-Path $WorkbookPath).Path
$SourceRoot = (Resolve-Path $SourceRoot).Path
$tempRoot = Join-Path $env:TEMP ("vba-import-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null

$utf8 = New-Object System.Text.UTF8Encoding($true)
$windows1251 = [System.Text.Encoding]::GetEncoding(1251)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($WorkbookPath)
    $project = $workbook.VBProject

    $priority = @{
        "modCore" = 1
        "modLevenshtein" = 2
    }

    $moduleFiles = Get-ChildItem -Path $SourceRoot -File |
        Where-Object { $_.Extension -in ".bas", ".cls", ".frm" } |
        Sort-Object @{ Expression = {
            $name = $_.BaseName
            if ($priority.ContainsKey($name)) { $priority[$name] } else { 100 }
        }}, Name

    foreach ($file in $moduleFiles) {
        $componentName = $file.BaseName
        $existingComponent = $null

        try {
            $existingComponent = $project.VBComponents.Item($componentName)
        }
        catch {
            $existingComponent = $null
        }

        if ($existingComponent) {
            if ($existingComponent.Type -in 1, 2, 3) {
                Write-Host "Удаление старой версии: $componentName"
                $project.VBComponents.Remove($existingComponent)
            }
            else {
                throw "Нельзя заменить встроенный модуль документа: $componentName"
            }
        }

        $importPath = Join-Path $tempRoot $file.Name
        $sourceText = [System.IO.File]::ReadAllText($file.FullName, $utf8)
        [System.IO.File]::WriteAllText($importPath, $sourceText, $windows1251)

        if ($file.Extension -eq ".frm") {
            $frxSource = [System.IO.Path]::ChangeExtension($file.FullName, ".frx")
            if (Test-Path $frxSource) {
                Copy-Item $frxSource ([System.IO.Path]::ChangeExtension($importPath, ".frx"))
            }
        }

        Write-Host "Импорт: $($file.Name)"
        $null = $project.VBComponents.Import($importPath)
    }

    $workbook.Save()
    Write-Host "Импорт завершён: $WorkbookPath"
}
finally {
    if ($workbook) {
        $workbook.Close($true)
    }

    $excel.Quit()

    if ($project) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($project) | Out-Null
    }
    if ($workbook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

    if (Test-Path $tempRoot) {
        Remove-Item -Recurse -Force $tempRoot
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
