param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $false)]
    [string]$DestinationRoot = ".\exported"
)

$ErrorActionPreference = "Stop"

$WorkbookPath = (Resolve-Path $WorkbookPath).Path
$DestinationRoot = [System.IO.Path]::GetFullPath($DestinationRoot)

New-Item -ItemType Directory -Force -Path $DestinationRoot | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($WorkbookPath)
    $project = $workbook.VBProject

    foreach ($component in $project.VBComponents) {
        $extension = switch ($component.Type) {
            1 { ".bas" }
            2 { ".cls" }
            3 { ".frm" }
            default { $null }
        }

        if ($extension) {
            $destination = Join-Path $DestinationRoot ($component.Name + $extension)
            Write-Host "Экспорт: $($component.Name)$extension"
            $component.Export($destination)
        }
    }

    Write-Host "Экспорт завершён: $DestinationRoot"
}
finally {
    if ($workbook) {
        $workbook.Close($false)
    }

    $excel.Quit()

    if ($project) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($project) | Out-Null
    }
    if ($workbook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
