# export_slides_to_images.ps1
# Exports PPTX slides to PNG images using PowerPoint COM automation.
# Usage: powershell -ExecutionPolicy Bypass -File export_slides_to_images.ps1 -InputFile "path.pptx" -OutputDir "dir" [-Width 1920] [-Height 1080]
# Output: JSON to stdout

param(
    [Parameter(Mandatory=$true)][string]$InputFile,
    [Parameter(Mandatory=$true)][string]$OutputDir,
    [int]$Width = 1920,
    [int]$Height = 1080
)

$ErrorActionPreference = 'Stop'

try {
    # Ensure output directory exists
    if (-not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }

    # Resolve paths
    $InputFile = (Resolve-Path $InputFile).Path
    $OutputDir = (Resolve-Path $OutputDir).Path

    # Start PowerPoint
    $ppt = $null
    try {
        $ppt = New-Object -ComObject PowerPoint.Application
    } catch {
        Write-Output (ConvertTo-Json @{ success = $false; error = "PowerPoint COM not available: $_" })
        exit 1
    }

    $presentation = $null
    try {
        # Open presentation (ReadOnly, Untitled, WithWindow=false)
        $presentation = $ppt.Presentations.Open($InputFile, $true, $false, $false)

        $slideCount = $presentation.Slides.Count
        $files = @()

        for ($i = 1; $i -le $slideCount; $i++) {
            $slide = $presentation.Slides.Item($i)
            $fileName = "slide_$($i.ToString('000')).png"
            $filePath = Join-Path $OutputDir $fileName

            # Export slide as PNG
            $slide.Export($filePath, "PNG", $Width, $Height)
            $files += $fileName
        }

        # Close presentation without saving
        $presentation.Close()

        $result = @{
            success = $true
            slideCount = $slideCount
            files = $files
        }
        Write-Output (ConvertTo-Json $result -Depth 3)
    }
    finally {
        if ($presentation) {
            try { $presentation.Close() } catch {}
        }
        # Quit PowerPoint
        try { $ppt.Quit() } catch {}
        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
catch {
    Write-Output (ConvertTo-Json @{ success = $false; error = $_.Exception.Message })
    exit 1
}
