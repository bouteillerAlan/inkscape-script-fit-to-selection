Add-Type -AssemblyName System.Windows.Forms

# Select input dir
$inputDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$inputDialog.Description = "Select the folder containing your SVG files"
$inputResult = $inputDialog.ShowDialog()

if ($inputResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $inputPath = $inputDialog.SelectedPath
    Write-Host "Input folder selected : $inputPath"

    # Select output dir
    $outputDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $outputDialog.Description = "Select the output folder"
    $outputResult = $outputDialog.ShowDialog()

    if ($outputResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputPath = $outputDialog.SelectedPath
        Write-Host "Output folder selected : $outputPath"

        $files = Get-ChildItem $inputPath -Filter *.svg
        $nbOfFiles = $files.Length
        $i = 0

        $files | Foreach-Object -Process {
            $inputFilename = $inputPath + "\" + $_.Name
            $outputFilename = $outputPath + "\" + $_.Name

            Write-Progress -Activity "Processing" -Status "Processing file : $_" -PercentComplete (($i/$nbOfFiles)*100)
            Write-Host "Processing $inputFilename..."
            inkscape --actions='page-fit-to-selection;export-plain-svg' --export-filename=$outputFilename --export-type='svg' $_.FullName
            Write-Host "Saving $outputFilename..." -ForegroundColor Green

            $i = $i + 1
        }
    }
}
