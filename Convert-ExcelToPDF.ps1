# Excel to PDF conversion script
# Shippingport
# Originally written March 3, 2019

<#
.Synopsis
   Converts Excel files to PDFs using an Office interop object.
.EXAMPLE
   ./Convert-ExcelToPDF.ps1
#>

Function Start-ExcelToPDFConversion()
    {
        $path = Get-FileName -initialDirectory "C:testfiles:excel-test1"
        Output-ToPDF -Files $path
    }

Function Get-FileName($initialDirectory)
{   
    [System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Multiselect = $true
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "Excel-files |*.xls;*.xlsx;*.xlsm"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileNames
}

Function Output-ToPDF($Files)
{
    if(!([string]::IsNullOrEmpty($path)))
    {
        $excelFiles = Get-ChildItem -Path $path
        $xlFixedFormat = “Microsoft.Office.Interop.Excel.xlFixedFormatType” -as [type]
        $objExcel = New-Object -ComObject excel.application 
        $objExcel.visible = $false
        
        $outFilePath = Split-Path -Path $excelFiles.Item(0) # Index is required because of weird path return values
        $pdfOutputPath = ($outFilePath + "\PDF")
        
        if (!(Test-Path $pdfOutputPath))
        {
            echo "Creating output directory..."
            New-Item -ItemType directory -Path ($pdfOutputPath) | Out-Null
        }

        $i = 0
        
        foreach($file in $excelFiles) 
        {
            $XLSXFileToPDF = ($file.FullName -replace '.xlsx?$', '.pdf')
            $PDFFileName = $XLSXFileToPDF -replace '.*\\' # Regex to remove everything except leading \
            $PDFOutputFilePath = $pdfOutputPath + "\$PDFFileName"
            
            $workbook = $objExcel.Workbooks.Open($file.fullname, 3) 
            $workbook.Saved = $true 
            Write-Host “Saving file $PDFFileName...” 
            $workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $PDFOutputFilePath) 
            $objExcel.Workbooks.Close()

            $i++
            Write-Progress -Activity "Converting..." -Status "Saved workbook $i of $($path.Count)" -PercentComplete (($i / $path.Count) * 100)
        }
        
        $objExcel.Quit()
        
        # Optionally open output directory on completion
        # Invoke-Item $pdfOutputPath
        
    } else {
    Write-Host "No files selected!"

    $retry = Read-Host "Retry? [y/N]"
    if ($retry -eq 'y|Y')
        {
            Start-ExcelToPDFConversion
        }
    }
}

Start-ExcelToPDFConversion
