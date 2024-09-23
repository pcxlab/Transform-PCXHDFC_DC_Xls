<#
.SYNOPSIS
This script converts an Excel file from .xls to .xlsx format if necessary, detects the header row dynamically, and transforms the data based on specific criteria.

.DESCRIPTION
The script performs the following steps:
1. Converts an .xls file to .xlsx format if necessary, adding the suffix '_ConvertedFromXls'.
2. Automatically detects if the input file is already in .xlsx format and skips the conversion step.
3. Dynamically detects the header row and column in the Excel file.
4. Imports the data starting from the detected header row.
5. Transforms the data by remapping columns and updating specific fields if certain conditions are met.
6. Exports the transformed data to a new Excel file with the suffix '_Transformed'. If the file was converted, the suffix '_ConvertedFromXls_Transformed' is used.

.PARAMETER inputFilePath
The path to the input Excel file (.xls or .xlsx).

.NOTES
Ensure the ImportExcel module is installed before running this script.
#>

param (
    [string]$inputFilePath
)

Clear-Host

# Function to validate parameters
function Validate-Parameters {
    param (
        [string]$inputFilePath
    )

    if (-not $inputFilePath) {
        throw "Please provide the input file path as a parameter."
    }

    if (-not (Test-Path $inputFilePath)) {
        throw "The file path '$inputFilePath' does not exist."
    }
}

# Function to convert .xls to .xlsx
function Convert-XlsToXlsx {
    param (
        [string]$inputFilePath
    )

    $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($inputFilePath)
    $directoryPath = [System.IO.Path]::GetDirectoryName($inputFilePath)
    $xlsxFilePath = "$directoryPath\$baseFileName`_ConvertedFromXls.xlsx"

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $workbook = $excel.Workbooks.Open($inputFilePath)
        $workbook.SaveAs($xlsxFilePath, 51)  # 51 is the Excel constant for .xlsx
        $workbook.Close()
        Write-Host "Conversion complete: $inputFilePath -> $xlsxFilePath"
    } catch {
        Write-Host "Error converting file '$inputFilePath' to .xlsx: $_"
    } finally {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
    }

    return $xlsxFilePath
}

# Function to detect header row and column
function Detect-Header {
    param (
        [string]$excelFilePath
    )

    $headerColumns = @('Date', 'Narration', 'Chq./Ref.No.', 'Value Dt', 'Withdrawal Amt.', 'Deposit Amt.', 'Closing Balance')

    $startRow = 0
    $startColumn = 0

    try {
        $excelPackage = Open-ExcelPackage -Path $excelFilePath
        foreach ($worksheet in $excelPackage.Workbook.Worksheets) {
            for ($rowIndex = 1; $rowIndex -le $worksheet.Dimension.End.Row; $rowIndex++) {
                $isRowBlank = $true
                for ($colIndex = 1; $colIndex -le $worksheet.Dimension.End.Column; $colIndex++) {
                    if ($worksheet.Cells[$rowIndex, $colIndex].Value -ne $null) {
                        $isRowBlank = $false
                        break
                    }
                }

                if ($isRowBlank) {
                    continue
                }

                for ($startColIndex = 1; $startColIndex -le 4; $startColIndex++) {
                    $headerMatch = $true
                    for ($colIndex = $startColIndex; $colIndex -le $headerColumns.Length + $startColIndex - 1; $colIndex++) {
                        $cellValue = $worksheet.Cells[$rowIndex, $colIndex].Value
                        if ($cellValue -ne $headerColumns[$colIndex - $startColIndex]) {
                            $headerMatch = $false
                            break
                        }
                    }

                    if ($headerMatch) {
                        $startRow = $rowIndex + 1
                        $startColumn = $startColIndex
                        break
                    }
                }

                if ($startRow -ne 0) {
                    break
                }
            }

            if ($startRow -ne 0) {
                break
            }
        }
    } catch {
        Write-Host "Error detecting header row in file '$excelFilePath': $_"
    } finally {
        $excelPackage.Dispose()
    }

    if ($startRow -eq 0) {
        Write-Host "Header row not found in file '$excelFilePath'. Please check the file format."
        return $null
    }

    return @{ StartRow = $startRow; StartColumn = $startColumn }
}

# Function to validate and convert dates from two-digit year format to four-digit year format
function Convert-DateToFourDigitYear {
    param (
        [string]$dateString,
        #[string]$separator = $null  # Default is null to retain the original separator
        [string]$separator = '-'  # Default is null to retain the original separator
    )

    # Adjust regex to support various separators (/, -, .)
    if ($dateString -match '(\d{1,2})([\/\-\.\s])(\d{1,2})\2(\d{2})') {
        $day = $matches[1]
        $originalSeparator = $matches[2]  # Capture the original separator
        $month = $matches[3]
        $yearPart = [int]$matches[4]

        $currentYear = [int](Get-Date).Year
        $currentCentury = [int]($currentYear / 100) * 100
        $cutoffYear = 50  # Cutoff year for 2021

        if ($yearPart -le $cutoffYear) {
            $yearPart = $currentCentury + $yearPart
        } else {
            $yearPart = ($currentCentury - 100) + $yearPart
        }

        # Use the provided separator, or fallback to the original separator
        $finalSeparator = if ($separator) { $separator } else { $originalSeparator }
        
        # Rebuild the date using the chosen separator
        return "$day$finalSeparator$month$finalSeparator$yearPart"
    }

    return $dateString  # Return the original date string if no match
}

# Function to transform data
# Function to transform data
function Transform-Data {
    param (
        [string]$inputFilePath,
        [int]$startRow,
        [int]$startColumn,
        [string]$MOP,
        [string]$transformedFilePath
    )

    $customHeaders = @('Date', 'Narration', 'Chq./Ref.No.', 'Value Dt', 'Withdrawal Amt.', 'Deposit Amt.', 'Closing Balance')

    try {
        $oldData = Import-Excel -Path $inputFilePath -StartRow $startRow -StartColumn $startColumn -HeaderName $customHeaders
        $newData = @()

        foreach ($row in $oldData) {
            $newRow = [PSCustomObject]@{
                Date = Convert-DateToFourDigitYear $row.Date
                Narration = $row.Narration
                Item = ""  # Default value
                Category = ""  # Default value
                Place = ""  # Default value
                Freq = ""  # Default value
                For = ""  # Default value
                MOP = $MOP  # Combined MOP from filename parts
                'Amt (Dr)' = $row.'Withdrawal Amt.'
                'Chq./Ref.No.' = $row.'Chq./Ref.No.'
                #'Value Dt' = $row.'Value Dt'
                'Value Dt' = Convert-DateToFourDigitYear $row.'Value Dt'
                'Amt (Cr)' = $row.'Deposit Amt.'
            }

            if ($row.Narration -match 'RESET') {
                $newRow.Item = "RESET"
                $newRow.Category = "RESET"
                $newRow.Place = "RESET"
                $newRow.Freq = "RESET"
                $newRow.For = "RESET"
            }

            $newData += $newRow
        }

        # Export the new data to Excel
        $newData | Export-Excel -Path $transformedFilePath -WorksheetName "FormattedData" -AutoSize

        # Set column widths for specific columns
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($transformedFilePath)
        $worksheet = $workbook.Sheets.Item("FormattedData")

        # Column widths
        $worksheet.Columns.Item(1).ColumnWidth = 10   # "Date"
        $worksheet.Columns.Item(2).ColumnWidth = 55   # "Narration"
        $worksheet.Columns.Item(3).ColumnWidth = 5    # "Item"
        $worksheet.Columns.Item(4).ColumnWidth = 5    # "Category"
        $worksheet.Columns.Item(5).ColumnWidth = 5    # "Place"
        $worksheet.Columns.Item(6).ColumnWidth = 5    # "Freq"
        $worksheet.Columns.Item(7).ColumnWidth = 5    # "For"
        $worksheet.Columns.Item(8).ColumnWidth = 10   # "Amt (Dr)"
        $worksheet.Columns.Item(9).ColumnWidth = 15   # "Chq./Ref.No."
        $worksheet.Columns.Item(10).ColumnWidth = 10  # "Value Dt"
        $worksheet.Columns.Item(11).ColumnWidth = 10  # "Amt (Cr)"

        # Save and close the workbook
        $workbook.Save()
        $workbook.Close()
        $excel.Quit()

        Write-Host "Transformation complete: $inputFilePath -> $transformedFilePath"
    } catch {
        Write-Host "Error transforming data in file '$inputFilePath': $_"
    } finally {
        # Clean up COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
    }
}

# Main script execution
try {
    # Process all XLS files in the current directory
    $xlsFiles = Get-ChildItem -Path . -Filter HDFC_DC_*.xls
    foreach ($xlsFile in $xlsFiles) {
        $xlsFilePath = $xlsFile.FullName
        $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($xlsFilePath)
        $convertedFilePath = "$($xlsFile.DirectoryName)\$baseFileName`_ConvertedFromXls.xlsx"
        $transformedFilePath = "$($xlsFile.DirectoryName)\$baseFileName`_ConvertedFromXls_Transformed.xlsx"

        try {
            # Convert .xls to .xlsx
            $xlsxFilePath = Convert-XlsToXlsx -inputFilePath $xlsFilePath

            # Process the newly created XLSX file
            $filenameParts = $baseFileName -split '_'
            $bank = $filenameParts[0]
            $accountType = $filenameParts[1]
            $name = $filenameParts[2]
            $MOP = "$bank`_$accountType`_$name"

            $headerInfo = Detect-Header -excelFilePath $xlsxFilePath
            if ($headerInfo -ne $null) {
                Transform-Data -inputFilePath $xlsxFilePath -startRow $headerInfo.StartRow -startColumn $headerInfo.StartColumn -MOP $MOP -transformedFilePath $transformedFilePath
            } else {
                Write-Host "Skipping transformation for file '$xlsxFilePath' due to header detection issues."
            }
        } catch {
            Write-Host "An error occurred while processing file '$xlsFilePath': $_"
        }
    }
} catch {
    Write-Error "An error occurred during script execution: $_"
}
