param(
    [Parameter(Mandatory)]
    [string]$Folder,

    [Parameter(Mandatory)]
    [string]$Pattern,

    [switch]$Csv,

    [string]$CsvPath = ".\SearchResult.csv"
)

# ImportExcel確認
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    throw "ImportExcelモジュールがインストールされていません。Install-Module ImportExcel"
}

$results = [System.Collections.Generic.List[object]]::new()

$excelFiles = Get-ChildItem `
    -Path $Folder `
    -File `
    -Recurse `
    -Include *.xlsx,*.xlsm

$total = $excelFiles.Count
$current = 0

foreach ($file in $excelFiles) {

    $current++

    Write-Verbose ("[{0}/{1}] {2}" -f `
        $current,
        $total,
        $file.FullName)

    try {

        $sheetInfos = Get-ExcelSheetInfo -Path $file.FullName

        foreach ($sheet in $sheetInfos) {

            Write-Verbose ("  Sheet: {0}" -f $sheet.Name)

            $rows = Import-Excel `
                -Path $file.FullName `
                -WorksheetName $sheet.Name `
                -NoHeader

            if (-not $rows) {
                continue
            }

            $rowNo = 0

            foreach ($row in $rows) {

                $rowNo++

                foreach ($property in $row.PSObject.Properties) {

                    $value = $property.Value

                    if ($null -eq $value) {
                        continue
                    }

                    $text = [string]$value

                    if ([string]::IsNullOrWhiteSpace($text)) {
                        continue
                    }

                    if ($text -match $Pattern) {

                        $colNumber =
                            [int]($property.Name -replace 'P','')

                        $colName = ""

                        $n = $colNumber

                        while ($n -gt 0) {
                            $n--
                            $colName =
                                [char](65 + ($n % 26)) + $colName
                            $n = [math]::Floor($n / 26)
                        }

                        $address = "$colName$rowNo"

                        $result = [PSCustomObject]@{
                            FilePath    = $file.FullName
                            SheetName   = $sheet.Name
                            CellAddress = $address
                            Value       = $text
                        }

                        $results.Add($result)

                        if (-not $Csv) {
                            "{0}:{1}!{2}" -f `
                                $file.FullName,
                                $sheet.Name,
                                $address
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-Warning ("処理失敗: {0}" -f $file.FullName)
        Write-Warning $_.Exception.Message
    }
}

if ($Csv) {

    $results |
        Export-Csv `
            -Path $CsvPath `
            -NoTypeInformation `
            -Encoding UTF8

    Write-Host "CSV出力完了: $CsvPath"
}
