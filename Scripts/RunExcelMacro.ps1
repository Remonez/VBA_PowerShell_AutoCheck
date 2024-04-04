# Define the path to the Excel file and the macro name
$excelFile = "C:\Path\To\Your\File.xlsm"
$macroName = "CheckExcelFile"

# Create a new instance of Excel application
$excel = New-Object -ComObject Excel.Application

# Hide Excel application window
$excel.Visible = $false

# Open the workbook
$workbook = $excel.Workbooks.Open($excelFile)

# Run the specified macro
$excel.Run($macroName)

# Close the workbook without saving changes
$workbook.Close($false)

# Quit Excel application
$excel.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Wait for a moment to ensure Excel closes properly
Start-Sleep -Seconds 1
