<#
===============================================================================
Employee Document Validation Utility
===============================================================================

Description
-----------
This utility validates employee records and document availability
before automated WhatsApp distribution.

Main Features
-------------
- Reads employee information from Excel
- Validates employee phone numbers
- Verifies document availability
- Detects employees without matching documents
- Detects documents without matching employees
- Prevents distribution errors
- Generates validation statistics

Privacy Notice
--------------
All employee names, phone numbers, file names, folder paths,
company information, and business-specific data contained in
this public repository are anonymized and provided for
demonstration purposes only.

===============================================================================
#>

Set-ExecutionPolicy -Scope Process Bypass

if (-not ("System.Windows.Forms.Form" -as [type])) {
    Add-Type -AssemblyName System.Windows.Forms
}

Import-Module ImportExcel

$ExcelPath = "C:\SampleCompany\Data\employee_list.xlsx"
$BasePath  = "C:\SampleCompany\Data"

# =====================================
# EXCEL FILE LOCK VALIDATION
# =====================================

try
{
    $Stream = [System.IO.File]::Open(
        $ExcelPath,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::ReadWrite,
        [System.IO.FileShare]::None
    )

    $Stream.Close()
}
catch
{
    [System.Windows.Forms.MessageBox]::Show(
        "employee_list.xlsx is currently open.`n`nClose the file and restart the application.",
        "Excel File Open"
    )
    exit
}

# =====================================
# RETRIEVE WEEK NUMBER
# =====================================

$ExcelPkg = Open-ExcelPackage $ExcelPath
$WeekNumber = $ExcelPkg.Workbook.Worksheets[1].Cells["G1"].Value
Close-ExcelPackage $ExcelPkg

$WeekFolder = Join-Path $BasePath "Week $WeekNumber"

Write-Host ""
Write-Host "Week Number : $WeekNumber"
Write-Host "Folder      : $WeekFolder"
Write-Host ""

# =====================================
# IMPORT EMPLOYEE DATA
# =====================================

$EmployeeData = Import-Excel $ExcelPath

# =====================================
# PHONE NUMBER VALIDATION
# =====================================

$InvalidPhoneNumbers = @()

foreach($Employee in $EmployeeData)
{
    if(-not $Employee.Name)
    {
        continue
    }

    $EmployeeName = $Employee.Name.Trim()

    $PhoneNumber = "$($Employee.PhoneNumber)"
    $PhoneNumber = $PhoneNumber.Replace("+","").Replace(" ","").Trim()

    if(-not $PhoneNumber)
    {
        $InvalidPhoneNumbers += "$EmployeeName - NO PHONE NUMBER"
    }
    elseif($PhoneNumber -notmatch '^\d{10,15}$')
    {
        $InvalidPhoneNumbers += "$EmployeeName - $PhoneNumber"
    }
}

# =====================================
# DOCUMENT VALIDATION
# =====================================

$MatchedDocuments = 0
$MissingDocuments = 0

foreach($Employee in $EmployeeData)
{
    if(-not $Employee.Name)
    {
        continue
    }

    $EmployeeName = $Employee.Name.Trim()

    $SearchName = ($EmployeeName -replace '\s+',' ').Trim()

    $Document = Get-ChildItem $WeekFolder -Recurse -File |
                Where-Object {
                    (($_.BaseName -replace '\s+',' ').Trim()) -like "*$SearchName*"
                } |
                Select-Object -First 1

    if($Document)
    {
        Write-Host "$EmployeeName >>> $($Document.Name)" -ForegroundColor Green
        $MatchedDocuments++
    }
    else
    {
        Write-Host "$EmployeeName >>> DOCUMENT NOT FOUND" -ForegroundColor Yellow
        $MissingDocuments++
    }
}

# =====================================
# DOCUMENTS WITHOUT EMPLOYEES
# =====================================

$DocumentsWithoutEmployees = @()

$PdfFiles = Get-ChildItem $WeekFolder -Recurse -Filter *.pdf

foreach($PdfFile in $PdfFiles)
{
    $PdfName = $PdfFile.BaseName

    $PdfName = $PdfName `
        -replace '\(.*?\)','' `
        -replace '_',' ' `
        -replace '-',' ' `
        -replace '\s+',' '

    $PdfName = $PdfName.Trim()

    $Found = $false

    foreach($Employee in $EmployeeData)
    {
        if(-not $Employee.Name)
        {
            continue
        }

        $EmployeeName = ($Employee.Name -replace '\s+',' ').Trim()

        if($PdfName -like "*$EmployeeName*")
        {
            $Found = $true
            break
        }
    }

    if(-not $Found)
    {
        $DocumentsWithoutEmployees += $PdfFile.Name
    }
}

# =====================================
# VALIDATION RESULTS
# =====================================

if($InvalidPhoneNumbers.Count -gt 0)
{
    Write-Host ""
    Write-Host "Invalid phone numbers:" -ForegroundColor Red
    Write-Host ""

    $InvalidPhoneNumbers |
    Sort-Object -Unique |
    ForEach-Object {
        Write-Host " - $_" -ForegroundColor Red
    }
}

if($DocumentsWithoutEmployees.Count -gt 0)
{
    Write-Host ""
    Write-Host "The following documents do not match any employee in employee_list.xlsx:" -ForegroundColor Red
    Write-Host ""

    $DocumentsWithoutEmployees |
    Sort-Object -Unique |
    ForEach-Object {
        Write-Host " - $_" -ForegroundColor Red
    }
}

# =====================================
# SUMMARY
# =====================================

Write-Host ""
Write-Host "--------------------------------------------"
Write-Host "SUMMARY"
Write-Host "--------------------------------------------"
Write-Host "Matched documents      : $MatchedDocuments"
Write-Host "Missing documents      : $MissingDocuments"
Write-Host "Documents not in Excel : $($DocumentsWithoutEmployees.Count)"
Write-Host "Invalid phone numbers  : $($InvalidPhoneNumbers.Count)"
Write-Host "--------------------------------------------"

if(($DocumentsWithoutEmployees.Count -eq 0) -and ($InvalidPhoneNumbers.Count -eq 0))
{
    Write-Host ""
    Write-Host "Validation completed successfully." -ForegroundColor Green
}

Write-Host ""
Read-Host "Press Enter to exit"