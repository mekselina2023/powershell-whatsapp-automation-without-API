<#
===============================================================================
PowerShell Employee WhatsApp Automation Without API
===============================================================================

Description
-----------
This project demonstrates how personalized employee notifications and
document delivery can be automated using PowerShell, Excel, and
WhatsApp Desktop without requiring paid messaging APIs.

Main Features
-------------
- Reads employee information from Excel
- Retrieves week-based document folders
- Validates employee records
- Verifies document availability
- Matches employees with their corresponding files
- Sends personalized WhatsApp messages
- Attaches employee-specific documents
- Generates delivery logs
- No paid API required

Workflow
--------
1. Read employee data from Excel
2. Retrieve the active week number
3. Validate employee records
4. Locate matching documents
5. Open WhatsApp Desktop
6. Send personalized messages
7. Attach employee-specific files
8. Record results in log files

Privacy Notice
--------------
All employee names, phone numbers, company information,
folder paths, and file names contained in this public
repository are anonymized and provided for demonstration
purposes only.

Technologies
------------
- PowerShell
- Excel
- ImportExcel Module
- WhatsApp Desktop
- Windows Forms

===============================================================================
#>

Set-ExecutionPolicy -Scope Process Bypass

Add-Type -AssemblyName System.Windows.Forms

# =====================================
# CONFIGURATION
# =====================================

$ExcelPath = "C:\SampleCompany\Data\employee_list.xlsx"
$BasePath  = "C:\SampleCompany\Data"

$LogFolder = Join-Path $BasePath "Logs"
$LogFile   = Join-Path $LogFolder ("Log-" + (Get-Date -Format "yyyy-MM-dd") + ".txt")

# =====================================
# STARTUP VALIDATION
# =====================================

if (-not (Get-Module -ListAvailable -Name ImportExcel))
{
    [System.Windows.Forms.MessageBox]::Show(
        "ImportExcel module not found.",
        "Error"
    )
    exit
}

Import-Module ImportExcel -ErrorAction Stop

if (!(Test-Path $ExcelPath))
{
    [System.Windows.Forms.MessageBox]::Show(
        "Excel file not found:`n`n$ExcelPath",
        "Error"
    )
    exit
}

if (!(Test-Path $LogFolder))
{
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

# =====================================
# RETRIEVE WEEK NUMBER
# =====================================

$ExcelPkg = Open-ExcelPackage $ExcelPath
$WeekNumber = $ExcelPkg.Workbook.Worksheets[1].Cells["G1"].Value
Close-ExcelPackage $ExcelPkg

if(:IsNullOrWhiteSpace($WeekNumber))
{
    [System.Windows.Forms.MessageBox]::Show(
        "Cell G1 does not contain a week number.",
        "Error"
    )
    exit
}

# =====================================
# EMPLOYEE DATA
# =====================================

$EmployeeData = Import-Excel $ExcelPath

$EmployeeData = $EmployeeData | Where-Object {
    $_.Name
}

# =====================================
# WEEK FOLDER
# =====================================

$WeekFolder = Join-Path $BasePath "Week $WeekNumber"

if (!(Test-Path $WeekFolder))
{
    [System.Windows.Forms.MessageBox]::Show(
        "Week folder not found:`n`n$WeekFolder",
        "Error"
    )
    exit
}

Write-Host "Week Number : $WeekNumber"
Write-Host "Folder      : $WeekFolder"

# =====================================
# DOCUMENT VALIDATION
# =====================================

$MissingEmployees = @()

$PdfFiles = Get-ChildItem $WeekFolder -Recurse -Filter *.pdf

foreach($PdfFile in $PdfFiles)
{
    $PdfName = ($PdfFile.BaseName -replace '\s+',' ').Trim()

    $MatchedEmployee = $EmployeeData | Where-Object {

        $EmployeeName = ($_.Name -replace '\s+',' ').Trim()

        $PdfName -like "*$EmployeeName*" -or
        $EmployeeName -like "*$PdfName*"
    }

    if(-not $MatchedEmployee)
    {
        $MissingEmployees += $PdfName
    }
}

if($MissingEmployees.Count -gt 0)
{
    $List = ($MissingEmployees |
        Sort-Object -Unique) -join "`n"

    [System.Windows.Forms.MessageBox]::Show(
        "The following people exist in the PDF files but not in Excel:`n`n" +
        $List +
        "`n`nPlease add them to Excel first.",
        "Missing Employees"
    )

    exit
}

# =====================================
# MESSAGE DELIVERY
# =====================================

foreach($Employee in $EmployeeData)
{
    $EmployeeName = $Employee.Name.Trim()

    $PhoneNumber = "$($Employee.PhoneNumber)"
    $PhoneNumber = $PhoneNumber.Replace("+","").Replace(" ","").Trim()

    $Message = $Employee.Message

    $SearchName = ($EmployeeName -replace '\s+',' ').Trim()

    $Document = Get-ChildItem $WeekFolder -Recurse -File |
                Where-Object {
                    (($_.BaseName -replace '\s+',' ').Trim()) -like "*$SearchName*"
                } |
                Select-Object -First 1

    if(-not $Document)
    {
        Write-Host "$EmployeeName >>> DOCUMENT NOT FOUND" -ForegroundColor Yellow

        Add-Content -Path $LogFile -Value "$EmployeeName | $PhoneNumber | DOCUMENT NOT FOUND | $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

        continue
    }

    Write-Host "$EmployeeName >>> $($Document.Name)" -ForegroundColor Green

    $ClipboardFiles = New-Object System.Collections.Specialized.StringCollection
    $ClipboardFiles.Add($Document.FullName)

    [System.Windows.Forms.Clipboard]::SetFileDropList($ClipboardFiles)

    Start-Process "whatsapp://send?phone=$PhoneNumber&text=$Message"

    Start-Sleep 2
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

    Start-Sleep 2
    [System.Windows.Forms.SendKeys]::SendWait("^v")

    Start-Sleep 2
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

    Start-Sleep 5

    Add-Content -Path $LogFile -Value "$EmployeeName | $PhoneNumber | $($Document.Name) | $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
}

Write-Host ""
Write-Host "--------------------------------------------"
Write-Host "Process completed."
Write-Host "Log file saved to: $LogFile"
Write-Host "--------------------------------------------"

Read-Host "Press Enter to exit"