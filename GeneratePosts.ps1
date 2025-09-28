<#
.SYNOPSIS
    Facebook Marketplace Post Generator with Parameters

.DESCRIPTION
    This script automates the creation of for sale posts for Facebook Marketplace.

.PARAMETER CsvPath

.PARAMETER ProfileName


.INPUTS
    CSV file

.OUTPUTS

.EXAMPLE
    .\GeneratePosts.ps1 -InputFolder "C:\path\toinput" -OutputFolder "C:\path\tooutput"

    .EXAMPLE
    .\GeneratePosts.ps1 `
    -InputFolder "C:\Users\JohnO\OneDrive\Heartland Auction" `
    -OutputFolder "C:\MarketplacePosts" `
    -TemplateFile "fb_template.txt" `
    -CsvFile "lots_to_sell.csv"
    
.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Date: 09/28/2025
    Version: 1.0
    Change Purpose: 

    Prerequisites:
                    PowerShell 5.1 or later
    
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$InputFolder = ".",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFolder = "marketplace_posts",
    
    [Parameter(Mandatory=$false)]
    [string]$TemplateFile = "post_template.txt",
    
    [Parameter(Mandatory=$false)]
    [string]$CsvFile = "inventory.csv"
)

# Display parameters being used
Write-Host "========== Facebook Marketplace Post Generator ==========" -ForegroundColor Cyan
Write-Host "Input Folder:    $InputFolder" -ForegroundColor White
Write-Host "Output Folder:   $OutputFolder" -ForegroundColor White
Write-Host "Template File:   $TemplateFile" -ForegroundColor White
Write-Host "CSV File:        $CsvFile" -ForegroundColor White
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host ""

# Resolve full paths
$InputFolder = Resolve-Path $InputFolder -ErrorAction SilentlyContinue
if (-not $InputFolder) {
    Write-Host "ERROR: Input folder does not exist!" -ForegroundColor Red
    exit 1
}

$TEMPLATE_PATH = Join-Path $InputFolder $TemplateFile
$CSV_PATH = Join-Path $InputFolder $CsvFile
$OUTPUT_PATH = $OutputFolder

# Handle relative vs absolute output path
if (-not [System.IO.Path]::IsPathRooted($OutputFolder)) {
    $OUTPUT_PATH = Join-Path $InputFolder $OutputFolder
}

# Create output folder if it doesn't exist
if (!(Test-Path $OUTPUT_PATH)) {
    New-Item -ItemType Directory -Path $OUTPUT_PATH | Out-Null
    Write-Host "Created output folder: $OUTPUT_PATH" -ForegroundColor Green
} else {
    Write-Host "Using existing output folder: $OUTPUT_PATH" -ForegroundColor Yellow
}

# Check if template file exists, create default if not
if (!(Test-Path $TEMPLATE_PATH)) {
    Write-Host "Template file not found at: $TEMPLATE_PATH" -ForegroundColor Yellow
    Write-Host "Creating default template..." -ForegroundColor Yellow
    
    $defaultTemplate = @"
FOR SALE: {Description}

Brand New - Never Installed!
Model: {ModelNo}
Lot #: {LotNo}

✅ Purchased from commercial RV manufacturer auction
✅ Perfect for RV upgrade, repair, or new build
✅ Professional grade equipment
✅ Can deliver locally for small fee

💰 Priced to sell fast - well below retail!
📦 Multiple units available - ask about bundle deals

Contact: {ContactPhone}
Text for fastest response

Cash or Zelle accepted
Available for pickup in Jefferson/Ashtabula area

#RVParts #RVLife #TrailerParts #RVUpgrade
"@
    
    $defaultTemplate | Out-File -FilePath $TEMPLATE_PATH -Encoding UTF8
    Write-Host "Created default template: $TEMPLATE_PATH" -ForegroundColor Green
}

# Read template
try {
    $template = Get-Content -Path $TEMPLATE_PATH -Raw -Encoding UTF8
    Write-Host "Template loaded successfully from: $TEMPLATE_PATH" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Could not read template file!" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

# Check if CSV file exists
if (!(Test-Path $CSV_PATH)) {
    Write-Host "ERROR: CSV file not found at: $CSV_PATH" -ForegroundColor Red
    
    # Create sample CSV for reference
    $samplePath = Join-Path $InputFolder "sample_inventory.csv"
    $sampleCSV = @"
LotNo,ModelNo,Description,ContactPhone,RetailPrice,AskingPrice
1601,Dexter 417167,Heavy Duty 6000lb RV Trailer Axle,(440) 813-6695,800,400
3,Furrion 2022120380,Arctic RV French Door Refrigerator 19.6 cu ft,(440) 813-6695,2400,1200
16,FLC03ACAFE-SP,Furrion 2.7 cu ft Washer Dryer Combo Ventless,(440) 813-6695,1800,900
"@
    
    $sampleCSV | Out-File -FilePath $samplePath -Encoding UTF8
    Write-Host "Created sample CSV file: $samplePath" -ForegroundColor Yellow
    Write-Host "Please create your $CsvFile with columns: LotNo, ModelNo, Description, ContactPhone" -ForegroundColor Yellow
    exit 1
}

# Read and process CSV
try {
    $csvData = Import-Csv -Path $CSV_PATH
    Write-Host "CSV loaded: Found $($csvData.Count) rows" -ForegroundColor Green
    
    if ($csvData.Count -eq 0) {
        Write-Host "WARNING: CSV file is empty!" -ForegroundColor Yellow
        exit 1
    }
    
    # Check for required columns
    $requiredColumns = @('LotNo', 'ModelNo', 'Description', 'ContactPhone')
    $csvColumns = $csvData[0].PSObject.Properties.Name
    
    Write-Host "CSV Columns found: $($csvColumns -join ', ')" -ForegroundColor Cyan
    
    $missingColumns = $requiredColumns | Where-Object { $_ -notin $csvColumns }
    if ($missingColumns) {
        Write-Host "WARNING: Missing recommended columns: $($missingColumns -join ', ')" -ForegroundColor Yellow
    }
    
    # Process each row
    $postsCreated = 0
    $skippedRows = 0
    
    foreach ($row in $csvData) {
        # Skip if no LotNo
        if ([string]::IsNullOrWhiteSpace($row.LotNo)) {
            $skippedRows++
            continue
        }
        
        # Start with template
        $post = $template
        
        # Replace placeholders with actual values
        foreach ($property in $row.PSObject.Properties) {
            $placeholder = "{$($property.Name)}"
            $value = if ([string]::IsNullOrWhiteSpace($property.Value)) { 
                "Not specified" 
            } else { 
                $property.Value.Trim() 
            }
            
            $post = $post.Replace($placeholder, $value)
        }
        
        # Handle any remaining placeholders (for columns not in CSV)
        $post = $post -replace '\{[^}]+\}', 'Not specified'
        
        # Create filename (remove invalid characters)
        $filename = "Lot_$($row.LotNo -replace '[^\w\-]', '_').txt"
        $filepath = Join-Path $OUTPUT_PATH $filename
        
        # Save the post
        $post | Out-File -FilePath $filepath -Encoding UTF8
        
        Write-Host "Created: $filename" -ForegroundColor Cyan
        $postsCreated++
    }
    
    # Summary
    Write-Host "`n========== COMPLETE ==========" -ForegroundColor Green
    Write-Host "Posts created: $postsCreated" -ForegroundColor Green
    if ($skippedRows -gt 0) {
        Write-Host "Rows skipped (no LotNo): $skippedRows" -ForegroundColor Yellow
    }
    Write-Host "Output folder: $OUTPUT_PATH" -ForegroundColor Green
    
    # Ask if user wants to open the output folder
    $openFolder = Read-Host "`nOpen output folder? (Y/N)"
    if ($openFolder -eq 'Y' -or $openFolder -eq 'y') {
        Start-Process explorer.exe $OUTPUT_PATH
    }
}
catch {
    Write-Host "ERROR processing CSV: $_" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")