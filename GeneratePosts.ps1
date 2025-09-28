<#
.SYNOPSIS
    Facebook Marketplace Post Generator with Parameters

.DESCRIPTION
    This script automates the creation of for sale posts for Facebook Marketplace.
    It reads a CSV file with item information and a template file, then generates
    individual post files for each item. It can also copy associated photos and
    create an HTML preview of all posts.

.PARAMETER InputFolder
    The folder where the input CSV and template files are located. Default is the current directory.

.PARAMETER OutputFolder
    The folder where the generated posts will be saved. Default is "marketplace_posts" in the input folder.

.PARAMETER TemplateFile
    The name of the template file. Default is "post_template.txt".

.PARAMETER CsvFile
    The name of the CSV file containing item data. Default is "inventory.csv".

.PARAMETER PhotosFolder
    The folder where photos are stored. If empty, uses InputFolder. Default is empty.

.INPUTS
    CSV file with columns: LotNo, ModelNo, Description, ContactPhone (minimum)
    Template file with placeholders like {LotNo}, {Description}, etc.
    Optional: Photos named with lot numbers (e.g., 1601.jpg)

.OUTPUTS
    Individual text files for each post (Lot_XXXX.txt)
    Copied photos with matching names
    HTML preview file (preview.html)

.EXAMPLE
    .\GeneratePosts.ps1 -InputFolder "C:\path\to\input" -OutputFolder "C:\path\to\output"

.EXAMPLE
    .\GeneratePosts.ps1 `
        -InputFolder "C:\ItemsToSell" `
        -OutputFolder "C:\MarketplacePosts" `
        -TemplateFile "posts_template.txt" `
        -CsvFile "items_to_sell.csv" `
        -PhotosFolder "C:\ItemPhotos"
    
.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Date: 09/28/2025
    Version: 1.1
    Change Purpose: Added PhotosFolder parameter, fixed photo path bug, enhanced HTML preview

    Prerequisites: PowerShell 5.1 or later
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$InputFolder = ".",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFolder = "marketplace_posts",
    
    [Parameter(Mandatory=$false)]
    [string]$TemplateFile = "post_template.txt",
    
    [Parameter(Mandatory=$false)]
    [string]$CsvFile = "inventory.csv",

    [Parameter(Mandatory=$false)]
    [string]$PhotosFolder = ""  # If empty, uses InputFolder
)

# Display parameters being used
Write-Host "========== Facebook Marketplace Post Generator ==========" -ForegroundColor Cyan
Write-Host "Input Folder:    $InputFolder" -ForegroundColor White
Write-Host "Output Folder:   $OutputFolder" -ForegroundColor White
Write-Host "Template File:   $TemplateFile" -ForegroundColor White
Write-Host "CSV File:        $CsvFile" -ForegroundColor White
if ($PhotosFolder) {
    Write-Host "Photos Folder:   $PhotosFolder" -ForegroundColor White
} else {
    Write-Host "Photos Folder:   (Using Input Folder)" -ForegroundColor Gray
}
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host ""

# Resolve full paths
$InputFolder = Resolve-Path $InputFolder -ErrorAction SilentlyContinue
if (-not $InputFolder) {
    Write-Host "ERROR: Input folder does not exist!" -ForegroundColor Red
    exit 1
}

# Resolve photos folder if specified
if ($PhotosFolder) {
    $PhotosFolder = Resolve-Path $PhotosFolder -ErrorAction SilentlyContinue
    if (-not $PhotosFolder) {
        Write-Host "WARNING: Photos folder does not exist! Will use Input Folder instead." -ForegroundColor Yellow
        $PhotosFolder = $InputFolder
    }
} else {
    $PhotosFolder = $InputFolder
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
1601,Dexter 417167,Heavy Duty 6000lb RV Trailer Axle,(440) 813-4765,800,400
3,Furrion 2022120380,Arctic RV French Door Refrigerator 19.6 cu ft,(440) 813-4765,2400,1200
16,FLC03ACAFE-SP,Furrion 2.7 cu ft Washer Dryer Combo Ventless,(440) 813-4765,1800,900
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
    $photosFound = 0
    
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
        $lotNumber = $row.LotNo -replace '[^\w\-]', '_'
        $filename = "Lot_$lotNumber.txt"
        $filepath = Join-Path $OUTPUT_PATH $filename
        
        # Save the post
        $post | Out-File -FilePath $filepath -Encoding UTF8
        
        # Check if corresponding photo exists and copy it
        $photoExtensions = @('.jpg', '.jpeg', '.png', '.gif', '.bmp')
        $photoFound = $false
        
        foreach ($ext in $photoExtensions) {
            $photoPath = Join-Path $PhotosFolder "$($row.LotNo)$ext"
            if (Test-Path $photoPath) {
                $photoDestination = Join-Path $OUTPUT_PATH "Lot_$lotNumber$ext"
                Copy-Item -Path $photoPath -Destination $photoDestination -Force
                Write-Host "Created: $filename (with photo)" -ForegroundColor Green
                $photoFound = $true
                $photosFound++
                break
            }
        }
        
        if (-not $photoFound) {
            Write-Host "Created: $filename (no photo found)" -ForegroundColor Cyan
        }
        
        $postsCreated++
    }
    
    # Generate HTML preview
    Write-Host "`nGenerating HTML preview..." -ForegroundColor Yellow
    
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Marketplace Posts Preview</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f0f2f5;
            padding: 20px;
        }
        h1 {
            color: #1877f2;
            text-align: center;
        }
        .summary {
            background: white;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .container {
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
        }
        .post {
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            padding: 15px;
            width: calc(50% - 10px);
            box-sizing: border-box;
        }
        .photo {
            width: 100%;
            max-width: 300px;
            height: auto;
            border-radius: 4px;
            margin-bottom: 15px;
        }
        .text {
            white-space: pre-wrap;
            font-size: 14px;
            line-height: 1.5;
            color: #333;
        }
        .no-photo {
            background: #f0f2f5;
            padding: 50px;
            text-align: center;
            color: #65676b;
            border-radius: 4px;
            margin-bottom: 15px;
        }
        @media (max-width: 768px) {
            .post {
                width: 100%;
            }
        }
    </style>
</head>
<body>
    <h1>Facebook Marketplace Posts Preview</h1>
    <div class="summary">
        <strong>Total Posts:</strong> $postsCreated | 
        <strong>With Photos:</strong> $photosFound | 
        <strong>Without Photos:</strong> $($postsCreated - $photosFound)
    </div>
    <div class="container">
"@

    foreach ($file in Get-ChildItem "$OUTPUT_PATH\*.txt" | Sort-Object Name) {
        $lotName = $file.BaseName
        $textContent = Get-Content $file.FullName -Raw
        $photoFile = Get-ChildItem "$OUTPUT_PATH\$lotName.*" | Where-Object { $_.Extension -match '\.(jpg|jpeg|png|gif|bmp)' } | Select-Object -First 1
        
        $htmlContent += "<div class='post'>"
        if ($photoFile) {
            $htmlContent += "<img src='$($photoFile.Name)' class='photo' alt='$lotName'/>"
        } else {
            $htmlContent += "<div class='no-photo'>No Photo Available</div>"
        }
        $htmlContent += "<pre class='text'>$textContent</pre></div>"
    }

    $htmlContent += @"
    </div>
</body>
</html>
"@

    $htmlContent | Out-File -FilePath "$OUTPUT_PATH\preview.html" -Encoding UTF8
    Write-Host "HTML preview created: preview.html" -ForegroundColor Green
    
    # Summary
    Write-Host "`n========== COMPLETE ==========" -ForegroundColor Green
    Write-Host "Posts created:   $postsCreated" -ForegroundColor Green
    Write-Host "Photos copied:   $photosFound" -ForegroundColor Green
    if ($skippedRows -gt 0) {
        Write-Host "Rows skipped (no LotNo): $skippedRows" -ForegroundColor Yellow
    }
    Write-Host "Output folder:   $OUTPUT_PATH" -ForegroundColor Green
    
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
