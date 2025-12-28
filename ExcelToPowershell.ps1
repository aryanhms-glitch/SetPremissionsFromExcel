param(
    [string]$BasePath = "C:\",  # this is the base path , the path in the excel goes after this .
    [switch]$WhatIf,
    [switch]$Backup
)


$ImportCsvPath = "C:\projectTemplate\ProjectsTemplate.csv"  # path of the generated CSV file


$ExportCsvPath = "C:\projectTemplate\PermissionsExport.csv"  # This is the path of the exported applied premissions , This will shows what changes are made


$csvData = Import-Csv -Path $ImportCsvPath -Encoding UTF8 -Header Column1,Column2,Column3,Column4,Column5,Column6,Column7,Column8,Column9,Column10,Column11,Column12,Column13,Column14,Column15,Column16,Column17,Column18,Column19,Column20 | Select-Object -Skip 0


$usersRow = $csvData[2]
$users = @()
$colIndex = 4  
while ($colIndex -le 20) {  
    $userProp = "Column$colIndex"
    $user = $usersRow.$userProp
    if (-not $user -or $user.Trim() -eq "") { break }
    $users += $user.Trim()
    $colIndex++
}

Write-Host "Found $($users.Count) users: $($users -join ', ')" -ForegroundColor Yellow
if ($users.Count -eq 0) {
    Write-Error "No users found! Check row 3 column D+ in Excel/CSV."
    exit
}


$permissions = @()
for ($i = 3; $i -lt $csvData.Count; $i++) {  
    $row = $csvData[$i]
    $folderName = $row.Column2  
    if (-not $folderName -or $folderName.Trim() -eq "") { continue }

    $fullPath = "$BasePath$($folderName.Trim())"

    Write-Host "Processing folder: $fullPath" -ForegroundColor Magenta

    $colIndex = 4  
    foreach ($user in $users) {
        $permProp = "Column$colIndex"
        $permCell = $row.$permProp
        if ($permCell -and $permCell.Trim() -ne "") {
            $perm = $permCell.Trim().ToUpper()
            $rights = switch ($perm) {
                "R" { "ReadAndExecute,ListDirectory" }                                                                 
                "W" { "ReadAndExecute,ListDirectory,Write" }                                                           
                "M" { "Modify" }                                                                                       

                default { 
                    Write-Warning "Invalid permission '$perm' for $user on $fullPath - skipping"
                    continue 
                }
            }

            $permissions += [PSCustomObject]@{
                FolderPath   = $fullPath
                Identity     = $user
                AccessRights = $rights
                AccessType   = "Allow"
            }

            Write-Host "  - $user = $perm ($rights)" -ForegroundColor Cyan
        }
        $colIndex++
    }
}  

# Check if the export CSV file already exists and delete it if it does
if (Test-Path $ExportCsvPath) {
    Write-Host "The file 'PermissionsExport.csv' already exists. Deleting it..." -ForegroundColor Red
    Remove-Item $ExportCsvPath -Force
}


foreach ($p in $permissions) {
    if (-not (Test-Path $p.FolderPath)) {
        Write-Warning "Folder not found: $($p.FolderPath) - Skipping"
        continue
    }

    if ($Backup) {
        $backupFile = "Backup_$(Split-Path $p.FolderPath -Leaf)_$(Get-Date -Format yyyyMMdd_HHmm).csv"
        (Get-Acl $p.FolderPath).Access | Export-Csv $backupFile -NoTypeInformation
        Write-Host "Backup: $backupFile"
    }

    $acl = Get-Acl $p.FolderPath

    # Remove any old permissions
    $rulesToRemove = $acl.Access | Where-Object {
        $_.AccessControlType -eq "Allow" -and 
        $_.IsInherited -eq $false -and 
        ($users | Where-Object { $_.IdentityReference.Value -match [regex]::Escape($_) })
    }
    foreach ($rule in $rulesToRemove) {
        $acl.RemoveAccessRule($rule)
        Write-Host "Removed old permission for $($rule.IdentityReference) on $($p.FolderPath)" -ForegroundColor Yellow
    }

    # Add new permission
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($p.Identity, $p.AccessRights, "ContainerInherit,ObjectInherit", "None", $p.AccessType)
    $acl.AddAccessRule($rule)

    Set-Acl -Path $p.FolderPath -AclObject $acl -WhatIf:$WhatIf

    Write-Host "APPLIED (overwrite): $($p.Identity) -> $($p.AccessRights) on $($p.FolderPath)" -ForegroundColor Green
}

# Export the permissions to a separate CSV file
$permissions | Export-Csv -Path $ExportCsvPath -NoTypeInformation -Encoding UTF8
Write-Host "Permissions exported to $ExportCsvPath"
