# ====================================================================
# Clean_Old_Excel_Addins.ps1 - Dọn dẹp triệt để các Add-in Excel cũ bị lỗi
# ====================================================================

Write-Host "1. Dang dung cac tien trinh Excel..." -ForegroundColor Yellow
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force

$addinsFolder = "$env:APPDATA\Microsoft\AddIns"
$xlstartFolder = "$env:APPDATA\Microsoft\Excel\XLSTART"

Write-Host "2. Dinh cau truc Add-ins trong APPDATA..." -ForegroundColor Yellow
if (Test-Path $addinsFolder) {
    Get-ChildItem -Path $addinsFolder -Filter "*.xlam" -ErrorAction SilentlyContinue | ForEach-Object {
        Write-Host "   - Xoa add-in cu: $($_.FullName)" -ForegroundColor Gray
        Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue
    }
}

if (Test-Path $xlstartFolder) {
    Get-ChildItem -Path $xlstartFolder -Filter "*.xlam" -ErrorAction SilentlyContinue | ForEach-Object {
        Write-Host "   - Xoa add-in trong XLSTART: $($_.FullName)" -ForegroundColor Gray
        Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue
    }
}

Write-Host "3. Don dep Registry Excel Options (Add-in Manager)..." -ForegroundColor Yellow
$excelVersions = @("16.0", "15.0", "14.0", "12.0")

foreach ($v in $excelVersions) {
    $regPath = "HKCU:\Software\Microsoft\Office\$v\Excel\Options"
    if (Test-Path $regPath) {
        $props = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        foreach ($prop in $props.PSObject.Properties) {
            if ($prop.Name -like "OPEN*") {
                Write-Host "   - Clear key Registry $regPath\$($prop.Name) = $($prop.Value)" -ForegroundColor Gray
                Remove-ItemProperty -Path $regPath -Name $prop.Name -ErrorAction SilentlyContinue
            }
        }
    }

    $mgrPath = "HKCU:\Software\Microsoft\Office\$v\Excel\Add-in Manager"
    if (Test-Path $mgrPath) {
        Write-Host "   - Resetting Add-in Manager at $mgrPath" -ForegroundColor Gray
        Remove-Item -Path $mgrPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host "4. Dang ki lai Add-in moi duy nhat tu C:\Users\Administrator\Downloads\NXTiengViet.xlam..." -ForegroundColor Green
$newAddinPath = "C:\Users\Administrator\Downloads\NXTiengViet.xlam"
if (Test-Path $newAddinPath) {
    # Set OPEN key in Excel 16.0 Options
    $regPath16 = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
    if (-not (Test-Path $regPath16)) {
        New-Item -Path $regPath16 -Force | Out-Null
    }
    Set-ItemProperty -Path $regPath16 -Name "OPEN" -Value "`"$newAddinPath`"" -ErrorAction SilentlyContinue
    Write-Host "   [OK] Da dang ky duy nhat file add-in moi: $newAddinPath" -ForegroundColor Green
}

Write-Host "`n=== HOAN TAT DON DEP TRIET DE ADD-IN CU! ===" -ForegroundColor Green
