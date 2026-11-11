<#
    build.ps1  -  Dong goi add-in NXTiengViet.xlam (chi cap Tieu hoc)

    Vienam:
        powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1
        powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Test

    -Test : sau khi build se mo add-in bang mot instance Excel rieng, tao
            workbook 3 sheet (Toan / Tieng Viet / TN-XH), chay NX_Run o che do
            khong hien hop thoai, roi kiem tra cot Nhan xet duoc dien va cac
            cot khac khong bi thay doi.
#>
[CmdletBinding()]
param(
    [string]$Root,
    [string]$Python,
    [string]$DataSource,
    [switch]$Test,
    [switch]$SkipInject,
    [switch]$KeepTemp,
    [switch]$Deploy,
    [string]$DeployTo
)

if ([string]::IsNullOrEmpty($Root)) {
    if ($PSScriptRoot) { $Root = $PSScriptRoot }
    elseif ($MyInvocation.MyCommand.Path) { $Root = Split-Path -Parent $MyInvocation.MyCommand.Path }
    else { $Root = (Get-Location).Path }
}

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

#================================================================== duong dan
$srcDir   = Join-Path $Root 'src'
$dataDir  = Join-Path $Root 'data'
$uiFile   = Join-Path $Root 'customUI\customUI.xml'
$toolsDir = Join-Path $Root 'tools'
$distDir  = Join-Path $Root 'dist'
$outXlam  = Join-Path $distDir 'NXTiengViet.xlam'
$dstData  = Join-Path $distDir 'NXTiengViet_data'

function Write-Step($s) { Write-Host ""; Write-Host "== $s" -ForegroundColor Cyan }
function Write-Ok($s)   { Write-Host "   [OK] $s" -ForegroundColor Green }
function Write-Warn2($s){ Write-Host "   [!!] $s" -ForegroundColor Yellow }

function Fail($s) { Write-Host "   [XX] $s" -ForegroundColor Red; throw $s }

function Assert-True($cond, $msg) {
    if (-not $cond) { Fail $msg } else { Write-Ok $msg }
}

function Move-ToRecycle([string[]]$paths) {
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
    foreach ($p in $paths) {
        if (Test-Path -LiteralPath $p) {
            try {
                if ((Get-Item -LiteralPath $p).PSIsContainer) {
                    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory(
                        $p, 'OnlyErrorDialogs', 'SendToRecycleBin')
                } else {
                    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
                        $p, 'OnlyErrorDialogs', 'SendToRecycleBin')
                }
            } catch {
                # thu muc mang / file dang khoa -> dung lai, khong xoa thang
                Write-Warn2 "Khong dua vao Thung rac duoc: $p"
            }
        }
    }
}

function Set-IniKey([string]$path, [string]$key, [string]$val) {
    # dat 1 khoa trong caidat.ini (UTF-8 khong BOM); them moi neu chua co
    $txt = ''
    if (Test-Path $path) { $txt = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) }
    $pat = '(?im)^\s*' + [regex]::Escape($key) + '\s*=.*$'
    if ([regex]::IsMatch($txt, $pat)) {
        $txt = [regex]::Replace($txt, $pat, "$key=$val")
    } else {
        $txt = $txt.TrimEnd() + "`r`n$key=$val`r`n"
    }
    [IO.File]::WriteAllText($path, $txt, (New-Object Text.UTF8Encoding($false)))
}

function Get-Msg([string]$file, [string]$key, [string]$default) {
    # doc 1 khoa tu thongdiep.txt (UTF-8), khong phu thuoc encoding cua .ps1
    $txt = ''
    try { $txt = [IO.File]::ReadAllText($file, [Text.Encoding]::UTF8) } catch { return $default }
    $m = [regex]::Match($txt, '(?im)^\s*' + [regex]::Escape($key) + '\s*=\s*(.*)$')
    if (-not $m.Success) { return $default }
    $v = $m.Groups[1].Value.Trim()
    return $v.Replace('\n', [Environment]::NewLine)
}

function Release-Com($o) {
    if ($null -ne $o) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { }
    }
}

#================================================================== kiem tra dau vao
Write-Step 'Kiem tra dau vao'

foreach ($p in @($srcDir, $dataDir, $toolsDir)) {
    if (-not (Test-Path $p)) { Fail "Thieu thu muc: $p" }
}
if (-not (Test-Path $uiFile)) { Fail "Thieu file: $uiFile" }

$basFiles = Get-ChildItem -Path $srcDir -Filter '*.bas' | Sort-Object Name
if (-not $basFiles) { Fail "Khong co module .bas nao trong $srcDir" }

# customUI phai la XML hop le
try {
    $null = [xml](Get-Content -Raw -Encoding UTF8 $uiFile)
    Write-Ok "customUI.xml hop le ($($basFiles.Count) module .bas)"
} catch {
    Fail "customUI.xml khong hop le: $($_.Exception.Message)"
}

# danh sach onAction <-> ten Sub trong modMain.bas
$uiText   = Get-Content -Raw -Encoding UTF8 $uiFile
$actions  = [regex]::Matches($uiText, 'onAction="([^"]+)"') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
$mainCode = Get-Content -Raw -Encoding UTF8 (Join-Path $srcDir 'modMain.bas')
foreach ($a in $actions) {
    if ($mainCode -notmatch ('(?im)^\s*Public\s+Sub\s+' + [regex]::Escape($a) + '\s*\(')) {
        Fail "onAction='$a' nhung khong co Sub $a trong modMain.bas"
    }
}
Write-Ok ("onAction khop modMain.bas: " + ($actions -join ', '))

#================================================================== tim python
if (-not $Python) {
    $cand = @(@(
        'C:\Users\Administrator\AppData\Local\Programs\Python\Python313\python.exe',
        'C:\Users\Administrator\AppData\Local\Programs\Python\Python312\python.exe'
    ) | Where-Object { Test-Path $_ })
    if ($cand.Count -gt 0) { $Python = $cand[0] }
    else {
        $cmd = Get-Command python -ErrorAction SilentlyContinue
        if ($cmd) { $Python = $cmd.Source }
    }
}
if (-not $Python) { Fail 'Khong tim thay python.exe (truyen -Python <path>)' }
Write-Ok "python: $Python"

#================================================================== Excel / AccessVBOM
Write-Step 'Chuan bi Excel COM'

$regPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$regName = 'AccessVBOM'
$oldVal  = $null
$hadProp = $false

if (Test-Path $regPath) {
    $props = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
    if ($props -and ($props.PSObject.Properties.Name -contains $regName)) {
        $oldVal  = $props.$regName
        $hadProp = $true
    }
}
if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
Set-ItemProperty -Path $regPath -Name $regName -Value 1 -Type DWord
Write-Ok "AccessVBOM=1 (tam thoi, se khoi phuc sau khi build)"

if (-not (Test-Path $distDir)) { New-Item -ItemType Directory -Path $distDir | Out-Null }

$xls = $null
$wb  = $null

try {
    $xls = New-Object -ComObject Excel.Application
    $xls.Visible         = $false
    $xls.DisplayAlerts   = $false
    $xls.EnableEvents    = $false
    $xls.ScreenUpdating  = $false
    $xls.AskToUpdateLinks = $false

    $wb = $xls.Workbooks.Add()
    $wb.VBProject.Name = 'NXTiengViet'

    #------------------------------------------------------------ nap module
    Write-Step "Nap $($basFiles.Count) module VBA"

    foreach ($f in $basFiles) {
        $name = [IO.Path]::GetFileNameWithoutExtension($f.Name)
        $code = [IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8)
        if ([string]::IsNullOrWhiteSpace($code)) { Fail "Module rong: $($f.Name)" }

        $comp = $wb.VBProject.VBComponents.Add(1)      # 1 = vbext_ct_StdModule
        $comp.Name = $name
        $comp.CodeModule.AddFromString($code)
        Write-Ok ("{0,-10} {1,6} dong" -f $name, $comp.CodeModule.CountOfLines)
    }

    $frmFiles = Get-ChildItem -Path $srcDir -Filter '*.frm' -ErrorAction SilentlyContinue | Sort-Object Name
    foreach ($f in $frmFiles) {
        $comp = $wb.VBProject.VBComponents.Import($f.FullName)
        Write-Ok ("{0,-10} {1,6} dong" -f $comp.Name, $comp.CodeModule.CountOfLines)
    }

    $twFile = Join-Path $srcDir 'ThisWorkbook.cls'
    if (Test-Path $twFile) {
        $tw   = $wb.VBProject.VBComponents.Item('ThisWorkbook')
        $code = [IO.File]::ReadAllText($twFile, [Text.Encoding]::UTF8)
        if ($tw.CodeModule.CountOfLines -gt 0) {
            $tw.CodeModule.DeleteLines(1, $tw.CodeModule.CountOfLines)
        }
        $tw.CodeModule.AddFromString($code)
        Write-Ok ("ThisWorkbook {0,4} dong" -f $tw.CodeModule.CountOfLines)
    }

    #------------------------------------------------------------ sheet thong tin
    while ($wb.Worksheets.Count -gt 1) {
        $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
    }
    $info = $wb.Worksheets.Item(1)
    $info.Name = 'NX_INFO'
    $msgFile   = Join-Path $dataDir 'thongdiep.txt'
    $infoLines = @(
        (Get-Msg $msgFile 'info.sheet1' 'NX Tieu hoc - add-in dien nhan xet'),
        (Get-Msg $msgFile 'info.sheet2' 'Phien ban 3.0'),
        (Get-Msg $msgFile 'info.sheet3' 'Dung tab NX Tieu hoc tren Ribbon khi mo file so diem.'),
        (Get-Msg $msgFile 'info.sheet4' 'Du lieu: thu muc NXTiengViet_data nam canh file .xlam.'),
        (Get-Msg $msgFile 'info.sheet5' 'File nay la add-in, khong can mo truc tiep.')
    )
    for ($i = 0; $i -lt $infoLines.Count; $i++) {
        $info.Cells.Item($i + 1, 1).Value2 = $infoLines[$i]
    }
    $info.Cells.Item(1, 1).Font.Bold = $true
    $info.Columns.Item(1).ColumnWidth = 90

    #------------------------------------------------------------ luu .xlam
    Write-Step 'Luu file .xlam'

    if (Test-Path $outXlam) { Move-ToRecycle @($outXlam) }
    $wb.SaveAs($outXlam, 55)          # 55 = xlOpenXMLAddIn
    $wb.Close($false)
    Release-Com $wb
    $wb = $null

    Write-Ok ("{0}  ({1:N0} bytes)" -f $outXlam, (Get-Item $outXlam).Length)
}
finally {
    if ($null -ne $wb)  { try { $wb.Close($false)  } catch { } ; Release-Com $wb }
    if ($null -ne $xls) { try { $xls.Quit() } catch { } ; Release-Com $xls }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($hadProp) {
        Set-ItemProperty -Path $regPath -Name $regName -Value $oldVal -Type DWord
    } else {
        Remove-ItemProperty -Path $regPath -Name $regName -ErrorAction SilentlyContinue
    }
    Write-Ok "Da khoi phuc AccessVBOM ve trang thai ban dau"
}

#================================================================== chen Ribbon
if (-not $SkipInject) {
    Write-Step 'Chen Ribbon (customUI)'
    & $Python (Join-Path $toolsDir 'inject_customui.py') $outXlam $uiFile
    if ($LASTEXITCODE -ne 0) { Fail "inject_customui.py thoat voi ma $LASTEXITCODE" }
}

#================================================================== copy du lieu
Write-Step 'Copy thu muc du lieu'

if ([string]::IsNullOrEmpty($DataSource)) { $DataSource = $dataDir }
if (-not (Test-Path $DataSource)) { Fail "Thieu thu muc du lieu: $DataSource" }

$srcMau = Join-Path $DataSource 'Mau_Nhan_Xet'
if (-not (Test-Path $srcMau)) { Fail "Thieu thu muc mau: $srcMau" }

$msgSrc = Join-Path $DataSource 'thongdiep.txt'
if (-not (Test-Path $msgSrc)) { Fail "Thieu file thongdiep.txt: $msgSrc" }

if (Test-Path $dstData) { Move-ToRecycle @($dstData) }
New-Item -ItemType Directory -Path $dstData -Force | Out-Null

Copy-Item -Path $srcMau -Destination (Join-Path $dstData 'Mau_Nhan_Xet') -Recurse -Force
Copy-Item -Path $msgSrc -Destination (Join-Path $dstData 'thongdiep.txt') -Force

# Khong copy caidat.ini / nhat_ky.txt / backup: de add-in tu tao moi
$nMau = @(Get-ChildItem (Join-Path $dstData 'Mau_Nhan_Xet') -Recurse -File -Filter '*.xlsx').Count
Assert-True ($nMau -ge 100) "Da copy $nMau file mau vao dist\NXTiengViet_data\Mau_Nhan_Xet"
Assert-True (Test-Path (Join-Path $dstData 'thongdiep.txt')) 'Da copy thongdiep.txt (UTF-8) vao dist'

#================================================================== kiem chung goi
Write-Step 'Kiem chung goi OOXML'

Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
$zip = [IO.Compression.ZipFile]::OpenRead($outXlam)
try {
    $names = @($zip.Entries | ForEach-Object { $_.FullName })
    Assert-True ($names -contains 'customUI/customUI.xml') 'Co part customUI/customUI.xml'
    Assert-True ($names -contains 'xl/vbaProject.bin')    'Co xl/vbaProject.bin (VBA)'
    Assert-True ($names -contains '_rels/.rels')          'Co _rels/.rels'

    $relsEntry = $zip.Entries | Where-Object { $_.FullName -eq '_rels/.rels' }
    $sr = New-Object IO.StreamReader($relsEntry.Open(), [Text.Encoding]::UTF8)
    $relsText = $sr.ReadToEnd()
    $sr.Close()
    Assert-True ($relsText -match 'customUI/customUI\.xml') 'Relationship tro toi customUI'
}
finally { $zip.Dispose() }

#================================================================== test chay that
if ($Test) {

    Write-Step 'TEST: chay add-in that'

    $tmpDir  = Join-Path $env:TEMP ('nx_test_' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
    $tmpBook = Join-Path $tmpDir 'SoDiem_Test.xlsx'

    # Tao file test bang Python/openpyxl (tranh loi COM ghi mang cua PS 5.1)
    & $Python (Join-Path $toolsDir 'make_test_book.py') $tmpBook
    if ($LASTEXITCODE -ne 0) { Fail "make_test_book.py thoat voi ma $LASTEXITCODE" }

    # danh sach sheet lay tu file JSON do make_test_book.py sinh ra (giu duoc tieng Viet)
    $metaPath = [IO.Path]::ChangeExtension($tmpBook, '.sheets.json')
    if (-not (Test-Path $metaPath)) { Fail "Thieu file mo ta sheet: $metaPath" }
    $meta   = [IO.File]::ReadAllText($metaPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $sheets = @($meta.sheets)
    $nHs        = [int]$meta.nHs
    $firstRow   = [int]$meta.firstRow
    $lastRow    = [int]$meta.lastRow
    $noteCol    = [int]$meta.noteCol
    $maxCol     = [int]$meta.maxCol
    $markerText = [string]$meta.marker
    Assert-True ($sheets.Count -eq 4) 'Doc duoc 4 sheet test tu JSON'

    $x2 = $null; $addin = $null; $tw2 = $null
    try {
        $x2 = New-Object -ComObject Excel.Application
        $x2.Visible            = $false
        $x2.DisplayAlerts      = $false
        $x2.EnableEvents       = $false
        $x2.ScreenUpdating     = $false
        $x2.AskToUpdateLinks   = $false
        $x2.AutomationSecurity = 1        # msoAutomationSecurityLow

        $addin = $x2.Workbooks.Open($outXlam)
        Write-Ok "Da nap add-in: $($addin.Name)"

        # thongdiep.txt phai duoc doc dung UTF-8 (truoc day bi mojibake)
        $title = [string]$x2.Run("'NXTiengViet.xlam'!NX_T", 'app.title', '')
        Assert-True ($title.Contains([string][char]0x1EC3)) ("Thong diep UTF-8 nap dung ({0} ky tu)" -f $title.Length)

        $tw2 = $x2.Workbooks.Open($tmpBook)
        $tw2.Activate()

        # ---- chup lai toan bo truoc khi ghi (tru cot Nhan xet)
        $before = @{}
        foreach ($s in $sheets) {
            $ws = $tw2.Worksheets.Item($s.Ten)
            $g = $ws.Range('A1').Resize($lastRow, $maxCol).Value2
            for ($r = 1; $r -le $lastRow; $r++) {
                for ($c = 1; $c -le $maxCol; $c++) {
                    if ($c -eq $noteCol) { continue }
                    $before[('{0}!{1},{2}' -f $s.Ten, $r, $c)] = [string]$g[$r, $c]
                }
            }
        }

        # ---- 1. XEM TRUOC tren sheet dau
        $tw2.Worksheets.Item($sheets[0].Ten).Activate()
        $x2.Run("'NXTiengViet.xlam'!NX_Run", $true, $true)
        $rep = [string]$x2.Run("'NXTiengViet.xlam'!NX_Report")
        Write-Host ("   report preview: {0}" -f $rep)
        Assert-True ($rep -like 'PREVIEW Lop_3 GK1 Toan*') "Xem truoc nhan dien dung Lop_3 / GK1 / Toan ($rep)"

        $hasPreview = $false
        foreach ($ws in $tw2.Worksheets) { if ($ws.Name -eq 'NX_XEMTRUOC') { $hasPreview = $true } }
        Assert-True $hasPreview 'Da tao sheet NX_XEMTRUOC'

        $pws   = $tw2.Worksheets.Item('NX_XEMTRUOC')
        $pRows = $pws.Cells.Item($pws.Rows.Count, 3).End(-4162).Row   # xlUp
        Assert-True ($pRows -ge ($nHs - 1 + 1)) ("Sheet xem truoc co $pRows dong (>= $nHs)")

        # ---- 2. GHI THAT tung sheet (moi sheet 1 dot khac nhau)
        foreach ($s in $sheets) {
            $tw2.Worksheets.Item($s.Ten).Activate()
            $x2.Run("'NXTiengViet.xlam'!NX_Run", $false, $true)
            $rep = [string]$x2.Run("'NXTiengViet.xlam'!NX_Report")
            Write-Host ("   {0,-28} -> {1}" -f $s.Ten, $rep)

            $expect = 'OK Lop_3 {0} {1} write={2} skip=1' -f $s.dotCode, $s.monCode, ($nHs - 1)
            Assert-True ($rep -like $expect) ("{0}: bao cao dung ({1})" -f $s.Ten, $expect)

            $ws = $tw2.Worksheets.Item($s.Ten)
            $g  = $ws.Range('A5').Resize($nHs, $maxCol).Value2
            $filled = 0
            for ($r = 1; $r -le $nHs - 1; $r++) {
                if ([string]$g[$r, $noteCol] -ne '') { $filled++ }
            }
            Assert-True ($filled -eq ($nHs - 1)) ("{0}: dien {1}/{2} o nhan xet" -f $s.Ten, $filled, ($nHs - 1))
            Assert-True ([string]$g[$nHs, $noteCol] -eq $markerText) ("{0}: GhiDe=0 giu nguyen o da co chu" -f $s.Ten)
        }

        # ---- 3. GhiDe=1: ghi de ca o da co chu
        $iniPath = Join-Path $dstData 'caidat.ini'
        Set-IniKey $iniPath 'GhiDe' '1'
        # nap lai cai dat tu dia (add-in cache trong bo nho)
        $null = $x2.Run("'NXTiengViet.xlam'!NX_IniReset")

        $tw2.Worksheets.Item($sheets[0].Ten).Activate()
        $x2.Run("'NXTiengViet.xlam'!NX_Run", $false, $true)
        $rep = [string]$x2.Run("'NXTiengViet.xlam'!NX_Report")
        Write-Host ("   GhiDe=1 -> {0}" -f $rep)
        Assert-True ($rep -like ('OK Lop_3 GK1 Toan write={0} skip=0' -f $nHs)) 'GhiDe=1 ghi de toan bo'

        $ws = $tw2.Worksheets.Item($sheets[0].Ten)
        $g  = $ws.Range('A5').Resize($nHs, $maxCol).Value2
        Assert-True ([string]$g[$nHs, $noteCol] -ne $markerText) 'GhiDe=1 da thay o da co chu'

        # tra lai GhiDe=0
        Set-IniKey $iniPath 'GhiDe' '0'
        $null = $x2.Run("'NXTiengViet.xlam'!NX_IniReset")

        # ---- 4. cac cot khac khong doi
        $changed = 0
        foreach ($s in $sheets) {
            $ws = $tw2.Worksheets.Item($s.Ten)
            $g  = $ws.Range('A1').Resize($lastRow, $maxCol).Value2
            for ($r = 1; $r -le $lastRow; $r++) {
                for ($c = 1; $c -le $maxCol; $c++) {
                    if ($c -eq $noteCol) { continue }
                    $key = ('{0}!{1},{2}' -f $s.Ten, $r, $c)
                    if ($before.ContainsKey($key)) {
                        if ($before[$key] -ne [string]$g[$r, $c]) {
                            $changed++
                            Write-Warn2 ("Doi o {0}: '{1}' -> '{2}'" -f $key, $before[$key], [string]$g[$r, $c])
                        }
                    }
                }
            }
        }
        Assert-True ($changed -eq 0) 'Cac cot khac (STT/Ten/Ngay sinh/KT/Xep loai) khong bi sua'

        # ---- 5. cau khong lap trong cung 1 sheet
        $ws0 = $tw2.Worksheets.Item($sheets[0].Ten)
        $g0  = $ws0.Range('A5').Resize($nHs, $maxCol).Value2
        $set = @{}
        $dup = 0
        for ($r = 1; $r -le $nHs; $r++) {
            $t = [string]$g0[$r, $noteCol]
            if ($set.ContainsKey($t)) { $dup++ } else { $set[$t] = 1 }
        }
        Write-Ok ("{0}: {1} cau khac nhau / {2} hoc sinh (trung {3})" -f $sheets[0].Ten, $set.Count, $nHs, $dup)

        # ---- 6. nhat ky
        $logPath = Join-Path $dstData 'nhat_ky.txt'
        Assert-True (Test-Path $logPath) 'Co file nhat_ky.txt'
        $log = [IO.File]::ReadAllText($logPath, [Text.Encoding]::UTF8)
        Assert-True ($log -match 'DIEN ') 'Nhat ky co ghi lan dien'

        Write-Host ""
        Write-Host '  --- 3 o dau tien (sheet 1) ---' -ForegroundColor DarkGray
        for ($r = 1; $r -le 3; $r++) {
            Write-Host ("  {0,-4} {1}" -f [string]$g0[$r, 2], [string]$g0[$r, $noteCol])
        }
    }
    catch {
        Write-Host ""
        Write-Host "   [XX] $($_.Exception.GetType().FullName)" -ForegroundColor Red
        Write-Host "   [XX] $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "   --- stack ---" -ForegroundColor DarkGray
        Write-Host ($_.ScriptStackTrace) -ForegroundColor DarkGray
        throw
    }
    finally {
        if ($null -ne $tw2)   { try { $tw2.Close($false) }   catch { } ; Release-Com $tw2 }
        if ($null -ne $addin) { try { $addin.Close($false) } catch { } ; Release-Com $addin }
        if ($null -ne $x2)    { try { $x2.Quit() }           catch { } ; Release-Com $x2 }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()

        if (-not $KeepTemp) { Move-ToRecycle @($tmpDir) }
        else { Write-Warn2 "Giu lai thu muc test: $tmpDir" }
    }
}

#================================================================== deploy (tuy chon)
if ($Deploy) {
    Write-Step 'Deploy sang thu muc su dung'

    if ([string]::IsNullOrEmpty($DeployTo)) {
        $DeployTo = Join-Path (Split-Path $Root -Parent) 'NX_MON'
    }
    if (-not (Test-Path $DeployTo)) { Fail "Khong tim thay thu muc deploy: $DeployTo" }

    $deployData = Join-Path $DeployTo 'NXTiengViet_data'
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $bkDir = Join-Path $DeployTo ('_backup_' + $stamp)
    New-Item -ItemType Directory -Path $bkDir -Force | Out-Null

    foreach ($n in @('NXTiengViet.xlam', 'NXTiengViet_data\Mau_Nhan_Xet', 'NXTiengViet_data\thongdiep.txt', 'NXTiengViet_data\caidat.ini')) {
        $srcItem = Join-Path $DeployTo $n
        if (Test-Path $srcItem) {
            $dstItem = Join-Path $bkDir $n
            $parent = Split-Path $dstItem -Parent
            if (-not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
            Copy-Item -Path $srcItem -Destination $dstItem -Recurse -Force
        }
    }
    Write-Ok "Da sao luu ban cu: $bkDir"

    if (-not (Test-Path $deployData)) { New-Item -ItemType Directory -Path $deployData -Force | Out-Null }

    Copy-Item -Path $outXlam -Destination (Join-Path $DeployTo 'NXTiengViet.xlam') -Force
    Copy-Item -Path (Join-Path $dstData 'thongdiep.txt') -Destination (Join-Path $deployData 'thongdiep.txt') -Force
    if (Test-Path (Join-Path $deployData 'Mau_Nhan_Xet')) { Move-ToRecycle @((Join-Path $deployData 'Mau_Nhan_Xet')) }
    Copy-Item -Path (Join-Path $dstData 'Mau_Nhan_Xet') -Destination (Join-Path $deployData 'Mau_Nhan_Xet') -Recurse -Force

    Write-Ok "Da deploy add-in + du lieu sang: $DeployTo"
    Write-Warn2 'Giu nguyen caidat.ini cua nguoi dung (chi bo sung thongdiep.txt + Mau_Nhan_Xet)'
}

#================================================================== ket thuc
Write-Step 'HOAN TAT'
Write-Host ("   add-in : {0}" -f $outXlam) -ForegroundColor Green
Write-Host ("   du lieu: {0}" -f $dstData) -ForegroundColor Green
Write-Host "   Nap add-in: Excel > File > Options > Add-ins > Browse > chon file .xlam"
