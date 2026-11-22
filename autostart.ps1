# ============================================================
#  خادم المكتبة — التشغيل التلقائي عند دخول Windows
#  ينشئ (أو يحذف) اختصاراً في مجلد بدء التشغيل الخاص بالمستخدم
#  المتغيّر المطلوب: APPDIR = مجلد المكتبة (ينتهي بشرطة مائلة)
#
#  ملاحظة تقنية: محرّك الاختصارات WSH على هذا النظام يحوّل
#  الحروف العربية إلى علامات استفهام عند الحفظ، لذلك:
#    - محتوى الاختصار كله إنجليزي (ASCII) ويستدعي ملف مساعد
#    - ملف المساعد يُكتب بترميز Unicode ويحمل المسار العربي
#    - يبدأ الخادم مصغّراً وبلا فتح المتصفح
# ============================================================
param([switch]$Remove)

$appDir = $env:APPDIR
if (-not $appDir -or -not (Test-Path $appDir)) {
    Write-Host "  X  لم يُحدَّد مجلد المكتبة (APPDIR)." -ForegroundColor Red
    exit 1
}
if (-not $appDir.EndsWith('\')) { $appDir = $appDir + '\' }

# ملف التشغيل الرئيسي هو ملف الـ bat الذي يستدعي server.py
$target = Get-ChildItem -Path $appDir -Filter *.bat -File -ErrorAction SilentlyContinue |
    Where-Object {
        (Get-Content -Path $_.FullName -Raw -ErrorAction SilentlyContinue) -match 'server\.py'
    } |
    Select-Object -First 1

if (-not $target) {
    Write-Host "  X  لم يُعثر على ملف تشغيل المكتبة (الذي يستدعي server.py)." -ForegroundColor Red
    exit 1
}

$startupDir = [Environment]::GetFolderPath('Startup')
$lnkName = 'خادم المكتبة - تشغيل تلقائي.lnk'
$lnkPath = Join-Path $startupDir $lnkName
$tmpLnk = Join-Path $startupDir 'library-server-start.tmp.lnk'

$helperDir = Join-Path $env:APPDATA 'LibraryServer'
$helperPs1 = Join-Path $helperDir 'start-library.ps1'

if ($Remove) {
    $removed = $false
    foreach ($p in @($lnkPath, $tmpLnk)) {
        if (Test-Path -LiteralPath $p) {
            Remove-Item -LiteralPath $p -Force
            $removed = $true
        }
    }
    if (Test-Path -LiteralPath $helperDir) {
        Remove-Item -LiteralPath $helperDir -Recurse -Force
        $removed = $true
    }
    if ($removed) {
        Write-Host "  OK  تم إلغاء التشغيل التلقائي — لن يبدأ الخادم عند تشغيل الحاسبة." -ForegroundColor Green
    } else {
        Write-Host "  --  التشغيل التلقائي غير مفعَّل أصلاً." -ForegroundColor Yellow
    }
    exit 0
}

# ---------- 1) ملف المساعد (يحمل المسار العربي — ترميز Unicode) ----------
New-Item -ItemType Directory -Path $helperDir -Force | Out-Null
$batPath = $target.FullName.Replace("'", "''")
$helper = @(
    '# يبدأ خادم المكتبة مصغّراً — يُستدعى من اختصار مجلد بدء التشغيل'
    "`$bat = '" + $batPath + "'"
    'if (Test-Path -LiteralPath $bat) {'
    "    Start-Process -FilePath `$bat -ArgumentList '--no-browser' -WindowStyle Minimized"
    '}'
    ''
) -join "`r`n"
[System.IO.File]::WriteAllText($helperPs1, $helper, (New-Object System.Text.UTF8Encoding $true))

if (-not (Test-Path -LiteralPath $helperPs1)) {
    Write-Host "  X  تعذّر إنشاء ملف المساعد: $helperPs1" -ForegroundColor Red
    exit 1
}

# ---------- 2) الاختصار في مجلد بدء التشغيل (محتوى إنجليزي فقط) ----------
$psExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }

$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut($tmpLnk)
$sc.TargetPath = $psExe
$sc.Arguments = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $helperPs1 + '"'
$sc.WorkingDirectory = $helperDir
$sc.WindowStyle = 7
$sc.Save()

if (-not (Test-Path -LiteralPath $tmpLnk)) {
    Write-Host "  X  تعذّر إنشاء الاختصار في مجلد بدء التشغيل." -ForegroundColor Red
    exit 1
}

# إزالة الاختصار القديم ثم تسمية الاختصار بالعربية (التسمية عبر PowerShell)
if (Test-Path -LiteralPath $lnkPath) { Remove-Item -LiteralPath $lnkPath -Force }
Rename-Item -LiteralPath $tmpLnk -NewName $lnkName

if (-not (Test-Path -LiteralPath $lnkPath)) {
    Write-Host "  X  تعذّر تسمية الاختصار." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "  OK  تم تفعيل التشغيل التلقائي للخادم." -ForegroundColor Green
Write-Host "    - يبدأ تلقائياً فور دخول المستخدم إلى هذه الحاسبة."
Write-Host "    - يعمل في الخلفية (نافذة مصغّرة في شريط المهام) بلا متصفح."
Write-Host "    - الرابط يبقى كما هو ويظهر للطلاب كالمعتاد."
Write-Host "    - للإيقاف: أغلق نافذة الخادم من شريط المهام،"
Write-Host "      أو شغّل ملف:  إلغاء-التشغيل-التلقائي.bat"
Write-Host ""
Write-Host "    الاختصار: $lnkPath"
Write-Host ""
exit 0
