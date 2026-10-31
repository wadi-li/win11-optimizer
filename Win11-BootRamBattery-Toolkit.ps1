#Requires -Version 5.1
<#
.SYNOPSIS
    Win11 Boot / RAM / Battery Toolkit v2 — замер влияния программ на загрузку, ОЗУ и батарею
    и обратимая оптимизация с журналом отката. Windows 10/11 (PowerShell 5.1 и 7.x).

.DESCRIPTION
    Режимы (-Mode):
      Menu              интерактивное меню (по умолчанию)
      Collect           только чтение: сбор телеметрии, отчёты JSON/CSV/MD (права администратора желательны)
      Plan              только чтение: какие пункты каталога применимы на этом ПК и что изменится
      Apply             применить выбранные пункты (точка восстановления + журнал отката); поддерживает -WhatIf
      Rollback          откат по журналу (по умолчанию — последний Journal_*.json в папке вывода)
      ScheduleMeasure   однократный замер через -MeasureDelaySec секунд после следующего входа в систему
      UnscheduleMeasure удалить запланированный замер
      BootTraceStart    включить трассировку следующей загрузки (wpr -boottrace -addboot GeneralProfile)
      BootTraceStop     сохранить трассировку загрузки в .etl (открывать в Windows Performance Analyzer)
      BootTraceCancel   отменить трассировку загрузки
      ListCatalog       показать каталог пунктов оптимизации

    Выбор пунктов: -Level Low (по умолчанию: только низкий риск) | Medium (низкий + средний) | None (только -Include).
    -Include / -Exclude принимают Id пунктов (через запятую). Пункты с риском «Высокий» не выбираются автоматически.

    Метрика влияния на запуск — классификация Microsoft (Task Manager / Compatibility Cookbook):
      Высокое: > 1000 мс CPU или > 3 МБ дискового ввода-вывода; Среднее: 300–1000 мс или 300 КБ–3 МБ; Низкое: меньше.
    Счётчики CPU/IO у процессов накопительные, поэтому точная оценка — сразу после входа (-Mode ScheduleMeasure).

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\Win11-BootRamBattery-Toolkit.ps1 -Mode Collect
.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\Win11-BootRamBattery-Toolkit.ps1 -Mode Apply -WhatIf
.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\Win11-BootRamBattery-Toolkit.ps1 -Mode Apply -Level Medium -Exclude SVC_NLSVC,STARTUP_GOOGLEDRIVE
.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\Win11-BootRamBattery-Toolkit.ps1 -Mode Rollback
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet('Menu', 'Collect', 'Plan', 'Apply', 'Rollback', 'ScheduleMeasure', 'UnscheduleMeasure',
                 'BootTraceStart', 'BootTraceStop', 'BootTraceCancel', 'ListCatalog')]
    [string]$Mode = 'Menu',
    [ValidateSet('None', 'Low', 'Medium')]
    [string]$Level = 'Low',
    [string[]]$Include = @(),
    [string[]]$Exclude = @(),
    [string]$OutputDir,
    [switch]$StopNow,
    [switch]$NoRestorePoint,
    [ValidateRange(0, 262144)][int]$PagefileInitialMB = 0,
    [ValidateRange(0, 262144)][int]$PagefileMaxMB = 0,
    [string]$RollbackFile,
    [ValidateRange(30, 1800)][int]$MeasureDelaySec = 120,
    [switch]$FromScheduledTask,
    [switch]$Yes,
    [Parameter(DontShow = $true)][switch]$Elevated
)

# ----------------------------------------------------------------------------------------------
# 0. Инициализация
# ----------------------------------------------------------------------------------------------
$ErrorActionPreference = 'Continue'   # Stop ломает перехват stderr нативных утилит в PS 5.1
$script:Version   = '2.0.0'
$script:Cmdlet    = $PSCmdlet
$script:CimProblems = @{}
$script:ProcFallback = $false
$script:ImpactWindowMin = 10   # окно, в котором накопленные CPU/диск ещё отражают запуск
$script:DryRun    = [bool]$WhatIfPreference   # собственный флаг: работает и из меню
$WhatIfPreference = $false   # иначе -WhatIf наследуют импорт модулей, запись отчётов и журнала; изменения отсекает $script:DryRun
$script:Bound     = @{} + $PSBoundParameters
$script:StartedAt = Get-Date
$script:Stamp     = $script:StartedAt.ToString('yyyyMMdd-HHmmss')
$script:PC        = $(if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName })
$script:MeasureTaskName = 'PostLogonMeasure'
$script:MeasureTaskPath = '\Win11Toolkit\'
$script:Utf8Bom   = New-Object System.Text.UTF8Encoding($true)
$script:Ru        = [Globalization.CultureInfo]::GetCultureInfo('ru-RU')

# -Include A,B через -File приходит одной строкой — нормализуем
$Include = @($Include | ForEach-Object { "$_" -split '[,; ]' } | Where-Object { $_ } | ForEach-Object { $_.Trim().ToUpperInvariant() })
$Exclude = @($Exclude | ForEach-Object { "$_" -split '[,; ]' } | Where-Object { $_ } | ForEach-Object { $_.Trim().ToUpperInvariant() })

if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    Write-Warning 'Запущен 32-битный PowerShell: реестр будет перенаправлен (WOW6432Node). Запустите 64-битный powershell.exe.'
    exit 2
}

if (-not ('W11Tk2.Native' -as [type])) {
    Add-Type -Namespace W11Tk2 -Name Native -MemberDefinition @'
[StructLayout(LayoutKind.Sequential)]
public struct PERFORMANCE_INFORMATION {
    public uint cb;
    public UIntPtr CommitTotal; public UIntPtr CommitLimit; public UIntPtr CommitPeak;
    public UIntPtr PhysicalTotal; public UIntPtr PhysicalAvailable; public UIntPtr SystemCache;
    public UIntPtr KernelTotal; public UIntPtr KernelPaged; public UIntPtr KernelNonpaged;
    public UIntPtr PageSize; public uint HandleCount; public uint ProcessCount; public uint ThreadCount;
}
[DllImport("psapi.dll", SetLastError = true)]
public static extern bool GetPerformanceInfo(out PERFORMANCE_INFORMATION pi, uint cb);
[DllImport("powrprof.dll", SetLastError = true)]
[return: MarshalAs(UnmanagedType.U1)]
public static extern bool GetPwrCapabilities([Out] byte[] lpspc);
[DllImport("kernel32.dll")]
public static extern ulong GetTickCount64();
'@
}

# ----------------------------------------------------------------------------------------------
# 1. Общие функции
# ----------------------------------------------------------------------------------------------
function Write-Info { param([string]$Text, [ConsoleColor]$Color = 'Gray') Write-Host $Text -ForegroundColor $Color }
function Write-Step { param([string]$Text) Write-Host ''; Write-Host "== $Text" -ForegroundColor Cyan }

# WMI/CIM может быть запрещён (политики, песочницы, сломанный репозиторий WMI) — не падаем, а помечаем
function Get-Cim {
    param([string]$Class, [string]$Filter, [string[]]$Property, [string]$Namespace = 'root\cimv2', [switch]$Quiet)
    $p = @{ ClassName = $Class; Namespace = $Namespace; ErrorAction = 'Stop' }
    if ($Filter) { $p.Filter = $Filter }
    if ($Property) { $p.Property = $Property }
    try { return @(Get-CimInstance @p) }
    catch {
        if (-not $Quiet) { $script:CimProblems["$Namespace\$Class"] = $_.Exception.Message }
        return @()
    }
}

function Test-IsAdmin {
    $p = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Resolve-OutDir {
    if ($OutputDir) { $d = $OutputDir }
    elseif ($PSScriptRoot) { $d = Join-Path $PSScriptRoot 'Win11Toolkit-Output' }
    else { $d = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Win11Toolkit-Output' }
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    return (Resolve-Path -LiteralPath $d).ProviderPath
}

function Invoke-SelfElevate {
    param([hashtable]$Override = @{})
    if (-not $PSCommandPath) { throw 'Скрипт вставлен в консоль, а не запущен из файла: сохраните его как .ps1 и запустите снова.' }
    $p = @{} + $script:Bound
    foreach ($k in $Override.Keys) { $p[$k] = $Override[$k] }
    if (-not $p.ContainsKey('OutputDir')) { $p['OutputDir'] = $script:OutDir }
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($s in '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath)) { $list.Add($s) }
    foreach ($k in $p.Keys) {
        $v = $p[$k]
        if ($v -is [System.Management.Automation.SwitchParameter] -or $v -is [bool]) { if ([bool]$v) { $list.Add("-$k") } }
        elseif ($v -is [array]) { if ($v.Count) { $list.Add("-$k"); $list.Add(('"{0}"' -f ($v -join ','))) } }
        elseif ($null -ne $v -and "$v" -ne '') { $list.Add("-$k"); $list.Add(('"{0}"' -f ("$v" -replace '\\$', '\\'))) }   # «C:\» → «"C:\\"», иначе \" съест кавычку
    }
    if ($script:DryRun -and -not $p.ContainsKey('WhatIf')) { $list.Add('-WhatIf') }
    $list.Add('-Elevated')
    $exe = Join-Path $PSHOME 'powershell.exe'
    if (-not (Test-Path -LiteralPath $exe)) { $exe = (Get-Process -Id $PID).Path }
    Write-Info 'Нужны права администратора — открываю окно с повышением (UAC)…' Yellow
    try { Start-Process -FilePath $exe -ArgumentList ($list -join ' ') -Verb RunAs -Wait -ErrorAction Stop }
    catch { Write-Warning "Повышение прав отменено или не удалось: $($_.Exception.Message)" }
}

function Format-Num { param($Value, [string]$Fmt = '0.#') if ($null -eq $Value) { return '' } ([double]$Value).ToString($Fmt, $script:Ru) }
function Format-Time { param($Dt) if ($null -eq $Dt) { return '' } ([datetime]$Dt).ToString('yyyy-MM-dd HH:mm:ss.fff') }

function Write-TextFile { param([string]$Path, [string]$Text) [IO.File]::WriteAllText($Path, $Text, $script:Utf8Bom) }

function Export-RuCsv {
    param($Rows, [string]$Path)
    $rowsArr = @($Rows)
    if ($rowsArr.Count -eq 0) { Write-TextFile $Path ''; return }
    $conv = foreach ($r in $rowsArr) {
        $o = [ordered]@{}
        foreach ($pr in $r.PSObject.Properties) {
            $v = $pr.Value
            if ($v -is [double] -or $v -is [single] -or $v -is [decimal]) { $v = $v.ToString('0.###', $script:Ru) }
            elseif ($v -is [datetime]) { $v = $v.ToString('yyyy-MM-dd HH:mm:ss') }
            elseif ($v -is [array]) { $v = ($v -join ', ') }
            $o[$pr.Name] = $v
        }
        [pscustomobject]$o
    }
    $lines = @($conv | ConvertTo-Csv -Delimiter ';' -NoTypeInformation)
    [IO.File]::WriteAllLines($Path, [string[]]$lines, $script:Utf8Bom)
}

function Invoke-Native {
    param([string]$File, [string[]]$Arguments)
    $out = & $File @Arguments 2>&1 | ForEach-Object { "$_" }
    return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = (@($out) -join "`n") }
}

function Get-EventsSafe {
    param([hashtable]$Filter, [int]$Max = 0)
    try {
        if ($Max -gt 0) { return @(Get-WinEvent -FilterHashtable $Filter -MaxEvents $Max -ErrorAction Stop) }
        return @(Get-WinEvent -FilterHashtable $Filter -ErrorAction Stop)
    } catch { return @() }
}

function Get-EventDataMap {
    param($Ev)
    $h = @{}
    try {
        $x = [xml]$Ev.ToXml()
        foreach ($d in @($x.Event.EventData.Data)) { if ($d -and $d.Name) { $h[$d.Name] = $d.'#text' } }
    } catch { }
    return $h
}

# --- реестр через .NET (литеральные имена, без подстановочных символов PowerShell) ---
function Open-RegKey {
    param([string]$Hive, [string]$SubKey, [switch]$Write, [switch]$Create)
    $base = if ($Hive -eq 'HKCU') { [Microsoft.Win32.Registry]::CurrentUser } else { [Microsoft.Win32.Registry]::LocalMachine }
    if ($Create) { return $base.CreateSubKey($SubKey) }
    return $base.OpenSubKey($SubKey, [bool]$Write)
}

function Get-RegValueInfo {
    param([string]$Hive, [string]$SubKey, [string]$Name)
    $k = Open-RegKey $Hive $SubKey
    if (-not $k) { return [pscustomobject]@{ Exists = $false; Kind = $null; Value = $null } }
    try {
        if (@($k.GetValueNames()) -notcontains $Name) { return [pscustomobject]@{ Exists = $false; Kind = $null; Value = $null } }
        $kind = $k.GetValueKind($Name).ToString()
        $v = $k.GetValue($Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        if ($kind -eq 'Binary') { $v = (@($v) | ForEach-Object { '{0:X2}' -f $_ }) -join '' }
        elseif ($kind -eq 'MultiString') { $v = @($v) }
        return [pscustomobject]@{ Exists = $true; Kind = $kind; Value = $v }
    } finally { $k.Close() }
}

function Set-RegValue {
    param([string]$Hive, [string]$SubKey, [string]$Name, [string]$Kind, $Value)
    $k = Open-RegKey $Hive $SubKey -Create
    if (-not $k) { throw "Не удалось открыть/создать $Hive\$SubKey" }
    try {
        switch ($Kind) {
            'Binary'      { $hex = "$Value"; $bytes = New-Object byte[] ($hex.Length / 2)
                            for ($i = 0; $i -lt $bytes.Length; $i++) { $bytes[$i] = [Convert]::ToByte($hex.Substring($i * 2, 2), 16) }
                            $k.SetValue($Name, $bytes, [Microsoft.Win32.RegistryValueKind]::Binary) }
            'DWord'       { $k.SetValue($Name, [int]$Value, [Microsoft.Win32.RegistryValueKind]::DWord) }
            'QWord'       { $k.SetValue($Name, [long]$Value, [Microsoft.Win32.RegistryValueKind]::QWord) }
            'MultiString' { $k.SetValue($Name, [string[]]@($Value), [Microsoft.Win32.RegistryValueKind]::MultiString) }
            'ExpandString'{ $k.SetValue($Name, [string]$Value, [Microsoft.Win32.RegistryValueKind]::ExpandString) }
            default       { $k.SetValue($Name, [string]$Value, [Microsoft.Win32.RegistryValueKind]::String) }
        }
    } finally { $k.Close() }
}

function Remove-RegValue {
    param([string]$Hive, [string]$SubKey, [string]$Name)
    $k = Open-RegKey $Hive $SubKey -Write
    if ($k) { try { $k.DeleteValue($Name, $false) } finally { $k.Close() } }
}

function Restore-RegValue {
    param([string]$Hive, [string]$SubKey, [string]$Name, $Before)
    if ($Before.Exists) { Set-RegValue $Hive $SubKey $Name $Before.Kind $Before.Value }
    else { Remove-RegValue $Hive $SubKey $Name }
}

# --- пути и издатели файлов ---
function Resolve-ImagePath {
    param([string]$Raw)
    if (-not $Raw) { return '' }
    $s = [Environment]::ExpandEnvironmentVariables($Raw.Trim())
    if ($s.StartsWith('"')) { $e = $s.IndexOf('"', 1); if ($e -gt 1) { $s = $s.Substring(1, $e - 1) } }
    else {
        $m = [regex]::Match($s, '^(.+?\.(exe|sys|dll|com|bat|cmd|scr))(\s|,|$)', 'IgnoreCase')
        if ($m.Success) { $s = $m.Groups[1].Value }
    }
    $s = $s -replace '^\\\?\?\\', ''
    if ($s -match '^(?i)\\SystemRoot\\') { $s = $env:windir + $s.Substring(11) }
    elseif ($s -match '^(?i)system32\\') { $s = Join-Path $env:windir $s }
    return $s
}

$script:CompanyCache = @{}
function Get-FileCompany {
    param([string]$Path)
    if (-not $Path) { return '' }
    $key = $Path.ToLowerInvariant()
    if ($script:CompanyCache.ContainsKey($key)) { return $script:CompanyCache[$key] }
    $c = ''
    try { $c = "$((Get-Item -LiteralPath $Path -ErrorAction Stop).VersionInfo.CompanyName)".Trim() } catch { }
    $script:CompanyCache[$key] = $c
    return $c
}
function Test-MicrosoftCompany { param([string]$Company) return ($Company -match 'Microsoft') }

# ----------------------------------------------------------------------------------------------
# 2. Каталог: пункты оптимизации (данные, а не код). Риск: Низкий / Средний / Высокий
# ----------------------------------------------------------------------------------------------
# Никогда не изменяются, даже по -Include (защита, фильтры томов, питание/охлаждение ThinkPad):
$script:Protected = @(
    'wesp', 'wespelam', 'WinDefend', 'WdFilter', 'WdBoot', 'mpssvc', 'BFE', 'SysMain', 'EventLog', 'RpcSs',
    'fltsrv', 'snapman', 'ekrn', 'efwd', 'ekrnEpfw', 'eamonm', 'edevmon', 'eelam', 'ehdrv', 'epfw', 'epfwwfp',
    'LITSSvc', 'PowerMgr', 'IBMPMDRV', 'TPPWRIF', 'Tbt*', 'WireGuardManager', 'AdskLicensingService'
)
$script:KnownNotes = [ordered]@{
    'wesp'  = 'wesp.sys / wesp_elam.sys — встроенный компонент Microsoft (Windows Endpoint Security Platform) из обновления KB5127936. Ошибки 7000/7026 — известная проблема этого обновления. НЕ отключать: ждать исправления от Microsoft.'
    'aehd'  = 'Android Emulator Hypervisor Driver: не загружается при работающем Hyper-V/VBS. Безопасно перевести в ручной запуск (пункт DRV_AEHD).'
    'CMigrationService' = 'Служба миграции Samsung Magician: нужна только для клонирования диска (пункт SVC_CMIGRATION).'
    'Adguard' = 'AdGuard: при сбое службы фильтрация трафика прерывается; проверьте версию и списки фильтров.'
}

function New-Cat {
    param([string]$Id, [string]$Group, [string]$Kind, [string]$Risk, [string]$Title, [string]$Reason, [string]$Effect, [hashtable]$P = @{})
    $o = [ordered]@{ Id = $Id; Group = $Group; Kind = $Kind; Risk = $Risk; Title = $Title; Reason = $Reason; Effect = $Effect }
    foreach ($k in $P.Keys) { $o[$k] = $P[$k] }
    if (-not $o.Contains('LaptopOnly')) { $o['LaptopOnly'] = $false }
    if (-not $o.Contains('Precheck'))   { $o['Precheck'] = $null }
    return [pscustomobject]$o
}

function Get-Catalog {
    $c = New-Object System.Collections.Generic.List[object]
    # --- Загрузка: драйверы и службы ---
    $c.Add((New-Cat 'DRV_AEHD' 'Загрузка' 'Service' 'Низкий' 'aehd → ручной запуск' `
        'Драйвер гипервизора Android Emulator стартует как System (Start=1) и не загружается при работающем Hyper-V/VBS (событие 7026).' `
        'Эмулятор Android через AEHD (без WHPX) не запустится, пока драйвер не вернут в System.' `
        @{ Targets = @('aehd'); Desired = 'Manual' }))
    $c.Add((New-Cat 'SVC_CMIGRATION' 'Загрузка' 'Service' 'Низкий' 'CMigrationService (Samsung) → вручную' `
        'Служба миграции Samsung Magician; на T14S аварийно завершалась при загрузке (7034).' `
        'Функция переноса ОС в Magician запустит службу при необходимости или потребует ручного старта.' `
        @{ Targets = @('CMigrationService'); Desired = 'Manual' }))
    $c.Add((New-Cat 'SVC_SAMSUNG_MAGICIAN' 'Загрузка' 'Service' 'Средний' 'SamsungMagicianSVC → вручную' `
        'Служба Samsung Magician при каждом старте ставит драйвер MagicianSataModeReader (7045). SSD работает на штатном драйвере Windows.' `
        'Samsung Magician может не показывать данные, пока служба не запущена.' `
        @{ Targets = @('SamsungMagicianSVC'); Desired = 'Manual' }))
    $c.Add((New-Cat 'SVC_OODEFRAG' 'Загрузка' 'Service' 'Средний' 'OODefragAgent → вручную' `
        'Фоновый агент O&O Defrag. Для SSD Windows сама выполняет TRIM по расписанию («Оптимизация дисков»).' `
        'Задания по расписанию O&O Defrag не будут выполняться.' `
        @{ Targets = @('OODefragAgent'); Desired = 'Manual' }))
    $c.Add((New-Cat 'SVC_CCLEANER' 'Загрузка' 'Service' 'Низкий' 'CCleaner7 → вручную' `
        'Фоновая служба CCleaner (мониторинг/обновления).' 'Фоновая «умная очистка» и уведомления CCleaner.' `
        @{ Targets = @('CCleaner*'); Desired = 'Manual' }))
    $c.Add((New-Cat 'SVC_ADSK_ACCESS' 'Загрузка' 'Service' 'Средний' 'Autodesk Access Service Host → вручную' `
        'Служба центра обновлений Autodesk Access; работает без запущенных продуктов Autodesk.' `
        'Обновления Autodesk не будут ставиться автоматически; лицензирование (AdskLicensingService) не затрагивается.' `
        @{ Targets = @('Autodesk Access Service Host'); Desired = 'Manual' }))
    $c.Add((New-Cat 'SVC_ADSK_CER' 'Загрузка' 'Service' 'Низкий' 'Autodesk CER Service → вручную' `
        'Служба отчётов о сбоях Autodesk.' 'Отчёты о сбоях Autodesk могут не отправляться.' `
        @{ Targets = @('Autodesk CER Service'); Desired = 'Manual' }))
    $c.Add((New-Cat 'SVC_FLEXNET' 'Загрузка' 'Service' 'Низкий' 'FlexNet Licensing Service 64 → вручную' `
        'Служба лицензирования FlexNet обычно запускается приложениями по требованию.' `
        'Если программа сообщит об ошибке лицензии — откатите пункт.' `
        @{ Targets = @('FlexNet Licensing Service*'); Desired = 'Manual' }))
    $c.Add((New-Cat 'SVC_INTEL_DSA' 'Загрузка' 'Service' 'Низкий' 'Intel Driver & Support Assistant → вручную' `
        'Служба проверки обновлений драйверов Intel.' 'DSA не проверяет обновления в фоне; при открытии сайта DSA службу можно запустить.' `
        @{ Targets = @('DSAService', 'DSAUpdateService'); Desired = 'Manual' }))
    $c.Add((New-Cat 'SVC_HP_DOCTOR' 'Загрузка' 'Service' 'Низкий' 'HP Print and Scan Doctor → вручную' `
        'Служба диагностики HP; печать через стандартную очередь не зависит от неё.' 'Нет фоновой диагностики HP.' `
        @{ Targets = @('HPPrintScanDoctorService'); Desired = 'Manual' }))
    $c.Add((New-Cat 'SVC_UPDATERS' 'Загрузка' 'Service' 'Низкий' 'Службы автообновления (Google/Foxit/Yandex/Logi) → вручную' `
        'Службы обновления не нужны с момента загрузки; программы проверяют обновления при запуске.' `
        'Обновления ставятся реже; некоторые установщики при обновлении вернут Auto.' `
        @{ Targets = @('GoogleUpdaterService*', 'GoogleUpdaterInternalService*', 'FoxitReaderUpdateService', 'YandexBrowserService', 'OptionsPlusUpdaterService'); Desired = 'Manual' }))
    $c.Add((New-Cat 'SVC_ANYDESK' 'Загрузка' 'Service' 'Средний' 'AnyDesk (служба) → вручную' `
        'Служба AnyDesk стартует с системой ради неконтролируемого доступа.' `
        'Удалённое подключение к этому ПК без запуска AnyDesk вручную станет невозможным.' `
        @{ Targets = @('AnyDesk'); Desired = 'Manual' }))
    $c.Add((New-Cat 'SVC_ADGUARD' 'Загрузка' 'Service' 'Низкий' 'Adguard Service → отложенный автозапуск' `
        'Служба фильтрации AdGuard уходит из критического пути загрузки (Delayed Auto ~1–2 мин).' `
        'Первые 1–2 минуты после старта реклама/трекеры не фильтруются.' `
        @{ Targets = @('Adguard Service'); Desired = 'DelayedAuto' }))
    $c.Add((New-Cat 'SVC_ADGUARDVPN' 'Загрузка' 'Service' 'Низкий' 'Adguard VPN Service → отложенный автозапуск' `
        'Служба VPN нужна клиенту, но не в первые секунды загрузки.' 'VPN-туннель доступен на 1–2 мин позже.' `
        @{ Targets = @('Adguard VPN Service'); Desired = 'DelayedAuto' }))
    $c.Add((New-Cat 'SVC_AMNEZIA' 'Загрузка' 'Service' 'Низкий' 'AmneziaVPN-service → отложенный автозапуск' `
        'Три VPN-службы одновременно стартуют при загрузке; эту можно отложить.' 'AmneziaVPN доступна на 1–2 мин позже.' `
        @{ Targets = @('AmneziaVPN-service'); Desired = 'DelayedAuto' }))
    $c.Add((New-Cat 'SVC_NLSVC' 'Загрузка' 'Service' 'Средний' 'NetLimiter (nlsvc) → отложенный автозапуск' `
        'Служба NetLimiter уходит из критического пути загрузки.' `
        'Правила/лимиты NetLimiter не действуют первые 1–2 минуты (если используете его как файрвол — не применяйте).' `
        @{ Targets = @('nlsvc'); Desired = 'DelayedAuto' }))
    $c.Add((New-Cat 'SVC_MAPSBROKER' 'Загрузка' 'Service' 'Низкий' 'MapsBroker → вручную' `
        'Диспетчер скачанных карт; нужен только офлайн-картам.' 'Офлайн-карты не обновляются автоматически.' `
        @{ Targets = @('MapsBroker'); Desired = 'Manual' }))
    $c.Add((New-Cat 'SVC_DIAGTRACK' 'Загрузка' 'Service' 'Средний' 'DiagTrack (телеметрия) → вручную' `
        'Служба диагностических данных Microsoft; небольшая фоновая нагрузка.' `
        'Диагностические данные и часть функций «Отзывы и диагностика» не работают; политика AllowTelemetry не меняется.' `
        @{ Targets = @('DiagTrack'); Desired = 'Manual' }))
    # --- Загрузка: элементы автозагрузки (как отключение в Диспетчере задач, обратимо) ---
    $c.Add((New-Cat 'STARTUP_OODEFRAGTRAY' 'Загрузка' 'Startup' 'Низкий' 'OODefragTray — отключить автозапуск' `
        'Иконка O&O Defrag в трее.' 'Нет значка O&O в трее.' @{ Targets = @('OODefragTray'); Location = '*' }))
    $c.Add((New-Cat 'STARTUP_ADSK_ACCESS' 'Загрузка' 'Startup' 'Низкий' 'Autodesk Access — отключить автозапуск' `
        'Интерфейс и служебный процесс Autodesk Access при входе.' 'Уведомления об обновлениях Autodesk.' `
        @{ Targets = @('Autodesk Access', 'Autodesk Access Service'); Location = '*' }))
    $c.Add((New-Cat 'STARTUP_FIIO' 'Загрузка' 'Startup' 'Низкий' 'FiiO Control Panel — отключить автозапуск' `
        'Панель FiiO нужна только для настройки ЦАП.' 'Запускайте панель вручную при смене настроек ЦАП.' `
        @{ Targets = @('FiiO Control Panel Autostart.lnk'); Location = '*' }))
    $c.Add((New-Cat 'STARTUP_ADGUARD_UI' 'Загрузка' 'Startup' 'Средний' 'Adguard (интерфейс) — отключить автозапуск' `
        'Окно/трей AdGuard; фильтрация выполняется службой Adguard Service.' `
        'Нет значка в трее; проверьте после применения, что фильтрация работает (у разных версий поведение может отличаться).' `
        @{ Targets = @('Adguard'); Location = '*' }))
    $c.Add((New-Cat 'STARTUP_ADGUARDVPN' 'Загрузка' 'Startup' 'Средний' 'AdGuard VPN (клиент) — отключить автозапуск' `
        'Клиент AdGuard VPN при входе (на T14S ~400 МБ Private).' 'VPN не подключится автоматически при входе — запускайте вручную.' `
        @{ Targets = @('AdGuardVpn'); Location = '*' }))
    $c.Add((New-Cat 'STARTUP_GOOGLEDRIVE' 'Загрузка' 'Startup' 'Средний' 'Google Drive — отключить автозапуск' `
        'Клиент синхронизации Google Drive при входе.' 'Синхронизация и диск G: появятся только после ручного запуска.' `
        @{ Targets = @('GoogleDriveFS'); Location = '*' }))
    $c.Add((New-Cat 'TASK_HDSENTINEL' 'Загрузка' 'Task' 'Средний' 'Hard Disk Sentinel — отключить задачу автозапуска' `
        'Постоянный фоновый опрос SMART/температуры.' 'Нет фонового мониторинга дисков и предупреждений HDSentinel.' `
        @{ Targets = @('*Sentinel*') }))
    # --- ОЗУ ---
    $c.Add((New-Cat 'PAGEFILE' 'ОЗУ' 'Pagefile' 'Низкий' 'Файл подкачки: минимум/максимум по пиковому commit' `
        'Фиксированный размер ограничивает предел выделения памяти (Commit Limit = ОЗУ + pagefile); при его достижении программы завершаются с нехваткой памяти.' `
        'Требуется перезагрузка. Файл может расти до максимума; на NVMe это не замедляет работу заметно.' @{ }))
    $c.Add((New-Cat 'MMA_PRELAUNCH' 'ОЗУ' 'MMAgent' 'Низкий' 'Отключить предзапуск приложений (ApplicationPreLaunch)' `
        'Windows заранее запускает часть UWP-приложений в фоне.' 'Первый запуск некоторых приложений Store чуть дольше. Сжатие памяти не затрагивается.' `
        @{ Feature = 'ApplicationPreLaunch' }))
    $c.Add((New-Cat 'REG_WIDGETS' 'ОЗУ' 'Reg' 'Низкий' 'Политика: отключить Виджеты' `
        'Групповая политика «Allow widgets» = 0: Widgets/WebView2 не работают в фоне.' 'Панель виджетов недоступна.' `
        @{ Hive = 'HKLM'; SubKey = 'SOFTWARE\Policies\Microsoft\Dsh'; Name = 'AllowNewsAndInterests'; ValueKind = 'DWord'; Value = 0 }))
    $c.Add((New-Cat 'REG_EDGE_STARTUPBOOST' 'ОЗУ' 'Reg' 'Низкий' 'Политика Edge: отключить Startup Boost' `
        'Edge предзагружает процессы при входе в систему даже без использования.' 'Edge открывается на долю секунды дольше; в Edge появится «Управляется организацией».' `
        @{ Hive = 'HKLM'; SubKey = 'SOFTWARE\Policies\Microsoft\Edge'; Name = 'StartupBoostEnabled'; ValueKind = 'DWord'; Value = 0 }))
    $c.Add((New-Cat 'REG_EDGE_BACKGROUND' 'ОЗУ' 'Reg' 'Низкий' 'Политика Edge: не работать в фоне после закрытия' `
        'Edge и его расширения продолжают работать после закрытия окна.' 'Расширения Edge не работают при закрытом браузере.' `
        @{ Hive = 'HKLM'; SubKey = 'SOFTWARE\Policies\Microsoft\Edge'; Name = 'BackgroundModeEnabled'; ValueKind = 'DWord'; Value = 0 }))
    # --- Батарея (только ноутбук) ---
    $c.Add((New-Cat 'PWR_NET_STANDBY' 'Батарея' 'Power' 'Низкий' 'Сеть в современном режиме сна на батарее: отключить' `
        'Параметр Microsoft «Allow networking during standby» (SUB_NONE / F15576E8…), значение 0 = отключать сеть во сне.' `
        'Почта и уведомления не синхронизируются во сне от батареи. Lenovo/Intel утилиты могут сбросить схему питания.' `
        @{ Sub = 'SUB_NONE'; Setting = 'F15576E8-98B7-4186-B944-EAFA664402D9'; Value = 0; LaptopOnly = $true
           Precheck = { if (-not $script:Sys.PowerCaps.AoAc) { 'нет Modern Standby (S0) — параметр не используется' } } }))
    $c.Add((New-Cat 'PWR_WAKE_TIMERS' 'Батарея' 'Power' 'Низкий' 'Таймеры пробуждения на батарее: отключить' `
        'Задачи (обслуживание, обновления) не будят ноутбук от батареи.' 'Обслуживание выполняется при работе от сети.' `
        @{ Sub = '238C9FA8-0AAD-41ED-83F4-97BE242C8F20'; Setting = 'BD3B718A-0680-4D9D-8AB2-E1D2B4AC806D'; Value = 0; LaptopOnly = $true }))
    $c.Add((New-Cat 'REG_FASTSTARTUP' 'Загрузка' 'Reg' 'Средний' 'Быстрый запуск (Fast Startup): включить' `
        'Холодный старт после «Завершения работы» использует гибернацию ядра — обычно заметно быстрее.' `
        'Не влияет на «Перезагрузку». Проверьте совместимость с шифрованием TrueCrypt/VeraCrypt и двойной загрузкой; диск остаётся «занятым» для других ОС.' `
        @{ Hive = 'HKLM'; SubKey = 'SYSTEM\CurrentControlSet\Control\Session Manager\Power'; Name = 'HiberbootEnabled'; ValueKind = 'DWord'; Value = 1
           Precheck = {
               $pol = Get-RegValueInfo HKLM 'SOFTWARE\Policies\Microsoft\Windows\System' 'HiberbootEnabled'
               if ($pol.Exists) { "задано групповой политикой (HiberbootEnabled=$($pol.Value))" }
               elseif (-not $script:Sys.PowerCaps.S4 -or -not $script:Sys.PowerCaps.HiberFilePresent) { 'гибернация выключена (powercfg /h on — создаст hiberfil.sys)' } } }))
    return $c
}

function Get-RiskRank { param([string]$Risk) switch ($Risk) { 'Низкий' { 1 } 'Средний' { 2 } default { 3 } } }
function Test-Selected {
    param($Item)
    if ($Exclude -contains $Item.Id) { return $false }
    if ($Include -contains $Item.Id) { return $true }
    $lvl = switch ($Level) { 'None' { 0 } 'Low' { 1 } 'Medium' { 2 } }
    return ((Get-RiskRank $Item.Risk) -le $lvl)
}
function Test-Protected {
    param([string]$Name)
    foreach ($p in $script:Protected) { if ($Name -like $p) { return $true } }
    return $false
}

# ----------------------------------------------------------------------------------------------
# 3. Сбор данных
# ----------------------------------------------------------------------------------------------
function Get-PowerCaps {
    $b = New-Object byte[] 256
    $ok = $false
    try { $ok = [W11Tk2.Native]::GetPwrCapabilities($b) } catch { }
    # смещения полей SYSTEM_POWER_CAPABILITIES (winnt.h)
    return [pscustomobject]@{
        Ok = $ok; LidPresent = [bool]$b[2]; S3 = [bool]$b[5]; S4 = [bool]$b[6]; HiberFilePresent = [bool]$b[8]
        Hiberboot = [bool]$b[18]; AoAc = [bool]$b[20]; AoAcConnectivity = [bool]$b[23]; BatteriesPresent = [bool]$b[30]
    }
}

function Get-SystemInfo {
    $cs   = Get-Cim Win32_ComputerSystem | Select-Object -First 1
    $os   = Get-Cim Win32_OperatingSystem | Select-Object -First 1
    $cpu  = Get-Cim Win32_Processor | Select-Object -First 1
    $bios = Get-Cim Win32_BIOS | Select-Object -First 1
    $cv   = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $hw   = Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -ErrorAction SilentlyContinue
    $cpuR = Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\0' -ErrorAction SilentlyContinue
    $build = [int]$cv.CurrentBuild
    $caps = Get-PowerCaps
    $bat  = @(Get-Cim Win32_Battery -Quiet)
    # время старта ядра без WMI: GetTickCount64 считает и время сна, как и LastBootUpTime
    $lastBoot = if ($os -and $os.LastBootUpTime) { $os.LastBootUpTime } else { (Get-Date).AddMilliseconds(-[double][W11Tk2.Native]::GetTickCount64()) }
    $osName = if ($build -ge 22000) { 'Windows 11' } else { 'Windows 10' }   # ProductName в реестре Win11 пишет «Windows 10»
    return [pscustomobject]@{
        Manufacturer = $(if ($cs) { $cs.Manufacturer } else { $hw.SystemManufacturer })
        Model = $(if ($cs) { $cs.Model } else { $hw.SystemProductName })
        SystemFamily = $(if ($cs) { $cs.SystemFamily } else { $hw.SystemFamily })
        CPU = $(if ($cpu) { "$($cpu.Name)".Trim() } else { "$($cpuR.ProcessorNameString)".Trim() })
        Cores = $(if ($cpu) { $cpu.NumberOfCores } else { $null }); Threads = $(if ($cpu) { $cpu.NumberOfLogicalProcessors } else { [Environment]::ProcessorCount })
        BIOS = $(if ($bios) { "$($bios.SMBIOSBIOSVersion)" } else { "$($hw.BIOSVersion)" }); OS = "$osName $($cv.EditionID)"; DisplayVersion = "$($cv.DisplayVersion)"
        Build = "$($cv.CurrentBuild).$($cv.UBR)"; LastBoot = $lastBoot
        IsLaptop = ($caps.BatteriesPresent -or $bat.Count -gt 0 -or $cs.PCSystemType -eq 2)
        PowerCaps = $caps; IsAdmin = (Test-IsAdmin)
    }
}

function Get-BootTimeline {
    $boot = $script:Sys.LastBoot
    $kernel = $boot
    $k12 = Get-EventsSafe @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Kernel-General'; Id = 12; StartTime = $boot.AddMinutes(-2) } |
           Sort-Object TimeCreated | Select-Object -First 1
    if ($k12) {
        $d = Get-EventDataMap $k12
        if ($d['StartTime']) { try { $kernel = ([datetime]::Parse($d['StartTime'], [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)).ToLocalTime() } catch { } }
    }
    $bootType = $null
    $kb = Get-EventsSafe @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Kernel-Boot'; Id = 27; StartTime = $kernel.AddMinutes(-2) } | Sort-Object TimeCreated | Select-Object -First 1
    if ($kb) { $bd = Get-EventDataMap $kb; if ($bd.ContainsKey('BootType')) { $bootType = [int]$bd['BootType'] } }
    $evlog = Get-EventsSafe @{ LogName = 'System'; ProviderName = 'EventLog'; Id = 6005; StartTime = $kernel } | Sort-Object TimeCreated | Select-Object -First 1
    $logon = Get-EventsSafe @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Winlogon'; Id = 7001; StartTime = $kernel } | Sort-Object TimeCreated | Select-Object -First 1
    $explorer = $script:Procs | Where-Object { $_.Name -eq 'explorer.exe' -and $_.Start -ge $kernel } | Sort-Object Start | Select-Object -First 1
    $logonTime = if ($logon) { $logon.TimeCreated } elseif ($explorer) { $explorer.Start } else { $null }
    $auto = Get-RegValueInfo HKLM 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' 'AutoAdminLogon'
    $diag = @(Get-EventsSafe @{ LogName = 'Microsoft-Windows-Diagnostics-Performance/Operational'; Id = 100 } 5)
    $diagRows = foreach ($e in $diag) {
        $m = Get-EventDataMap $e
        [pscustomobject]@{ TimeCreated = Format-Time $e.TimeCreated; Id = 100; Friendly = 'Загрузка (Event 100)'; FileName = ''
                           BootTimeMs = $m['BootTime']; MainPathMs = $m['MainPathBootTime']; PostBootMs = $m['BootPostBootTime']; DegradMs = '' }
    }
    $deg = @(Get-EventsSafe @{ LogName = 'Microsoft-Windows-Diagnostics-Performance/Operational'; Id = 101, 102, 103, 106, 107, 109, 110 } 40)
    $degRows = foreach ($e in $deg) {
        $m = Get-EventDataMap $e
        [pscustomobject]@{ TimeCreated = Format-Time $e.TimeCreated; Id = $e.Id; Friendly = $m['FriendlyName']; FileName = $m['Name']
                           BootTimeMs = $m['TotalTime']; MainPathMs = ''; PostBootMs = ''; DegradMs = $m['DegradationTime'] }
    }
    $cut = (Get-Date).AddDays(-30)
    $unexp = @(Get-EventsSafe @{ LogName = 'System'; ProviderName = 'EventLog'; Id = 6008; StartTime = $cut }).Count
    $kp41  = @(Get-EventsSafe @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Kernel-Power'; Id = 41; StartTime = $cut }).Count
    $preLogon = if ($logonTime) { [math]::Round(($logonTime - $kernel).TotalSeconds, 1) } else { $null }
    return [pscustomobject]@{
        KernelStart = $kernel; EventLogStart = $(if ($evlog) { $evlog.TimeCreated } else { $null }); Logon = $logonTime
        Explorer = $(if ($explorer) { $explorer.Start } else { $null }); PreLogonSec = $preLogon
        ScmSec = $(if ($evlog) { [math]::Round(($evlog.TimeCreated - $kernel).TotalSeconds, 1) } else { $null })
        DesktopAfterLogonSec = $(if ($explorer -and $logonTime) { [math]::Round(($explorer.Start - $logonTime).TotalSeconds, 1) } else { $null })
        AutoLogon = ($auto.Exists -and "$($auto.Value)" -eq '1')
        BootType = $bootType; BootTypeText = $(switch ($bootType) { 0 { 'холодный старт' } 1 { 'быстрый запуск (Fast Startup)' } 2 { 'выход из гибернации' } default { 'нет данных' } })
        DiagPerf = @($diagRows) + @($degRows); Unexpected30d = $unexp; KernelPower41_30d = $kp41
    }
}

function Get-ProcSnapshot {
    $props = 'ProcessId', 'ParentProcessId', 'Name', 'ExecutablePath', 'CreationDate', 'KernelModeTime', 'UserModeTime',
             'ReadTransferCount', 'WriteTransferCount', 'WorkingSetSize', 'PrivatePageCount', 'SessionId'
    $list = @(Get-Cim Win32_Process -Property $props)
    if (-not $list.Count) {
        # запасной путь без WMI: нет родителя и счётчиков ввода-вывода, но есть время старта, CPU и память
        $script:ProcFallback = $true
        $res = foreach ($p in @(Get-Process -ErrorAction SilentlyContinue)) {
            $st = $null; $cpuMs = 0; $path = ''
            try { $st = $p.StartTime } catch { }
            try { $cpuMs = $p.TotalProcessorTime.TotalMilliseconds } catch { }
            try { $path = "$($p.Path)" } catch { }
            [pscustomobject]@{ Pid = $p.Id; PPid = -1; Name = "$($p.ProcessName).exe"; Path = $path; Start = $st; CpuMs = $cpuMs; IoBytes = 0
                               WSMB = [math]::Round($p.WorkingSet64 / 1MB, 1); PrivMB = [math]::Round($p.PrivateMemorySize64 / 1MB, 1); Session = $p.SessionId }
        }
        $script:Children = @{}
        return @($res)
    }
    $res = foreach ($p in $list) {
        [pscustomobject]@{
            Pid = [int]$p.ProcessId; PPid = [int]$p.ParentProcessId; Name = "$($p.Name)"; Path = "$($p.ExecutablePath)"
            Start = $p.CreationDate; CpuMs = ([double]$p.KernelModeTime + [double]$p.UserModeTime) / 10000.0
            IoBytes = [double]$p.ReadTransferCount + [double]$p.WriteTransferCount
            WSMB = [math]::Round([double]$p.WorkingSetSize / 1MB, 1); PrivMB = [math]::Round([double]$p.PrivatePageCount / 1MB, 1)
            Session = $p.SessionId
        }
    }
    $script:Children = @{}
    foreach ($p in $res) {
        if (-not $script:Children.ContainsKey($p.PPid)) { $script:Children[$p.PPid] = New-Object System.Collections.Generic.List[object] }
        $script:Children[$p.PPid].Add($p)
    }
    return @($res)
}

function Get-TreeStats {
    param([object[]]$Roots)
    $seen = @{}; $stack = New-Object System.Collections.Stack
    foreach ($r in $Roots) { $stack.Push($r) }
    $cpu = 0.0; $io = 0.0; $ws = 0.0; $priv = 0.0; $n = 0
    while ($stack.Count) {
        $p = $stack.Pop()
        if ($seen.ContainsKey($p.Pid)) { continue }
        $seen[$p.Pid] = $true; $n++
        $cpu += $p.CpuMs; $io += $p.IoBytes; $ws += $p.WSMB; $priv += $p.PrivMB
        if ($p.Pid -ne 0 -and $script:Children.ContainsKey($p.Pid)) {
            foreach ($ch in $script:Children[$p.Pid]) { if ($ch.Start -and $p.Start -and $ch.Start -ge $p.Start) { $stack.Push($ch) } }
        }
    }
    return [pscustomobject]@{ Count = $n; CpuMs = $cpu; IoBytes = $io; WSMB = $ws; PrivMB = $priv }
}

function Get-ImpactClass {
    param([double]$CpuMs, [double]$IoBytes)
    if ($CpuMs -gt 1000 -or $IoBytes -gt 3MB) { return 'Высокое' }
    if ($CpuMs -ge 300 -or $IoBytes -ge 300KB) { return 'Среднее' }
    return 'Низкое'
}

$script:RunLocations = @(
    @{ Loc = 'HKCU\Run';   Hive = 'HKCU'; Key = 'Software\Microsoft\Windows\CurrentVersion\Run';             Approved = 'Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run' },
    @{ Loc = 'HKLM\Run';   Hive = 'HKLM'; Key = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Run';             Approved = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run' },
    @{ Loc = 'HKLM\Run32'; Hive = 'HKLM'; Key = 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'; Approved = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32' }
)
$script:FolderLocations = @(
    @{ Loc = 'HKCU\StartupFolder'; Hive = 'HKCU'; Folder = [Environment]::GetFolderPath('Startup');       Approved = 'Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder' },
    @{ Loc = 'HKLM\StartupFolder'; Hive = 'HKLM'; Folder = [Environment]::GetFolderPath('CommonStartup'); Approved = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder' }
)

function Get-ApprovedState {
    param([string]$Hive, [string]$Approved, [string]$Name)
    $v = Get-RegValueInfo $Hive $Approved $Name
    # байт 0: чётный (02/06) — включено, нечётный (03/07) — отключено; нет значения — включено
    $enabled = $true
    if ($v.Exists -and $v.Kind -eq 'Binary' -and "$($v.Value)".Length -ge 2) { $enabled = (([Convert]::ToByte("$($v.Value)".Substring(0, 2), 16) -band 1) -eq 0) }
    return [pscustomobject]@{ Enabled = $enabled; Raw = $v }
}

function Get-StartupEntries {
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($l in $script:RunLocations) {
        $k = Open-RegKey $l.Hive $l.Key
        if (-not $k) { continue }
        try {
            foreach ($n in $k.GetValueNames()) {
                if (-not $n) { continue }
                $cmd = "$($k.GetValue($n))"
                $st = Get-ApprovedState $l.Hive $l.Approved $n
                $list.Add([pscustomobject]@{ Kind = 'Run'; Loc = $l.Loc; Hive = $l.Hive; Approved = $l.Approved; Name = $n
                                             Command = $cmd; Exe = (Resolve-ImagePath $cmd); Enabled = $st.Enabled })
            }
        } finally { $k.Close() }
    }
    $shell = $null
    try { $shell = New-Object -ComObject WScript.Shell } catch { }
    foreach ($l in $script:FolderLocations) {
        if (-not $l.Folder -or -not (Test-Path -LiteralPath $l.Folder)) { continue }
        foreach ($f in Get-ChildItem -LiteralPath $l.Folder -File -ErrorAction SilentlyContinue) {
            if ($f.Name -eq 'desktop.ini') { continue }
            $target = $f.FullName
            if ($f.Extension -eq '.lnk' -and $shell) { try { $t = $shell.CreateShortcut($f.FullName).TargetPath; if ($t) { $target = $t } } catch { } }
            $st = Get-ApprovedState $l.Hive $l.Approved $f.Name
            $list.Add([pscustomobject]@{ Kind = 'StartupFolder'; Loc = $l.Loc; Hive = $l.Hive; Approved = $l.Approved; Name = $f.Name
                                         Command = $target; Exe = $target; Enabled = $st.Enabled })
        }
    }
    foreach ($t in @(Get-ScheduledTask -ErrorAction SilentlyContinue)) {
        if ($t.TaskPath -like '\Microsoft\*') { continue }
        $trig = @($t.Triggers | ForEach-Object { if ($_) { $_.CimClass.CimClassName } })
        if (-not ($trig -match 'Logon|Boot')) { continue }
        $act = @($t.Actions | Where-Object { $_.Execute }) | Select-Object -First 1
        $exe = if ($act) { Resolve-ImagePath ('"' + [Environment]::ExpandEnvironmentVariables($act.Execute.Trim('"')) + '"') } else { '' }
        $list.Add([pscustomobject]@{ Kind = 'Task'; Loc = $t.TaskPath; Hive = ''; Approved = ''; Name = $t.TaskName
                                     Command = $(if ($act) { "$($act.Execute) $($act.Arguments)" } else { '' }); Exe = $exe
                                     Enabled = ($t.State -ne 'Disabled') })
    }
    return $list
}

function Get-ServiceTable {
    $root = 'SYSTEM\CurrentControlSet\Services'
    $wmi = @{}
    foreach ($s in @(Get-Cim Win32_Service)) { $wmi[$s.Name.ToLowerInvariant()] = $s }
    $reg = @{}; $rev = @{}
    $base = Open-RegKey HKLM $root
    try {
        foreach ($n in $base.GetSubKeyNames()) {
            $k = $base.OpenSubKey($n)
            if (-not $k) { continue }
            try {
                $type = $k.GetValue('Type')
                if ($null -eq $type) { continue }
                $deps = @($k.GetValue('DependOnService', @()))
                $trigKey = $k.OpenSubKey('TriggerInfo'); $hasTrig = [bool]$trigKey; if ($trigKey) { $trigKey.Close() }
                $reg[$n.ToLowerInvariant()] = [pscustomobject]@{ Name = $n; Type = [int]$type; Start = $k.GetValue('Start'); Delayed = [int]$k.GetValue('DelayedAutoStart', 0)
                    Image = "$($k.GetValue('ImagePath', '', [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames))"; Group = "$($k.GetValue('Group'))"
                    DependOn = $deps; Trigger = $hasTrig; Display = "$($k.GetValue('DisplayName'))" }
                foreach ($d in $deps) { if ($d) { $dk = $d.ToLowerInvariant(); if (-not $rev.ContainsKey($dk)) { $rev[$dk] = New-Object System.Collections.Generic.List[string] }; $rev[$dk].Add($n) } }
            } finally { $k.Close() }
        }
    } finally { $base.Close() }
    $rows = foreach ($key in $reg.Keys) {
        $r = $reg[$key]
        if ($null -eq $r.Start) { continue }
        $w = $wmi[$key]
        $img = Resolve-ImagePath $r.Image
        $company = Get-FileCompany $img
        $isDriver = ($r.Type -band 0x0B) -ne 0 -and ($r.Type -band 0x30) -eq 0
        [pscustomobject]@{
            Name = $r.Name; Display = $(if ($w) { $w.DisplayName } else { $r.Display }); IsDriver = $isDriver
            Start = [int]$r.Start; Delayed = $r.Delayed; StartText = (Get-StartText ([int]$r.Start) $r.Delayed); Trigger = $r.Trigger
            State = $(if ($w) { "$($w.State)" } else { '' }); ProcessId = $(if ($w) { [int]$w.ProcessId } else { 0 })
            Image = $img; Company = $company; IsMicrosoft = (Test-MicrosoftCompany $company); Group = $r.Group
            DependOn = @($r.DependOn | Where-Object { $_ }); DependedBy = @($(if ($rev.ContainsKey($key)) { $rev[$key] } else { @() }))
            Shared = ($img -match '(?i)\\svchost\.exe$'); Protected = (Test-Protected $r.Name)
        }
    }
    return @($rows)
}

function Get-StartText {
    param([int]$Start, [int]$Delayed)
    switch ($Start) { 0 { 'Boot' } 1 { 'System' } 2 { if ($Delayed) { 'Auto (отложенно)' } else { 'Auto' } } 3 { 'Вручную' } 4 { 'Отключено' } default { "$Start" } }
}

function Get-DriverStatus {
    $sd = @{}
    foreach ($d in @(Get-Cim Win32_SystemDriver)) { $sd[$d.Name.ToLowerInvariant()] = $d }
    $rows = foreach ($s in $script:Services) {
        if (-not $s.IsDriver -or $s.Start -gt 2) { continue }
        $w = $sd[$s.Name.ToLowerInvariant()]
        $state = if ($w) { "$($w.State)" } else { '' }
        $failed = ($state -eq 'Stopped')
        if ($s.IsMicrosoft -and -not $failed) { continue }
        $note = ''
        foreach ($k in $script:KnownNotes.Keys) { if ($s.Name -like "$k*") { $note = $script:KnownNotes[$k]; break } }
        [pscustomobject]@{ Name = $s.Name; Display = $s.Display; Start = $s.StartText; State = $state; Company = $s.Company
                           Group = $s.Group; Image = $s.Image; NotLoaded = $failed; Protected = $s.Protected; Note = $note }
    }
    return @($rows | Sort-Object @{ e = { -not $_.NotLoaded } }, Name)
}

function Get-ScmIssues {
    $ids = 7000, 7001, 7009, 7011, 7022, 7023, 7024, 7026, 7031, 7034, 7038, 7043, 7045
    $ev = Get-EventsSafe @{ LogName = 'System'; ProviderName = 'Service Control Manager'; Id = $ids; StartTime = $script:Boot.KernelStart.AddMinutes(-1) } 300
    $rows = foreach ($e in ($ev | Sort-Object TimeCreated)) {
        $msg = ("$($e.Message)" -replace '\s+', ' ').Trim()
        $note = ''
        foreach ($k in $script:KnownNotes.Keys) { if ($msg -match [regex]::Escape($k)) { $note = $script:KnownNotes[$k]; break } }
        [pscustomobject]@{ Time = Format-Time $e.TimeCreated; EventID = $e.Id; Source = 'Service Control Manager'; Message = $msg; Note = $note }
    }
    return @($rows)
}

function Get-MemoryInfo {
    $pi = New-Object 'W11Tk2.Native+PERFORMANCE_INFORMATION'
    $size = [uint32][Runtime.InteropServices.Marshal]::SizeOf([type]'W11Tk2.Native+PERFORMANCE_INFORMATION')
    [void][W11Tk2.Native]::GetPerformanceInfo([ref]$pi, $size)
    $pg = [double]$pi.PageSize.ToUInt64()
    function _mb([UIntPtr]$v) { [math]::Round([double]$v.ToUInt64() * $pg / 1MB, 0) }
    $total = _mb $pi.PhysicalTotal; $avail = _mb $pi.PhysicalAvailable
    $cs = Get-Cim Win32_ComputerSystem | Select-Object -First 1   # без WMI: «?:\pagefile.sys» в PagingFiles = размер выбирает система
    $pf = Get-RegValueInfo HKLM 'SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management' 'PagingFiles'
    $usage = @(Get-Cim Win32_PageFileUsage | ForEach-Object {
        [pscustomobject]@{ File = $_.Name; AllocatedMB = $_.AllocatedBaseSize; CurrentMB = $_.CurrentUsage; PeakMB = $_.PeakUsage } })
    $sysDrive = $env:SystemDrive
    $disk = Get-Cim Win32_LogicalDisk -Filter ("DeviceID='{0}'" -f $sysDrive) | Select-Object -First 1
    $freeMB = if ($disk) { [math]::Round($disk.FreeSpace / 1MB, 0) } else { try { [math]::Round((New-Object IO.DriveInfo $sysDrive).AvailableFreeSpace / 1MB, 0) } catch { $null } }
    $mm = $null
    try { $mm = Get-MMAgent -ErrorAction Stop } catch { }
    $mc = Get-Process -Name 'Memory Compression', 'MemCompression' -ErrorAction SilentlyContinue | Select-Object -First 1
    $dg = $null
    $dg = Get-Cim Win32_DeviceGuard -Namespace 'root\Microsoft\Windows\DeviceGuard' | Select-Object -First 1
    $dump = Get-RegValueInfo HKLM 'SYSTEM\CurrentControlSet\Control\CrashControl' 'CrashDumpEnabled'
    $hb = Get-RegValueInfo HKLM 'SYSTEM\CurrentControlSet\Control\Session Manager\Power' 'HiberbootEnabled'
    $sysmain = Get-Service -Name SysMain -ErrorAction SilentlyContinue
    return [pscustomobject]@{
        TotalPhysMB = $total; AvailPhysMB = $avail; UsedPhysMB = $total - $avail
        LoadPct = $(if ($total) { [math]::Round(100 * ($total - $avail) / $total, 0) } else { 0 })
        CommitUsedMB = (_mb $pi.CommitTotal); CommitTotalMB = (_mb $pi.CommitLimit); CommitPeakMB = (_mb $pi.CommitPeak)
        SystemCacheMB = (_mb $pi.SystemCache); KernelPagedMB = (_mb $pi.KernelPaged); KernelNonpagedMB = (_mb $pi.KernelNonpaged)
        AutoPagefile = $(if ($cs) { [bool]$cs.AutomaticManagedPagefile } else { [bool]((@($pf.Value) -join ' ') -match '^\?:\\') }); PagingFilesReg = @($pf.Value); PagefileUsage = $usage
        SystemDriveFreeMB = $freeMB
        MemoryCompression = $(if ($mm) { [bool]$mm.MemoryCompression } else { $null })
        ApplicationPreLaunch = $(if ($mm) { [bool]$mm.ApplicationPreLaunch } else { $null })
        PageCombining = $(if ($mm) { [bool]$mm.PageCombining } else { $null })
        CompressionWSMB = $(if ($mc) { [math]::Round($mc.WorkingSet64 / 1MB, 0) } else { $null })
        SysMain = $(if ($sysmain) { "$($sysmain.Status) / $($sysmain.StartType)" } else { 'нет' })
        VBS_Enabled = $(if ($dg) { [int]$dg.VirtualizationBasedSecurityStatus } else { $null })
        HVCI_Enabled = $(if ($dg) { [int](@($dg.SecurityServicesRunning) -contains 2) } else { $null })
        CrashDump = $(if ($dump.Exists) { switch ([int]$dump.Value) { 0 { 'нет' } 1 { 'полный' } 2 { 'ядра' } 3 { 'малый' } 7 { 'автоматический' } default { "$($dump.Value)" } } } else { '' })
        HiberbootEnabled = $(if ($hb.Exists) { [int]$hb.Value } else { $null })
    }
}

function Get-PagefileRecommendation {
    param($Mem)
    $ram = [double]$Mem.TotalPhysMB; $peak = [double]$Mem.CommitPeakMB
    $curInit = 0; $curMax = 0
    foreach ($line in @($Mem.PagingFilesReg)) {
        $m = [regex]::Match("$line", '^(?<f>\S+)\s+(?<i>\d+)\s+(?<x>\d+)')
        if ($m.Success) { $curInit += [int]$m.Groups['i'].Value; $curMax += [int]$m.Groups['x'].Value }
    }
    $round = { param($v) [int]([math]::Ceiling($v / 1024.0) * 1024) }
    $init = [math]::Max(4096, (& $round ($peak * 1.1 - $ram)))
    $init = [math]::Max($init, $curInit)
    $max = [math]::Max((& $round ($ram * 1.5)), (& $round ($peak * 1.5 - $ram)))
    $max = [math]::Min($max, (& $round ($ram * 3)))
    $note = "ОЗУ $([int]$ram) МБ, пиковый commit с загрузки $([int]$peak) МБ, предел $($Mem.CommitTotalMB) МБ."
    if ($PagefileInitialMB -gt 0) { $init = $PagefileInitialMB; $note += ' Минимум задан параметром.' }
    if ($PagefileMaxMB -gt 0) { $max = $PagefileMaxMB; $note += ' Максимум задан параметром.' }
    if ($max -lt $init) { $max = $init }
    $free = [double]$Mem.SystemDriveFreeMB
    $skip = $null
    if ($free -and ($free - $max) -lt 10240) { $skip = "на $env:SystemDrive мало места: свободно $([int]$free) МБ, нужно максимум + 10 ГБ" }
    if ($peak -gt 0.85 * $Mem.CommitTotalMB) { $note += ' Пик близок к пределу — риск нехватки памяти.' }
    return [pscustomobject]@{ InitialMB = $init; MaxMB = $max; CurrentInitMB = $curInit; CurrentMaxMB = $curMax; Note = $note; Skip = $skip }
}

function Get-TopProcesses {
    $g = $script:Procs | Group-Object Name | ForEach-Object {
        $m = $_.Group | Measure-Object -Property WSMB, PrivMB, CpuMs -Sum
        [pscustomobject]@{ Name = $_.Name; Count = $_.Count
            WS_MB = [math]::Round(($m | Where-Object Property -eq 'WSMB').Sum, 1)
            Priv_MB = [math]::Round(($m | Where-Object Property -eq 'PrivMB').Sum, 1)
            CPU_Sec = [math]::Round(($m | Where-Object Property -eq 'CpuMs').Sum / 1000, 1) }
    }
    return @($g | Sort-Object Priv_MB -Descending | Select-Object -First 30)
}

function Get-StartupImpact {
    $now = Get-Date
    $logon = $script:Boot.Logon
    $kernel = $script:Boot.KernelStart
    $rows = New-Object System.Collections.Generic.List[object]
    $add = {
        param($Kind, $Loc, $Name, $Exe, $Enabled, $Roots, $Anchor, $Extra)
        $Roots = @($Roots | Where-Object { $null -ne $_ })   # пустой return из scriptblock = $null, а @($null).Count = 1
        $st = Get-TreeStats $Roots
        $first = $Roots | Where-Object { $_.Start } | Sort-Object Start | Select-Object -First 1
        $offset = if ($first -and $Anchor) { [math]::Round(($first.Start - $Anchor).TotalSeconds, 1) } else { $null }
        $mins = if ($first) { [math]::Round(($now - $first.Start).TotalMinutes, 0) } else { $null }
        # пороги Microsoft относятся к ресурсам, потраченным при запуске; через час работы накопленный CPU уже ничего не говорит о старте
        $impact = if (@($Roots).Count -eq 0) { 'Не запущено' }
                  elseif ($null -ne $mins -and $mins -gt $script:ImpactWindowMin) { 'Нет оценки' }
                  elseif ($script:ProcFallback -and $st.IoBytes -eq 0) { (Get-ImpactClass $st.CpuMs 0) + ' (только CPU)' }
                  else { Get-ImpactClass $st.CpuMs $st.IoBytes }
        $notes = @($Extra)
        if ($impact -eq 'Нет оценки') { $notes += "накоплено за $mins мин работы — для оценки запуска используйте -Mode ScheduleMeasure" }
        if ($null -ne $offset -and $offset -lt 0) { $notes += 'процесс запущен до входа (вероятно, его служба), а не этим пунктом автозагрузки' }
        $Extra = (@($notes | Where-Object { $_ }) -join '; ')
        $rows.Add([pscustomobject]@{
            Kind = $Kind; Location = $Loc; Name = $Name; Enabled = $Enabled; Running = (@($Roots).Count -gt 0); Impact = $impact
            StartOffsetSec = $offset; CPU_ms = [math]::Round($st.CpuMs, 0); DiskIO_MB = [math]::Round($st.IoBytes / 1MB, 1)
            WS_MB = [math]::Round($st.WSMB, 1); Priv_MB = [math]::Round($st.PrivMB, 1); Processes = $st.Count; MinutesRunning = $mins
            Company = (Get-FileCompany $Exe); Exe = $Exe; Note = $Extra })
    }
    $byPath = @{}; $byName = @{}
    foreach ($p in $script:Procs) {
        if ($p.Path) { $k = $p.Path.ToLowerInvariant(); if (-not $byPath.ContainsKey($k)) { $byPath[$k] = New-Object System.Collections.Generic.List[object] }; $byPath[$k].Add($p) }
        $n = $p.Name.ToLowerInvariant(); if (-not $byName.ContainsKey($n)) { $byName[$n] = New-Object System.Collections.Generic.List[object] }; $byName[$n].Add($p)
    }
    $findRoots = {
        param([string]$Exe)
        if (-not $Exe) { return @() }
        $cand = @()
        $k = $Exe.ToLowerInvariant()
        if ($byPath.ContainsKey($k)) { $cand = $byPath[$k].ToArray() }   # .ToArray(): @(List[object]) в PS 5.1 даёт «Argument types do not match»
        else { $leaf = [IO.Path]::GetFileName($k); if ($byName.ContainsKey($leaf)) { $cand = @($byName[$leaf].ToArray() | Where-Object { -not $_.Path }) } }
        $pids = @{}; foreach ($c in $cand) { $pids[$c.Pid] = $true }
        return @($cand | Where-Object { -not $pids.ContainsKey($_.PPid) })
    }
    foreach ($e in $script:Startup) {
        if ($e.Exe -match '(?i)\\(rundll32|cmd|powershell|wscript|cscript|explorer)\.exe$') { & $add $e.Kind $e.Loc $e.Name $e.Exe $e.Enabled @() $logon 'запуск через системный хост — не сопоставляется'; continue }
        $roots = & $findRoots $e.Exe
        & $add $e.Kind $e.Loc $e.Name $e.Exe $e.Enabled $roots $logon ''
    }
    foreach ($s in $script:Services) {
        if ($s.IsDriver -or $s.IsMicrosoft -or $s.Start -ne 2) { continue }
        if ($s.Shared) { & $add 'Служба' $s.StartText $s.Name $s.Image $true @() $kernel 'общий svchost — не измеряется отдельно'; continue }
        $roots = @($script:Procs | Where-Object { $_.Pid -eq $s.ProcessId -and $s.ProcessId -gt 0 })
        & $add 'Служба' $s.StartText $s.Name $s.Image $true $roots $kernel $(if ($s.Protected) { 'защищено — не изменять' } else { '' })
    }
    $rank = @{ 'Высокое' = 0; 'Среднее' = 1; 'Низкое' = 2; 'Нет оценки' = 3; 'Не запущено' = 4 }
    return @($rows | Sort-Object @{ e = { $r = $rank[($_.Impact -replace ' \(.*$', '')]; if ($null -eq $r) { 5 } else { $r } } }, @{ e = { $_.CPU_ms }; Descending = $true })
}

function Get-RunGaps {
    $seq = @($script:Impact | Where-Object { ($_.Kind -eq 'Run' -or $_.Kind -eq 'StartupFolder') -and $null -ne $_.StartOffsetSec -and $_.StartOffsetSec -ge 0 } | Sort-Object StartOffsetSec)
    $res = for ($i = 1; $i -lt $seq.Count; $i++) {
        $gap = [math]::Round($seq[$i].StartOffsetSec - $seq[$i - 1].StartOffsetSec, 1)
        [pscustomobject]@{ From = $seq[$i - 1].Name; To = $seq[$i].Name; AtSec = $seq[$i].StartOffsetSec; GapSec = $gap; Suspicious = ($gap -ge 25) }
    }
    return @($res)
}

function Get-PowerIndex {
    param([string]$Sub, [string]$Setting, [string]$Scheme = 'SCHEME_CURRENT')
    $r = Invoke-Native 'powercfg.exe' @('/qh', $Scheme, $Sub, $Setting)   # /qh показывает и скрытые параметры (например, F15576E8…)
    if ($r.ExitCode -ne 0 -or -not $r.Output) { return $null }
    $hex = [regex]::Matches($r.Output, '0x([0-9a-fA-F]{8})')
    if ($hex.Count -lt 2) { return $null }
    return [pscustomobject]@{ AC = [Convert]::ToInt32($hex[$hex.Count - 2].Groups[1].Value, 16); DC = [Convert]::ToInt32($hex[$hex.Count - 1].Groups[1].Value, 16) }
}
function Get-ActiveSchemeGuid {
    $r = Invoke-Native 'powercfg.exe' @('/getactivescheme')
    $m = [regex]::Match($r.Output, '[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}')
    if ($m.Success) { return $m.Value } else { return $null }
}

function Get-PowerInfo {
    $bat = $null
    if ($script:Sys.IsLaptop) {
        $b = Get-Cim Win32_Battery -Quiet | Select-Object -First 1
        $design = $null; $full = $null; $cycles = $null
        try { $design = (Get-CimInstance -Namespace root\wmi -ClassName BatteryStaticData -ErrorAction Stop | Select-Object -First 1).DesignedCapacity } catch { }
        try { $full = (Get-CimInstance -Namespace root\wmi -ClassName BatteryFullChargedCapacity -ErrorAction Stop | Select-Object -First 1).FullChargedCapacity } catch { }
        try { $cycles = (Get-CimInstance -Namespace root\wmi -ClassName BatteryCycleCount -ErrorAction Stop | Select-Object -First 1).CycleCount } catch { }
        $bat = [pscustomobject]@{ ChargePct = $(if ($b) { $b.EstimatedChargeRemaining } else { $null }); DesignWh = $(if ($design) { [math]::Round($design / 1000, 1) } else { $null })
            FullWh = $(if ($full) { [math]::Round($full / 1000, 1) } else { $null }); Cycles = $cycles
            HealthPct = $(if ($design -and $full) { [math]::Round(100 * $full / $design, 0) } else { $null }) }
    }
    $net = Get-PowerIndex 'SUB_NONE' 'F15576E8-98B7-4186-B944-EAFA664402D9'
    $wake = Get-PowerIndex '238C9FA8-0AAD-41ED-83F4-97BE242C8F20' 'BD3B718A-0680-4D9D-8AB2-E1D2B4AC806D'
    $resets = @(Get-EventsSafe @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-UserModePowerService'; Id = 12; StartTime = $script:Boot.KernelStart } 100)
    $resetApps = @($resets | ForEach-Object { $m = [regex]::Match("$($_.Message)", '[^\\\s]+\.exe'); if ($m.Success) { $m.Value } } | Group-Object | ForEach-Object { "$($_.Name) ×$($_.Count)" })
    return [pscustomobject]@{
        Scheme = (Get-ActiveSchemeGuid); ModernStandby = $script:Sys.PowerCaps.AoAc; S3 = $script:Sys.PowerCaps.S3
        NetStandbyDC = $(if ($net) { $net.DC } else { $null }); WakeTimersDC = $(if ($wake) { $wake.DC } else { $null })
        Battery = $bat; SchemeChangers = $resetApps
        StandbyText = (Invoke-Native 'powercfg.exe' @('/a')).Output
    }
}

function Invoke-Collect {
    Write-Step "Сбор данных ($script:PC, $(if ($script:Sys.IsAdmin) { 'администратор' } else { 'без прав администратора — часть данных будет недоступна' }))"
    $tag = if ($FromScheduledTask) { 'PostLogon' } else { 'Manual' }
    $base = "{0}_{1}_{2}" -f $script:PC, $script:Stamp, $tag
    Initialize-Data
    $mem = Get-MemoryInfo
    $pfr = Get-PagefileRecommendation $mem
    $pwr = Get-PowerInfo
    $scm = Get-ScmIssues
    $drv = Get-DriverStatus
    $top = Get-TopProcesses
    $gaps = Get-RunGaps
    $plan = Get-Plan
    $b = $script:Boot

    $json = [ordered]@{
        Tool = "Win11-BootRamBattery-Toolkit $script:Version"; ComputerName = $script:PC; CollectedAt = Format-Time (Get-Date); Tag = $tag
        IsAdmin = $script:Sys.IsAdmin
        Hardware = [ordered]@{ Manufacturer = $script:Sys.Manufacturer; ProductName = $script:Sys.Model; Family = $script:Sys.SystemFamily; CPU = $script:Sys.CPU
                               BIOS = $script:Sys.BIOS; OS = $script:Sys.OS; DisplayVersion = $script:Sys.DisplayVersion; OSBuild = $script:Sys.Build; IsLaptop = $script:Sys.IsLaptop }
        BootAnchors = [ordered]@{ KernelStart = Format-Time $b.KernelStart; EventLogStart = Format-Time $b.EventLogStart; Logon = Format-Time $b.Logon
                                  Explorer = Format-Time $b.Explorer; PreLogonSec = $b.PreLogonSec; ScmSec = $b.ScmSec; DesktopAfterLogonSec = $b.DesktopAfterLogonSec
                                  AutoLogon = $b.AutoLogon; BootType = $b.BootTypeText; Unexpected30d = $b.Unexpected30d; KernelPower41_30d = $b.KernelPower41_30d }
        Memory = $mem; PagefileRecommendation = $pfr
        Power = $pwr; DiagPerfEvents = $b.DiagPerf; SCMErrors = $scm; Drivers = $drv
        StartupImpact = $script:Impact; RunGaps = $gaps; TopProcesses = $top
        Plan = @($plan | Select-Object Id, Group, Risk, Selected, Status, Target, Current, Desired, Title)
        Limitations = [ordered]@{ WmiUnavailable = @($script:CimProblems.Keys); ProcessFallback = $script:ProcFallback }
    }
    $jsonPath = Join-Path $script:OutDir "Win11-Telemetry_$base.json"
    Write-TextFile $jsonPath ($json | ConvertTo-Json -Depth 6)
    Export-RuCsv $script:Impact (Join-Path $script:OutDir "StartupImpact_$base.csv")
    Export-RuCsv ($script:Services | Where-Object { -not $_.IsMicrosoft -or $_.Start -le 2 } | Select-Object Name, Display, IsDriver, StartText, Trigger, State, Company, IsMicrosoft, Protected, Shared, DependOn, DependedBy, Image) (Join-Path $script:OutDir "Services_$base.csv")
    Export-RuCsv $drv (Join-Path $script:OutDir "Drivers_$base.csv")
    Export-RuCsv ($plan | Select-Object Id, Group, Risk, Selected, Status, Target, Current, Desired, Title, Reason, Effect) (Join-Path $script:OutDir "Plan_$base.csv")
    if ($script:Sys.IsLaptop) {
        $br = Join-Path $script:OutDir "battery-report_$base.html"
        [void](Invoke-Native 'powercfg.exe' @('/batteryreport', '/output', $br))
        if ($script:Sys.IsAdmin -and $script:Sys.PowerCaps.AoAc) { [void](Invoke-Native 'powercfg.exe' @('/sleepstudy', '/output', (Join-Path $script:OutDir "sleepstudy_$base.html"))) }
    }
    $md = New-SummaryMarkdown $mem $pfr $pwr $scm $drv $gaps $plan
    $mdPath = Join-Path $script:OutDir "Summary_$base.md"
    Write-TextFile $mdPath $md
    Write-Info "Готово. Папка: $script:OutDir" Green
    Write-Info "  Отчёт: $(Split-Path $mdPath -Leaf); JSON для HTML-анализатора: $(Split-Path $jsonPath -Leaf)" Green
    Show-ShortSummary $mem $pfr
    if ($FromScheduledTask) {
        try { Unregister-ScheduledTask -TaskName $script:MeasureTaskName -TaskPath $script:MeasureTaskPath -Confirm:$false -ErrorAction Stop } catch { }
    }
}

function Show-ShortSummary {
    param($Mem, $Pfr)
    $b = $script:Boot
    Write-Step 'Кратко'
    if ($script:CimProblems.Count) { Write-Warning ("WMI недоступен для: {0}. Использованы запасные источники (реестр, Get-Process); часть метрик пустая." -f (@($script:CimProblems.Keys) -join ', ')) }
    if ($script:ProcFallback) { Write-Warning 'Процессы прочитаны без WMI: нет дерева родитель→потомок и счётчиков диска, влияние на запуск оценено грубее.' }
    Write-Info ("Ядро → вход: {0} с{1}; вход → Проводник: {2} с; тип старта: {3}" -f $b.PreLogonSec, $(if (-not $b.AutoLogon) { ' (включает ожидание ввода пароля)' } else { '' }), $b.DesktopAfterLogonSec, $b.BootTypeText)
    Write-Info ("ОЗУ: {0}/{1} МБ ({2}%), commit {3}/{4} МБ, пик {5} МБ" -f $Mem.UsedPhysMB, $Mem.TotalPhysMB, $Mem.LoadPct, $Mem.CommitUsedMB, $Mem.CommitTotalMB, $Mem.CommitPeakMB)
    Write-Info ("Файл подкачки: сейчас {0}; рекомендуется {1}–{2} МБ" -f (($Mem.PagingFilesReg) -join ', '), $Pfr.InitialMB, $Pfr.MaxMB)
    $hi = @($script:Impact | Where-Object { $_.Impact -like 'Высокое*' } | Select-Object -First 10)
    if ($hi.Count) {
        Write-Info 'Высокое влияние на запуск (пороги Microsoft: CPU > 1 с или диск > 3 МБ):' Yellow
        $hi | Format-Table Kind, Name, Impact, StartOffsetSec, CPU_ms, DiskIO_MB, Priv_MB -AutoSize | Out-String -Width 200 | Write-Host
    }
    $late = @($script:Impact | Where-Object { $_.Impact -eq 'Нет оценки' } | Sort-Object CPU_ms -Descending | Select-Object -First 8)
    if ($late.Count) {
        Write-Info ("Замер сделан позже {0} мин после входа — влияние на запуск не оценивается. Самые затратные пункты автозагрузки за время работы:" -f $script:ImpactWindowMin) Yellow
        $late | Format-Table Kind, Name, StartOffsetSec, CPU_ms, Priv_MB, MinutesRunning -AutoSize | Out-String -Width 200 | Write-Host
        Write-Info 'Точная оценка: -Mode ScheduleMeasure (замер через 2 мин после следующего входа) или -Mode BootTraceStart (трасса WPR).' Gray
    }
}

function Initialize-Data {
    if ($script:DataReady) { return }
    Write-Info 'Чтение системы, процессов, автозагрузки, служб и журналов…'
    $script:Procs    = Get-ProcSnapshot
    $script:Boot     = Get-BootTimeline
    $script:Startup  = Get-StartupEntries
    $script:Services = Get-ServiceTable
    $script:Impact   = Get-StartupImpact
    $script:DataReady = $true
}

function New-SummaryMarkdown {
    param($Mem, $Pfr, $Pwr, $Scm, $Drv, $Gaps, $Plan)
    $s = $script:Sys; $b = $script:Boot
    $sb = New-Object System.Text.StringBuilder
    $L = { param([string]$t) [void]$sb.AppendLine($t) }
    $esc = { param($v) ("$v" -replace '\|', '/' -replace '[\r\n]+', ' ') }
    & $L "# Отчёт Win11 Boot/RAM/Battery Toolkit $script:Version — $script:PC"
    & $L ''
    & $L "Собрано: $(Format-Time (Get-Date)); права администратора: $(if ($s.IsAdmin) { 'да' } else { 'нет (часть данных недоступна)' }); замер: $(if ($FromScheduledTask) { "через $MeasureDelaySec с после входа" } else { 'вручную' })."
    & $L ''
    & $L '## Система'
    & $L "- $($s.Manufacturer) $($s.Model) ($($s.SystemFamily)), BIOS $($s.BIOS); $($s.CPU), $(if ($s.Cores) { "$($s.Cores) ядер / " })$($s.Threads) потоков"
    & $L "- $($s.OS) $($s.DisplayVersion), сборка $($s.Build); ноутбук: $(if ($s.IsLaptop) { 'да' } else { 'нет' }); Modern Standby: $(if ($s.PowerCaps.AoAc) { 'да' } else { 'нет' }); S3: $(if ($s.PowerCaps.S3) { 'да' } else { 'нет' })"
    & $L "- VBS: $(if ($null -ne $Mem.VBS_Enabled) { "$($Mem.VBS_Enabled) (2 = работает)" } else { 'нет данных (WMI)' }), HVCI: $(if ($null -ne $Mem.HVCI_Enabled) { $Mem.HVCI_Enabled } else { 'нет данных' }); Fast Startup (HiberbootEnabled): $($Mem.HiberbootEnabled)"
    & $L ''
    & $L '## Хронология последней загрузки'
    & $L "| Этап | Время | От старта ядра, с |"
    & $L '|---|---|---|'
    & $L "| Старт ядра (Kernel-General 12) | $(Format-Time $b.KernelStart) | 0 |"
    & $L "| Журнал событий (6005) | $(Format-Time $b.EventLogStart) | $($b.ScmSec) |"
    & $L "| Вход пользователя (Winlogon 7001) | $(Format-Time $b.Logon) | $($b.PreLogonSec) |"
    & $L "| Проводник | $(Format-Time $b.Explorer) | $(if ($b.Explorer) { [math]::Round(($b.Explorer - $b.KernelStart).TotalSeconds, 1) }) |"
    & $L ''
    & $L "Тип старта: $($b.BootTypeText). $(if (-not $b.AutoLogon) { 'Автовход выключен — интервал «ядро → вход» включает время на экране входа/ввод пароля.' })"
    & $L "Неожиданных завершений за 30 дней: $($b.Unexpected30d) (6008), Kernel-Power 41: $($b.KernelPower41_30d)."
    if (@($b.DiagPerf).Count) {
        & $L ''
        & $L '### Diagnostics-Performance (официальные замеры Windows)'
        & $L '| Время | ID | Объект | Всего, мс | Деградация, мс |'
        & $L '|---|---|---|---|---|'
        foreach ($d in $b.DiagPerf) { & $L "| $($d.TimeCreated) | $($d.Id) | $(& $esc $(if ($d.Friendly) { $d.Friendly } else { $d.FileName })) | $($d.BootTimeMs) | $($d.DegradMs) |" }
    } else {
        & $L ''
        & $L 'Журнал Diagnostics-Performance пуст или недоступен (нужны права администратора; журнал заполняется не после каждой загрузки).'
    }
    & $L ''
    & $L '## Влияние элементов автозапуска'
    & $L 'Классы по критериям Microsoft: высокое > 1 с CPU или > 3 МБ диска; среднее 0,3–1 с или 0,3–3 МБ. Значения накопительные с момента запуска процесса — для точной оценки используйте режим ScheduleMeasure.'
    & $L ''
    & $L '| Влияние | Тип | Имя | Вкл. | Старт, с* | CPU, мс | Диск, МБ | Private, МБ | Мин. работы | Примечание |'
    & $L '|---|---|---|---|---|---|---|---|---|---|'
    foreach ($r in ($script:Impact | Where-Object { $_.Running } | Select-Object -First 40)) {
        & $L "| $($r.Impact) | $($r.Kind) | $(& $esc $r.Name) | $(if ($r.Enabled) { 'да' } else { 'нет' }) | $($r.StartOffsetSec) | $($r.CPU_ms) | $($r.DiskIO_MB) | $($r.Priv_MB) | $($r.MinutesRunning) | $(& $esc $r.Note) |"
    }
    & $L ''
    & $L '\* для Run/папки/задач — от входа пользователя, для служб — от старта ядра.'
    $sus = @($Gaps | Where-Object Suspicious)
    if ($sus.Count) {
        & $L ''
        & $L "### Паузы между элементами автозапуска ≥ 25 с"
        foreach ($g in $sus) { & $L "- $(& $esc $g.From) → $(& $esc $g.To): $($g.GapSec) с (на +$($g.AtSec) с)" }
        & $L ''
        & $L 'Это наблюдение по времени старта процессов. Причина (последовательный запуск с ожиданием, задержки самих программ) этим методом не доказывается — для проверки используйте BootTraceStart/BootTraceStop и Windows Performance Analyzer.'
    }
    & $L ''
    & $L '## Драйверы: сторонние ранние и незагрузившиеся'
    & $L '| Драйвер | Старт | Состояние | Производитель | Примечание |'
    & $L '|---|---|---|---|---|'
    foreach ($d in $Drv) { & $L "| $($d.Name) | $($d.Start) | $(if ($d.NotLoaded) { '**не загружен**' } else { $d.State }) | $(& $esc $d.Company) | $(& $esc $(if ($d.Protected -and -not $d.Note) { 'не изменять' } else { $d.Note })) |" }
    if (@($Scm).Count) {
        & $L ''
        & $L '## Ошибки диспетчера служб с момента загрузки'
        foreach ($e in $Scm) { & $L "- $($e.Time) — **$($e.EventID)**: $(& $esc $e.Message)$(if ($e.Note) { " — _$($e.Note)_" })" }
    }
    & $L ''
    & $L '## Память'
    & $L "- Физическая: $($Mem.UsedPhysMB) из $($Mem.TotalPhysMB) МБ ($($Mem.LoadPct)%), системный кэш $($Mem.SystemCacheMB) МБ"
    & $L "- Commit: $($Mem.CommitUsedMB) / $($Mem.CommitTotalMB) МБ, пик с загрузки $($Mem.CommitPeakMB) МБ"
    & $L "- Файл подкачки: $(if ($Mem.AutoPagefile) { 'системный' } else { 'вручную' }) — $((@($Mem.PagingFilesReg)) -join ', '); использование: $((@($Mem.PagefileUsage) | ForEach-Object { "$($_.File) $($_.CurrentMB)/$($_.AllocatedMB) МБ (пик $($_.PeakMB))" }) -join '; ')"
    & $L "- Рекомендация: минимум $($Pfr.InitialMB) МБ, максимум $($Pfr.MaxMB) МБ. $($Pfr.Note) $(if ($Pfr.Skip) { 'Пропуск: ' + $Pfr.Skip })"
    & $L "- Сжатие памяти: $($Mem.MemoryCompression) (процесс Memory Compression $($Mem.CompressionWSMB) МБ); SysMain: $($Mem.SysMain). SysMain не отключать: по опыту пользователей и проверкам, без него перестаёт работать сжатие памяти."
    & $L ''
    & $L '| Процесс | Кол-во | WS, МБ | Private, МБ | CPU, с |'
    & $L '|---|---|---|---|---|'
    foreach ($p in (Get-TopProcesses | Select-Object -First 20)) { & $L "| $($p.Name) | $($p.Count) | $($p.WS_MB) | $($p.Priv_MB) | $($p.CPU_Sec) |" }
    if ($s.IsLaptop) {
        & $L ''
        & $L '## Батарея и сон'
        if ($Pwr.Battery) { & $L "- Заряд $($Pwr.Battery.ChargePct)%, ёмкость $($Pwr.Battery.FullWh) из $($Pwr.Battery.DesignWh) Вт·ч ($($Pwr.Battery.HealthPct)%), циклов: $($Pwr.Battery.Cycles)" }
        & $L "- Сеть во сне (батарея): $($Pwr.NetStandbyDC) (0 — выкл., 1 — вкл., 2 — решает Windows); таймеры пробуждения (батарея): $($Pwr.WakeTimersDC)"
        if (@($Pwr.SchemeChangers).Count) { & $L "- Схему питания с момента загрузки меняли: $((@($Pwr.SchemeChangers)) -join ', ') — эти утилиты могут перезаписать настройки." }
        if ($s.PowerCaps.AoAc) { & $L '- Для ThinkPad T14s Gen 1 (Intel) в BIOS есть Config → Power → Sleep State = Linux (S3); переключение отключает Modern Standby. Проверяйте после смены: powercfg /a.' }
    }
    & $L ''
    & $L "## План (уровень -Level $Level)"
    & $L '| Id | Риск | Выбран | Статус | Цель | Сейчас → Будет |'
    & $L '|---|---|---|---|---|---|'
    foreach ($p in $Plan) { & $L "| $($p.Id) | $($p.Risk) | $(if ($p.Selected) { 'да' } else { '' }) | $($p.Status) | $(& $esc $p.Target) | $(& $esc $p.Current) → $(& $esc $p.Desired) |" }
    return $sb.ToString()
}

# ----------------------------------------------------------------------------------------------
# 4. Движок: состояние → план → применение → журнал → откат
# ----------------------------------------------------------------------------------------------
function Get-ServiceCode {
    param([string]$Desired)
    switch ($Desired) { 'Manual' { @{ Start = 3; Delayed = 0; Sc = 'demand' } } 'DelayedAuto' { @{ Start = 2; Delayed = 1; Sc = 'delayed-auto' } }
                        'Disabled' { @{ Start = 4; Delayed = 0; Sc = 'disabled' } } 'Auto' { @{ Start = 2; Delayed = 0; Sc = 'auto' } } }
}
function Get-ScFromState {
    param([int]$Start, [int]$Delayed)
    switch ($Start) { 0 { 'boot' } 1 { 'system' } 2 { if ($Delayed) { 'delayed-auto' } else { 'auto' } } 3 { 'demand' } 4 { 'disabled' } }
}
function Get-SvcState {
    param([string]$Name)
    $k = Open-RegKey HKLM "SYSTEM\CurrentControlSet\Services\$Name"
    if (-not $k) { return $null }
    try { return [pscustomobject]@{ Start = [int]$k.GetValue('Start', -1); Delayed = [int]$k.GetValue('DelayedAutoStart', 0) } } finally { $k.Close() }
}
function Set-SvcStart {
    param([string]$Name, [string]$Sc)
    $r = Invoke-Native 'sc.exe' @('config', $Name, 'start=', $Sc)
    if ($r.ExitCode -ne 0) { throw "sc.exe config $Name start= $Sc → код $($r.ExitCode): $($r.Output)" }
}

function Resolve-Targets {
    param($Item)
    switch ($Item.Kind) {
        'Service' {
            $names = @()
            foreach ($t in $Item.Targets) { $names += @($script:Services | Where-Object { $_.Name -like $t } | ForEach-Object { $_.Name }) }
            return @($names | Select-Object -Unique)
        }
        'Startup' {
            return @($script:Startup | Where-Object { ($_.Kind -eq 'Run' -or $_.Kind -eq 'StartupFolder') -and $Item.Targets -contains $_.Name -and ($Item.Location -eq '*' -or $_.Loc -eq $Item.Location) } |
                     ForEach-Object { [pscustomobject]@{ Hive = $_.Hive; SubKey = $_.Approved; Name = $_.Name; Loc = $_.Loc } })
        }
        'Task' {
            $all = @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskPath -notlike '\Microsoft\*' })
            $res = foreach ($t in $Item.Targets) { $all | Where-Object { $_.TaskName -like $t } | ForEach-Object { [pscustomobject]@{ TaskPath = $_.TaskPath; TaskName = $_.TaskName } } }
            return @($res)
        }
        'Reg'      { return @([pscustomobject]@{ Hive = $Item.Hive; SubKey = $Item.SubKey; Name = $Item.Name }) }
        'Power'    { return @([pscustomobject]@{ Sub = $Item.Sub; Setting = $Item.Setting }) }
        'MMAgent'  { return @($Item.Feature) }
        'Pagefile' { return @($env:SystemDrive) }
    }
    return @()
}

function Get-TargetText {
    param($Item, $T)
    switch ($Item.Kind) {
        'Startup' { "$($T.Loc)\$($T.Name)" } 'Task' { "$($T.TaskPath)$($T.TaskName)" } 'Reg' { "$($T.Hive)\$($T.SubKey)\$($T.Name)" }
        'Power' { "powercfg DC $($T.Sub) $($T.Setting)" } 'MMAgent' { "MMAgent $T" } 'Pagefile' { "$T\pagefile.sys" } default { "$T" }
    }
}

# Возвращает @{ Exists; Data (снимок для отката); Text; IsDesired; DesiredText }
function Get-ItemState {
    param($Item, $T)
    switch ($Item.Kind) {
        'Service' {
            $st = Get-SvcState $T
            if (-not $st) { return [pscustomobject]@{ Exists = $false } }
            $want = Get-ServiceCode $Item.Desired
            $ok = ($st.Start -eq $want.Start) -and ($want.Start -ne 2 -or $st.Delayed -eq $want.Delayed)
            # «Вручную» не ухудшаем, если уже «Отключено»
            if ($Item.Desired -eq 'Manual' -and $st.Start -eq 4) { $ok = $true }
            if ($Item.Desired -eq 'DelayedAuto' -and $st.Start -ge 3) { $ok = $true }
            return [pscustomobject]@{ Exists = $true; Data = $st; Text = (Get-StartText $st.Start $st.Delayed); IsDesired = $ok; DesiredText = (Get-StartText $want.Start $want.Delayed) }
        }
        'Startup' {
            $a = Get-ApprovedState $T.Hive $T.SubKey $T.Name
            return [pscustomobject]@{ Exists = $true; Data = $a.Raw; Text = $(if ($a.Enabled) { 'включён' } else { 'отключён' }); IsDesired = (-not $a.Enabled); DesiredText = 'отключён' }
        }
        'Task' {
            $t = Get-ScheduledTask -TaskPath $T.TaskPath -TaskName $T.TaskName -ErrorAction SilentlyContinue
            if (-not $t) { return [pscustomobject]@{ Exists = $false } }
            return [pscustomobject]@{ Exists = $true; Data = [pscustomobject]@{ State = "$($t.State)" }; Text = "$($t.State)"; IsDesired = ("$($t.State)" -eq 'Disabled'); DesiredText = 'Disabled' }
        }
        'Reg' {
            $v = Get-RegValueInfo $T.Hive $T.SubKey $T.Name
            return [pscustomobject]@{ Exists = $true; Data = $v; Text = $(if ($v.Exists) { "$($v.Value)" } else { 'нет значения' })
                                      IsDesired = ($v.Exists -and "$($v.Value)" -eq "$($Item.Value)"); DesiredText = "$($Item.Value)" }
        }
        'Power' {
            $scheme = Get-ActiveSchemeGuid
            $v = Get-PowerIndex $T.Sub $T.Setting
            if (-not $v -or -not $scheme) { return [pscustomobject]@{ Exists = $false; Why = 'параметр скрыт или не поддерживается этой схемой питания' } }
            return [pscustomobject]@{ Exists = $true; Data = [pscustomobject]@{ Scheme = $scheme; Value = $v.DC }; Text = "DC=$($v.DC)"; IsDesired = ($v.DC -eq [int]$Item.Value); DesiredText = "DC=$($Item.Value)" }
        }
        'MMAgent' {
            $mm = $null; try { $mm = Get-MMAgent -ErrorAction Stop } catch { }
            if (-not $mm) { return [pscustomobject]@{ Exists = $false; Why = $(if (-not $script:Sys.IsAdmin) { 'Get-MMAgent требует прав администратора' } else { 'Get-MMAgent недоступен' }) } }
            $en = [bool]$mm.$T
            return [pscustomobject]@{ Exists = $true; Data = [pscustomobject]@{ Enabled = $en }; Text = $(if ($en) { 'вкл.' } else { 'выкл.' }); IsDesired = (-not $en); DesiredText = 'выкл.' }
        }
        'Pagefile' {
            $mem = Get-MemoryInfo
            $rec = Get-PagefileRecommendation $mem
            $line = "$T\pagefile.sys $($rec.InitialMB) $($rec.MaxMB)"
            $ok = (-not $mem.AutoPagefile) -and $rec.CurrentInitMB -ge $rec.InitialMB -and $rec.CurrentMaxMB -ge $rec.MaxMB
            return [pscustomobject]@{ Exists = $true; Data = [pscustomobject]@{ Auto = $mem.AutoPagefile; PagingFiles = @($mem.PagingFilesReg) }
                                      Text = $(if ($mem.AutoPagefile) { 'системный' } else { (@($mem.PagingFilesReg) -join ', ') }); IsDesired = $ok; DesiredText = $line
                                      Line = $line; Skip = $rec.Skip }
        }
    }
}

function Get-Plan {
    Initialize-Data
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($it in $script:Catalog) {
        $sel = Test-Selected $it
        $skip = $null
        if ($it.LaptopOnly -and -not $script:Sys.IsLaptop) { $skip = 'не ноутбук' }
        elseif ($it.Precheck) { $skip = & $it.Precheck }
        $targets = @(Resolve-Targets $it)
        if (-not $targets.Count) {
            $rows.Add([pscustomobject]@{ Id = $it.Id; Group = $it.Group; Kind = $it.Kind; Risk = $it.Risk; Selected = $sel; Status = 'Нет на ПК'; Target = (@($it.Targets) -join ', ')
                                         Current = ''; Desired = ''; Title = $it.Title; Reason = $it.Reason; Effect = $it.Effect; Item = $it; T = $null; Before = $null })
            continue
        }
        foreach ($t in $targets) {
            $st = Get-ItemState $it $t
            $tt = Get-TargetText $it $t
            $status = 'Изменить'; $note = $skip
            if ($it.Kind -eq 'Service' -and (Test-Protected "$t")) { $status = 'Пропуск'; $note = 'защищённая служба' }
            elseif ($it.Kind -eq 'Startup' -and $t.Hive -eq 'HKCU' -and $script:HkcuMismatch) { $status = 'Пропуск'; $note = 'HKCU другого пользователя (повышение от другой учётной записи)' }
            elseif (-not $st.Exists) { if ($skip) { $status = 'Пропуск' } elseif ($st.PSObject.Properties['Why'] -and $st.Why) { $status = 'Пропуск'; $note = $st.Why } else { $status = 'Нет на ПК' } }
            elseif ($st.IsDesired) { $status = 'Уже так' }
            elseif ($skip) { $status = 'Пропуск' }
            elseif ($it.Kind -eq 'Pagefile' -and $st.Skip) { $status = 'Пропуск'; $note = $st.Skip }
            $rows.Add([pscustomobject]@{ Id = $it.Id; Group = $it.Group; Kind = $it.Kind; Risk = $it.Risk; Selected = $sel
                Status = $(if ($status -eq 'Пропуск') { "Пропуск: $note" } else { $status }); Target = $tt
                Current = "$($st.Text)"; Desired = "$($st.DesiredText)"; Title = $it.Title; Reason = $it.Reason; Effect = $it.Effect
                Item = $it; T = $t; Before = $st.Data; Line = $(if ($st.PSObject.Properties['Line']) { $st.Line } else { $null }) })
        }
    }
    return $rows
}

function Show-Plan {
    param($Plan)
    $Plan | Select-Object @{ n = 'Id'; e = { $_.Id } }, @{ n = 'Риск'; e = { $_.Risk } }, @{ n = 'Выбр.'; e = { if ($_.Selected) { 'да' } else { '' } } },
        @{ n = 'Статус'; e = { $_.Status } }, @{ n = 'Цель'; e = { $_.Target } }, @{ n = 'Сейчас → Будет'; e = { if ($_.Status -eq 'Уже так') { "$($_.Current)" } elseif ($_.Current -or $_.Desired) { "$($_.Current) → $($_.Desired)" } else { '' } } } |
        Format-Table -AutoSize -Wrap | Out-String -Width 220 | Write-Host
}

function Set-ItemDesired {
    param($Row)
    $it = $Row.Item; $t = $Row.T
    switch ($it.Kind) {
        'Service' {
            Set-SvcStart $t (Get-ServiceCode $it.Desired).Sc
            if ($StopNow -and $it.Desired -ne 'DelayedAuto') {
                $svc = Get-Service -Name $t -ErrorAction SilentlyContinue
                if ($svc -and $svc.Status -eq 'Running') { try { Stop-Service -Name $t -Force -ErrorAction Stop } catch { Write-Warning "Не удалось остановить ${t}: $($_.Exception.Message)" } }
            }
        }
        'Startup' {
            $bytes = [byte[]](3, 0, 0, 0) + [BitConverter]::GetBytes([DateTime]::UtcNow.ToFileTimeUtc())
            Set-RegValue $t.Hive $t.SubKey $t.Name 'Binary' ((@($bytes) | ForEach-Object { '{0:X2}' -f $_ }) -join '')
        }
        'Task'     { Disable-ScheduledTask -TaskPath $t.TaskPath -TaskName $t.TaskName -ErrorAction Stop | Out-Null }
        'Reg'      { Set-RegValue $t.Hive $t.SubKey $t.Name $it.ValueKind $it.Value }
        'Power'    {
            $r = Invoke-Native 'powercfg.exe' @('/setdcvalueindex', 'SCHEME_CURRENT', $t.Sub, $t.Setting, "$($it.Value)")
            if ($r.ExitCode -ne 0) { throw "powercfg /setdcvalueindex → код $($r.ExitCode): $($r.Output)" }
            [void](Invoke-Native 'powercfg.exe' @('/setactive', 'SCHEME_CURRENT'))
        }
        'MMAgent'  {
            $p = @{ $t = $true }
            Disable-MMAgent @p -ErrorAction Stop
        }
        'Pagefile' {
            if (-not $Row.Line) { throw 'нет рассчитанной строки PagingFiles' }
            $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
            if ($cs.AutomaticManagedPagefile) { Set-CimInstance -InputObject $cs -Property @{ AutomaticManagedPagefile = $false } -ErrorAction Stop }
            Set-RegValue HKLM 'SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management' 'PagingFiles' 'MultiString' @($Row.Line)
            $script:NeedReboot = $true
        }
    }
    if ($it.Kind -eq 'Reg' -and $it.Name -eq 'HiberbootEnabled') { $script:NeedReboot = $true }
    if ($it.Kind -eq 'Service' -and $it.Id -like 'DRV_*') { $script:NeedReboot = $true }
}

function Restore-ItemState {
    param([string]$Kind, $T, $Before, [string]$Feature)
    switch ($Kind) {
        'Service' {
            if (-not (Get-SvcState "$T")) { Write-Warning "Служба $T больше не существует — пропуск"; return }
            Set-SvcStart "$T" (Get-ScFromState ([int]$Before.Start) ([int]$Before.Delayed))
        }
        'Startup'  { Restore-RegValue $T.Hive $T.SubKey $T.Name $Before }
        'Reg'      { Restore-RegValue $T.Hive $T.SubKey $T.Name $Before }
        'Task'     { if ("$($Before.State)" -ne 'Disabled') { Enable-ScheduledTask -TaskPath $T.TaskPath -TaskName $T.TaskName -ErrorAction Stop | Out-Null } }
        'Power'    {
            $r = Invoke-Native 'powercfg.exe' @('/setdcvalueindex', "$($Before.Scheme)", $T.Sub, $T.Setting, "$($Before.Value)")
            if ($r.ExitCode -ne 0) { throw "powercfg → код $($r.ExitCode): $($r.Output)" }
            [void](Invoke-Native 'powercfg.exe' @('/setactive', 'SCHEME_CURRENT'))
        }
        'MMAgent'  { if ($Before.Enabled) { $p = @{ "$T" = $true }; Enable-MMAgent @p -ErrorAction Stop } }
        'Pagefile' {
            Set-RegValue HKLM 'SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management' 'PagingFiles' 'MultiString' @($Before.PagingFiles)
            if ([bool]$Before.Auto) {
                $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
                Set-CimInstance -InputObject $cs -Property @{ AutomaticManagedPagefile = $true } -ErrorAction Stop
            }
            $script:NeedReboot = $true
        }
    }
}

# --- точка восстановления: честная проверка результата ---
function New-RestorePointSafe {
    Write-Step 'Точка восстановления'
    $drive = "$env:SystemDrive\"
    try { Enable-ComputerRestore -Drive $drive -ErrorAction Stop } catch { Write-Warning "Не удалось включить защиту системы на ${drive}: $($_.Exception.Message)" }
    $srKey = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore'
    $freqName = 'SystemRestorePointCreationFrequency'
    $old = Get-RegValueInfo HKLM $srKey $freqName
    $created = $false; $err = $null
    try {
        Set-RegValue HKLM $srKey $freqName 'DWord' 0      # снимаем лимит «одна точка в 24 ч» на время вызова
        $t0 = (Get-Date).AddMinutes(-1)
        Checkpoint-Computer -Description ("Win11Toolkit {0} {1}" -f $script:Version, (Get-Date -Format 'yyyy-MM-dd HH:mm')) -RestorePointType MODIFY_SETTINGS -ErrorAction Stop
        $rp = @(Get-ComputerRestorePoint -ErrorAction SilentlyContinue) | Sort-Object SequenceNumber | Select-Object -Last 1
        if ($rp) { $created = ([Management.ManagementDateTimeConverter]::ToDateTime($rp.CreationTime) -ge $t0) }
    } catch { $err = $_.Exception.Message }
    finally { try { Restore-RegValue HKLM $srKey $freqName $old } catch { } }
    if ($created) { Write-Info 'Точка восстановления создана и найдена в списке.' Green; return $true }
    Write-Warning "Точка восстановления НЕ создана$(if ($err) { ": $err" }). Откат всё равно возможен по журналу (-Mode Rollback)."
    if ($Yes) { return $true }
    $a = Read-Host 'Продолжить без точки восстановления? (да/нет)'
    return ($a -match '^(д|да|y|yes)$')
}

function Save-Journal { param($Journal, [string]$Path) Write-TextFile $Path ($Journal | ConvertTo-Json -Depth 8) }

function Test-HkcuMismatch {
    # если окно повышено под другой учётной записью, HKCU — не тот профиль, где работает Проводник
    try {
        $me = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        $ex = Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" -ErrorAction Stop | Select-Object -First 1
        if (-not $ex) { return $false }
        $sid = (Invoke-CimMethod -InputObject $ex -MethodName GetOwnerSid -ErrorAction Stop).Sid
        return ($sid -and $sid -ne $me)
    } catch { return $false }
}

function Invoke-Apply {
    $plan = Get-Plan
    $todo = @($plan | Where-Object { $_.Selected -and $_.Status -eq 'Изменить' })
    Write-Step "План применения (уровень $Level$(if ($Include.Count) { ', +' + ($Include -join ',') })$(if ($Exclude.Count) { ', −' + ($Exclude -join ',') }))"
    Show-Plan ($plan | Where-Object { $_.Selected })
    if (-not $todo.Count) { Write-Info 'Изменять нечего: выбранные пункты уже настроены или отсутствуют на этом ПК.' Green; return }
    $whatIf = $script:DryRun
    if (-not $whatIf -and -not $script:Sys.IsAdmin) { Invoke-SelfElevate @{ Mode = 'Apply'; Level = $Level }; return }
    if (-not $whatIf -and -not $Yes) {
        $a = Read-Host "Применить $($todo.Count) изменений? (да/нет)"
        if ($a -notmatch '^(д|да|y|yes)$') { Write-Info 'Отменено.'; return }
    }
    if (-not $whatIf -and -not $NoRestorePoint) { if (-not (New-RestorePointSafe)) { Write-Info 'Отменено.'; return } }
    $jPath = Join-Path $script:OutDir ("Journal_{0}_{1}.json" -f $script:PC, $script:Stamp)
    $journal = [ordered]@{ Tool = "Win11-BootRamBattery-Toolkit $script:Version"; Computer = $script:PC; User = [Environment]::UserName
                           Created = Format-Time (Get-Date); Level = $Level; Items = (New-Object System.Collections.ArrayList) }
    $script:NeedReboot = $false
    foreach ($row in $todo) {
        $desc = "$($row.Id): $($row.Current) → $($row.Desired)"
        if ($whatIf) { Write-Info "  Что если: $desc" Yellow; continue }
        if (-not $script:Cmdlet.ShouldProcess($row.Target, $desc)) { continue }
        $entry = [ordered]@{ Id = $row.Id; Kind = $row.Item.Kind; Target = $row.T; TargetText = $row.Target; Before = $row.Before
                             After = $null; Ok = $false; Error = $null; Time = Format-Time (Get-Date) }
        try {
            Set-ItemDesired $row
            $after = Get-ItemState $row.Item $row.T
            $entry.After = $after.Data; $entry.Ok = [bool]$after.IsDesired
            if (-not $after.IsDesired -and $row.Item.Kind -ne 'Pagefile') { $entry.Error = 'после применения состояние не совпало с целевым' }
            if ($row.Item.Kind -eq 'Pagefile') { $entry.Ok = $true }
        } catch { $entry.Error = $_.Exception.Message }
        [void]$journal.Items.Add([pscustomobject]$entry)
        Save-Journal $journal $jPath                    # журнал пишется после каждого шага
        if ($entry.Ok) { Write-Info "  OK   $desc" Green } else { Write-Info "  ОШИБКА $desc — $($entry.Error)" Red }
    }
    if ($whatIf) { Write-Info 'Пробный прогон (-WhatIf): ничего не изменено.' Yellow; return }
    Write-Info "Журнал отката: $jPath" Green
    Write-Info "Откат: powershell -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Mode Rollback -RollbackFile `"$jPath`"" Gray
    Write-Info 'Службы с новым типом запуска продолжают работать до перезагрузки (без -StopNow).' Gray
    if ($script:NeedReboot) { Write-Info 'Для части изменений (файл подкачки, драйверы, Fast Startup) нужна перезагрузка.' Yellow }
    Write-Info 'После перезагрузки запустите -Mode Collect (или ScheduleMeasure) и сравните отчёты.' Gray
}

function Invoke-Rollback {
    $file = $RollbackFile
    if (-not $file) {
        $file = Get-ChildItem -LiteralPath $script:OutDir -Filter 'Journal_*.json' -File -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
    }
    if (-not $file -or -not (Test-Path -LiteralPath $file)) { Write-Warning "Журнал не найден в $script:OutDir. Укажите -RollbackFile."; return }
    if (-not $script:DryRun -and -not (Test-IsAdmin)) { Invoke-SelfElevate @{ Mode = 'Rollback'; RollbackFile = $file }; return }
    $j = Get-Content -LiteralPath $file -Raw -Encoding UTF8 | ConvertFrom-Json
    $items = @($j.Items); [array]::Reverse($items)
    Write-Step "Откат по журналу $(Split-Path $file -Leaf) от $($j.Created): $($items.Count) пунктов"
    $log = New-Object System.Collections.ArrayList
    foreach ($e in $items) {
        if ($null -eq $e.Before) { continue }
        $desc = "$($e.Id): вернуть $($e.TargetText)"
        if ($script:DryRun) { Write-Info "  Что если: $desc" Yellow; continue }
        if (-not $script:Cmdlet.ShouldProcess($e.TargetText, $desc)) { continue }
        $res = [ordered]@{ Id = $e.Id; Target = $e.TargetText; Ok = $false; Error = $null }
        try { Restore-ItemState $e.Kind $e.Target $e.Before; $res.Ok = $true } catch { $res.Error = $_.Exception.Message }
        [void]$log.Add([pscustomobject]$res)
        if ($res.Ok) { Write-Info "  OK   $desc" Green } else { Write-Info "  ОШИБКА $desc — $($res.Error)" Red }
    }
    if (-not $script:DryRun) {
        Write-TextFile ([IO.Path]::ChangeExtension($file, ".rollback-$script:Stamp.json")) ($log | ConvertTo-Json -Depth 4)
        Write-Info 'Откат выполнен. Перезагрузите компьютер, если менялись файл подкачки, драйверы или Fast Startup.' Green
    }
}

# --- однократный замер после следующего входа ---
function Register-PostLogonMeasure {
    if ($script:DryRun) { Write-Info "Что если: задача \Win11Toolkit\$script:MeasureTaskName — однократный сбор через $MeasureDelaySec с после входа, копия скрипта в $env:ProgramData\Win11Toolkit" Yellow; return }
    if (-not (Test-IsAdmin)) { Invoke-SelfElevate @{ Mode = 'ScheduleMeasure' }; return }
    if (-not $PSCommandPath) { throw 'Запустите скрипт из файла .ps1.' }
    $dir = Join-Path $env:ProgramData 'Win11Toolkit'
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    # задача выполняется с наивысшими правами — копию защищаем от изменения обычными пользователями
    [void](Invoke-Native 'icacls.exe' @($dir, '/inheritance:r', '/grant:r', '*S-1-5-32-544:(OI)(CI)F', '*S-1-5-18:(OI)(CI)F', '*S-1-5-32-545:(OI)(CI)RX'))
    $copy = Join-Path $dir 'Win11-BootRamBattery-Toolkit.ps1'
    Copy-Item -LiteralPath $PSCommandPath -Destination $copy -Force
    $user = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $arg = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -Mode Collect -FromScheduledTask -MeasureDelaySec {1} -OutputDir "{2}"' -f $copy, $MeasureDelaySec, $script:OutDir
    $action = New-ScheduledTaskAction -Execute (Join-Path $PSHOME 'powershell.exe') -Argument $arg
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User $user
    $trigger.Delay = 'PT{0}S' -f $MeasureDelaySec
    $principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 15)
    Register-ScheduledTask -TaskName $script:MeasureTaskName -TaskPath $script:MeasureTaskPath -Action $action -Trigger $trigger `
        -Principal $principal -Settings $settings -Force -ErrorAction Stop | Out-Null
    Write-Info "Замер запланирован: через $MeasureDelaySec с после следующего входа $user. Отчёт появится в $script:OutDir, задача удалится сама." Green
}
function Unregister-PostLogonMeasure {
    if ($script:DryRun) { Write-Info 'Что если: удаление задачи замера' Yellow; return }
    if (-not (Test-IsAdmin)) { Invoke-SelfElevate @{ Mode = 'UnscheduleMeasure' }; return }
    try { Unregister-ScheduledTask -TaskName $script:MeasureTaskName -TaskPath $script:MeasureTaskPath -Confirm:$false -ErrorAction Stop; Write-Info 'Задача замера удалена.' Green }
    catch { Write-Info 'Задача замера не найдена.' }
}

# --- трассировка загрузки штатным WPR (Windows Performance Recorder) ---
function Invoke-BootTrace {
    param([ValidateSet('Start', 'Stop', 'Cancel')][string]$Action)
    if ($script:DryRun) { Write-Info "Что если: wpr.exe -boottrace ($Action)" Yellow; return }
    if (-not (Test-IsAdmin)) { Invoke-SelfElevate @{ Mode = "BootTrace$Action" }; return }
    $wpr = Join-Path $env:windir 'System32\wpr.exe'
    if (-not (Test-Path -LiteralPath $wpr)) { Write-Warning 'wpr.exe не найден (нужен Windows Performance Toolkit).'; return }
    switch ($Action) {
        'Start' {
            $tmp = Join-Path $script:OutDir 'wpr-temp'; New-Item -ItemType Directory -Path $tmp -Force | Out-Null
            $r = Invoke-Native $wpr @('-boottrace', '-addboot', 'GeneralProfile', '-filemode', '-recordtempto', $tmp)
            if ($r.ExitCode -eq 0) { Write-Info 'Трассировка следующей загрузки включена. Перезагрузитесь, войдите, подождите 2–3 минуты и выполните -Mode BootTraceStop.' Green }
            else { Write-Warning "wpr: код $($r.ExitCode). $($r.Output)" }
        }
        'Stop' {
            $etl = Join-Path $script:OutDir ("BootTrace_{0}_{1}.etl" -f $script:PC, $script:Stamp)
            $r = Invoke-Native $wpr @('-boottrace', '-stopboot', $etl, 'Win11Toolkit slow boot')
            if ($r.ExitCode -eq 0) { Write-Info "Трассировка сохранена: $etl (открыть в Windows Performance Analyzer из Windows ADK: Boot Phases, CPU Usage, Disk Usage)." Green }
            else { Write-Warning "wpr: код $($r.ExitCode). $($r.Output)" }
        }
        'Cancel' { $r = Invoke-Native $wpr @('-boottrace', '-cancelboot'); Write-Info "wpr -cancelboot: код $($r.ExitCode)" }
    }
}

function Show-Catalog {
    $script:Catalog | Select-Object Id, Group, Risk, Kind, Title | Format-Table -AutoSize -Wrap | Out-String -Width 200 | Write-Host
    Write-Info ('Не изменяются никогда: ' + ($script:Protected -join ', ')) DarkGray
}

# ----------------------------------------------------------------------------------------------
# 5. Меню и точка входа
# ----------------------------------------------------------------------------------------------
function Show-Menu {
    :menu while ($true) {
        Write-Host ''
        Write-Host "Win11 Boot/RAM/Battery Toolkit $script:Version — $script:PC  [$(if ($script:Sys.IsAdmin) { 'администратор' } else { 'обычные права' })]" -ForegroundColor Cyan
        Write-Host "Папка отчётов: $script:OutDir" -ForegroundColor DarkGray
        Write-Host '  1  Сбор диагностики (только чтение)'
        Write-Host '  2  План изменений (только чтение)'
        Write-Host '  3  Пробный прогон оптимизации (-WhatIf)'
        Write-Host '  4  Применить пункты низкого риска'
        Write-Host '  5  Применить пункты низкого и среднего риска'
        Write-Host '  6  Откат по последнему журналу'
        Write-Host "  7  Запланировать замер через $MeasureDelaySec с после следующего входа"
        Write-Host '  8  Трассировка загрузки WPR: включить'
        Write-Host '  9  Трассировка загрузки WPR: сохранить'
        Write-Host '  L  Каталог пунктов'
        Write-Host '  0  Выход'
        $c = (Read-Host 'Выбор').Trim().ToUpperInvariant()
        switch ($c) {
            '1' { Invoke-Collect }
            '2' { Show-Plan (Get-Plan) }
            '3' { $old = $script:DryRun; $script:DryRun = $true; try { Invoke-Apply } finally { $script:DryRun = $old } }
            '4' { $Level = 'Low'; Invoke-Apply }
            '5' { $Level = 'Medium'; Invoke-Apply }
            '6' { Invoke-Rollback }
            '7' { Register-PostLogonMeasure }
            '8' { Invoke-BootTrace Start }
            '9' { Invoke-BootTrace Stop }
            'L' { Show-Catalog }
            '0' { break menu }
            default { Write-Info 'Нет такого пункта.' Yellow }
        }
    }
}

$script:OutDir = Resolve-OutDir
$script:Sys = Get-SystemInfo
$script:Catalog = Get-Catalog
$script:HkcuMismatch = $false
if ($script:Sys.IsAdmin -and ($Mode -in @('Apply', 'Plan', 'Menu'))) { $script:HkcuMismatch = Test-HkcuMismatch }

$transcript = $false
if (($Mode -in @('Apply', 'Rollback')) -and -not $script:DryRun -and $script:Sys.IsAdmin) {
    try { Start-Transcript -LiteralPath (Join-Path $script:OutDir ("Log_{0}_{1}_{2}.txt" -f $script:PC, $Mode, $script:Stamp)) -ErrorAction Stop | Out-Null; $transcript = $true } catch { }
}
try {
    switch ($Mode) {
        'Menu'              { Show-Menu }
        'Collect'           { Invoke-Collect }
        'Plan'              { $p = Get-Plan; Show-Plan $p; Export-RuCsv ($p | Select-Object Id, Group, Risk, Selected, Status, Target, Current, Desired, Title, Reason, Effect) (Join-Path $script:OutDir ("Plan_{0}_{1}.csv" -f $script:PC, $script:Stamp)) }
        'Apply'             { Invoke-Apply }
        'Rollback'          { Invoke-Rollback }
        'ScheduleMeasure'   { Register-PostLogonMeasure }
        'UnscheduleMeasure' { Unregister-PostLogonMeasure }
        'BootTraceStart'    { Invoke-BootTrace Start }
        'BootTraceStop'     { Invoke-BootTrace Stop }
        'BootTraceCancel'   { Invoke-BootTrace Cancel }
        'ListCatalog'       { Show-Catalog }
    }
} finally {
    if ($transcript) { try { Stop-Transcript | Out-Null } catch { } }
    if ($Elevated -and -not $FromScheduledTask) { [void](Read-Host 'Готово. Нажмите Enter, чтобы закрыть окно') }
}
