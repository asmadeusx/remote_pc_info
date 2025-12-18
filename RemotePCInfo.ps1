# RemotePCInfo.ps1 - Оптимизированный скрипт для сбора информации о удаленных ПК

# Определяем путь к папке скрипта
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$computersFilePath = Join-Path $scriptPath "computers.txt"

# Поиск PsExec (предпочтение отдается 64-битной версии)
$psexecPath = $null
$psexec64Path = Join-Path $scriptPath "PsExec64.exe"
$psexec32Path = Join-Path $scriptPath "PsExec.exe"

if (Test-Path $psexec64Path) {
    $psexecPath = $psexec64Path
}
elseif (Test-Path $psexec32Path) {
    $psexecPath = $psexec32Path
}
else {
    Write-Host "ОШИБКА: PsExec не найден!" -ForegroundColor Red
    Write-Host "Поместите один из файлов в папку со скриптом ($scriptPath):" -ForegroundColor Yellow
    Write-Host "  - PsExec64.exe (предпочтительно)" -ForegroundColor Yellow
    Write-Host "  - PsExec.exe" -ForegroundColor Yellow
    exit 1
}

# Функция проверки доступности компьютера по ping
function Test-ComputerOnline {
    param([string]$ComputerIP)
    
    try {
        $pingResult = Test-Connection -ComputerName $ComputerIP -Count 1 -Quiet -ErrorAction Stop
        return $pingResult
    }
    catch {
        return $false
    }
}

# Основная функция сбора информации
function Get-RemotePCInfo {
    param([string]$RemotePCIP)

    Write-Host ""
    Write-Host "=== СБОР ИНФОРМАЦИИ О УДАЛЕННОМ ПК ($RemotePCIP) ===" -ForegroundColor Green
    Write-Host ""

    # 1. Общая информация о системе
    Write-Host "1. ОБЩАЯ ИНФОРМАЦИЯ О СИСТЕМЕ:" -ForegroundColor Yellow
    $systemInfo = & $psexecPath \\$RemotePCIP -accepteula -nobanner powershell @'
$system = Get-WmiObject -Class Win32_ComputerSystem
Write-Host "Имя компьютера: $($system.Name)"
Write-Host "   Домен: $($system.Domain)"
'@ 2>$null
    $systemInfo
    Write-Host ""

    # 2. Информация о возрасте ОС
    Write-Host "2. ИНФОРМАЦИЯ О ВОЗРАСТЕ ОС:" -ForegroundColor Yellow
    $osInfo = & $psexecPath \\$RemotePCIP -accepteula -nobanner powershell @'
$os = Get-WmiObject -Class Win32_OperatingSystem
$installDate = [Management.ManagementDateTimeConverter]::ToDateTime($os.InstallDate)
$lastBootTime = [Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
$localDateTime = [Management.ManagementDateTimeConverter]::ToDateTime($os.LocalDateTime)

$currentDate = $localDateTime
$daysSinceInstall = [math]::Round(($currentDate - $installDate).TotalDays, 0)

Write-Host "Операционная система: $($os.Caption)"
Write-Host "   Версия: $($os.Version)"
Write-Host "   Дата установки: $($installDate.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "   Последняя загрузка: $($lastBootTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "   Текущее время: $($localDateTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "   Дней с установки: $daysSinceInstall"
'@ 2>$null
    $osInfo
    Write-Host ""

    # 3. Информация о материнской плате
    Write-Host "3. ИНФОРМАЦИЯ О МАТЕРИНСКОЙ ПЛАТЕ:" -ForegroundColor Yellow
    $motherboard = & $psexecPath \\$RemotePCIP -accepteula -nobanner powershell @'
$board = Get-WmiObject -Class Win32_BaseBoard
Write-Host "Производитель: $($board.Manufacturer)"
Write-Host "   Модель: $($board.Product)"
Write-Host "   Серийный номер: $($board.SerialNumber)"
'@ 2>$null
    $motherboard
    Write-Host ""

    # 4. Информация о процессоре
    Write-Host "4. ИНФОРМАЦИЯ О ПРОЦЕССОРЕ:" -ForegroundColor Yellow
    $processor = & $psexecPath \\$RemotePCIP -accepteula -nobanner powershell @'
$cpu = Get-WmiObject -Class Win32_Processor
Write-Host "Модель: $($cpu.Name)"
Write-Host "   Количество ядер: $($cpu.NumberOfCores)"
Write-Host "   Максимальная частота: $($cpu.MaxClockSpeed) MHz"
'@ 2>$null
    $processor
    Write-Host ""

    # 5. Информация о оперативной памяти
    Write-Host "5. ИНФОРМАЦИЯ О ОПЕРАТИВНОЙ ПАМЯТИ:" -ForegroundColor Yellow
    $memory = & $psexecPath \\$RemotePCIP -accepteula -nobanner powershell @'
$memoryModules = Get-WmiObject -Class Win32_PhysicalMemory
$totalGB = 0

foreach ($module in $memoryModules) {
    $sizeGB = [math]::Round($module.Capacity / 1GB, 2)
    $totalGB += $sizeGB
    Write-Host "Модуль: $($module.DeviceLocator)"
    Write-Host "   Размер: $sizeGB GB"
    Write-Host "   Производитель: $($module.Manufacturer)"
    Write-Host "   Серийный номер: $($module.SerialNumber)"
    Write-Host "   Партномер: $($module.PartNumber)"
    Write-Host ""
}

Write-Host "Общий объем: $totalGB GB"
'@ 2>$null
    $memory
    Write-Host ""

    # 6. Информация о дисках
    Write-Host "6. ИНФОРМАЦИЯ О ДИСКАХ:" -ForegroundColor Yellow
    $disks = & $psexecPath \\$RemotePCIP -accepteula -nobanner powershell @'
$disksInfo = Get-WmiObject -Class Win32_DiskDrive

foreach ($disk in $disksInfo) {
    $sizeGB = [math]::Round($disk.Size / 1GB, 2)
    Write-Host "Модель: $($disk.Model)"
    Write-Host "   Размер: $sizeGB GB"
    Write-Host "   Интерфейс: $($disk.InterfaceType)"
    Write-Host "   Тип носителя: $($disk.MediaType)"
    Write-Host "   Серийный номер: $($disk.SerialNumber)"
    Write-Host ""
}
'@ 2>$null
    $disks
    Write-Host ""

    # 7. Информация о видеоадаптере
    Write-Host "7. ИНФОРМАЦИЯ О ВИДЕОАДАПТЕРЕ:" -ForegroundColor Yellow
    $gpu = & $psexecPath \\$RemotePCIP -accepteula -nobanner powershell @'
$videoCards = Get-WmiObject -Class Win32_VideoController

foreach ($card in $videoCards) {
    $vramGB = [math]::Round($card.AdapterRAM / 1GB, 2)
    Write-Host "Видеокарта: $($card.Name)"
    Write-Host "   Видеопамять: $vramGB GB"
    Write-Host "   Версия драйвера: $($card.DriverVersion)"
    Write-Host "   Видеопроцессор: $($card.VideoProcessor)"
    Write-Host ""
}
'@ 2>$null
    $gpu
    Write-Host ""

    Write-Host "=== СБОР ИНФОРМАЦИИ ЗАВЕРШЕН ===" -ForegroundColor Green
}

# --- ОСНОВНОЙ КОД СКРИПТА ---

Write-Host "=== RemotePCInfo Скрипт ===" -ForegroundColor Cyan
Write-Host ""

# Вывод информации о используемой версии PsExec ПОСЛЕ заголовка
if (Test-Path $psexec64Path) {
    Write-Host "Используется: PsExec64.exe" -ForegroundColor Green
}
elseif (Test-Path $psexec32Path) {
    Write-Host "Используется: PsExec.exe" -ForegroundColor Green
}
Write-Host "Путь к скрипту: $scriptPath" -ForegroundColor Gray
Write-Host ""

# Проверка наличия файла computers.txt
if (-not (Test-Path $computersFilePath)) {
    Write-Host "Файл computers.txt не найден в папке со скриптом." -ForegroundColor Yellow
    Write-Host "Создаю файл computers.txt с примером..." -ForegroundColor Yellow
    
    @"
# Файл со списком компьютеров для сканирования
# Каждый IP-адрес на новой строке
# Строки, начинающиеся с #, игнорируются

# Примеры:
192.168.1.100
192.168.1.101
# 192.168.1.102 - этот компьютер закомментирован
"@ | Out-File -FilePath $computersFilePath -Encoding UTF8
    
    Write-Host "Файл создан: $computersFilePath" -ForegroundColor Green
    Write-Host "Добавьте IP-адреса в файл и запустите скрипт снова." -ForegroundColor Yellow
    exit 0
}

# Чтение файла computers.txt
Write-Host "Чтение списка компьютеров из: $computersFilePath" -ForegroundColor Cyan
$computers = Get-Content $computersFilePath | Where-Object {
    $_ -notmatch '^\s*#' -and $_ -match '\S' -and $_ -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$'
}

if ($computers.Count -eq 0) {
    Write-Host "ОШИБКА: В файле computers.txt не найдено корректных IP-адресов!" -ForegroundColor Red
    Write-Host "Формат: каждый IP-адрес на новой строке (например: 192.168.1.100)" -ForegroundColor Yellow
    exit 1
}

Write-Host "Найдено компьютеров для проверки: $($computers.Count)" -ForegroundColor Green
Write-Host ""

# Обработка каждого компьютера
$successCount = 0
$skippedCount = 0

foreach ($computer in $computers) {
    $computer = $computer.Trim()
    
    Write-Host "Проверка компьютера: $computer" -ForegroundColor Cyan
    
    # Проверка доступности по ping
    if (Test-ComputerOnline -ComputerIP $computer) {
        Write-Host "Компьютер доступен, начинаю сбор информации..." -ForegroundColor Green
        Write-Host ""
        
        try {
            Get-RemotePCInfo -RemotePCIP $computer
            $successCount++
        }
        catch {
            Write-Host "ОШИБКА при сборе информации с $computer : $_" -ForegroundColor Red
            Write-Host ""
        }
    }
    else {
        Write-Host "ПРЕДУПРЕЖДЕНИЕ: Компьютер $computer недоступен (ping), пропускаю..." -ForegroundColor Yellow
        Write-Host ""
        $skippedCount++
    }
    
    # Пауза между компьютерами (опционально)
    if ($computer -ne $computers[-1]) {
        Start-Sleep -Seconds 1
    }
}

# Итоговая статистика
Write-Host "=== ИТОГОВАЯ СТАТИСТИКА ===" -ForegroundColor Cyan
Write-Host "Успешно обработано: $successCount компьютеров" -ForegroundColor Green
Write-Host "Пропущено (недоступно): $skippedCount компьютеров" -ForegroundColor Yellow
Write-Host "Всего в списке: $($computers.Count) компьютеров" -ForegroundColor White