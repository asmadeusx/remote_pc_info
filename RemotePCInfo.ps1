# RemotePCInfo.ps1 - Оптимизированный скрипт для сбора информации о удаленных ПК

# Определяем путь к папке скрипта
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$computersFilePath = Join-Path $scriptPath "computers.txt"
$outputFilePath = Join-Path $scriptPath "CollectedPCInfo.txt"

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

# Функция записи в файл и консоль одновременно
function Write-OutputBoth {
    param([string]$Message, [string]$FilePath, [switch]$NoNewLine)
    
    # Запись в консоль
    if ($NoNewLine) {
        Write-Host $Message -NoNewline
    }
    else {
        Write-Host $Message
    }
    
    # Запись в файл
    if ($NoNewLine) {
        $Message | Out-File -FilePath $FilePath -Encoding UTF8 -Append -NoNewline
    }
    else {
        $Message | Out-File -FilePath $FilePath -Encoding UTF8 -Append
    }
}

# Функция для захвата вывода PsExec и записи в файл
function Invoke-RemoteCommand {
    param([string]$RemotePCIP, [string]$Command)
    
    $output = & $psexecPath \\$RemotePCIP -accepteula -nobanner powershell $Command 2>&1
    
    # Фильтруем сообщения об ошибках PsExec
    $cleanOutput = $output | Where-Object { $_ -is [string] } | ForEach-Object {
        if ($_ -notmatch '^PsExec' -and $_ -notmatch '^Copyright' -and $_ -notmatch '^Sysinternals') {
            $_
        }
    }
    
    return $cleanOutput -join "`n"
}

# Основная функция сбора информации
function Get-RemotePCInfo {
    param([string]$RemotePCIP, [string]$OutputFile)

    Write-OutputBoth "`n=== СБОР ИНФОРМАЦИИ О УДАЛЕННОМ ПК ($RemotePCIP) ===" -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 1. Общая информация о системе
    Write-OutputBoth "1. ОБЩАЯ ИНФОРМАЦИЯ О СИСТЕМЕ:" -FilePath $OutputFile
    $systemInfo = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
$system = Get-WmiObject -Class Win32_ComputerSystem
Write-Host "Имя компьютера: $($system.Name)"
Write-Host "   Домен: $($system.Domain)"
'@
    Write-OutputBoth $systemInfo -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 2. Информация о возрасте ОС
    Write-OutputBoth "2. ИНФОРМАЦИЯ О ВОЗРАСТЕ ОС:" -FilePath $OutputFile
    $osInfo = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
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
'@
    Write-OutputBoth $osInfo -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 3. Информация о материнской плате
    Write-OutputBoth "3. ИНФОРМАЦИЯ О МАТЕРИНСКОЙ ПЛАТЕ:" -FilePath $OutputFile
    $motherboard = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
$board = Get-WmiObject -Class Win32_BaseBoard
Write-Host "Производитель: $($board.Manufacturer)"
Write-Host "   Модель: $($board.Product)"
Write-Host "   Серийный номер: $($board.SerialNumber)"
'@
    Write-OutputBoth $motherboard -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 4. Информация о процессоре
    Write-OutputBoth "4. ИНФОРМАЦИЯ О ПРОЦЕССОРЕ:" -FilePath $OutputFile
    $processor = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
$cpu = Get-WmiObject -Class Win32_Processor
Write-Host "Модель: $($cpu.Name)"
Write-Host "   Количество ядер: $($cpu.NumberOfCores)"
Write-Host "   Максимальная частота: $($cpu.MaxClockSpeed) MHz"
'@
    Write-OutputBoth $processor -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 5. Информация о оперативной памяти
    Write-OutputBoth "5. ИНФОРМАЦИЯ О ОПЕРАТИВНОЙ ПАМЯТИ:" -FilePath $OutputFile
    $memory = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
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
'@
    Write-OutputBoth $memory -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 6. Информация о дисках
    Write-OutputBoth "6. ИНФОРМАЦИЯ О ДИСКАХ:" -FilePath $OutputFile
    $disks = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
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
'@
    Write-OutputBoth $disks -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 7. Информация о видеоадаптере
    Write-OutputBoth "7. ИНФОРМАЦИЯ О ВИДЕОАДАПТЕРЕ:" -FilePath $OutputFile
    $gpu = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
$videoCards = Get-WmiObject -Class Win32_VideoController

foreach ($card in $videoCards) {
    $vramGB = [math]::Round($card.AdapterRAM / 1GB, 2)
    Write-Host "Видеокарта: $($card.Name)"
    Write-Host "   Видеопамять: $vramGB GB"
    Write-Host "   Версия драйвера: $($card.DriverVersion)"
    Write-Host "   Видеопроцессор: $($card.VideoProcessor)"
    Write-Host ""
}
'@
    Write-OutputBoth $gpu -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 8. Информация о BIOS/UEFI
    Write-OutputBoth "8. ИНФОРМАЦИЯ О БИОС (UEFI):" -FilePath $OutputFile
    $biosInfo = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
$bios = Get-WmiObject -Class Win32_BIOS
$releaseDate = [Management.ManagementDateTimeConverter]::ToDateTime($bios.ReleaseDate)
$currentDate = Get-Date
$daysSinceRelease = [math]::Round(($currentDate - $releaseDate).TotalDays, 0)

Write-Host "Производитель BIOS: $($bios.Manufacturer)"
Write-Host "   Версия: $($bios.SMBIOSBIOSVersion)"
Write-Host "   Серийный номер: $($bios.SerialNumber)"
Write-Host "   Версия прошивки: $($bios.Version)"
Write-Host "   Дата выпуска: $($releaseDate.ToString('yyyy-MM-dd'))"
Write-Host "   Дней с выпуска: $daysSinceRelease"
'@
    Write-OutputBoth $biosInfo -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    Write-OutputBoth "=== СБОР ИНФОРМАЦИИ ЗАВЕРШЕН ===" -FilePath $OutputFile
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

# Очистка файла вывода перед началом работы
if (Test-Path $outputFilePath) {
    Write-Host "Очистка предыдущего файла вывода: $outputFilePath" -ForegroundColor Yellow
}
else {
    Write-Host "Создание файла вывода: $outputFilePath" -ForegroundColor Green
}

# Создаем/очищаем файл и записываем заголовок
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"=== RemotePCInfo - Отчет от $timestamp ===" | Out-File -FilePath $outputFilePath -Encoding UTF8
"=== Сгенерировано скриптом: $($MyInvocation.MyCommand.Name)" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append

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
Write-Host "Результаты будут сохранены в: $outputFilePath" -ForegroundColor Green
Write-Host ""

# Записываем информацию о сканировании в файл
"=== Начало сканирования ===" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"Время начала: $timestamp" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"Количество компьютеров для проверки: $($computers.Count)" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append

# Обработка каждого компьютера
$successCount = 0
$skippedCount = 0

foreach ($computer in $computers) {
    $computer = $computer.Trim()
    
    Write-Host "Проверка компьютера: $computer" -ForegroundColor Cyan
    "`n--- Проверка компьютера: $computer ---" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
    
    # Проверка доступности по ping
    if (Test-ComputerOnline -ComputerIP $computer) {
        Write-Host "Компьютер доступен, начинаю сбор информации..." -ForegroundColor Green
        "Статус: доступен, сбор информации..." | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
        Write-Host ""
        
        try {
            Get-RemotePCInfo -RemotePCIP $computer -OutputFile $outputFilePath
            $successCount++
            "Статус: успешно обработан" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
        }
        catch {
            $errorMsg = "ОШИБКА при сборе информации с $computer : $_"
            Write-Host $errorMsg -ForegroundColor Red
            $errorMsg | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
            Write-Host ""
        }
    }
    else {
        $warningMsg = "ПРЕДУПРЕЖДЕНИЕ: Компьютер $computer недоступен (ping), пропускаю..."
        Write-Host $warningMsg -ForegroundColor Yellow
        $warningMsg | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
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
Write-Host ""
Write-Host "Результаты сохранены в файл: $outputFilePath" -ForegroundColor Green

# Записываем итоговую статистику в файл
"`n=== ИТОГОВАЯ СТАТИСТИКА ===" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"Успешно обработано: $successCount компьютеров" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"Пропущено (недоступно): $skippedCount компьютеров" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"Всего в списке: $($computers.Count) компьютеров" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"`n=== Сканирование завершено ===" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
$endTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"Время окончания: $endTime" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append