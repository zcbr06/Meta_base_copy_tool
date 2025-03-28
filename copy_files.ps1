# Запуск с правами администратора (если не запущен)
$currProc = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-Not $currProc.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Функция выбора папки через Проводник
function Select-Folder {
    param ([string]$message)
    Add-Type -AssemblyName System.Windows.Forms
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $message
    $folderBrowser.ShowNewFolderButton = $true
    if ($folderBrowser.ShowDialog() -eq "OK") {
        return $folderBrowser.SelectedPath
    } else {
        return $null
    }
}

# Переменные для путей и автора
$sourceDrive = "Not selected"
$destinationDrive = "Not selected"
$authorName = "Not set"

# Функция для отображения меню
function Show-Menu {
    Clear-Host
    Write-Host "-----------------------------------------"
    Write-Host "         FILE COPY TOOL MENU             "
    Write-Host "-----------------------------------------"
    Write-Host "[1] Select source folder    | $sourceDrive"
    Write-Host "[2] Select destination folder | $destinationDrive"
    Write-Host "[3] Enter author name      | $authorName"
    Write-Host "[4] Start copy process"
    Write-Host "[5] Exit"
    Write-Host "-----------------------------------------"
    Write-Host "Press a number key (1-5) to select an option..."
}

# Флаг для выхода из меню
$exitMenu = $false

# Цикл меню с обработкой нажатий клавиш
do {
    Show-Menu
    $choice = [System.Console]::ReadKey($true).KeyChar

    switch ($choice) {
        "1" {
            $selectedFolder = Select-Folder -message "Select the source folder"
            if ($selectedFolder) { $sourceDrive = $selectedFolder }
        }
        "2" {
            $selectedFolder = Select-Folder -message "Select the destination folder"
            if ($selectedFolder) { $destinationDrive = $selectedFolder }
        }
        "3" {
            Write-Host "`nEnter author name: " -NoNewline
            $authorName = Read-Host
        }
        "4" {
            if ($sourceDrive -eq "Not selected" -or $destinationDrive -eq "Not selected" -or $authorName -eq "Not set") {
                Write-Host "`nERROR: Please complete all fields before starting the process!"
                Start-Sleep 2
            } else {
                $exitMenu = $true  # Завершаем меню
            }
        }
        "5" {
            Write-Host "Exiting..."
            exit
        }
        default {
            Write-Host "Invalid option! Try again."
            Start-Sleep 1
        }
    }
} while (-Not $exitMenu)

# Создаем папку назначения, если её нет
if (!(Test-Path $destinationDrive)) {
    New-Item -ItemType Directory -Path $destinationDrive | Out-Null
}

# Создаем объект Shell.Application для получения метаданных
$shell = New-Object -ComObject Shell.Application

# Получаем список всех файлов
$files = Get-ChildItem -Path $sourceDrive -Recurse -File
$totalFiles = $files.Count
$processedFiles = 0
$copiedFiles = @()

# Функция получения автора файла
function Get-FileAuthor($filePath) {
    $folder = $shell.Namespace((Get-Item $filePath).DirectoryName)
    $item = $folder.ParseName((Get-Item $filePath).Name)
    return $folder.GetDetailsOf($item, 20)  # 20 — индекс свойства "Автор"
}

Write-Host "`n-----------------------------------------"
Write-Host "Starting file copy process..."
Write-Host "-----------------------------------------"
Start-Sleep 1

# Ищем файлы с нужным автором и копируем их
foreach ($file in $files) {
    $fileAuthor = Get-FileAuthor $file.FullName
    if ($fileAuthor -eq $authorName) {
        $destinationPath = Join-Path -Path $destinationDrive -ChildPath $file.Name
        Copy-Item -Path $file.FullName -Destination $destinationPath -Force
        $copiedFiles += $file.FullName
        Write-Host "[COPIED] $($file.FullName)"
    }

    # Обновляем прогресс
    $processedFiles++
    $percentComplete = [math]::Round(($processedFiles / $totalFiles) * 100, 2)
    Write-Progress -Activity "Copying files" -Status "$percentComplete% completed" -PercentComplete $percentComplete
}

Write-Host "-----------------------------------------"
Write-Host "PROCESS COMPLETED"
Write-Host "Files by author '$authorName' copied to: $destinationDrive"
Write-Host "Total copied: $($copiedFiles.Count) files"
Write-Host "-----------------------------------------"
pause
