# -----------------------------------------
# ФУНКЦІЇ
# -----------------------------------------

# Функція вибору папки через Проводник
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

# Функція відображення меню
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

# Функція отримання автора файлу
function Get-FileAuthor($filePath) {
    $folder = $shell.Namespace((Get-Item $filePath).DirectoryName)
    $item = $folder.ParseName((Get-Item $filePath).Name)
    return $folder.GetDetailsOf($item, 20)  # 20 — індекс властивості "Автор"
}

# Функція виведення повідомлення про початок копіювання
function Show-StartMessage {
    Write-Host "-----------------------------------------"
    Write-Host "Starting file copy process..."
    Write-Host "-----------------------------------------"
    Start-Sleep 1
}

# Функція виведення повідомлення про завершення копіювання
function Show-CompletionMessage {
    Write-Host "-----------------------------------------"
    Write-Host "PROCESS COMPLETED"
    Write-Host "Files copied to: $destinationDrive"
    Write-Host "Total copied: $($copiedFiles.Count) files"
    Write-Host "-----------------------------------------"
    Write-Host "Thank you for using our tool!"
    Write-Host "We appreciate your trust."
    Write-Host "Have a great day! <3"
    Write-Host "-----------------------------------------"
    pause
}

# Функція виведення прощального повідомлення
function Show-ExitMessage {
    Write-Host "-----------------------------------------"
    Write-Host "Thank you for using our tool!"
    Write-Host "We appreciate your trust."
    Write-Host "Have a great day! <3"
    Write-Host "-----------------------------------------"
    Start-Sleep 5
}

# -----------------------------------------
# ОСНОВНИЙ КОД
# -----------------------------------------

# Запуск з правами адміністратора
$currProc = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-Not $currProc.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Ініціалізація змінних
$sourceDrive = "Not selected"
$destinationDrive = "Not selected"
$authorName = "Not set"
$exitMenu = $false
$shell = New-Object -ComObject Shell.Application

# Цикл меню
do {
    Show-Menu
    $choice = [System.Console]::ReadKey($true).KeyChar

    switch ($choice) {
        "1" { $selectedFolder = Select-Folder -message "Select the source folder"; if ($selectedFolder) { $sourceDrive = $selectedFolder } }
        "2" { $selectedFolder = Select-Folder -message "Select the destination folder"; if ($selectedFolder) { $destinationDrive = $selectedFolder } }
        "3" { Write-Host "\nEnter author name: " -NoNewline; $authorName = Read-Host }
        "4" {
            if ($sourceDrive -eq "Not selected" -or $destinationDrive -eq "Not selected" -or $authorName -eq "Not set") {
                Write-Host "\nERROR: Please complete all fields before starting the process!"
                Start-Sleep 2
            } else {
                $exitMenu = $true
            }
        }
        "5" { Show-ExitMessage; exit }
        default { Write-Host "Invalid option! Try again."; Start-Sleep 1 }
    }
} while (-Not $exitMenu)

# Переконуємось, що папка призначення існує
if (!(Test-Path $destinationDrive)) {
    New-Item -ItemType Directory -Path $destinationDrive | Out-Null
}

# Отримуємо список файлів
$files = Get-ChildItem -Path $sourceDrive -Recurse -File
$totalFiles = $files.Count
$processedFiles = 0
$copiedFiles = @()

Show-StartMessage

# Копіювання файлів за автором
foreach ($file in $files) {
    $fileAuthor = Get-FileAuthor $file.FullName
    if ($fileAuthor -eq $authorName) {
        $destinationPath = Join-Path -Path $destinationDrive -ChildPath $file.Name
        Copy-Item -Path $file.FullName -Destination $destinationPath -Force
        $copiedFiles += $file.FullName
        Write-Host "[COPIED] $($file.FullName)"
    }
    
    # Оновлення прогресу
    $processedFiles++
    $percentComplete = [math]::Round(($processedFiles / $totalFiles) * 100, 2)
    Write-Progress -Activity "Copying files" -Status "$percentComplete% completed" -PercentComplete $percentComplete
}

Show-CompletionMessage
