# Variables
$outputCsv = "$env:TEMP\system_info.csv"
$emailRecipient = "stat@kogus.io"

# Affichage du logo
Write-Host "  _  __                       _       " -ForegroundColor Cyan
Write-Host " | |/ /___   __ _ _   _ ___  (_) ___  " -ForegroundColor Cyan
Write-Host " | ' // _ \ / _` | | | / __| | |/ _ \ " -ForegroundColor Cyan
Write-Host " | . \ (_) | (_| | |_| \__ \_| | (_)| " -ForegroundColor Cyan
Write-Host " |_|\_\___/ \__, |\__,_|___(_)_|\___/ " -ForegroundColor Cyan
Write-Host "            |___/                     " -ForegroundColor Cyan
Write-Host "                                      " -ForegroundColor Cyan
Write-Host "**************************************" -ForegroundColor Cyan

# Affichage du message de patience et lancement du chronomètre
Write-Host "Analyse du systeme en cours" -ForegroundColor Yellow
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Barre de chargement dynamique
$totalSteps = 20
for ($i = 0; $i -le $totalSteps; $i++) {
    $progress = "[{0}{1}] {2}%" -f ('#' * $i), ('-' * ($totalSteps - $i)), [math]::Round(($i / $totalSteps) * 100)
    Write-Host "`r$progress" -NoNewline -ForegroundColor Green
    Start-Sleep -Milliseconds 300
}
Write-Host ""  # Pour sauter une ligne après la barre de chargement

# Récupération des informations de l'ordinateur
$processor = Get-CimInstance Win32_Processor | Select-Object -ExpandProperty Name
$ram = [math]::round((Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum / 1GB, 2)
$disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$diskSize = [math]::round($disk.Size / 1GB, 2)
$diskFreeSpace = [math]::round($disk.FreeSpace / 1GB, 2)

# Détermination du type de disque (SSD ou HDD)
$physicalDisk = Get-CimInstance Win32_DiskDrive | Where-Object { $_.DeviceID -like '*PHYSICALDRIVE0*' }
$diskType = if ($physicalDisk.Model -match 'SSD' -or $physicalDisk.MediaType -like '*SSD*' -or $physicalDisk.MediaType -eq 'Removable Media') { 'SSD : ok' } else { 'SSD : no' }

# Création du fichier CSV
$csvContent = @(
    "Processeur",
    "$processor",
    "Quantité de RAM (GB)",
    "$ram",
    "Taille du disque (GB)",
    "$diskSize",
    "Espace libre (GB)",
    "$diskFreeSpace",
    "Type de disque",
    "$diskType"
)
$csvContent | Out-File -FilePath $outputCsv -Encoding UTF8

# Arrêt du chronomètre et affichage du temps 
$stopwatch.Stop()
elapsed_time = $stopwatch.Elapsed
Write-Host "Analyse terminée en $($elapsed_time.Hours) heures, $($elapsed_time.Minutes) minutes, $($elapsed_time.Seconds) secondes." -ForegroundColor Green

# Création d'un email dans Outlook avec le fichier CSV en pièce jointe
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.To = $emailRecipient
$mail.Subject = "Informations de configuration de l'ordinateur"
$mail.Body = "Veuillez trouver en pièce jointe les informations sur la configuration de cet ordinateur."
$mail.Attachments.Add($outputCsv)
$mail.Display() # Ouvre le mail dans Outlook prêt à être envoyé

# Nettoyage (optionnel)
# Remove-Item -Path $outputCsv
