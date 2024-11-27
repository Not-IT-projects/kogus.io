import psutil
import platform
import os
import time
import win32com.client as win32

# Variables
output_csv = os.path.join(os.getenv('TEMP'), 'system_info.csv')
email_recipient = "stat@kogus.io"

# Affichage d'un message d'attente
print("""
  _  __                       _       
 | |/ /___   __ _ _   _ ___  (_) ___  
 | ' // _ \ / _` | | | / __| | |/ _ \ 
 | . \ (_) | (_| | |_| \__ \_| | (_) |
 |_|\_\___/ \__, |\__,_|___(_)_|\___/ 
            |___/                      
""")
print("Analyse du système en cours...\n")

# Chronomètre
start_time = time.time()

# Récupération des informations de l'ordinateur
processor = platform.processor()
ram = round(psutil.virtual_memory().total / (1024 ** 3), 2)
disk = psutil.disk_usage('/')
disk_size = round(disk.total / (1024 ** 3), 2)
disk_free_space = round(disk.free / (1024 ** 3), 2)

# Détermination du type de disque (SSD ou HDD)
disk_type = "SSD : no"
for disk_part in psutil.disk_partitions():
    if 'C:' in disk_part.device:
        try:
            if "ssd" in psutil.disk_io_counters(perdisk=True).keys():
                disk_type = "SSD : ok"
        except Exception as e:
            pass

# Création du fichier CSV
csv_content = [
    "Processeur",
    processor,
    "Quantité de RAM (GB)",
    str(ram),
    "Taille du disque (GB)",
    str(disk_size),
    "Espace libre (GB)",
    str(disk_free_space),
    "Type de disque",
    disk_type
]

with open(output_csv, 'w') as f:
    for i in range(0, len(csv_content), 2):
        f.write(f"{csv_content[i]},{csv_content[i + 1]}\n")

# Arrêt du chronomètre et affichage du temps
elapsed_time = time.time() - start_time
print(f"\nAnalyse terminée en {int(elapsed_time // 3600)} heures, {int((elapsed_time % 3600) // 60)} minutes, {int(elapsed_time % 60)} secondes.\n")

# Création de l'email dans Outlook avec le fichier CSV en pièce jointe
try:
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = email_recipient
    mail.Subject = "Informations de configuration de l'ordinateur"
    mail.Body = "Veuillez trouver en pièce jointe les informations sur la configuration de cet ordinateur."
    mail.Attachments.Add(output_csv)
    mail.Display()  # Ouvre le mail dans Outlook prêt à être envoyé
    print("Email prêt à être envoyé dans Outlook.")
except Exception as e:
    print(f"Erreur lors de la création de l'email dans Outlook : {e}")
