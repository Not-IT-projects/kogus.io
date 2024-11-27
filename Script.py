# SCript kogus en python
# vincent BLIN 27/11/24

import psutil
import platform
import smtplib
import os
from email.message import EmailMessage
import time

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

# Création et envoi de l'email
msg = EmailMessage()
msg['Subject'] = "Informations de configuration de l'ordinateur"
msg['From'] = "your_email@example.com"
msg['To'] = email_recipient
msg.set_content("Veuillez trouver en pièce jointe les informations sur la configuration de cet ordinateur.")

with open(output_csv, 'rb') as f:
    msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename='system_info.csv')

# Configurer l'envoi d'email (à adapter avec vos informations SMTP)
try:
    with smtplib.SMTP('smtp.example.com', 587) as server:
        server.starttls()
        server.login('your_email@example.com', 'your_password')
        server.send_message(msg)
        print("Email envoyé avec succès.")
except Exception as e:
    print(f"Erreur lors de l'envoi de l'email : {e}")
