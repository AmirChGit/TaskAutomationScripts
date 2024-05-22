import pyautogui
import time

time.sleep(5)

#Go To Desktop
pyautogui.moveTo(1917, 1049)
pyautogui.click()

#Open FNG
time.sleep(2)
pyautogui.moveTo(60, 625)
pyautogui.doubleClick()

#Login FNG
time.sleep(2)
pyautogui.write('Cesi2023')
pyautogui.press('enter')
time.sleep(5)

#open Bora
pyautogui.moveTo(1731, 888)
pyautogui.doubleClick()

#Etats et courriers
time.sleep(1)
pyautogui.moveTo(169, 33)
pyautogui.click()

#Gestion des requestes
time.sleep(1)
pyautogui.moveTo(159, 75)
pyautogui.click()
time.sleep(1)
pyautogui.moveTo(134, 153)
pyautogui.click()

pyautogui.write('0329')
pyautogui.press('enter')
time.sleep(2)

#Search
pyautogui.moveTo(516, 137)
pyautogui.click()
time.sleep(1)
pyautogui.moveTo(221, 236)
pyautogui.click()
time.sleep(1)
pyautogui.moveTo(242, 254)
pyautogui.click()
time.sleep(1)
pyautogui.write('Gestion de projet')
pyautogui.moveTo(573, 260)
pyautogui.click()
pyautogui.press('enter')
time.sleep(1)
pyautogui.moveTo(227, 201)
pyautogui.click()
time.sleep(5)

#Export to Excel
pyautogui.moveTo(67, 91)
pyautogui.click()
time.sleep(2)
pyautogui.moveTo(1027, 652)
pyautogui.click()
time.sleep(2)
pyautogui.moveTo(1037, 588)
pyautogui.click()

#done

print(pyautogui.position())


# # Exemple pour envoyer un email
# def envoyer_email(destinataire, sujet, message):
#     serveur = smtplib.SMTP('amirchachoui@gmail.com', 587)
#     serveur.starttls()
#     serveur.login("votre_email@example.com", "votre_mot_de_passe")
#     email = f"Subject: {sujet}\n\n{message}"
#     serveur.sendmail("votre_email@example.com", destinataire, email)
#     serveur.quit()

# # Exemple pour automatiser une tâche de bureau
# def automatiser_tache():
#     pyautogui.click(x=100, y=200)  # Clic à une position spécifique
#     pyautogui.write('Recherche d\'expert')  # Écrire du texte
#     pyautogui.press('enter')  # Appuyer sur la touche 'Entrée'

# # Exemple pour enregistrer des données dans un fichier Excel
# def enregistrer_excel(fichier, donnees):
#     wb = openpyxl.load_workbook(fichier)
#     ws = wb.active
#     ws.append(donnees)
#     wb.save(fichier)

# # Exemple pour générer un rapport
# def generer_rapport(donnees):
#     date_rapport = datetime.now().strftime("%Y-%m-%d")
#     rapport = f"Rapport du {date_rapport}\n\n{donnees}"
#     return rapport

# # Utilisation des fonctions
# envoyer_email("expert@example.com", "Demande d'intervention", "Bonjour, nous souhaitons organiser une session.")
# automatiser_tache()
# enregistrer_excel("sessions.xlsx", ["Expert", "Date", "Heure", "Sujet"])
# rapport = generer_rapport("Session organisée avec succès.")
# print(rapport)
