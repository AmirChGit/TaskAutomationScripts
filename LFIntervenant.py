import pyautogui
import time
import tkinter as tk
from tkinter import simpledialog
import keyboard
import sys

def get_user_input():
    root = tk.Tk()
    root.withdraw()  # Masquer la fenêtre principale
    # Utiliser simpledialog pour obtenir l'entrée utilisateur
    matiere = simpledialog.askstring("Input", "Quelle matière cherchez-vous ?")
    root.destroy()  # Détruire la fenêtre après l'obtention de l'entrée
    return matiere

def check_escape():
    if keyboard.is_pressed('esc'):
        print("Échap pressé. Arrêt du script.")
        sys.exit()

# Obtenir la matière recherchée de l'utilisateur
matiere = get_user_input()

# Vérifiez si l'utilisateur a annulé l'entrée
if not matiere:
    print("Aucune matière saisie, arrêt du script.")
    sys.exit()

# Mettre une pause avant de commencer l'automatisation
time.sleep(5)

# Vérifier si Échap est pressé à chaque étape critique
check_escape()

# Aller sur le bureau
pyautogui.moveTo(1917, 1049)
pyautogui.click()
check_escape()

# Ouvrir FNG
time.sleep(2)
pyautogui.moveTo(60, 625)
pyautogui.doubleClick()
check_escape()

# Login FNG
time.sleep(2)
pyautogui.write('Cesi2023')
pyautogui.press('enter')
time.sleep(5)
check_escape()

# Ouvrir Bora
pyautogui.moveTo(1731, 888)
pyautogui.doubleClick()
check_escape()

# États et courriers
time.sleep(1)
pyautogui.moveTo(169, 33)
pyautogui.click()
check_escape()

# Gestion des requêtes
time.sleep(1)
pyautogui.moveTo(159, 75)
pyautogui.click()
time.sleep(1)
pyautogui.moveTo(134, 153)
pyautogui.click()
check_escape()

pyautogui.write('0614')
#pyautogui.write('0329')
pyautogui.press('enter')
time.sleep(2)
check_escape()

# Recherche
pyautogui.moveTo(516, 137)
pyautogui.click()
time.sleep(1)
pyautogui.moveTo(221, 236)
pyautogui.click()
time.sleep(1)
pyautogui.moveTo(242, 254)
pyautogui.click()
time.sleep(1)
pyautogui.write(matiere)
pyautogui.moveTo(573, 260)
pyautogui.click()
pyautogui.press('enter')
time.sleep(1)
pyautogui.moveTo(227, 201)
pyautogui.click()
time.sleep(5)
check_escape()

# Exporter vers Excel
pyautogui.moveTo(67, 91)
pyautogui.click()
time.sleep(2)
pyautogui.moveTo(1027, 652)
pyautogui.click()
time.sleep(2)
pyautogui.moveTo(1037, 588)
pyautogui.click()
check_escape()

# Fin

# Pour les tests
# print(pyautogui.position())
