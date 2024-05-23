import pyautogui
import time
import tkinter as tk
from tkinter import simpledialog
import keyboard
import sys
import win32com.client as win32
import os
import pandas as pd
from pywinauto import Application, Desktop
import psutil
import subprocess

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

def save_current_excel_file(destination_directory, matiere):
    # Initialiser l'application Excel via COM
    excel = win32.Dispatch('Excel.Application')
    
    # Vérifier si des classeurs sont ouverts
    if excel.Workbooks.Count == 0:
        print("Aucun fichier Excel ouvert.")
        return
    
    # Obtenir le classeur actif
    workbook = excel.ActiveWorkbook
    
    # Créer le nouveau chemin de sauvegarde avec le nom basé sur la matière
    new_file_name = f"{matiere}_Intervenant.xlsx"
    new_path = os.path.join(destination_directory, new_file_name)
    
    # Enregistrer le fichier dans le nouveau répertoire
    workbook.SaveAs(new_path)
    print(f"Fichier enregistré sous: {new_path}")
    return new_path

def show_filter_popup(matiere, file_path):
    root = tk.Tk()
    root.title("Critères de sélection")
    
    tk.Label(root, text="Nom").grid(row=0)
    tk.Label(root, text="Prénom").grid(row=1)
    tk.Label(root, text="Email").grid(row=2)
    tk.Label(root, text="Numéro").grid(row=3)
    tk.Label(root, text="Ville(s)").grid(row=4)
    tk.Label(root, text="Tarif Horaire Maximal").grid(row=5)
    tk.Label(root, text="Nombre d'Interventions").grid(row=6)
    tk.Label(root, text="Date de la Dernière Intervention").grid(row=7)
    
    nom = tk.Entry(root)
    prenom = tk.Entry(root)
    email = tk.Entry(root)
    numero = tk.Entry(root)
    ville = tk.Entry(root)
    tarif_horaire = tk.Entry(root)
    nombre_interventions = tk.Entry(root)
    derniere_intervention = tk.Entry(root)
    
    nom.grid(row=0, column=1)
    prenom.grid(row=1, column=1)
    email.grid(row=2, column=1)
    numero.grid(row=3, column=1)
    ville.grid(row=4, column=1)
    tarif_horaire.grid(row=5, column=1)
    nombre_interventions.grid(row=6, column=1)
    derniere_intervention.grid(row=7, column=1)
    
    def submit_filters():
        filters = {
            'NomPersonne': nom.get(),
            'PrenomPersonne': prenom.get(),
            'EmailPersonne': email.get(),
            'PortablePersonne': numero.get(),
            'VillePersonne': ville.get(),
            'TarifHoraire': tarif_horaire.get(),
            'NombreAnimationMoins2AnsToutesMatieresConfondues': nombre_interventions.get(),
            'DerniereAnimationToutesMatieresConfondues': derniere_intervention.get()
        }
        root.destroy()
        filter_excel_file(file_path, filters)
    
    tk.Button(root, text='Soumettre', command=submit_filters).grid(row=8, column=1, sticky=tk.W, pady=4)
    root.mainloop()

def filter_excel_file(file_path, filters):
    df = pd.read_excel(file_path)
    
    # Appliquer les filtres
    for key, value in filters.items():
        if value:
            if key == 'VillePersonne':
                villes = [v.strip() for v in value.split(',')]
                df = df[df[key].str.contains('|'.join(villes), case=False, na=False)]
            elif key == 'TarifHoraire':
                df = df[df[key] <= float(value)]
            elif key == 'NombreAnimationMoins2AnsToutesMatieresConfondues':
                df = df[df[key] >= int(value)]
            else:
                df = df[df[key].str.contains(value, case=False, na=False)]
    
    top_5 = df.head(5)
    print("Top 5 Intervenants selon les critères :")
    print(top_5)
    
    output_file = file_path.replace('.xlsx', '_Filtré.xlsx')
    top_5.to_excel(output_file, index=False)
    print(f"Fichier filtré enregistré sous: {output_file}")

    # Ouvrir le fichier Excel filtré
    os.startfile(output_file)

    # Appeler le script de la fenêtre de demande après la génération du fichier Excel filtré
    from SendEmail import create_popup
    create_popup(output_file)

def wait_for_export_done():
    # Utiliser pywinauto pour attendre l'apparition de la fenêtre
    print("En attente de la fenêtre contextuelle d'exportation...")
    while True:
        try:
            # Trouver la fenêtre avec le titre contenant "bora"
            app = Desktop(backend="uia").window(title_re=".*bora.*")
            if app.exists():
                app.wait('visible', timeout=60)
                print("Export Excel terminée avec succès.")
                app.OK.click()  # Cliquer sur le bouton OK pour fermer la fenêtre
                break
        except Exception as e:
            time.sleep(1)
            continue

def wait_for_bora_window():
    # Utiliser pywinauto pour attendre l'apparition de la fenêtre "bora"
    print("En attente de la fenêtre 'bora'...")
    while True:
        try:
            app = Desktop(backend="uia").window(title_re=".*bora.*")
            if app.exists():
                app.wait('visible', timeout=60)
                print("Fenêtre 'bora' détectée.")
                break
        except Exception as e:
            time.sleep(1)
            continue

# Obtenir la matière recherchée de l'utilisateur
matiere = get_user_input()

# Vérifier si Échap est pressé à chaque étape critique
check_escape()

# Aller sur le bureau
pyautogui.moveTo(1917, 1049)
pyautogui.click()
check_escape()

# Lancer l'application après être allé sur le bureau
subprocess.Popen(r"C:\Program Files (x86)\CESI\Fng\gpForm.exe")

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

# Vérifier si la fenêtre 'bora' est ouverte avant de continuer
wait_for_bora_window()

# États et courriers
time.sleep(2)
pyautogui.moveTo(169, 33)
pyautogui.click()
check_escape()

# Gestion des requêtes
time.sleep(2)
pyautogui.moveTo(159, 75)
pyautogui.click()
time.sleep(1)
pyautogui.moveTo(134, 153)
pyautogui.click()
check_escape()

pyautogui.write('0614')
time.sleep(1)
pyautogui.press('enter')
#time.sleep(2)
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

# Attendre la fenêtre contextuelle d'exportation
wait_for_export_done()

# Sauvegarder le fichier Excel ouvert avec le nom basé sur la matière
destination_directory = r"C:\Users\achachoui\Documents\Saved_Excel_Files"

# S'assurer que le répertoire de destination existe
if not os.path.exists(destination_directory):
    os.makedirs(destination_directory)

file_path = save_current_excel_file(destination_directory, matiere)

# Afficher la fenêtre de filtre après la sauvegarde du fichier
show_filter_popup(matiere, file_path)

# Fin

# Pour les tests
# print(pyautogui.position())
