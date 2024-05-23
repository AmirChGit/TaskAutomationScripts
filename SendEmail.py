import tkinter as tk
import pandas as pd
import win32com.client as win32
import os

def send_email(email):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = 'Request for Intervention'
    mail.Body = 'Dear Sir/Madam,\n\nWe would like to request your intervention services for our upcoming project.\n\nBest regards,\nYour Name'
    mail.Display()

def create_popup(file_path):
    df = pd.read_excel(file_path)
    
    root = tk.Tk()
    root.title("Filtered Results")

    canvas = tk.Canvas(root)
    scrollbar = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    for index, row in df.iterrows():
        frame = tk.Frame(scrollable_frame, borderwidth=1, relief="solid")
        frame.pack(fill="x", pady=5)

        tk.Label(frame, text=f"Nom: {row['NomPersonne']}, Prénom: {row['PrenomPersonne']}, Email: {row['EmailPersonne']}, Ville: {row['VillePersonne']}, Tarif: {row['TarifHoraire']}, Interventions: {row['NombreAnimationMoins2AnsToutesMatieresConfondues']}, Dernière Intervention: {row['DerniereAnimationToutesMatieresConfondues']}").pack(side="left", padx=5)
        
        btn = tk.Button(frame, text="Send Request", command=lambda email=row['EmailPersonne']: send_email(email))
        btn.pack(side="right", padx=5)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    root.mainloop()
