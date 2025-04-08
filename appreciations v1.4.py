import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def transform_excel(file_path):
    try:
        # Lecture du fichier Excel
        df = pd.read_excel(file_path, engine="openpyxl")
        # Transformation en format texte pour un prompt AI
        transformed_data = df.to_string(index=False)
        return df, transformed_data
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur s'est produite : {e}")
        return None, None

def save_to_file(df):
    try:
        # Générer les données transformées avec séparateur "|"
        transformed_data = df.to_csv(sep="|", index=False)
        
        # Prompt personnalisé
        prompt = (
            "je veux que tu analyses les notes des élèves en considérant la 2ème ligne du tableau comme "
            "étant les coefficients des notes de la colonne correspondante et en tenant compte des mots clefs "
            "de la dernière colonne. Une fois l'analyse faite, je veux que tu génères une appréciation de 250 "
            "caractères maximum dans un ton académique qui respecte les recommandations françaises de "
            "l'éducation nationale.\n\n"
        )

        # Sauvegarder le tout dans un fichier texte
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Fichier texte", "*.txt")])
        if file_path:
            with open(file_path, "w", encoding="utf-8") as file:
                file.write(prompt)
                file.write(transformed_data)
            messagebox.showinfo("Succès", f"Fichier sauvegardé : {file_path}")
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur s'est produite : {e}")

def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xls *.xlsx")])
    if file_path:
        df, transformed_data = transform_excel(file_path)
        if transformed_data:
            result_text.delete(1.0, tk.END)
            result_text.insert(tk.END, transformed_data)
            save_button.config(state=tk.NORMAL)
            save_button.config(command=lambda: save_to_file(df))

# Création de la fenêtre principale
root = tk.Tk()
root.title("Transformateur Excel pour IA")

# Bouton pour charger un fichier
load_button = tk.Button(root, text="Charger un fichier Excel", command=open_file)
load_button.pack(pady=10)

# Bouton pour sauvegarder le fichier (désactivé par défaut)
save_button = tk.Button(root, text="Sauvegarder en fichier texte", state=tk.DISABLED)
save_button.pack(pady=10)

# Zone de texte pour afficher les données transformées
result_text = tk.Text(root, wrap="word", height=20, width=60)
result_text.pack(padx=10, pady=10)

# Lancement de l'interface
root.mainloop()
