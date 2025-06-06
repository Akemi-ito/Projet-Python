import tkinter as tk
from tkinter import ttk, messagebox
import threading
import main

class ReportApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Générateur de rapport Word")
        self.geometry("400x250")

        ttk.Label(self, text="Générateur de rapport Project Gutenberg", font=("Arial", 14)).pack(pady=10)

        ttk.Label(self, text="Nom de l'auteur du rapport :").pack(pady=(10, 0))
        self.nom_var = tk.StringVar(value="Kervoelen Erwann")
        ttk.Entry(self, textvariable=self.nom_var, width=40).pack(pady=5)

        self.bouton = ttk.Button(self, text="Générer le rapport", command=self.lancer_generation)
        self.bouton.pack(pady=20)

        self.status = ttk.Label(self, text="", foreground="green")
        self.status.pack()

    def lancer_generation(self):
        self.bouton.config(state="disabled")
        self.status.config(text="Génération en cours...")
        threading.Thread(target=self.generation_worker).start()

    def generation_worker(self):
        try:
            main.AUTEUR_RAPPORT = self.nom_var.get()
            main.main()
            self.status.config(text="✅ Rapport généré avec succès.")
            messagebox.showinfo("Succès", f"Le rapport a été généré :\n{main.DOCX_PATH}")
        except Exception as e:
            messagebox.showerror("Erreur", str(e))
            self.status.config(text="Erreur lors de la génération.")
        finally:
            self.bouton.config(state="normal")

if __name__ == "__main__":
    app = ReportApp()
    app.mainloop()
