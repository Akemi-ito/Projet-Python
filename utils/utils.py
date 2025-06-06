import tkinter as tk
from tkinter import messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

def setup_matplotlib_style():
    """Configure le style des graphiques"""
    plt.style.use('ggplot')
    params = {
        'axes.facecolor': '#FFFFFF',
        'figure.facecolor': '#FFFFFF',
        'axes.edgecolor': '#3E5C76',
        'axes.labelcolor': '#0C0C0C',
        'xtick.color': '#3E5C76',
        'ytick.color': '#3E5C76',
        'axes.titlecolor': '#00BFFF',
        'axes.prop_cycle': plt.cycler(color=['#00BFFF', '#3E5C76', '#B0B3B8']),
        'grid.color': '#B0B3B8'
    }
    plt.rcParams.update(params)

def create_matplotlib_graph(parent, data, title=""):
    """Crée un graphique"""
    setup_matplotlib_style()

    window = tk.Toplevel(parent)
    window.title(f"Statistiques: {title}")
    window.configure(bg='#FFFFFF')  # Blanc


    fig, ax = plt.subplots(figsize=(10, 6), facecolor='#FFFFFF')  # Blanc

    # Graphique à barres
    bars = ax.bar(
        data['labels'],
        data['values'],
        color='#00BFFF',  # Bleu cyan
        edgecolor='#3E5C76',  # Bleu acier
        linewidth=1.5,
        width=0.7
    )

    # Configuration des titres et labels
    ax.set_title(title, pad=20, color='#00BFFF', fontweight='bold')  # Titre en bleu cyan
    ax.set_xlabel('Type de Mission', color='#3E5C76', fontweight='bold')  # Texte en bleu acier
    ax.set_ylabel('Nombre de Lancements', color='#3E5C76', fontweight='bold')

    # Configuration des barres
    ax.tick_params(axis='x', colors='#3E5C76')  # Couleur bleu acier
    ax.tick_params(axis='y', colors='#3E5C76')

    # Ajout des valeurs sur les barres
    for bar in bars:
        height = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width()/2., height,
            f'{height}',
            ha='center', va='bottom',
            color='#0C0C0C',  # Noir
            fontsize=10,
            fontweight='bold'
        )

    # Style du cadre
    for spine in ax.spines.values():
        spine.set_color('#B0B3B8')  # Gris métal
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    # Fond de la zone du graphique
    ax.set_facecolor('#FFFFFF')  # Blanc

    # Grille
    ax.grid(True, color='#B0B3B8', linestyle='--', alpha=0.7)

    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()

    # Intégration dans Tkinter
    canvas = FigureCanvasTkAgg(fig, master=window)
    canvas.draw()
    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    return window

def show_error(message):
    """Affiche une boîte de dialogue d'erreur"""
    messagebox.showerror(
        "Erreur",
        message,
        icon="error",
        parent=tk._default_root
    )

def show_info(message):
    """Affiche une boîte de dialogue d'information"""
    messagebox.showinfo(
        "Information",
        message,
        parent=tk._default_root
    )

def show_warning(message):
    """Affiche une boîte de dialogue d'avertissement"""
    messagebox.showwarning(
        "Avertissement",
        message,
        parent=tk._default_root
    )

def ask_confirmation(message):
    """Demande une confirmation à l'utilisateur """
    return messagebox.askyesno(
        "Confirmation",
        message,
        parent=tk._default_root
    )

