import tkinter as tk
from tkinter import ttk

def apply_style(root):
    # Estilo general de la ventana
    root.configure(bg="#F5F5F5")  # Fondo gris claro
    root.option_add("*Font", "Arial 11")  # Fuente predeterminada
    root.option_add("*Foreground", "#333333")  # Texto en color gris oscuro
    root.option_add("*Button.Background", "#4C8BF5")  # Fondo de botones
    root.option_add("*Button.Foreground", "#FFFFFF")  # Texto de botones
    root.option_add("*TCombobox*Listbox*Background", "#FFFFFF")  # Fondo del combobox desplegable
    root.option_add("*TCombobox*Listbox*Foreground", "#333333")  # Texto del combobox desplegable

    # Etiqueta superior
    label = tk.Label(
        root,
        text="Select a script to execute or manage configuration:",
        font=("Arial", 13),
        bg="#F5F5F5",
        fg="#333333",
        pady=10,
    )
    label.pack(pady=(20, 10))

def create_buttons_and_dropdown(root, open_excel, run_script, script_var):
    # Configuración del estilo para combobox
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TCombobox", fieldbackground="#FFFFFF", background="#FFFFFF", foreground="#333333", padding=5)

    # Efecto hover para botones
    def on_enter(button, hover_color):
        button.config(bg=hover_color)

    def on_leave(button, original_color):
        button.config(bg=original_color)

    # Selector de script
    script_menu = ttk.Combobox(
        root,
        textvariable=script_var,
        values=["CIFS", "NFS"],
        state="readonly",
        font=("Arial", 11),
        justify="center",
    )
    script_menu.pack(pady=15)

    # Botón para abrir Excel
    open_button = tk.Button(
        root,
        text="Open Excel",
        font=("Arial", 11, "bold"),
        command=open_excel,
        width=18,
        height=2,
        bg="#6CC24A",  # Verde suave
        fg="white",
        activebackground="#5EAC3C",
        relief="groove",
    )
    open_button.pack(pady=10)
    open_button.bind("<Enter>", lambda e: on_enter(open_button, "#5EAC3C"))  # Hover color más oscuro
    open_button.bind("<Leave>", lambda e: on_leave(open_button, "#6CC24A"))

    # Botón para ejecutar el script
    run_button = tk.Button(
        root,
        text="Run Script",
        font=("Arial", 11, "bold"),
        command=lambda: run_script(script_var.get()),
        width=18,
        height=2,
        bg="#4C8BF5",  # Azul
        fg="white",
        activebackground="#3A73C9",
        relief="groove",
    )
    run_button.pack(pady=10)
    run_button.bind("<Enter>", lambda e: on_enter(run_button, "#3A73C9"))  # Hover color más oscuro
    run_button.bind("<Leave>", lambda e: on_leave(run_button, "#4C8BF5"))

    return script_menu, open_button, run_button
