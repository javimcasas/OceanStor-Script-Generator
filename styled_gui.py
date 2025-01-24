import tkinter as tk
from tkinter import ttk

def apply_style(root):
    root.configure(bg="#F5F5F5")  # Fondo general de la ventana
    root.option_add("*Font", "Arial 11")  # Fuente predeterminada
    root.option_add("*Foreground", "#333333")  # Color de texto
    root.option_add("*Button.Background", "#D9D9D9")  # Fondo de botones más neutro
    root.option_add("*Button.Foreground", "#333333")  # Texto de botones más oscuro

    label = tk.Label(
        root,
        text="Select a script to execute or manage configuration:",
        font=("Arial", 13),
        bg="#F5F5F5",
        fg="#333333",
        pady=10,
    )
    label.pack(pady=(20, 10))


def create_buttons_and_dropdown(root, open_excel, run_script, script_var, command_var):
    # Configuración del estilo del combobox
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TCombobox", fieldbackground="#FFFFFF", background="#FFFFFF", foreground="#333333", padding=5)

    # Selector de script
    script_menu = ttk.Combobox(
        root,
        textvariable=script_var,
        values=["CIFS", "NFS"],
        state="readonly",
        font=("Arial", 11),
        justify="center",
    )
    script_menu.pack(pady=10)

    # Frame para los botones de comando (Create, Change, Show)
    command_buttons_frame = tk.Frame(root, bg="#F5F5F5")
    command_buttons_frame.pack(pady=(20, 20))

    # Efecto hover para botones
    def on_enter(button, hover_color):
        button.config(bg=hover_color)

    def on_leave(button, original_color):
        button.config(bg=original_color)

    # Botones de comando (Create, Change, Show)
    def select_command(command):
        command_var.set(command)
        for key, button in command_buttons.items():
            if key == command:
                button.config(relief="sunken", bg="#CCCCCC")  # Color más oscuro para botón activo
            else:
                button.config(relief="raised", bg="#D9D9D9")  # Restaurar color neutro

    command_buttons = {
        "Create": tk.Button(
            command_buttons_frame,
            text="Create",
            command=lambda: select_command("Create"),
            width=10,
            height=1,
            bg="#D9D9D9",  # Color neutro
            fg="#333333",
            activebackground="#CCCCCC",  # Color neutro más oscuro
        ),
        "Change": tk.Button(
            command_buttons_frame,
            text="Change",
            command=lambda: select_command("Change"),
            width=10,
            height=1,
            bg="#D9D9D9",
            fg="#333333",
            activebackground="#CCCCCC",
        ),
        "Show": tk.Button(
            command_buttons_frame,
            text="Show",
            command=lambda: select_command("Show"),
            width=10,
            height=1,
            bg="#D9D9D9",
            fg="#333333",
            activebackground="#CCCCCC",
        ),
    }

    # Agregar efecto hover y colocar botones en el frame
    for key, button in command_buttons.items():
        button.pack(side="left", padx=5)
        button.bind("<Enter>", lambda e, b=button: on_enter(b, "#E6E6E6"))  # Hover color más claro
        button.bind("<Leave>", lambda e, b=button: on_leave(b, "#D9D9D9"))

    # Botón para abrir Excel
    open_button = tk.Button(
        root,
        text="Open Excel",
        font=("Arial", 11, "bold"),
        command=lambda: open_excel(script_var.get(), command_var.get()),
        width=18,
        height=2,
        bg="#6CC24A",  # Verde suave
        fg="white",
        activebackground="#5EAC3C",
    )
    open_button.pack(pady=(20, 10))
    # Efecto hover para Open Excel
    open_button.bind("<Enter>", lambda e: on_enter(open_button, "#5EAC3C"))
    open_button.bind("<Leave>", lambda e: on_leave(open_button, "#6CC24A"))

    # Botón para ejecutar el script
    run_button = tk.Button(
        root,
        text="Run Script",
        font=("Arial", 11, "bold"),
        command=lambda: run_script(script_var.get(), command_var.get()),
        width=18,
        height=2,
        bg="#4C8BF5",  # Azul
        fg="white",
        activebackground="#3A73C9",
    )
    run_button.pack(pady=(10, 20))
    # Efecto hover para Run Script
    run_button.bind("<Enter>", lambda e: on_enter(run_button, "#3A73C9"))
    run_button.bind("<Leave>", lambda e: on_leave(run_button, "#4C8BF5"))

    return script_menu, open_button, run_button, command_buttons
