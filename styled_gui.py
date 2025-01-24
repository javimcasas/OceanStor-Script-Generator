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

def create_buttons_and_dropdown(root, open_excel, run_script, script_var, command_var, clear_excel, config):
    # Configuración del estilo del combobox
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TCombobox", fieldbackground="#FFFFFF", background="#FFFFFF", foreground="#333333", padding=5)

    # Selector de script
    script_menu = ttk.Combobox(
        root,
        textvariable=script_var,
        values=list(config.keys()),  # Use keys from JSON (e.g., CIFS, NFS)
        state="readonly",
        font=("Arial", 11),
        justify="center",
    )
    script_menu.pack(pady=10)

    # Frame para los botones de comando (Create, Change, Show, etc.)
    command_buttons_frame = tk.Frame(root, bg="#F5F5F5")
    command_buttons_frame.pack(pady=(20, 20))

    # Efecto hover para botones
    def on_enter(button, hover_color):
        button.config(bg=hover_color)

    def on_leave(button, original_color):
        button.config(bg=original_color)

    # Function to update command buttons based on selected script type
    def update_command_buttons():
        # Clear existing buttons
        for button in command_buttons_frame.winfo_children():
            button.destroy()

        # Get the selected script type
        selected_script = script_var.get()

        # Get the operations for the selected script type
        operations = config.get(selected_script, {}).keys()

        # Create buttons for each operation
        for operation in operations:
            button = tk.Button(
                command_buttons_frame,
                text=operation,
                command=lambda op=operation: select_command(op),
                width=10,
                height=1,
                bg="#D9D9D9",  # Color neutro
                fg="#333333",
                activebackground="#CCCCCC",  # Color neutro más oscuro
            )
            button.pack(side="left", padx=5)
            button.bind("<Enter>", lambda e, b=button: on_enter(b, "#E6E6E6"))  # Hover color más claro
            button.bind("<Leave>", lambda e, b=button: on_leave(b, "#D9D9D9"))

        # Visually select the default button ("Create")
        default_button = [button for button in command_buttons_frame.winfo_children() if button.cget("text") == "Create"]
        if default_button:
            select_command("Create")

    # Botones de comando (Create, Change, Show, etc.)
    def select_command(command):
        command_var.set(command)
        for button in command_buttons_frame.winfo_children():
            if button.cget("text") == command:
                button.config(relief="sunken", bg="#CCCCCC")  # Color más oscuro para botón activo
            else:
                button.config(relief="raised", bg="#D9D9D9")  # Restaurar color neutro

    # Update command buttons when script type changes
    script_var.trace_add("write", lambda *args: update_command_buttons())

    # Initial update of command buttons
    update_command_buttons()

    # Frame para los botones de Excel (Open Excel y Clear Excel)
    excel_buttons_frame = tk.Frame(root, bg="#F5F5F5")
    excel_buttons_frame.pack(pady=(20, 10))

    # Botón para abrir Excel
    open_button = tk.Button(
        excel_buttons_frame,
        text="Open Excel",
        font=("Arial", 11, "bold"),
        command=lambda: open_excel(script_var.get(), command_var.get()),
        width=18,
        height=2,
        bg="#6CC24A",  # Verde suave
        fg="white",
        activebackground="#5EAC3C",
    )
    open_button.pack(side="left", padx=5)
    # Efecto hover para Open Excel
    open_button.bind("<Enter>", lambda e: on_enter(open_button, "#5EAC3C"))
    open_button.bind("<Leave>", lambda e: on_leave(open_button, "#6CC24A"))

    # Botón para borrar el archivo Excel
    clear_button = tk.Button(
        excel_buttons_frame,
        text="Clear Excel",
        font=("Arial", 11, "bold"),
        command=lambda: clear_excel(script_var.get()),
        width=18,
        height=2,
        bg="#FF6B6B",  # Rojo suave
        fg="white",
        activebackground="#E65A5A",
    )
    clear_button.pack(side="left", padx=5)
    # Efecto hover para Clear Excel
    clear_button.bind("<Enter>", lambda e: on_enter(clear_button, "#E65A5A"))
    clear_button.bind("<Leave>", lambda e: on_leave(clear_button, "#FF6B6B"))

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

    return script_menu, open_button, run_button, clear_button