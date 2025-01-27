import tkinter as tk
from tkinter import ttk
from tkinter import font as tkfont
from tkinter import messagebox
import json
import webbrowser
import os

def apply_style(root):
    """Apply a professional and consistent style to the GUI."""
    root.configure(bg="#FFFFFF")  # White background
    root.option_add("*Font", "Arial 11")  # Use Arial as the default font
    root.option_add("*Foreground", "#333333")  # Dark gray text
    root.option_add("*Button.Background", "#F0F0F0")  # Light gray background for buttons
    root.option_add("*Button.Foreground", "#333333")  # Dark gray text for buttons
    root.option_add("*Button.Relief", "flat")  # Flat button style
    root.option_add("*Button.BorderWidth", 1)  # Thin border
    root.option_add("*Button.Padding", 8)  # Padding for buttons

    # Custom font for the title
    title_font = tkfont.Font(family="Arial", size=14, weight="bold")  # Use Arial for the title
    label = tk.Label(
        root,
        text="Select a script to execute or manage configuration:",
        font=title_font,
        bg="#FFFFFF",
        fg="#333333",
        pady=10,
    )
    label.pack(pady=(20, 10))

def create_buttons_and_dropdown(root, open_excel, run_script, script_var, command_var, clear_excel, config):
    """Create buttons and dropdowns for the GUI with a professional look."""
    style = ttk.Style()
    style.theme_use("clam")
    
    # Style for the Combobox
    style.configure("TCombobox",
                    fieldbackground="#FFFFFF",
                    background="#FFFFFF",
                    foreground="#333333",
                    padding=5,
                    bordercolor="#CCCCCC",
                    lightcolor="#CCCCCC",
                    darkcolor="#CCCCCC")
    style.map("TCombobox",
              fieldbackground=[("readonly", "#FFFFFF")],
              selectbackground=[("readonly", "#F0F0F0")],
              selectforeground=[("readonly", "#333333")])

    # Script selection dropdown
    script_menu = ttk.Combobox(
        root,
        textvariable=script_var,
        values=list(config.keys()),
        state="readonly",
        font=("Arial", 11),  # Use Arial for the dropdown
        justify="center",
    )
    script_menu.pack(pady=10)

    # Frame for command buttons
    command_buttons_frame = tk.Frame(root, bg="#FFFFFF")
    command_buttons_frame.pack(pady=(20, 20))

    def on_enter(button, hover_color):
        button.config(bg=hover_color)

    def on_leave(button, original_color):
        button.config(bg=original_color)

    def update_command_buttons():
        for button in command_buttons_frame.winfo_children():
            button.destroy()

        selected_script = script_var.get()
        operations = config.get(selected_script, {}).keys()

        for operation in operations:
            button = tk.Button(
                command_buttons_frame,
                text=operation,
                command=lambda op=operation: select_command(op),
                width=10,
                height=1,
                bg="#F0F0F0",  # Light gray background
                fg="#333333",  # Dark gray text
                activebackground="#E0E0E0",  # Slightly darker gray on click
                relief="flat",  # Flat button style
                bd=1,  # Thin border
                padx=8,  # Padding
                pady=5,  # Padding
            )
            button.pack(side="left", padx=5)
            button.bind("<Enter>", lambda e, b=button: on_enter(b, "#E0E0E0"))  # Hover effect
            button.bind("<Leave>", lambda e, b=button: on_leave(b, "#F0F0F0"))

        default_button = [button for button in command_buttons_frame.winfo_children() if button.cget("text") == "Create"]
        if default_button:
            select_command("Create")

    def select_command(command):
        command_var.set(command)
        for button in command_buttons_frame.winfo_children():
            if button.cget("text") == command:
                button.config(relief="sunken", bg="#E0E0E0")  # Slightly darker gray for selected button
            else:
                button.config(relief="flat", bg="#F0F0F0")  # Default light gray for unselected buttons

    script_var.trace_add("write", lambda *args: update_command_buttons())
    update_command_buttons()

    # Frame for Excel buttons
    excel_buttons_frame = tk.Frame(root, bg="#FFFFFF")
    excel_buttons_frame.pack(pady=(20, 10))

    # Open Excel button (Green)
    open_button = tk.Button(
        excel_buttons_frame,
        text="Open Excel",
        font=("Arial", 11, "bold"),  # Use Arial for the button
        command=lambda: open_excel(script_var.get(), command_var.get()),
        width=18,
        height=2,
        bg="#4CAF50",  # Green background
        fg="white",  # White text
        activebackground="#45A049",  # Darker green on click
        relief="flat",  # Flat button style
        bd=1,  # Thin border
        padx=8,  # Padding
        pady=5,  # Padding
    )
    open_button.pack(side="left", padx=5)
    open_button.bind("<Enter>", lambda e: on_enter(open_button, "#45A049"))  # Hover effect
    open_button.bind("<Leave>", lambda e: on_leave(open_button, "#4CAF50"))

    # Clear Excel button (Red)
    clear_button = tk.Button(
        excel_buttons_frame,
        text="Clear Excel",
        font=("Arial", 11, "bold"),  # Use Arial for the button
        command=lambda: clear_excel(script_var.get()),
        width=18,
        height=2,
        bg="#F44336",  # Red background
        fg="white",  # White text
        activebackground="#D32F2F",  # Darker red on click
        relief="flat",  # Flat button style
        bd=1,  # Thin border
        padx=8,  # Padding
        pady=5,  # Padding,
    )
    clear_button.pack(side="left", padx=5)
    clear_button.bind("<Enter>", lambda e: on_enter(clear_button, "#D32F2F"))  # Hover effect
    clear_button.bind("<Leave>", lambda e: on_leave(clear_button, "#F44336"))

    # Run Script button (Blue)
    run_button = tk.Button(
        root,
        text="Run Script",
        font=("Arial", 11, "bold"),  # Use Arial for the button
        command=lambda: run_script(script_var.get(), command_var.get()),
        width=18,
        height=2,
        bg="#2196F3",  # Blue background
        fg="white",  # White text
        activebackground="#1976D2",  # Darker blue on click
        relief="flat",  # Flat button style
        bd=1,  # Thin border
        padx=8,  # Padding
        pady=5,  # Padding,
    )
    run_button.pack(pady=(10, 20))
    run_button.bind("<Enter>", lambda e: on_enter(run_button, "#1976D2"))  # Hover effect
    run_button.bind("<Leave>", lambda e: on_leave(run_button, "#2196F3"))

    # Frame for additional buttons (Open Results Folder and Help)
    additional_buttons_frame = tk.Frame(root, bg="#FFFFFF")
    additional_buttons_frame.pack(pady=(10, 20))

    # Open Results Folder button
    def open_results_folder():
        results_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Results")
        if os.path.exists(results_folder):
            os.startfile(results_folder)
        else:
            messagebox.showinfo("Info", "Results folder does not exist.")

    open_results_label = tk.Label(
        additional_buttons_frame,
        text="Open Results Folder",
        fg="#333333",
        bg="#FFFFFF",
        font=("Arial", 11, "underline"),
        cursor="hand2",
    )
    open_results_label.pack(side="left", padx=10)
    open_results_label.bind("<Button-1>", lambda e: open_results_folder())
    open_results_label.bind("<Enter>", lambda e: open_results_label.config(fg="#000000"))
    open_results_label.bind("<Leave>", lambda e: open_results_label.config(fg="#333333"))

    # Help button
    def open_help_link():
        try:
            with open("help_links.json", "r") as file:
                help_links = json.load(file)
            selected_script = script_var.get()
            selected_command = command_var.get()
            link = help_links.get(selected_script, {}).get(selected_command)
            if link:
                webbrowser.open(link)
            else:
                messagebox.showinfo("Info", "No help link available for the selected command.")
        except FileNotFoundError:
            messagebox.showerror("Error", "Help links file not found.")
        except json.JSONDecodeError:
            messagebox.showerror("Error", "Invalid JSON format in help links file.")

    help_label = tk.Label(
        additional_buttons_frame,
        text="Help",
        fg="#333333",
        bg="#FFFFFF",
        font=("Arial", 11, "underline"),
        cursor="hand2",
    )
    help_label.pack(side="left", padx=10)
    help_label.bind("<Button-1>", lambda e: open_help_link())
    help_label.bind("<Enter>", lambda e: help_label.config(fg="#000000"))
    help_label.bind("<Leave>", lambda e: help_label.config(fg="#333333"))

    return script_menu, open_button, run_button, clear_button