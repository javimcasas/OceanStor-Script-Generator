import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk
from excel_operations import open_excel, clear_excel, import_excel
from gui_helpers import apply_style, darken_color, toggle_loading
from command_generator import main as generate_commands
from utils import load_config
from tooltip_manager import ToolTipManager
import webbrowser

tooltip_manager = ToolTipManager()

def run_script(script_type, command_type, device_type):
    try:
        config = load_config(device_type)
        if script_type not in config:
            raise ValueError(f"Invalid script type: {script_type}. Valid types are: {list(config.keys())}")

        documents_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Documents")
        excel_path = os.path.join(documents_folder, f'{device_type}_{script_type}_commands.xlsx')

        if not os.path.exists(excel_path):
            messagebox.showinfo("Info", f"Please create the Excel file first by pressing the 'Create Excel' button.")
            return

        # Show loading indicator
        toggle_loading(root, True, f"Generating {command_type} commands...")
        root.update()  # Force UI update
        
        # Call the command generator
        output_file_path = generate_commands(script_type, device_type)
        
        # Hide loading indicator
        toggle_loading(root, False)

        if os.path.exists(output_file_path):
            messagebox.showinfo("Success", f"Commands generated successfully!\nOutput saved to:\n{output_file_path}")
            os.startfile(output_file_path)
        else:
            messagebox.showerror("Error", f"Output file not found at:\n{output_file_path}")
    except Exception as e:
        toggle_loading(root, False)
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

def main():
    global root
    root = tk.Tk()
    root.title("OceanStor Command Generator")
    
    try:
        icon_path = os.path.join("Icons", "exe_icon.ico")
        root.iconbitmap(icon_path)
    except Exception as e:
        try:
            base_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.abspath(".")
            icon_path = os.path.join(base_path, "Icons", "exe_icon.ico")
            root.iconbitmap(icon_path)
        except Exception:
            pass

    # Window sizing and positioning
    window_width = 680
    window_height = 640
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_x = (screen_width - window_width) // 2
    position_y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
    root.resizable(False, False)

    # Apply modern style
    apply_style(root)

    # Create main container
    main_container = tk.Frame(root, bg="#f8f9fa")
    main_container.pack(fill="both", expand=True, padx=20, pady=20)

    # Header section
    header_frame = tk.Frame(main_container, bg="#ffffff", height=110)
    header_frame.pack(fill="x", pady=(0, 20))
    header_frame.pack_propagate(False)
    
    # Logo and title
    logo_frame = tk.Frame(header_frame, bg="#ffffff")
    logo_frame.pack(side="left", padx=20, pady=(15, 15))
    
    try:
        logo_path = os.path.join("Icons", "huawei_logo.png")
        logo_img = tk.PhotoImage(file=logo_path)
        logo_label = tk.Label(logo_frame, image=logo_img, bg="#ffffff")
        logo_label.image = logo_img
        logo_label.pack(side="left", padx=(0, 15))
    except:
        # Fallback if no logo
        pass
    
    title_frame = tk.Frame(logo_frame, bg="#ffffff")
    title_frame.pack(side="left")
    
    title_label = tk.Label(title_frame, text="OceanStor", font=("Segoe UI", 24, "bold"), 
                          bg="#ffffff", fg="#1a237e")
    title_label.pack(anchor="w")
    
    subtitle_label = tk.Label(title_frame, text="Command Generator", font=("Segoe UI", 12), 
                            bg="#ffffff", fg="#5c6bc0")
    subtitle_label.pack(anchor="w", pady=(0, 5))

    # Configuration section
    config_frame = tk.LabelFrame(main_container, text="Configuration", 
                                font=("Segoe UI", 12, "bold"),
                                bg="#ffffff", fg="#37474f",
                                relief="flat", bd=1,
                                highlightbackground="#e0e0e0", highlightthickness=1)
    config_frame.pack(fill="x", pady=(0, 20))

    # Device selection
    device_var = tk.StringVar(value="OceanStor Dorado")
    script_var = tk.StringVar()
    command_var = tk.StringVar()

    device_label = tk.Label(config_frame, text="Device Type:", font=("Segoe UI", 10), 
                           bg="#ffffff", fg="#546e7a")
    device_label.grid(row=0, column=0, padx=20, pady=15, sticky="w")

    device_menu = ttk.Combobox(
        config_frame,
        textvariable=device_var,
        values=["OceanStor Dorado", "OceanStor Pacific"],
        state="readonly",
        font=("Segoe UI", 10),
        width=25
    )
    device_menu.grid(row=0, column=1, padx=20, pady=15, sticky="ew")
    config_frame.columnconfigure(1, weight=1)
    tooltip_manager.add_tooltip(device_menu, "device", "selectors")

    # Script selection
    script_label = tk.Label(config_frame, text="Resource Type:", font=("Segoe UI", 10), 
                           bg="#ffffff", fg="#546e7a")
    script_label.grid(row=1, column=0, padx=20, pady=15, sticky="w")

    # Load initial config
    def load_current_config():
        return load_config(device_var.get())

    config = load_current_config()
    
    script_menu = ttk.Combobox(
        config_frame,
        textvariable=script_var,
        values=list(config.keys()),
        state="readonly",
        font=("Segoe UI", 10),
        width=25
    )
    script_menu.grid(row=1, column=1, padx=20, pady=15, sticky="ew")
    tooltip_manager.add_tooltip(script_menu, "script", "selectors")
    
    # Set initial script value
    if config:
        script_var.set(list(config.keys())[0])

    # Command selection section
    command_frame = tk.LabelFrame(main_container, text="Command Type", 
                                 font=("Segoe UI", 12, "bold"),
                                 bg="#ffffff", fg="#37474f",
                                 relief="flat", bd=1,
                                 highlightbackground="#e0e0e0", highlightthickness=1)
    command_frame.pack(fill="x", pady=(0, 25))

    # Create command buttons container
    command_buttons_frame = tk.Frame(command_frame, bg="#ffffff")
    command_buttons_frame.pack(padx=20, pady=20)

    def update_command_buttons(config):
        """Update command buttons based on current script selection"""
        for widget in command_buttons_frame.winfo_children():
            widget.destroy()
        
        current_script = script_var.get()
        if current_script in config:
            operations = config[current_script]['operations'].keys()
            max_buttons_per_row = 4
            row = 0
            col = 0
            
            for i, operation in enumerate(operations):
                if col == 0:
                    button_row = tk.Frame(command_buttons_frame, bg="#ffffff")
                    button_row.pack(pady=(0 if i == 0 else 10))
                
                btn = tk.Button(
                    button_row,
                    text=operation,
                    width=14,
                    height=2,
                    bg="#f5f5f5",
                    fg="#37474f",
                    activebackground="#e0e0e0",
                    relief="flat",
                    bd=0,
                    font=("Segoe UI", 10),
                    cursor="hand2"
                )
                btn.config(
                    command=lambda op=operation: select_command(op, command_var, command_buttons_frame),
                    **button_style(operation == command_var.get())
                )
                btn.pack(side="left", padx=5)
                
                # Setup hover effects
                btn.bind("<Enter>", lambda e, b=btn: b.config(bg="#e3f2fd"))
                btn.bind("<Leave>", lambda e, b=btn: b.config(bg="#f5f5f5" if btn.cget("text") != command_var.get() else "#2196f3"))
                
                tooltip_manager.add_tooltip(btn, operation)
                
                col += 1
                if col >= max_buttons_per_row:
                    col = 0
                    row += 1

    def button_style(active):
        """Return style properties based on active state"""
        return {
            'bg': "#2196f3" if active else "#f5f5f5",
            'fg': "white" if active else "#37474f"
        }

    def select_command(command, command_var, parent_frame):
        """Handle command button selection"""
        command_var.set(command)
        for row_frame in parent_frame.winfo_children():
            for btn in row_frame.winfo_children():
                btn.config(**button_style(btn.cget("text") == command))

    # Initialize command buttons
    if config:
        update_command_buttons(config)

    # Action buttons section
    action_frame = tk.Frame(main_container, bg="#f8f9fa")
    action_frame.pack(fill="x", pady=(20, 0))  # Increased top padding

    # Create a container for buttons to ensure proper alignment
    buttons_container = tk.Frame(action_frame, bg="#f8f9fa")
    buttons_container.pack()

    def create_action_button(parent, text, color, command):
        """Create a modern action button with centered text and fixed width"""
        # Calculate the maximum width needed
        btn = tk.Button(
            parent,
            text=text,
            bg=color,
            fg="white",
            activebackground=color,
            activeforeground="white",
            relief="flat",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            height=2,
            cursor="hand2",
            command=command,
            anchor="center",
            justify="center"
        )
        btn.pack(side="left", padx=8, fill="x", expand=True)  # Use fill and expand for equal width
        
        # Add hover effect
        btn.bind("<Enter>", lambda e, b=btn: b.config(bg=darken_color(color)))
        btn.bind("<Leave>", lambda e, b=btn: b.config(bg=color))
        
        return btn

    # Then update the buttons_container to use grid for better control:
    buttons_container = tk.Frame(action_frame, bg="#f8f9fa")
    buttons_container.pack(fill="x", expand=True, padx=20)  # Add padding to match command box

    # Create action buttons using grid
    buttons = [
        ("📂 Open Excel", "#4caf50", lambda: open_excel(script_var.get(), command_var.get(), device_var.get())),
        ("🗑️ Clear Excel", "#f44336", lambda: clear_excel(script_var.get(), device_var.get())),
        ("📥 Import Excel", "#ff9800", lambda: import_excel(root)),
        ("🚀 Run Script", "#2196f3", lambda: run_script(script_var.get(), command_var.get(), device_var.get()))
    ]

    for i, (text, color, command) in enumerate(buttons):
        btn = tk.Button(
            buttons_container,
            text=text,
            bg=color,
            fg="white",
            activebackground=color,
            activeforeground="white",
            relief="flat",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            height=2,
            cursor="hand2",
            command=command,
            anchor="center",
            justify="center"
        )
        btn.grid(row=0, column=i, padx=4, sticky="nsew")  # Use grid with sticky
        buttons_container.grid_columnconfigure(i, weight=1, uniform="action_buttons")  # Equal columns
        
        # Add hover effect
        btn.bind("<Enter>", lambda e, b=btn, c=color: b.config(bg=darken_color(c)))
        btn.bind("<Leave>", lambda e, b=btn, c=color: b.config(bg=c))

    # Footer section
    footer_frame = tk.Frame(main_container, bg="#f8f9fa")
    footer_frame.pack(fill="x", side="bottom", pady=(20, 0))

    # Quick links
    links_frame = tk.Frame(footer_frame, bg="#f8f9fa")
    links_frame.pack(pady=(10, 0))

    def create_link(parent, text, command):
        """Create a text link with proper vertical alignment"""
        link = tk.Label(
            parent,
            text=text,
            font=("Segoe UI", 10),
            fg="#2196f3",
            bg="#f8f9fa",
            cursor="hand2",
            height=1
        )
        link.pack(side="left", padx=12, pady=2)
        
        # Bind events
        link.bind("<Button-1>", lambda e: command())
        link.bind("<Enter>", lambda e: link.config(font=("Segoe UI", 10, "underline")))
        link.bind("<Leave>", lambda e: link.config(font=("Segoe UI", 10)))
        
        return link
    
    # Update the links_frame to use proper packing:
    links_frame = tk.Frame(footer_frame, bg="#f8f9fa")
    links_frame.pack(pady=(15, 0))

    create_link(links_frame, "📁 Results", lambda: os.startfile("Results"))
    create_link(links_frame, "📁 Imported", lambda: os.startfile("Imported_Results"))
    create_link(links_frame, "📁 Logs", lambda: os.startfile("Logs"))
    create_link(links_frame, "❓ Help", lambda: webbrowser.open("https://info.support.huawei.com/storageinfo/refer/#/home"))

    # Version label
    version_label = tk.Label(footer_frame, text="Version 0.3.2", font=("Segoe UI", 9), 
                            fg="#9e9e9e", bg="#f8f9fa")
    version_label.pack(side="right", padx=20, pady=(10, 0))

    # Update UI when device changes
    def on_device_change(*args):
        new_config = load_current_config()
        script_menu['values'] = list(new_config.keys())
        if script_var.get() not in new_config:
            script_var.set(list(new_config.keys())[0] if new_config else "")
        update_command_buttons(new_config)

    device_var.trace_add("write", on_device_change)

    # Update command buttons when script changes
    def on_script_change(*args):
        update_command_buttons(load_current_config())

    script_var.trace_add("write", on_script_change)

    root.mainloop()

if __name__ == "__main__":
    main()