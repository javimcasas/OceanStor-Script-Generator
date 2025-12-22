import tkinter as tk
from tkinter import ttk, font as tkfont

def apply_style(root):
    """Apply modern styling to the application"""
    # Configure ttk styles
    style = ttk.Style()
    style.theme_use('clam')
    
    # Configure colors
    colors = {
        'primary': '#1a237e',
        'secondary': '#5c6bc0',
        'success': '#4caf50',
        'danger': '#f44336',
        'warning': '#ff9800',
        'info': '#2196f3',
        'light': '#f8f9fa',
        'dark': '#212529',
        'white': '#ffffff',
        'gray': '#6c757d',
        'light_gray': '#e9ecef'
    }
    
    # Configure combobox style - FIXED FOCUS ISSUE
    style.configure("TCombobox",
                    font=("Segoe UI", 10),
                    padding=8,
                    relief="flat",
                    fieldbackground=colors['white'],
                    background=colors['white'],
                    borderwidth=1,
                    focusthickness=0,  # Set to 0 to remove focus border
                    focuscolor='')  # Empty string removes focus color
    
    style.map('TCombobox',
              fieldbackground=[('readonly', colors['white'])],
              selectbackground=[('readonly', colors['light_gray'])],
              selectforeground=[('readonly', colors['dark'])],
              highlightcolor=[('focus', '')],  # Remove highlight on focus
              highlightbackground=[('focus', '')])  # Remove highlight background
    
    # Configure label frame to prevent focus border
    style.configure("TLabelframe",
                    background=colors['white'],
                    foreground=colors['dark'],
                    relief="flat",
                    borderwidth=1,
                    focusthickness=0)  # Add this
    
    style.configure("TLabelframe.Label",
                    font=("Segoe UI", 11, "bold"),
                    background=colors['white'],
                    foreground=colors['dark'])
    
    # Set root background
    root.configure(bg=colors['light'])
    
    # Also configure button focus
    style.configure("TButton",
                    focusthickness=0,
                    focuscolor='')
    
    return style, colors  # Return these for use elsewhere

def create_modern_combobox(parent, var, values, label_text, tooltip_text=""):
    """Create a modern combobox with label"""
    frame = tk.Frame(parent, bg="white")
    frame.pack(fill="x", padx=20, pady=10)
    
    label = tk.Label(frame, text=label_text, font=("Segoe UI", 10), 
                    bg="white", fg="#546e7a", anchor="w")
    label.pack(fill="x")
    
    combobox = ttk.Combobox(
        frame,
        textvariable=var,
        values=values,
        state="readonly",
        font=("Segoe UI", 10)
    )
    combobox.pack(fill="x", pady=(5, 0))
    
    return combobox

def create_section_header(parent, text):
    """Create a section header"""
    header = tk.Label(parent, text=text, font=("Segoe UI", 12, "bold"),
                     bg=parent.cget('bg'), fg="#37474f", anchor="w")
    header.pack(fill="x", padx=20, pady=(15, 10))
    return header

def create_card(parent, title=""):
    """Create a card container"""
    card = tk.Frame(parent, bg="white", relief="flat", bd=1,
                   highlightbackground="#e0e0e0", highlightthickness=1)
    
    if title:
        title_label = tk.Label(card, text=title, font=("Segoe UI", 11, "bold"),
                              bg="white", fg="#37474f", anchor="w")
        title_label.pack(fill="x", padx=20, pady=(15, 10))
    
    return card

def create_modern_button(parent, text, command, color="#2196f3", icon=None):
    """Create a modern styled button"""
    btn = tk.Button(
        parent,
        text=text,
        command=command,
        bg=color,
        fg="white",
        activebackground=color,
        activeforeground="white",
        relief="flat",
        bd=0,
        font=("Segoe UI", 10, "bold"),
        height=2,
        cursor="hand2",
        padx=20
    )
    
    # Add hover effect
    def on_enter(e):
        btn['bg'] = darken_color(color)
    
    def on_leave(e):
        btn['bg'] = color
    
    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)
    
    return btn

def darken_color(hex_color, factor=0.8):
    """Darken a hex color"""
    if hex_color.startswith('#'):
        hex_color = hex_color[1:]
    
    rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    darker = tuple(max(0, int(c * factor)) for c in rgb)
    return f"#{darker[0]:02x}{darker[1]:02x}{darker[2]:02x}"

def create_command_button(parent, text, is_selected=False, command=None):
    """Create a command selection button"""
    bg_color = "#2196f3" if is_selected else "#f5f5f5"
    fg_color = "white" if is_selected else "#37474f"
    
    btn = tk.Button(
        parent,
        text=text,
        width=14,
        height=2,
        bg=bg_color,
        fg=fg_color,
        activebackground="#e3f2fd",
        activeforeground="#37474f",
        relief="flat",
        bd=0,
        font=("Segoe UI", 10),
        cursor="hand2",
        command=command
    )
    
    # Add hover effect for non-selected buttons
    if not is_selected:
        def on_enter(e):
            btn['bg'] = "#e3f2fd"
        
        def on_leave(e):
            btn['bg'] = bg_color
        
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
    
    return btn

# Add this function after the imports and before run_script
def toggle_loading(root, state, message="Processing..."):
    """Modern styled loading indicator"""
    global loading_window, loading_label, loading_progress
    
    if state:  # Show loading window
        if not hasattr(toggle_loading, 'loading_window') or toggle_loading.loading_window is None:
            # Create loading overlay
            toggle_loading.loading_window = tk.Toplevel(root)
            toggle_loading.loading_window.title("Processing")
            toggle_loading.loading_window.overrideredirect(True)
            toggle_loading.loading_window.attributes('-alpha', 0.95)
            
            # Center window
            root.update_idletasks()
            root_x = root.winfo_x()
            root_y = root.winfo_y()
            root_width = root.winfo_width()
            root_height = root.winfo_height()
            
            width, height = 400, 150
            x = root_x + (root_width - width) // 2
            y = root_y + (root_height - height) // 2
            
            toggle_loading.loading_window.geometry(f"{width}x{height}+{x}+{y}")
            toggle_loading.loading_window.configure(bg="#ffffff", 
                                                   highlightbackground="#2196f3", 
                                                   highlightthickness=2)
            
            # Create content
            content_frame = tk.Frame(toggle_loading.loading_window, bg="#ffffff")
            content_frame.pack(expand=True, fill="both", padx=30, pady=30)
            
            # Title
            title_label = tk.Label(content_frame, text="Processing", 
                                  font=("Segoe UI", 14, "bold"),
                                  bg="#ffffff", fg="#2196f3")
            title_label.pack(pady=(0, 15))
            
            # Message
            toggle_loading.loading_label = tk.Label(content_frame, text=message,
                                                   font=("Segoe UI", 10),
                                                   bg="#ffffff", fg="#546e7a")
            toggle_loading.loading_label.pack(pady=(0, 20))
            
            # Progress bar
            style = ttk.Style()
            style.theme_use('clam')
            style.configure("Modern.Horizontal.TProgressbar",
                           troughcolor="#f0f0f0",
                           background="#2196f3",
                           borderwidth=0)
            
            toggle_loading.loading_progress = ttk.Progressbar(
                content_frame,
                mode="indeterminate",
                length=300,
                style="Modern.Horizontal.TProgressbar"
            )
            toggle_loading.loading_progress.pack()
            toggle_loading.loading_progress.start(10)
            
            toggle_loading.loading_window.update_idletasks()
            
        else:
            # Update existing window
            toggle_loading.loading_label.config(text=message)
            toggle_loading.loading_window.update_idletasks()

    elif hasattr(toggle_loading, 'loading_window') and toggle_loading.loading_window:  # Hide loading window
        toggle_loading.loading_progress.stop()
        toggle_loading.loading_window.destroy()
        toggle_loading.loading_window = None
        root.update_idletasks()