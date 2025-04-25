import json
import os
import tkinter as tk

class ToolTipManager:
    def __init__(self):
        self.tooltips = self._load_tooltips()
        
    def _load_tooltips(self):
        """Load tooltips from JSON file"""
        try:
            with open(os.path.join(os.path.dirname(__file__), 'tooltips.json')) as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return {"buttons": {}, "selectors": {}}
    
    def add_tooltip(self, widget, key, category="buttons"):
        """Add tooltip to widget if defined in config"""
        if key in self.tooltips.get(category, {}):
            _ToolTip(widget, self.tooltips[category][key])

class _ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, _):
        """Display tooltip"""
        if self.tip_window or not self.text:
            return
            
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(tw, text=self.text, justify='left',
                        background="#ffffe0", relief='solid', borderwidth=1,
                        font=("Arial", 10))
        label.pack(ipadx=1)

    def hide_tip(self, _):
        """Hide tooltip"""
        if self.tip_window:
            self.tip_window.destroy()
        self.tip_window = None