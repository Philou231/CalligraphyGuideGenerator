"""
Calligraphy Guide Sheet Generator
---------------------------------
A Windows-optimized Desktop Application to design and generate custom
Calligraphy guide sheets in PostScript (.ps) format.

Features:
- High-DPI Awareness for crisp rendering on 4K/Windows displays.
- Dynamic Object-Oriented Guidelines Engine (arbitrary horizontal lines & slants).
- Live vector preview using Tkinter Canvas.
- State persistence via embedded JSON "Ghost Metadata" in PostScript files.
- Direct Windows printing.

Author: Gemini Pro
  Date: 2026-03-01
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import math
import json
import os
import re
import ctypes

# -----------------------------------------------------------------------------
# Configuration & Defaults
# -----------------------------------------------------------------------------
CONFIG = {
    "page_width_mm": 210.0,       # A4 Width
    "page_height_mm": 297.0,      # A4 Height
    "margin_mm": 15.0,            # Page margins
    "mm_to_pts": 2.83465,         # PostScript points conversion factor
    
    # Defaults (Startup State: Italic Preset)
    "default_pen_width": 2.0,
    "default_group_gap": 8.0,
    "default_lines": "Ascender: 7\nX-Height: 5\nBase: 0\nDescender: -5",
    "default_slants": "10: 15",   # Angle: Spacing_mm
    
    # UI Styling
    "bg_color": "#242424",        # customtkinter default dark bg
    "page_color": "#FFFFFF",
    "line_color": "#A0A0A0",
    "base_color": "#000000",
    "slant_color": "#D3D3D3",
}

# -----------------------------------------------------------------------------
# High-DPI Awareness Initialization
# -----------------------------------------------------------------------------
try:
    # Windows 8.1 and later
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    # Fallback for older Windows versions or non-Windows OS
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


# -----------------------------------------------------------------------------
# Main Application Class
# -----------------------------------------------------------------------------
class CalligraphyApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # App Configuration
        self.title("Calligraphy Guide Sheet Generator")
        self.geometry("1100x800")
        self.minsize(900, 600)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # Internal State Tracker for Debouncing
        self._update_job = None
        
        # Build UI
        self._setup_layout()
        self._build_sidebar()
        self._build_canvas()
        
        # Initialize default state
        self._set_ui_state({
            "pen_width": str(CONFIG["default_pen_width"]),
            "group_gap": str(CONFIG["default_group_gap"]),
            "lines": CONFIG["default_lines"],
            "slants": CONFIG["default_slants"]
        })
        
        # Bind resize event to re-render preview
        self.bind("<Configure>", self._on_resize)
        
        # Initial Render
        self.after(200, self.update_preview)

    # --- UI Setup ---

    def _setup_layout(self):
        """Configures the main grid layout."""
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.sidebar_frame = ctk.CTkFrame(self, width=300, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(6, weight=1)
        
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=CONFIG["bg_color"])
        self.main_frame.grid(row=0, column=1, sticky="nsew")

    def _build_sidebar(self):
        """Constructs the sidebar controls for the guidelines engine."""
        # Header
        lbl_title = ctk.CTkLabel(self.sidebar_frame, text="Guide Parameters", font=ctk.CTkFont(size=20, weight="bold"))
        lbl_title.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        
        # Pen Width & Gap
        frame_metrics = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        frame_metrics.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        
        ctk.CTkLabel(frame_metrics, text="Pen Width (mm):").grid(row=0, column=0, sticky="w", pady=5)
        self.ent_pen_width = ctk.CTkEntry(frame_metrics, width=80)
        self.ent_pen_width.grid(row=0, column=1, sticky="e", pady=5)
        self.ent_pen_width.bind("<KeyRelease>", self._debounce_update)
        
        ctk.CTkLabel(frame_metrics, text="Group Gap (mm):").grid(row=1, column=0, sticky="w", pady=5)
        self.ent_group_gap = ctk.CTkEntry(frame_metrics, width=80)
        self.ent_group_gap.grid(row=1, column=1, sticky="e", pady=5)
        self.ent_group_gap.bind("<KeyRelease>", self._debounce_update)
        
        # Horizontal Lines Configuration
        ctk.CTkLabel(self.sidebar_frame, text="Horizontal Lines (Name: PenWidths)", font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, padx=20, pady=(10, 0), sticky="w")
        self.txt_lines = ctk.CTkTextbox(self.sidebar_frame, height=120)
        self.txt_lines.grid(row=3, column=0, padx=20, pady=5, sticky="ew")
        self.txt_lines.bind("<KeyRelease>", self._debounce_update)
        
        # Slants Configuration
        ctk.CTkLabel(self.sidebar_frame, text="Slant Overlays (Angle° : Spacing mm)", font=ctk.CTkFont(weight="bold")).grid(row=4, column=0, padx=20, pady=(10, 0), sticky="w")
        self.txt_slants = ctk.CTkTextbox(self.sidebar_frame, height=80)
        self.txt_slants.grid(row=5, column=0, padx=20, pady=5, sticky="ew")
        self.txt_slants.bind("<KeyRelease>", self._debounce_update)
        
        # Actions
        frame_actions = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        frame_actions.grid(row=7, column=0, padx=20, pady=20, sticky="ew")
        frame_actions.grid_columnconfigure(0, weight=1)
        
        btn_save = ctk.CTkButton(frame_actions, text="Save PostScript (.ps)", command=self.save_postscript, fg_color="#1E90FF", hover_color="#104E8B")
        btn_save.grid(row=0, column=0, pady=5, sticky="ew")
        
        btn_load = ctk.CTkButton(frame_actions, text="Load Template", command=self.load_postscript, fg_color="#555555", hover_color="#333333")
        btn_load.grid(row=1, column=0, pady=5, sticky="ew")
        
        btn_print = ctk.CTkButton(frame_actions, text="Print Now", command=self.print_postscript, fg_color="#2E8B57", hover_color="#1E5C3A")
        btn_print.grid(row=2, column=0, pady=(20, 5), sticky="ew")

    def _build_canvas(self):
        """Initializes the live vector preview canvas."""
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)
        
        self.canvas = tk.Canvas(self.main_frame, bg=CONFIG["bg_color"], highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

    # --- Data Parsing & State Management ---

    def _get_ui_state(self):
        """Returns the current UI parameters as a dictionary."""
        return {
            "pen_width": self.ent_pen_width.get().strip(),
            "group_gap": self.ent_group_gap.get().strip(),
            "lines": self.txt_lines.get("1.0", tk.END).strip(),
            "slants": self.txt_slants.get("1.0", tk.END).strip()
        }

    def _set_ui_state(self, state):
        """Populates the UI fields with the provided dictionary state."""
        self.ent_pen_width.delete(0, tk.END)
        self.ent_pen_width.insert(0, state.get("pen_width", ""))
        
        self.ent_group_gap.delete(0, tk.END)
        self.ent_group_gap.insert(0, state.get("group_gap", ""))
        
        self.txt_lines.delete("1.0", tk.END)
        self.txt_lines.insert("1.0", state.get("lines", ""))
        
        self.txt_slants.delete("1.0", tk.END)
        self.txt_slants.insert("1.0", state.get("slants", ""))
        
        self.update_preview()

    def _parse_inputs(self):
        """Parses UI inputs into strongly typed variables for rendering."""
        try:
            pen_width = float(self.ent_pen_width.get())
            group_gap = float(self.ent_group_gap.get())
        except ValueError:
            return None # Invalid numeric input
            
        lines_data = []
        for line in self.txt_lines.get("1.0", tk.END).split('\n'):
            if ':' in line:
                try:
                    name, pw = line.split(':')
                    lines_data.append((name.strip(), float(pw.strip())))
                except ValueError:
                    continue
                    
        slants_data = []
        for line in self.txt_slants.get("1.0", tk.END).split('\n'):
            if ':' in line:
                try:
                    angle, spacing = line.split(':')
                    slants_data.append((float(angle.strip()), float(spacing.strip())))
                except ValueError:
                    continue
                    
        return pen_width, group_gap, lines_data, slants_data

    # --- Live Preview Engine ---

    def _debounce_update(self, event=None):
        """Delays the update_preview call to avoid stuttering while typing."""
        if self._update_job is not None:
            self.after_cancel(self._update_job)
        self._update_job = self.after(300, self.update_preview)

    def _on_resize(self, event):
        """Handles canvas resizing efficiently."""
        if event.widget == self:
            self._debounce_update()

    def update_preview(self):
        """Renders the simplified vector view of the guidelines on the Tkinter Canvas."""
        self.canvas.delete("all")
        
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw <= 1 or ch <= 1:
            return
            
        parsed = self._parse_inputs()
        if not parsed:
            return
        pen_width, group_gap, lines_data, slants_data = parsed
        
        # Calculate Canvas Scaling (Fit A4 inside canvas with padding)
        pad = 20
        scale_h = (ch - pad*2) / CONFIG["page_height_mm"]
        scale_w = (cw - pad*2) / CONFIG["page_width_mm"]
        scale = min(scale_h, scale_w)
        
        offset_x = (cw - (CONFIG["page_width_mm"] * scale)) / 2
        offset_y = (ch - (CONFIG["page_height_mm"] * scale)) / 2
        
        # Coordinate mapping helper
        def map_coord(x_mm, y_mm):
            return offset_x + (x_mm * scale), offset_y + (y_mm * scale)

        # Draw Page
        p_x1, p_y1 = map_coord(0, 0)
        p_x2, p_y2 = map_coord(CONFIG["page_width_mm"], CONFIG["page_height_mm"])
        self.canvas.create_rectangle(p_x1, p_y1, p_x2, p_y2, fill=CONFIG["page_color"], outline="#111111")
        
        margin = CONFIG["margin_mm"]
        
        # Render Slants (Underneath horizontal lines)
        for angle_deg, spacing in slants_data:
            if spacing <= 0: continue
            rad = math.radians(angle_deg)
            # Distance horizontally to maintain perpendicular spacing
            dx_step = spacing / math.cos(rad) if math.cos(rad) != 0 else spacing
            
            x_start = -CONFIG["page_height_mm"] * math.tan(abs(rad))
            x_end = CONFIG["page_width_mm"] + CONFIG["page_height_mm"] * math.tan(abs(rad))
            
            x_curr = x_start
            while x_curr <= x_end:
                y_bottom = CONFIG["page_height_mm"] - margin
                x_bottom = x_curr
                y_top = margin
                x_top = x_curr + (CONFIG["page_height_mm"] - 2*margin) * math.tan(rad)
                
                pt1 = map_coord(x_bottom, y_bottom)
                pt2 = map_coord(x_top, y_top)
                
                self.canvas.create_line(*pt1, *pt2, fill=CONFIG["slant_color"], width=1)
                x_curr += dx_step

        # Render Horizontal Lines
        if lines_data:
            pw_values = [pw for _, pw in lines_data]
            max_pw = max(pw_values)
            min_pw = min(pw_values)
            
            current_y = margin + (max_pw * pen_width)
            
            while current_y + (abs(min_pw) * pen_width) <= CONFIG["page_height_mm"] - margin:
                for name, pw in lines_data:
                    y_line = current_y - (pw * pen_width)
                    x1, y1 = map_coord(margin, y_line)
                    x2, y2 = map_coord(CONFIG["page_width_mm"] - margin, y_line)
                    
                    color = CONFIG["base_color"] if name.lower() == 'base' else CONFIG["line_color"]
                    dash = () if name.lower() == 'base' else (4, 4)
                    width = 2 if name.lower() == 'base' else 1
                    
                    self.canvas.create_line(x1, y1, x2, y2, fill=color, dash=dash, width=width)
                
                current_y += ((max_pw - min_pw) * pen_width) + group_gap

        # Create clipping mask over the edges using thick rectangles
        mask_color = CONFIG["bg_color"]
        m_x1, m_y1 = map_coord(margin, margin)
        m_x2, m_y2 = map_coord(CONFIG["page_width_mm"] - margin, CONFIG["page_height_mm"] - margin)
        
        # Mask Top, Bottom, Left, Right to hide slant spillover outside margins
        self.canvas.create_rectangle(0, 0, cw, m_y1, fill=mask_color, outline="")
        self.canvas.create_rectangle(0, m_y2, cw, ch, fill=mask_color, outline="")
        self.canvas.create_rectangle(0, 0, m_x1, ch, fill=mask_color, outline="")
        self.canvas.create_rectangle(m_x2, 0, cw, ch, fill=mask_color, outline="")


    # --- PostScript Generation & IO ---

    def generate_postscript_string(self):
        """
        Generates standard PostScript code.
        Translates metric calculations to PS points, handling the bottom-left PS origin.
        Injects JSON representation of the GUI state as Ghost Metadata.
        """
        parsed = self._parse_inputs()
        if not parsed:
            raise ValueError("Invalid parameters.")
        pen_width, group_gap, lines_data, slants_data = parsed
        
        state_json = json.dumps(self._get_ui_state())
        
        width_pts = CONFIG["page_width_mm"] * CONFIG["mm_to_pts"]
        height_pts = CONFIG["page_height_mm"] * CONFIG["mm_to_pts"]
        
        ps_lines = [
            "%!PS-Adobe-3.0",
            f"%%BoundingBox: 0 0 {int(width_pts)} {int(height_pts)}",
            "%%Creator: Python Calligraphy Guide Generator",
            "%%EndComments",
            "% BEGIN_METADATA",
            f"% {state_json}",
            "% END_METADATA",
            "",
            "/mm {2.83465 mul} def",
            ""
        ]

        # Draw Slants
        margin = CONFIG["margin_mm"]
        if slants_data:
            # Set up clipping path for margins so slants don't draw to edge of paper
            ps_lines.extend([
                "gsave",
                "newpath",
                f"{margin} mm {margin} mm moveto",
                f"{CONFIG['page_width_mm'] - margin} mm {margin} mm lineto",
                f"{CONFIG['page_width_mm'] - margin} mm {CONFIG['page_height_mm'] - margin} mm lineto",
                f"{margin} mm {CONFIG['page_height_mm'] - margin} mm lineto",
                "closepath clip",
                "0.8 setgray",
                "0.3 setlinewidth"
            ])
            
            for angle_deg, spacing in slants_data:
                if spacing <= 0: continue
                rad = math.radians(angle_deg)
                dx_step = spacing / math.cos(rad) if math.cos(rad) != 0 else spacing
                
                x_start = -CONFIG["page_height_mm"] * math.tan(abs(rad))
                x_end = CONFIG["page_width_mm"] + CONFIG["page_height_mm"] * math.tan(abs(rad))
                
                x_curr = x_start
                while x_curr <= x_end:
                    y_bottom_mm = margin
                    x_bottom_mm = x_curr
                    y_top_mm = CONFIG["page_height_mm"] - margin
                    x_top_mm = x_curr + (CONFIG["page_height_mm"] - 2*margin) * math.tan(rad)
                    
                    ps_lines.extend([
                        "newpath",
                        f"{x_bottom_mm:.2f} mm {y_bottom_mm:.2f} mm moveto",
                        f"{x_top_mm:.2f} mm {y_top_mm:.2f} mm lineto",
                        "stroke"
                    ])
                    x_curr += dx_step
            ps_lines.append("grestore\n")

        # Draw Horizontal Lines
        if lines_data:
            pw_values = [pw for _, pw in lines_data]
            max_pw = max(pw_values)
            min_pw = min(pw_values)
            
            current_y = margin + (max_pw * pen_width)
            
            while current_y + (abs(min_pw) * pen_width) <= CONFIG["page_height_mm"] - margin:
                for name, pw in lines_data:
                    # In Tkinter, Y is from top. In PS, Y is from bottom.
                    y_line_mm = current_y - (pw * pen_width)
                    y_ps_mm = CONFIG["page_height_mm"] - y_line_mm
                    
                    # Formatting based on line type
                    if name.lower() == 'base':
                        ps_lines.append("0 setgray 0.6 setlinewidth [] 0 setdash")
                    else:
                        ps_lines.append("0.4 setgray 0.3 setlinewidth [2 2] 0 setdash")
                    
                    ps_lines.extend([
                        "newpath",
                        f"{margin} mm {y_ps_mm:.2f} mm moveto",
                        f"{CONFIG['page_width_mm'] - margin} mm {y_ps_mm:.2f} mm lineto",
                        "stroke"
                    ])
                current_y += ((max_pw - min_pw) * pen_width) + group_gap

        ps_lines.append("showpage")
        return "\n".join(ps_lines)

    def save_postscript(self):
        """Prompts the user to save the generated PostScript file."""
        try:
            ps_code = self.generate_postscript_string()
        except ValueError:
            messagebox.showerror("Validation Error", "Please ensure all inputs are correctly formatted.")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".ps",
            filetypes=[("PostScript Files", "*.ps")],
            title="Save Guide Sheet"
        )
        if filepath:
            try:
                with open(filepath, 'w') as f:
                    f.write(ps_code)
                messagebox.showinfo("Success", f"PostScript file saved successfully to:\n{filepath}")
            except Exception as e:
                messagebox.showerror("Save Error", str(e))

    def load_postscript(self):
        """Reads a previously generated .ps file, extracting embedded JSON metadata to update UI."""
        filepath = filedialog.askopenfilename(
            filetypes=[("PostScript Files", "*.ps")],
            title="Load Template"
        )
        if filepath:
            try:
                with open(filepath, 'r') as f:
                    content = f.read()
                
                # Regex search for the custom metadata block
                match = re.search(r"% BEGIN_METADATA\n% (.*?)\n% END_METADATA", content)
                if match:
                    state_json = match.group(1)
                    state = json.loads(state_json)
                    self._set_ui_state(state)
                    messagebox.showinfo("Success", "Template loaded successfully.")
                else:
                    messagebox.showwarning("Warning", "No template metadata found in this PostScript file.")
            except Exception as e:
                messagebox.showerror("Load Error", f"Failed to load file:\n{e}")

    def print_postscript(self):
        """
        Sends the file directly to the default Windows printer.
        Relies on system associations for the 'print' shell verb.
        """
        try:
            ps_code = self.generate_postscript_string()
        except ValueError:
            messagebox.showerror("Validation Error", "Please ensure all inputs are correctly formatted.")
            return
            
        temp_path = os.path.join(os.environ.get("TEMP", "."), "calligraphy_temp_print.ps")
        try:
            with open(temp_path, 'w') as f:
                f.write(ps_code)
            
            # Use Windows native shell execute to print
            os.startfile(temp_path, "print")
        except Exception as e:
            messagebox.showerror("Print Error", f"Could not send to printer.\nEnsure a PostScript driver (e.g., Ghostscript) is installed and associated with .ps files.\n\nError: {e}")

# -----------------------------------------------------------------------------
# Bootstrapper
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    app = CalligraphyApp()
    app.mainloop()