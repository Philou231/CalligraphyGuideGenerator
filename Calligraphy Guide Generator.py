"""
Calligraphy Guide Sheet Generator - Pro Edition
-----------------------------------------------
Updates:
- Per-row slants starting strictly at the left margin.
- Individual line style controls (Width and Dash).
- Auto-generated 'x' height markers.
- Complete groups only, maximized to top margin.
- Native Windows Print Dialog integration via Ghostscript.
- Undo/Redo support and Zero-value crash protection.
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import math
import json
import os
import re
import ctypes
import subprocess

# -----------------------------------------------------------------------------
# Configuration & Defaults
# -----------------------------------------------------------------------------
CONFIG = {
    "page_width_mm": 210.0,       # A4 Width
    "page_height_mm": 297.0,      # A4 Height
    "mm_to_pts": 2.83465,         # PostScript points conversion factor
    
    # Startup State
    "default_margins": "15",      # Top, Bottom, Left, Right
    "default_pen_width": 2.0,
    "default_group_gap": 8.0,
    # Format: Name : PenWidths : LineWidth : DashPattern (solid or space-separated numbers)
    "default_lines": (
        "Ascender : 7 : 0.5 : 4 4\n"
        "X-Height : 5 : 0.5 : 4 4\n"
        "Base : 0 : 1.5 : solid\n"
        "Descender : -5 : 0.5 : 4 4"
    ),
    # Format: Angle : Spacing_mm : LineWidth : DashPattern
    "default_slants": "10 : 15 : 0.3 : solid",   
    
    # UI Styling
    "bg_color": "#242424",
    "page_color": "#FFFFFF",
    "line_color": "#000000",
}

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


class CalligraphyApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Calligraphy Guide Sheet Generator")
        self.geometry("1200x850")
        self.minsize(1000, 700)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self._update_job = None
        
        self._setup_layout()
        self._build_sidebar()
        self._build_canvas()
        
        self._set_ui_state({
            "margin": CONFIG["default_margins"],
            "pen_width": str(CONFIG["default_pen_width"]),
            "group_gap": str(CONFIG["default_group_gap"]),
            "lines": CONFIG["default_lines"],
            "slants": CONFIG["default_slants"]
        })
        
        self.bind("<Configure>", self._on_resize)
        self.after(200, self.update_preview)

    def _setup_layout(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.sidebar_frame = ctk.CTkScrollableFrame(self, width=350, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=CONFIG["bg_color"])
        self.main_frame.grid(row=0, column=1, sticky="nsew")

    def _build_sidebar(self):
        row_idx = 0
        
        lbl_title = ctk.CTkLabel(self.sidebar_frame, text="Parameters", font=ctk.CTkFont(size=20, weight="bold"))
        lbl_title.grid(row=row_idx, column=0, padx=20, pady=(20, 10), sticky="w"); row_idx += 1
        
        frame_metrics = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        frame_metrics.grid(row=row_idx, column=0, padx=20, pady=5, sticky="ew"); row_idx += 1
        
        # Inputs
        ctk.CTkLabel(frame_metrics, text="Margins (mm):").grid(row=0, column=0, sticky="w", pady=5)
        self.ent_margin = ctk.CTkEntry(frame_metrics, width=80)
        self.ent_margin.grid(row=0, column=1, sticky="e", pady=5)
        self.ent_margin.bind("<KeyRelease>", self._debounce_update)

        ctk.CTkLabel(frame_metrics, text="Pen Width (mm):").grid(row=1, column=0, sticky="w", pady=5)
        self.ent_pen_width = ctk.CTkEntry(frame_metrics, width=80)
        self.ent_pen_width.grid(row=1, column=1, sticky="e", pady=5)
        self.ent_pen_width.bind("<KeyRelease>", self._debounce_update)
        
        ctk.CTkLabel(frame_metrics, text="Group Gap (mm):").grid(row=2, column=0, sticky="w", pady=5)
        self.ent_group_gap = ctk.CTkEntry(frame_metrics, width=80)
        self.ent_group_gap.grid(row=2, column=1, sticky="e", pady=5)
        self.ent_group_gap.bind("<KeyRelease>", self._debounce_update)
        
        # Horizontal Lines
        ctk.CTkLabel(self.sidebar_frame, text="Horizontal Lines", font=ctk.CTkFont(weight="bold")).grid(row=row_idx, column=0, padx=20, pady=(15, 0), sticky="w"); row_idx += 1
        ctk.CTkLabel(self.sidebar_frame, text="Format: Name : Position : LineWidth : Dash (e.g. 4 4 or solid)", font=ctk.CTkFont(size=11, slant="italic")).grid(row=row_idx, column=0, padx=20, sticky="w"); row_idx += 1
        
        self.txt_lines = ctk.CTkTextbox(self.sidebar_frame, height=120)
        self.txt_lines.grid(row=row_idx, column=0, padx=20, pady=5, sticky="ew"); row_idx += 1
        self.txt_lines.bind("<KeyRelease>", self._debounce_update)
        self.txt_lines._textbox.configure(undo=True) # Enable Undo/Redo
        
        # Slants
        ctk.CTkLabel(self.sidebar_frame, text="Slant Overlays", font=ctk.CTkFont(weight="bold")).grid(row=row_idx, column=0, padx=20, pady=(15, 0), sticky="w"); row_idx += 1
        ctk.CTkLabel(self.sidebar_frame, text="Format: Angle° : Spacing : LineWidth : Dash", font=ctk.CTkFont(size=11, slant="italic")).grid(row=row_idx, column=0, padx=20, sticky="w"); row_idx += 1
        
        self.txt_slants = ctk.CTkTextbox(self.sidebar_frame, height=80)
        self.txt_slants.grid(row=row_idx, column=0, padx=20, pady=5, sticky="ew"); row_idx += 1
        self.txt_slants.bind("<KeyRelease>", self._debounce_update)
        self.txt_slants._textbox.configure(undo=True) # Enable Undo/Redo
        
        # Actions
        frame_actions = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        frame_actions.grid(row=row_idx, column=0, padx=20, pady=30, sticky="ew"); row_idx += 1
        frame_actions.grid_columnconfigure(0, weight=1)
        
        ctk.CTkButton(frame_actions, text="Save PostScript (.ps)", command=self.save_postscript, fg_color="#1E90FF", hover_color="#104E8B").grid(row=0, column=0, pady=5, sticky="ew")
        ctk.CTkButton(frame_actions, text="Load Template", command=self.load_postscript, fg_color="#555555", hover_color="#333333").grid(row=1, column=0, pady=5, sticky="ew")
        ctk.CTkButton(frame_actions, text="Print Now...", command=self.print_postscript, fg_color="#2E8B57", hover_color="#1E5C3A").grid(row=2, column=0, pady=(20, 5), sticky="ew")

    def _build_canvas(self):
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(self.main_frame, bg=CONFIG["bg_color"], highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

    # --- Data Parsing ---

    def _get_ui_state(self):
        return {
            "margin": self.ent_margin.get().strip(),
            "pen_width": self.ent_pen_width.get().strip(),
            "group_gap": self.ent_group_gap.get().strip(),
            "lines": self.txt_lines.get("1.0", tk.END).strip(),
            "slants": self.txt_slants.get("1.0", tk.END).strip()
        }

    def _set_ui_state(self, state):
        self.ent_margin.delete(0, tk.END); self.ent_margin.insert(0, state.get("margin", ""))
        self.ent_pen_width.delete(0, tk.END); self.ent_pen_width.insert(0, state.get("pen_width", ""))
        self.ent_group_gap.delete(0, tk.END); self.ent_group_gap.insert(0, state.get("group_gap", ""))
        self.txt_lines.delete("1.0", tk.END); self.txt_lines.insert("1.0", state.get("lines", ""))
        self.txt_slants.delete("1.0", tk.END); self.txt_slants.insert("1.0", state.get("slants", ""))
        self.update_preview()

    def _parse_inputs(self):
        try:
            margin = float(self.ent_margin.get())
            pen_width = float(self.ent_pen_width.get())
            group_gap = float(self.ent_group_gap.get())
            if pen_width <= 0 or margin < 0 or group_gap < 0:
                return None
        except ValueError:
            return None 
            
        def parse_dash(dash_str):
            if dash_str.lower() in ["solid", ""]: return ()
            try: return tuple(map(int, dash_str.split()))
            except: return (4, 4)

        lines_data = []
        for line in self.txt_lines.get("1.0", tk.END).split('\n'):
            parts = [p.strip() for p in line.split(':')]
            if len(parts) >= 2:
                try:
                    name = parts[0]
                    pos = float(parts[1])
                    lw = float(parts[2]) if len(parts) > 2 else 1.0
                    dash = parse_dash(parts[3]) if len(parts) > 3 else (4, 4)
                    lines_data.append({"name": name, "pos": pos, "lw": lw, "dash": dash})
                except ValueError:
                    continue
                    
        slants_data = []
        for line in self.txt_slants.get("1.0", tk.END).split('\n'):
            parts = [p.strip() for p in line.split(':')]
            if len(parts) >= 2:
                try:
                    angle = float(parts[0])
                    spacing = float(parts[1])
                    lw = float(parts[2]) if len(parts) > 2 else 0.5
                    dash = parse_dash(parts[3]) if len(parts) > 3 else ()
                    if spacing > 0:
                        slants_data.append({"angle": angle, "spacing": spacing, "lw": lw, "dash": dash})
                except ValueError:
                    continue
                    
        return margin, pen_width, group_gap, lines_data, slants_data

    # --- Live Preview Engine ---

    def _debounce_update(self, event=None):
        if self._update_job is not None:
            self.after_cancel(self._update_job)
        self._update_job = self.after(300, self.update_preview)

    def _on_resize(self, event):
        if event.widget == self:
            self._debounce_update()

    def update_preview(self):
        self.canvas.delete("all")
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        if cw <= 1 or ch <= 1: return
            
        parsed = self._parse_inputs()
        if not parsed: return
        margin, pen_width, group_gap, lines_data, slants_data = parsed
        
        pad = 20
        scale = min((ch - pad*2) / CONFIG["page_height_mm"], (cw - pad*2) / CONFIG["page_width_mm"])
        offset_x = (cw - (CONFIG["page_width_mm"] * scale)) / 2
        offset_y = (ch - (CONFIG["page_height_mm"] * scale)) / 2
        
        def map_c(x_mm, y_mm): return offset_x + (x_mm * scale), offset_y + (y_mm * scale)

        # Draw Page
        p_x1, p_y1 = map_c(0, 0)
        p_x2, p_y2 = map_c(CONFIG["page_width_mm"], CONFIG["page_height_mm"])
        self.canvas.create_rectangle(p_x1, p_y1, p_x2, p_y2, fill=CONFIG["page_color"], outline="#111111")
        
        if not lines_data: return
        
        pw_values = [ld["pos"] for ld in lines_data]
        max_pw, min_pw = max(pw_values), min(pw_values)
        group_h_mm = (max_pw - min_pw) * pen_width
        
        # Calculate Group Math
        base_y_mm = margin + (max_pw * pen_width)
        
        while base_y_mm - (max_pw * pen_width) + group_h_mm <= CONFIG["page_height_mm"] - margin:
            top_y_mm = base_y_mm - (max_pw * pen_width)
            bottom_y_mm = base_y_mm - (min_pw * pen_width)
            
            # 1. Draw Slants (Per Group)
            for s in slants_data:
                rad = math.radians(s["angle"])
                dx_spacing = s["spacing"] / math.cos(rad) if math.cos(rad) != 0 else s["spacing"]
                dx_offset = group_h_mm * math.tan(rad)
                
                x_curr = margin
                while x_curr <= CONFIG["page_width_mm"] - margin:
                    pt1 = map_c(x_curr, bottom_y_mm)
                    pt2 = map_c(x_curr + dx_offset, top_y_mm)
                    self.canvas.create_line(*pt1, *pt2, fill="#C0C0C0", width=max(1, s["lw"]*scale), dash=s["dash"])
                    x_curr += dx_spacing

            # 2. Draw Horizontal Lines
            y_base, y_xheight = None, None
            for ld in lines_data:
                y_line = base_y_mm - (ld["pos"] * pen_width)
                x1, y1 = map_c(margin, y_line)
                x2, y2 = map_c(CONFIG["page_width_mm"] - margin, y_line)
                self.canvas.create_line(x1, y1, x2, y2, fill=CONFIG["line_color"], dash=ld["dash"], width=max(1, ld["lw"]*scale))
                
                if ld["name"].lower() == "base": y_base = y_line
                if ld["name"].lower() == "x-height": y_xheight = y_line

            # 3. Draw "x" marker
            if y_base is not None and y_xheight is not None:
                mid_y = (y_base + y_xheight) / 2
                h = abs(y_base - y_xheight) * 0.5
                
                # Start slightly to the right of the margin so it's readable
                x_start_mm = margin + 1 
                x_start, m_y = map_c(x_start_mm, mid_y)
                scaled_h = h * scale
                
                self.canvas.create_line(x_start, m_y - scaled_h/2, x_start + scaled_h, m_y + scaled_h/2, fill="#FF0000", width=max(1, scale*0.5))
                self.canvas.create_line(x_start, m_y + scaled_h/2, x_start + scaled_h, m_y - scaled_h/2, fill="#FF0000", width=max(1, scale*0.5))

            # Move to next group
            base_y_mm += group_h_mm + group_gap

        # Mask margins visually
        mask_color = CONFIG["bg_color"]
        m_x1, m_y1 = map_c(margin, margin)
        m_x2, m_y2 = map_c(CONFIG["page_width_mm"] - margin, CONFIG["page_height_mm"] - margin)
        self.canvas.create_rectangle(0, 0, cw, m_y1, fill=mask_color, outline="")
        self.canvas.create_rectangle(0, m_y2, cw, ch, fill=mask_color, outline="")
        self.canvas.create_rectangle(0, 0, m_x1, ch, fill=mask_color, outline="")
        self.canvas.create_rectangle(m_x2, 0, cw, ch, fill=mask_color, outline="")


    # --- PostScript Generation ---

    def generate_postscript_string(self):
        parsed = self._parse_inputs()
        if not parsed: raise ValueError("Invalid parameters.")
        margin, pen_width, group_gap, lines_data, slants_data = parsed
        
        state_json = json.dumps(self._get_ui_state())
        width_pts = CONFIG["page_width_mm"] * CONFIG["mm_to_pts"]
        height_pts = CONFIG["page_height_mm"] * CONFIG["mm_to_pts"]
        
        ps = [
            "%!PS-Adobe-3.0",
            f"%%BoundingBox: 0 0 {int(width_pts)} {int(height_pts)}",
            "%%Creator: Python Calligraphy Guide Generator Pro",
            "%%EndComments",
            "% BEGIN_METADATA",
            f"% {state_json}",
            "% END_METADATA",
            "",
            "/mm {2.83465 mul} def",
            ""
        ]

        def format_dash(dash_tuple):
            if not dash_tuple: return "[] 0 setdash"
            return f"[{' '.join(map(str, dash_tuple))}] 0 setdash"

        if not lines_data:
            ps.append("showpage")
            return "\n".join(ps)

        pw_values = [ld["pos"] for ld in lines_data]
        max_pw, min_pw = max(pw_values), min(pw_values)
        group_h_mm = (max_pw - min_pw) * pen_width
        base_y_mm = margin + (max_pw * pen_width)
        
        # Clipping path for margins
        ps.extend([
            "gsave",
            "newpath",
            f"{margin} mm {margin} mm moveto",
            f"{CONFIG['page_width_mm'] - margin} mm {margin} mm lineto",
            f"{CONFIG['page_width_mm'] - margin} mm {CONFIG['page_height_mm'] - margin} mm lineto",
            f"{margin} mm {CONFIG['page_height_mm'] - margin} mm lineto",
            "closepath clip"
        ])

        while base_y_mm - (max_pw * pen_width) + group_h_mm <= CONFIG["page_height_mm"] - margin:
            top_y_mm = base_y_mm - (max_pw * pen_width)
            bottom_y_mm = base_y_mm - (min_pw * pen_width)
            
            ps_top_y = CONFIG["page_height_mm"] - top_y_mm
            ps_bottom_y = CONFIG["page_height_mm"] - bottom_y_mm

            # 1. Slants
            for s in slants_data:
                rad = math.radians(s["angle"])
                dx_spacing = s["spacing"] / math.cos(rad) if math.cos(rad) != 0 else s["spacing"]
                dx_offset = group_h_mm * math.tan(rad)
                
                ps.append(f"0.7 setgray {s['lw']} setlinewidth {format_dash(s['dash'])}")
                
                x_curr = margin
                while x_curr <= CONFIG["page_width_mm"] - margin:
                    ps.extend([
                        "newpath",
                        f"{x_curr:.2f} mm {ps_bottom_y:.2f} mm moveto",
                        f"{(x_curr + dx_offset):.2f} mm {ps_top_y:.2f} mm lineto",
                        "stroke"
                    ])
                    x_curr += dx_spacing

            # 2. Horizontal Lines
            y_base, y_xheight = None, None
            for ld in lines_data:
                y_line_mm = base_y_mm - (ld["pos"] * pen_width)
                ps_y = CONFIG["page_height_mm"] - y_line_mm
                
                ps.append(f"0 setgray {ld['lw']} setlinewidth {format_dash(ld['dash'])}")
                ps.extend([
                    "newpath",
                    f"{margin} mm {ps_y:.2f} mm moveto",
                    f"{CONFIG['page_width_mm'] - margin} mm {ps_y:.2f} mm lineto",
                    "stroke"
                ])
                
                if ld["name"].lower() == "base": y_base = y_line_mm
                if ld["name"].lower() == "x-height": y_xheight = y_line_mm

            # 3. 'x' Marker
            if y_base is not None and y_xheight is not None:
                mid_y = (y_base + y_xheight) / 2
                h = abs(y_base - y_xheight) * 0.5
                ps_mid = CONFIG["page_height_mm"] - mid_y
                x_start = margin + 1
                
                ps.append("1 0 0 setrgbcolor 0.5 setlinewidth [] 0 setdash") # Red color
                ps.extend([
                    "newpath",
                    f"{x_start:.2f} mm {(ps_mid - h/2):.2f} mm moveto",
                    f"{(x_start + h):.2f} mm {(ps_mid + h/2):.2f} mm lineto",
                    "stroke",
                    "newpath",
                    f"{x_start:.2f} mm {(ps_mid + h/2):.2f} mm moveto",
                    f"{(x_start + h):.2f} mm {(ps_mid - h/2):.2f} mm lineto",
                    "stroke"
                ])

            base_y_mm += group_h_mm + group_gap

        ps.append("grestore") # Remove clipping path
        ps.append("showpage")
        return "\n".join(ps)

    # --- IO & Printing ---

    def save_postscript(self):
        try: ps_code = self.generate_postscript_string()
        except ValueError:
            messagebox.showerror("Error", "Invalid parameters.")
            return

        filepath = filedialog.asksaveasfilename(defaultextension=".ps", filetypes=[("PostScript", "*.ps")])
        if filepath:
            with open(filepath, 'w') as f: f.write(ps_code)
            messagebox.showinfo("Success", "File saved!")

    def load_postscript(self):
        filepath = filedialog.askopenfilename(filetypes=[("PostScript", "*.ps")])
        if filepath:
            with open(filepath, 'r') as f: content = f.read()
            match = re.search(r"% BEGIN_METADATA\n% (.*?)\n% END_METADATA", content)
            if match:
                self._set_ui_state(json.loads(match.group(1)))
            else:
                messagebox.showwarning("Warning", "No metadata found.")

    def _find_ghostscript(self):
        """Hunts the Windows registry/directories for Ghostscript execution paths."""
        paths = [r"C:\Program Files\gs", r"C:\Program Files (x86)\gs"]
        for base in paths:
            if os.path.exists(base):
                for folder in os.listdir(base):
                    bin_path = os.path.join(base, folder, "bin")
                    for exe in ["gswin64c.exe", "gswin32c.exe"]:
                        exe_path = os.path.join(bin_path, exe)
                        if os.path.exists(exe_path): return exe_path
        return None

    def print_postscript(self):
        try: ps_code = self.generate_postscript_string()
        except ValueError:
            messagebox.showerror("Error", "Invalid parameters.")
            return
            
        gs_path = self._find_ghostscript()
        if not gs_path:
            messagebox.showerror(
                "Ghostscript Required", 
                "Could not locate Ghostscript.\n\nTo use the native Print Dialog, please install Ghostscript (64-bit) in its default directory (C:\\Program Files\\gs\\)."
            )
            return

        temp_path = os.path.join(os.environ.get("TEMP", "."), "calligraphy_temp_print.ps")
        with open(temp_path, 'w') as f:
            f.write(ps_code)
            
        # -sDEVICE=mswinpr2 tells Ghostscript to open the standard Windows Print Dialog.
        try:
            subprocess.run([gs_path, "-sDEVICE=mswinpr2", "-dBATCH", "-dNOPAUSE", temp_path], check=True)
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Print Cancelled or Error", f"Printing process ended.\n{e}")

if __name__ == "__main__":
    app = CalligraphyApp()
    app.mainloop()