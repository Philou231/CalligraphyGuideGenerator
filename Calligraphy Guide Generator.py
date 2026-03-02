"""
Calligraphy Guide Sheet Generator - Pro Edition v2
--------------------------------------------------
Updates:
- Adjustable Paper Size (Width/Height in mm).
- Independent Vertical and Horizontal Margins.
- Paper-accurate Canvas Preview (white margins, fixed page bounds).
- PostScript Dash Array scaling to match Tkinter's relative scaling.
- Pure black lines everywhere, thinner black 'x' mark.
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
    "mm_to_pts": 2.83465,         # PostScript points conversion factor
    
    # Startup State Defaults
    "default_page_width": 210.0,
    "default_page_height": 297.0,
    "default_margin_v": 15.0,
    "default_margin_h": 15.0,
    "default_pen_width": 2.0,
    "default_group_gap": 8.0,
    
    # Format: Name : Position : LineWidth : DashPattern (solid or space-separated numbers)
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
    "line_color": "#000000",      # All lines are now perfectly black
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
        self.geometry("1200x900")
        self.minsize(1050, 750)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self._update_job = None
        
        self._setup_layout()
        self._build_sidebar()
        self._build_canvas()
        
        self._set_ui_state({
            "page_width": str(CONFIG["default_page_width"]),
            "page_height": str(CONFIG["default_page_height"]),
            "margin_v": str(CONFIG["default_margin_v"]),
            "margin_h": str(CONFIG["default_margin_h"]),
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
        
        self.sidebar_frame = ctk.CTkScrollableFrame(self, width=380, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=CONFIG["bg_color"])
        self.main_frame.grid(row=0, column=1, sticky="nsew")

    def _build_sidebar(self):
        row_idx = 0
        
        lbl_title = ctk.CTkLabel(self.sidebar_frame, text="Parameters", font=ctk.CTkFont(size=20, weight="bold"))
        lbl_title.grid(row=row_idx, column=0, padx=20, pady=(20, 10), sticky="w"); row_idx += 1
        
        # --- Page & Margins ---
        ctk.CTkLabel(self.sidebar_frame, text="Page & Layout (mm)", font=ctk.CTkFont(weight="bold")).grid(row=row_idx, column=0, padx=20, pady=(5, 0), sticky="w"); row_idx += 1
        frame_page = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        frame_page.grid(row=row_idx, column=0, padx=20, pady=5, sticky="ew"); row_idx += 1
        
        ctk.CTkLabel(frame_page, text="Page Width:").grid(row=0, column=0, sticky="w", pady=5)
        self.ent_page_width = ctk.CTkEntry(frame_page, width=70)
        self.ent_page_width.grid(row=0, column=1, sticky="e", pady=5)
        self.ent_page_width.bind("<KeyRelease>", self._debounce_update)
        
        ctk.CTkLabel(frame_page, text="Page Height:").grid(row=0, column=2, sticky="w", padx=(15, 0), pady=5)
        self.ent_page_height = ctk.CTkEntry(frame_page, width=70)
        self.ent_page_height.grid(row=0, column=3, sticky="e", pady=5)
        self.ent_page_height.bind("<KeyRelease>", self._debounce_update)

        ctk.CTkLabel(frame_page, text="Vert. Margin:").grid(row=1, column=0, sticky="w", pady=5)
        self.ent_margin_v = ctk.CTkEntry(frame_page, width=70)
        self.ent_margin_v.grid(row=1, column=1, sticky="e", pady=5)
        self.ent_margin_v.bind("<KeyRelease>", self._debounce_update)
        
        ctk.CTkLabel(frame_page, text="Horz. Margin:").grid(row=1, column=2, sticky="w", padx=(15, 0), pady=5)
        self.ent_margin_h = ctk.CTkEntry(frame_page, width=70)
        self.ent_margin_h.grid(row=1, column=3, sticky="e", pady=5)
        self.ent_margin_h.bind("<KeyRelease>", self._debounce_update)

        # --- Calligraphy Metrics ---
        ctk.CTkLabel(self.sidebar_frame, text="Calligraphy Metrics (mm)", font=ctk.CTkFont(weight="bold")).grid(row=row_idx, column=0, padx=20, pady=(15, 0), sticky="w"); row_idx += 1
        frame_metrics = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        frame_metrics.grid(row=row_idx, column=0, padx=20, pady=5, sticky="ew"); row_idx += 1
        
        ctk.CTkLabel(frame_metrics, text="Pen Width:").grid(row=0, column=0, sticky="w", pady=5)
        self.ent_pen_width = ctk.CTkEntry(frame_metrics, width=70)
        self.ent_pen_width.grid(row=0, column=1, sticky="e", pady=5)
        self.ent_pen_width.bind("<KeyRelease>", self._debounce_update)
        
        ctk.CTkLabel(frame_metrics, text="Group Gap:").grid(row=0, column=2, sticky="w", padx=(15, 0), pady=5)
        self.ent_group_gap = ctk.CTkEntry(frame_metrics, width=70)
        self.ent_group_gap.grid(row=0, column=3, sticky="e", pady=5)
        self.ent_group_gap.bind("<KeyRelease>", self._debounce_update)
        
        # --- Horizontal Lines ---
        ctk.CTkLabel(self.sidebar_frame, text="Horizontal Lines", font=ctk.CTkFont(weight="bold")).grid(row=row_idx, column=0, padx=20, pady=(15, 0), sticky="w"); row_idx += 1
        ctk.CTkLabel(self.sidebar_frame, text="Format: Name : Position : LineWidth : Dash\nExample: Ascender : 7 : 0.5 : 4 4", justify="left", font=ctk.CTkFont(size=11, slant="italic")).grid(row=row_idx, column=0, padx=20, sticky="w"); row_idx += 1
        
        self.txt_lines = ctk.CTkTextbox(self.sidebar_frame, height=120)
        self.txt_lines.grid(row=row_idx, column=0, padx=20, pady=5, sticky="ew"); row_idx += 1
        self.txt_lines.bind("<KeyRelease>", self._debounce_update)
        self.txt_lines._textbox.configure(undo=True)
        
        # --- Slants ---
        ctk.CTkLabel(self.sidebar_frame, text="Slant Overlays", font=ctk.CTkFont(weight="bold")).grid(row=row_idx, column=0, padx=20, pady=(15, 0), sticky="w"); row_idx += 1
        ctk.CTkLabel(self.sidebar_frame, text="Format: Angle° : Spacing : LineWidth : Dash\nExample: 10 : 15 : 0.3 : solid", justify="left", font=ctk.CTkFont(size=11, slant="italic")).grid(row=row_idx, column=0, padx=20, sticky="w"); row_idx += 1
        
        self.txt_slants = ctk.CTkTextbox(self.sidebar_frame, height=80)
        self.txt_slants.grid(row=row_idx, column=0, padx=20, pady=5, sticky="ew"); row_idx += 1
        self.txt_slants.bind("<KeyRelease>", self._debounce_update)
        self.txt_slants._textbox.configure(undo=True)
        
        # --- Actions ---
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
            "page_width": self.ent_page_width.get().strip(),
            "page_height": self.ent_page_height.get().strip(),
            "margin_v": self.ent_margin_v.get().strip(),
            "margin_h": self.ent_margin_h.get().strip(),
            "pen_width": self.ent_pen_width.get().strip(),
            "group_gap": self.ent_group_gap.get().strip(),
            "lines": self.txt_lines.get("1.0", tk.END).strip(),
            "slants": self.txt_slants.get("1.0", tk.END).strip()
        }

    def _set_ui_state(self, state):
        self.ent_page_width.delete(0, tk.END); self.ent_page_width.insert(0, state.get("page_width", ""))
        self.ent_page_height.delete(0, tk.END); self.ent_page_height.insert(0, state.get("page_height", ""))
        self.ent_margin_v.delete(0, tk.END); self.ent_margin_v.insert(0, state.get("margin_v", ""))
        self.ent_margin_h.delete(0, tk.END); self.ent_margin_h.insert(0, state.get("margin_h", ""))
        self.ent_pen_width.delete(0, tk.END); self.ent_pen_width.insert(0, state.get("pen_width", ""))
        self.ent_group_gap.delete(0, tk.END); self.ent_group_gap.insert(0, state.get("group_gap", ""))
        self.txt_lines.delete("1.0", tk.END); self.txt_lines.insert("1.0", state.get("lines", ""))
        self.txt_slants.delete("1.0", tk.END); self.txt_slants.insert("1.0", state.get("slants", ""))
        self.update_preview()

    def _parse_inputs(self):
        try:
            pw = float(self.ent_page_width.get())
            ph = float(self.ent_page_height.get())
            margin_v = float(self.ent_margin_v.get())
            margin_h = float(self.ent_margin_h.get())
            pen_width = float(self.ent_pen_width.get())
            group_gap = float(self.ent_group_gap.get())
            if any(val <= 0 for val in [pw, ph, pen_width]) or margin_v < 0 or margin_h < 0 or group_gap < 0:
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
                    
        return pw, ph, margin_v, margin_h, pen_width, group_gap, lines_data, slants_data

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
        page_width, page_height, margin_v, margin_h, pen_width, group_gap, lines_data, slants_data = parsed
        
        pad = 20
        scale = min((ch - pad*2) / page_height, (cw - pad*2) / page_width)
        offset_x = (cw - (page_width * scale)) / 2
        offset_y = (ch - (page_height * scale)) / 2
        
        def map_c(x_mm, y_mm): return offset_x + (x_mm * scale), offset_y + (y_mm * scale)

        # 1. Draw Physical Paper Base (White)
        p_x1, p_y1 = map_c(0, 0)
        p_x2, p_y2 = map_c(page_width, page_height)
        self.canvas.create_rectangle(p_x1, p_y1, p_x2, p_y2, fill=CONFIG["page_color"], outline="")
        
        if lines_data:
            pw_values = [ld["pos"] for ld in lines_data]
            max_pw, min_pw = max(pw_values), min(pw_values)
            group_h_mm = (max_pw - min_pw) * pen_width
            
            base_y_mm = margin_v + (max_pw * pen_width)
            
            # Loop per Group
            while base_y_mm - (max_pw * pen_width) + group_h_mm <= page_height - margin_v:
                top_y_mm = base_y_mm - (max_pw * pen_width)
                bottom_y_mm = base_y_mm - (min_pw * pen_width)
                
                # 2. Draw Slants
                for s in slants_data:
                    rad = math.radians(s["angle"])
                    dx_spacing = s["spacing"] / math.cos(rad) if math.cos(rad) != 0 else s["spacing"]
                    dx_offset = group_h_mm * math.tan(rad)
                    
                    # Ensure first visible line starts exactly at margin_h
                    x_curr = margin_h if dx_offset >= 0 else margin_h - dx_offset
                    
                    while min(x_curr, x_curr + dx_offset) <= page_width - margin_h:
                        pt1 = map_c(x_curr, bottom_y_mm)
                        pt2 = map_c(x_curr + dx_offset, top_y_mm)
                        self.canvas.create_line(*pt1, *pt2, fill=CONFIG["line_color"], width=max(1, s["lw"]*scale), dash=s["dash"])
                        x_curr += dx_spacing

                # 3. Draw Horizontal Lines
                y_base, y_xheight = None, None
                for ld in lines_data:
                    y_line = base_y_mm - (ld["pos"] * pen_width)
                    x1, y1 = map_c(margin_h, y_line)
                    x2, y2 = map_c(page_width - margin_h, y_line)
                    self.canvas.create_line(x1, y1, x2, y2, fill=CONFIG["line_color"], dash=ld["dash"], width=max(1, ld["lw"]*scale))
                    
                    if ld["name"].lower() == "base": y_base = y_line
                    if ld["name"].lower() == "x-height": y_xheight = y_line

                # 4. Draw Black, Thinner "x" marker
                if y_base is not None and y_xheight is not None:
                    mid_y = (y_base + y_xheight) / 2
                    h = abs(y_base - y_xheight) * 0.5
                    
                    x_start_mm = margin_h + 1 
                    x_start, m_y = map_c(x_start_mm, mid_y)
                    scaled_h = h * scale
                    
                    # Thin, pure black marker
                    self.canvas.create_line(x_start, m_y - scaled_h/2, x_start + scaled_h, m_y + scaled_h/2, fill=CONFIG["line_color"], width=1)
                    self.canvas.create_line(x_start, m_y + scaled_h/2, x_start + scaled_h, m_y - scaled_h/2, fill=CONFIG["line_color"], width=1)

                base_y_mm += group_h_mm + group_gap

        # 5. Mask Margins (Draw white rectangles over horizontal/slant spillover)
        m_x1, m_y1 = map_c(margin_h, margin_v)
        m_x2, m_y2 = map_c(page_width - margin_h, page_height - margin_v)
        
        # Left and Right Masks (White)
        self.canvas.create_rectangle(p_x1, p_y1, m_x1, p_y2, fill=CONFIG["page_color"], outline="")
        self.canvas.create_rectangle(m_x2, p_y1, p_x2, p_y2, fill=CONFIG["page_color"], outline="")
        
        # Finally, draw a subtle border around the physical paper edge
        self.canvas.create_rectangle(p_x1, p_y1, p_x2, p_y2, fill="", outline="#888888")


    # --- PostScript Generation ---

    def generate_postscript_string(self):
        parsed = self._parse_inputs()
        if not parsed: raise ValueError("Invalid parameters.")
        page_width, page_height, margin_v, margin_h, pen_width, group_gap, lines_data, slants_data = parsed
        
        state_json = json.dumps(self._get_ui_state())
        width_pts = page_width * CONFIG["mm_to_pts"]
        height_pts = page_height * CONFIG["mm_to_pts"]
        
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

        def format_dash(dash_tuple, line_width):
            """Scales the PostScript point dash array by the line width to match Tkinter's relative behavior."""
            if not dash_tuple: return "[] 0 setdash"
            scaled_dashes = [str(d * line_width) for d in dash_tuple]
            return f"[{' '.join(scaled_dashes)}] 0 setdash"

        if not lines_data:
            ps.append("showpage")
            return "\n".join(ps)

        pw_values = [ld["pos"] for ld in lines_data]
        max_pw, min_pw = max(pw_values), min(pw_values)
        group_h_mm = (max_pw - min_pw) * pen_width
        base_y_mm = margin_v + (max_pw * pen_width)
        
        # Clipping path for physical margins
        ps.extend([
            "gsave",
            "newpath",
            f"{margin_h} mm {margin_v} mm moveto",
            f"{page_width - margin_h} mm {margin_v} mm lineto",
            f"{page_width - margin_h} mm {page_height - margin_v} mm lineto",
            f"{margin_h} mm {page_height - margin_v} mm lineto",
            "closepath clip"
        ])

        while base_y_mm - (max_pw * pen_width) + group_h_mm <= page_height - margin_v:
            top_y_mm = base_y_mm - (max_pw * pen_width)
            bottom_y_mm = base_y_mm - (min_pw * pen_width)
            
            ps_top_y = page_height - top_y_mm
            ps_bottom_y = page_height - bottom_y_mm

            # 1. Slants (Absolute Black: 0 setgray)
            for s in slants_data:
                rad = math.radians(s["angle"])
                dx_spacing = s["spacing"] / math.cos(rad) if math.cos(rad) != 0 else s["spacing"]
                dx_offset = group_h_mm * math.tan(rad)
                
                ps.append(f"0 setgray {s['lw']} setlinewidth {format_dash(s['dash'], s['lw'])}")
                
                x_curr = margin_h if dx_offset >= 0 else margin_h - dx_offset
                while min(x_curr, x_curr + dx_offset) <= page_width - margin_h:
                    ps.extend([
                        "newpath",
                        f"{x_curr:.2f} mm {ps_bottom_y:.2f} mm moveto",
                        f"{(x_curr + dx_offset):.2f} mm {ps_top_y:.2f} mm lineto",
                        "stroke"
                    ])
                    x_curr += dx_spacing

            # 2. Horizontal Lines (Absolute Black: 0 setgray)
            y_base, y_xheight = None, None
            for ld in lines_data:
                y_line_mm = base_y_mm - (ld["pos"] * pen_width)
                ps_y = page_height - y_line_mm
                
                ps.append(f"0 setgray {ld['lw']} setlinewidth {format_dash(ld['dash'], ld['lw'])}")
                ps.extend([
                    "newpath",
                    f"{margin_h} mm {ps_y:.2f} mm moveto",
                    f"{page_width - margin_h} mm {ps_y:.2f} mm lineto",
                    "stroke"
                ])
                
                if ld["name"].lower() == "base": y_base = y_line_mm
                if ld["name"].lower() == "x-height": y_xheight = y_line_mm

            # 3. 'x' Marker (Black and thin)
            if y_base is not None and y_xheight is not None:
                mid_y = (y_base + y_xheight) / 2
                h = abs(y_base - y_xheight) * 0.5
                ps_mid = page_height - mid_y
                x_start = margin_h + 1
                
                # 0 setgray = black. 0.2 line width makes it very thin.
                ps.append("0 setgray 0.2 setlinewidth [] 0 setdash") 
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
            
        try:
            subprocess.run([gs_path, "-sDEVICE=mswinpr2", "-dBATCH", "-dNOPAUSE", temp_path], check=True)
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Print Cancelled or Error", f"Printing process ended.\n{e}")

if __name__ == "__main__":
    app = CalligraphyApp()
    app.mainloop()