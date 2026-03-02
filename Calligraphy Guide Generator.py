"""
Calligraphy Guide Sheet Generator - Pro Edition v6
--------------------------------------------------
Updates:
- Adobe-style Smart Input Parsing: Mix 'in', 'inch', '"', or 'mm' in any field.
- Eradicated global unit toggle; the UI now respects granular, per-field strings.
- Restored custom slant line clearance anchors (1.0mm and 0.01mm).
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
    "mm_to_pts": 72.0 / 25.4,
    "in_to_mm": 25.4,
    
    # Startup State Defaults (Notice the mixed strings)
    "default_page_width": "8.5 in",
    "default_page_height": "11 in",
    "default_margin_v":  "5 mm",
    "default_margin_h":  "5 mm",
    "default_pen_width": "1.0",      # defaults to mm
    "default_group_gap": "5.0",      # defaults to mm
    
    "default_lines": (
        "Ascender : 7 : 0.10 : solid\n"
        "X-Height : 5 : 0.10 : solid\n"
        "Base : 0 : 0.30 : solid\n"
        "Descender : -5 : 0.10 : solid"
    ),
    "default_slants": "10 : 5 : 0.10 : 1 1",   
    
    "bg_color": "#242424",
    "page_color": "#FFFFFF",
    "line_color": "#000000",
}

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try: ctypes.windll.user32.SetProcessDPIAware()
    except Exception: pass

# -----------------------------------------------------------------------------
# Core Architecture: Data & Engine
# -----------------------------------------------------------------------------

class RenderData:
    def __init__(self, page_width, page_height, margin_v, margin_h):
        self.page_width = page_width
        self.page_height = page_height
        self.margin_v = margin_v
        self.margin_h = margin_h
        
        self.slants = []
        self.horizontals = []
        self.markers = []

class GeometryEngine:
    @staticmethod
    def calculate(page_width, page_height, margin_v, margin_h, pen_width, group_gap, lines_data, slants_data):
        rd = RenderData(page_width, page_height, margin_v, margin_h)
        if not lines_data: return rd

        pw_values = [ld["pos"] for ld in lines_data]
        max_pw, min_pw = max(pw_values), min(pw_values)
        group_h_mm = (max_pw - min_pw) * pen_width
        
        base_pos, xheight_pos = 0.0, None
        for ld in lines_data:
            if ld["name"].lower() == "base": base_pos = ld["pos"]
            if ld["name"].lower() == "x-height": xheight_pos = ld["pos"]

        x_start_mm = margin_h + 1.0
        if xheight_pos is not None:
            h_mm = abs(base_pos - xheight_pos) * pen_width * 0.5
            slant_anchor_x_mm = x_start_mm + h_mm + 1.0
        else:
            slant_anchor_x_mm = margin_h + 0.01

        current_top_y_mm = margin_v
        
        while current_top_y_mm + group_h_mm <= page_height - margin_v:
            top_y_mm = current_top_y_mm
            bottom_y_mm = current_top_y_mm + group_h_mm
            base_y_mm = current_top_y_mm + (max_pw - base_pos) * pen_width
            
            y_xheight_mm = None
            if xheight_pos is not None:
                y_xheight_mm = current_top_y_mm + (max_pw - xheight_pos) * pen_width

            for s in slants_data:
                rad = math.radians(s["angle"])
                spacing = s["spacing"]
                
                n_min = math.floor((margin_h - slant_anchor_x_mm) / spacing) - 2
                n_max = math.ceil((page_width - margin_h - slant_anchor_x_mm) / spacing) + 2
                
                for n in range(n_min, n_max):
                    x_cross = slant_anchor_x_mm + n * spacing
                    
                    if y_xheight_mm is not None:
                        x_at_xheight = x_cross + (base_y_mm - y_xheight_mm) * math.tan(rad)
                        if min(x_cross, x_at_xheight) < slant_anchor_x_mm and max(x_cross, x_at_xheight) > margin_h:
                            continue
                    
                    x_top = x_cross + (base_y_mm - top_y_mm) * math.tan(rad)
                    x_bottom = x_cross + (base_y_mm - bottom_y_mm) * math.tan(rad)
                    
                    if max(x_top, x_bottom) > margin_h and min(x_top, x_bottom) < page_width - margin_h:
                        rd.slants.append((x_top, top_y_mm, x_bottom, bottom_y_mm, s["lw"], s["dash"]))

            for ld in lines_data:
                y_line = current_top_y_mm + (max_pw - ld["pos"]) * pen_width
                rd.horizontals.append((y_line, ld["lw"], ld["dash"]))

            if xheight_pos is not None:
                mid_y_mm = (base_y_mm + y_xheight_mm) / 2.0
                h_mm = abs(base_y_mm - y_xheight_mm) * 0.5
                rd.markers.append((x_start_mm, mid_y_mm, h_mm))

            current_top_y_mm += group_h_mm + group_gap

        return rd

class PostScriptExporter:
    @staticmethod
    def generate(render_data, state_json):
        rd = render_data
        width_pts = rd.page_width * CONFIG["mm_to_pts"]
        height_pts = rd.page_height * CONFIG["mm_to_pts"]
        
        ps = [
            "%!PS-Adobe-3.0",
            f"%%BoundingBox: 0 0 {width_pts:.4f} {height_pts:.4f}",
            "%%Creator: Python Calligraphy Guide Generator Pro",
            "%%EndComments",
            "% BEGIN_METADATA",
            f"% {state_json}",
            "% END_METADATA",
            "",
            f"/mm {{{CONFIG['mm_to_pts']:.6f} mul}} def",
            ""
        ]

        def format_dash(dash_tuple):
            if not dash_tuple: return "[] 0 setdash"
            return f"[{' '.join(map(str, dash_tuple))}] 0 setdash"

        ps.extend([
            "gsave", "newpath",
            f"{rd.margin_h:.4f} mm {rd.margin_v:.4f} mm moveto",
            f"{(rd.page_width - rd.margin_h):.4f} mm {rd.margin_v:.4f} mm lineto",
            f"{(rd.page_width - rd.margin_h):.4f} mm {(rd.page_height - rd.margin_v):.4f} mm lineto",
            f"{rd.margin_h:.4f} mm {(rd.page_height - rd.margin_v):.4f} mm lineto",
            "closepath clip"
        ])

        for (x_top, y_top, x_bot, y_bot, lw, dash) in rd.slants:
            ps.append(f"0 setgray {lw:.4f} setlinewidth {format_dash(dash)}")
            ps.extend([
                "newpath",
                f"{x_bot:.4f} mm {(rd.page_height - y_bot):.4f} mm moveto",
                f"{x_top:.4f} mm {(rd.page_height - y_top):.4f} mm lineto",
                "stroke"
            ])

        for (y, lw, dash) in rd.horizontals:
            ps.append(f"0 setgray {lw:.4f} setlinewidth {format_dash(dash)}")
            ps.extend([
                "newpath",
                f"{rd.margin_h:.4f} mm {(rd.page_height - y):.4f} mm moveto",
                f"{(rd.page_width - rd.margin_h):.4f} mm {(rd.page_height - y):.4f} mm lineto",
                "stroke"
            ])

        ps.append("0 setgray 0.2 setlinewidth [] 0 setdash") 
        for (x_c, y_c, h) in rd.markers:
            ps_y_c = rd.page_height - y_c
            ps.extend([
                "newpath",
                f"{x_c:.4f} mm {(ps_y_c - h/2):.4f} mm moveto",
                f"{(x_c + h):.4f} mm {(ps_y_c + h/2):.4f} mm lineto",
                "stroke",
                "newpath",
                f"{x_c:.4f} mm {(ps_y_c + h/2):.4f} mm moveto",
                f"{(x_c + h):.4f} mm {(ps_y_c - h/2):.4f} mm lineto",
                "stroke"
            ])

        ps.append("grestore\nshowpage")
        return "\n".join(ps)

# -----------------------------------------------------------------------------
# User Interface & Interaction
# -----------------------------------------------------------------------------

class CalligraphyApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Calligraphy Guide Sheet Generator")
        self.geometry("1200x900")
        self.minsize(1050, 750)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.backend_state = None
        self._update_job = None
        
        self._setup_layout()
        self._build_sidebar()
        self._build_canvas()
        
        self._set_ui_state({
            "page_width": CONFIG["default_page_width"],
            "page_height": CONFIG["default_page_height"],
            "margin_v": CONFIG["default_margin_v"],
            "margin_h": CONFIG["default_margin_h"],
            "pen_width": CONFIG["default_pen_width"],
            "group_gap": CONFIG["default_group_gap"],
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
        self.lbl_page = ctk.CTkLabel(self.sidebar_frame, text="Page & Layout (mm or in)", font=ctk.CTkFont(weight="bold"))
        self.lbl_page.grid(row=row_idx, column=0, padx=20, pady=(5, 0), sticky="w"); row_idx += 1
        
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
        self.lbl_metrics = ctk.CTkLabel(self.sidebar_frame, text="Calligraphy Metrics (mm or in)", font=ctk.CTkFont(weight="bold"))
        self.lbl_metrics.grid(row=row_idx, column=0, padx=20, pady=(15, 0), sticky="w"); row_idx += 1
        
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
        ctk.CTkLabel(self.sidebar_frame, text="Format: Name : Position : LineWidth : Dash", justify="left", font=ctk.CTkFont(size=11, slant="italic")).grid(row=row_idx, column=0, padx=20, sticky="w"); row_idx += 1
        
        self.txt_lines = ctk.CTkTextbox(self.sidebar_frame, height=120)
        self.txt_lines.grid(row=row_idx, column=0, padx=20, pady=5, sticky="ew"); row_idx += 1
        self.txt_lines.bind("<KeyRelease>", self._debounce_update)
        self.txt_lines._textbox.configure(undo=True)
        
        # --- Slants ---
        ctk.CTkLabel(self.sidebar_frame, text="Slant Overlays", font=ctk.CTkFont(weight="bold")).grid(row=row_idx, column=0, padx=20, pady=(15, 0), sticky="w"); row_idx += 1
        ctk.CTkLabel(self.sidebar_frame, text="Format: Angle° : Spacing : LineWidth : Dash", justify="left", font=ctk.CTkFont(size=11, slant="italic")).grid(row=row_idx, column=0, padx=20, sticky="w"); row_idx += 1
        
        self.txt_slants = ctk.CTkTextbox(self.sidebar_frame, height=80)
        self.txt_slants.grid(row=row_idx, column=0, padx=20, pady=5, sticky="ew"); row_idx += 1
        self.txt_slants.bind("<KeyRelease>", self._debounce_update)
        self.txt_slants._textbox.configure(undo=True)
        
        # --- Actions ---
        frame_actions = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        frame_actions.grid(row=row_idx, column=0, padx=20, pady=30, sticky="ew"); row_idx += 1
        frame_actions.grid_columnconfigure(0, weight=1)
        
        ctk.CTkButton(frame_actions, text="💾 Save PostScript (.ps)", command=self.save_postscript, fg_color="#1E90FF", hover_color="#104E8B").grid(row=0, column=0, pady=5, sticky="ew")
        ctk.CTkButton(frame_actions, text="📂 Load Template", command=self.load_postscript, fg_color="#555555", hover_color="#333333").grid(row=1, column=0, pady=5, sticky="ew")
        ctk.CTkButton(frame_actions, text="🖨️ Print Now...", command=self.print_postscript, fg_color="#2E8B57", hover_color="#1E5C3A").grid(row=2, column=0, pady=(20, 5), sticky="ew")

    def _build_canvas(self):
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(self.main_frame, bg=CONFIG["bg_color"], highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

    # --- Data Parsing Engine ---

    def _parse_val(self, val_str):
        """Intelligently extracts mm conversions from mixed string formats."""
        val_str = str(val_str).strip().lower()
        if val_str.endswith(('in', 'inch', '"')):
            num_str = re.sub(r'[a-z"]', '', val_str).strip()
            try: return float(num_str) * CONFIG["in_to_mm"]
            except ValueError: return None
        elif val_str.endswith('mm'):
            num_str = re.sub(r'mm', '', val_str).strip()
            try: return float(num_str)
            except ValueError: return None
        else:
            try: return float(val_str)
            except ValueError: return None

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

    def _parse_inputs_to_mm(self):
        pw = self._parse_val(self.ent_page_width.get())
        ph = self._parse_val(self.ent_page_height.get())
        margin_v = self._parse_val(self.ent_margin_v.get())
        margin_h = self._parse_val(self.ent_margin_h.get())
        pen_width = self._parse_val(self.ent_pen_width.get())
        group_gap = self._parse_val(self.ent_group_gap.get())
        
        if None in (pw, ph, margin_v, margin_h, pen_width, group_gap):
            return None
        if any(val <= 0 for val in [pw, ph, pen_width]) or margin_v < 0 or margin_h < 0 or group_gap < 0:
            return None 
            
        def parse_dash(dash_str):
            if dash_str.lower() in ["solid", ""]: return ()
            try: return tuple(map(int, dash_str.split()))
            except: return (4, 4)

        lines_data, slants_data = [], []
        
        for line in self.txt_lines.get("1.0", tk.END).split('\n'):
            parts = [p.strip() for p in line.split(':')]
            if len(parts) >= 2:
                try:
                    lw = self._parse_val(parts[2]) if len(parts) > 2 else 1.0
                    if lw is None: continue
                    lines_data.append({
                        "name": parts[0], 
                        "pos": float(parts[1]), # Relative ratio, purely unitless
                        "lw": lw, 
                        "dash": parse_dash(parts[3]) if len(parts) > 3 else (4, 4)
                    })
                except ValueError: continue
                    
        for line in self.txt_slants.get("1.0", tk.END).split('\n'):
            parts = [p.strip() for p in line.split(':')]
            if len(parts) >= 2:
                try:
                    spacing = self._parse_val(parts[1])
                    if spacing is None or spacing <= 0: continue
                    
                    lw = self._parse_val(parts[2]) if len(parts) > 2 else 0.5
                    if lw is None: continue
                    
                    slants_data.append({
                        "angle": float(parts[0]), # Degrees, purely unitless
                        "spacing": spacing, 
                        "lw": lw, 
                        "dash": parse_dash(parts[3]) if len(parts) > 3 else ()
                    })
                except ValueError: continue
                    
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
            
        parsed = self._parse_inputs_to_mm()
        if parsed: self.backend_state = parsed
        else: return

        rd = GeometryEngine.calculate(*self.backend_state)
        
        pad = 20
        scale = min((ch - pad*2) / rd.page_height, (cw - pad*2) / rd.page_width)
        offset_x = (cw - (rd.page_width * scale)) / 2
        offset_y = (ch - (rd.page_height * scale)) / 2
        
        def map_c(x_mm, y_mm): return offset_x + (x_mm * scale), offset_y + (y_mm * scale)

        p_x1, p_y1 = map_c(0, 0)
        p_x2, p_y2 = map_c(rd.page_width, rd.page_height)
        self.canvas.create_rectangle(p_x1, p_y1, p_x2, p_y2, fill=CONFIG["page_color"], outline="")
        
        for (x1, y1, x2, y2, lw, dash) in rd.slants:
            pt1, pt2 = map_c(x1, y1), map_c(x2, y2)
            self.canvas.create_line(*pt1, *pt2, fill=CONFIG["line_color"], width=max(1, lw*scale), dash=dash)

        for (y, lw, dash) in rd.horizontals:
            pt1, pt2 = map_c(rd.margin_h, y), map_c(rd.page_width - rd.margin_h, y)
            self.canvas.create_line(*pt1, *pt2, fill=CONFIG["line_color"], dash=dash, width=max(1, lw*scale))

        for (x_c, y_c, h) in rd.markers:
            cx, cy = map_c(x_c, y_c)
            scaled_h = h * scale
            self.canvas.create_line(cx, cy - scaled_h/2, cx + scaled_h, cy + scaled_h/2, fill=CONFIG["line_color"], width=1)
            self.canvas.create_line(cx, cy + scaled_h/2, cx + scaled_h, cy - scaled_h/2, fill=CONFIG["line_color"], width=1)

        m_x1, m_y1 = map_c(rd.margin_h, rd.margin_v)
        m_x2, m_y2 = map_c(rd.page_width - rd.margin_h, rd.page_height - rd.margin_v)
        self.canvas.create_rectangle(p_x1, p_y1, m_x1, p_y2, fill=CONFIG["page_color"], outline="")
        self.canvas.create_rectangle(m_x2, p_y1, p_x2, p_y2, fill=CONFIG["page_color"], outline="")
        self.canvas.create_rectangle(p_x1, p_y1, p_x2, p_y2, fill="", outline="#888888")

    # --- IO & Printing ---

    def _get_render_data(self):
        parsed = self._parse_inputs_to_mm()
        if not parsed: raise ValueError("Invalid parameters.")
        self.backend_state = parsed
        return GeometryEngine.calculate(*parsed)

    def save_postscript(self):
        try:
            rd = self._get_render_data()
            ps_code = PostScriptExporter.generate(rd, json.dumps(self._get_ui_state()))
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
            if match: self._set_ui_state(json.loads(match.group(1)))
            else: messagebox.showwarning("Warning", "No metadata found.")

    def print_postscript(self):
        try:
            rd = self._get_render_data()
            ps_code = PostScriptExporter.generate(rd, json.dumps(self._get_ui_state()))
        except ValueError:
            messagebox.showerror("Error", "Invalid parameters.")
            return
            
        paths = [r"C:\Program Files\gs", r"C:\Program Files (x86)\gs"]
        gs_path = next((os.path.join(b, f, "bin", e) for b in paths if os.path.exists(b) 
                        for f in os.listdir(b) for e in ["gswin64c.exe", "gswin32c.exe"] 
                        if os.path.exists(os.path.join(b, f, "bin", e))), None)
        
        if not gs_path:
            messagebox.showerror("Ghostscript Required", "Could not locate Ghostscript installation.")
            return

        temp_path = os.path.join(os.environ.get("TEMP", "."), "calligraphy_temp_print.ps")
        with open(temp_path, 'w') as f: f.write(ps_code)
            
        try: subprocess.run([gs_path, "-sDEVICE=mswinpr2", "-dBATCH", "-dNOPAUSE", temp_path], check=True)
        except subprocess.CalledProcessError as e: messagebox.showerror("Print Cancelled", f"Process ended.\n{e}")

if __name__ == "__main__":
    app = CalligraphyApp()
    app.mainloop()