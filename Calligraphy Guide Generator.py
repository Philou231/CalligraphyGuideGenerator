"""
Calligraphy Guide Sheet Generator
---------------------------------
"""

import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
import customtkinter as ctk
import math
import json
import os
import ctypes
import subprocess
import re
from dataclasses import dataclass, field
from typing import List, Tuple, Dict, Optional

# -----------------------------------------------------------------------------
# Configuration & Defaults
# -----------------------------------------------------------------------------
CONFIG = {
    "mm_to_pts": 72.0 / 25.4,
    "in_to_mm": 25.4,
    
    "default_page_width": "8.5 in",
    "default_page_height": "11 in",
    "default_margin_v":  "5 mm",
    "default_margin_h":  "5 mm",
    "default_pen_width": "1.0",
    "default_group_gap": "5.0",
    
    "style_map_ui": {"Solid": (), "Dashed": (1, 0.5), "Dotted": (0.5, 0.5)},
    "style_map_svg": {"Solid": (), "Dashed": (1, 0.5), "Dotted": (0.5, 0.5)},
    
    "bg_color": "#242424",
    "page_color": "#FFFFFF",
    "default_line_color": "#808080", 
    "default_dot_color": "#C0C0C0"
}

PRESETS = {
    "Italic (10°)": {
        "pen_width": "1.0", "group_gap": "0.0",
        "lines": [
            {"name": "Ascender", "pos": "10", "lw": "0.10", "style": "Solid"},
            {"name": "Capital", "pos": "7", "lw": "0.10", "style": "Solid"},
            {"name": "X-Height", "pos": "5", "lw": "0.10", "style": "Solid"},
            {"name": "Base", "pos": "0", "lw": "0.30", "style": "Solid"},
            {"name": "Descender", "pos": "-5", "lw": "0.10", "style": "Solid"}
        ],
        "slants": [{"angle": "10", "spacing": "3.4 mm", "lw": "0.10", "style": "Dotted"}],
        "oval_enabled": False, "oval_top": "X-Height", "oval_bot": "Base", "oval_ratio": "0.4"
    },
    "Copperplate (55°)": {
        "pen_width": "1.0", "group_gap": "5.0",
        "lines": [
            {"name": "Ascender", "pos": "10", "lw": "0.10", "style": "Solid"},
            {"name": "X-Height", "pos": "5", "lw": "0.10", "style": "Solid"},
            {"name": "Base", "pos": "0", "lw": "0.30", "style": "Solid"},
            {"name": "Descender", "pos": "-5", "lw": "0.10", "style": "Solid"}
        ],
        "slants": [{"angle": "35", "spacing": "5 mm", "lw": "0.10", "style": "Dotted"}], 
        "oval_enabled": True, "oval_top": "X-Height", "oval_bot": "Base", "oval_ratio": "0.5"
    },
    "Spencerian (52°)": {
        "pen_width": "1.0", "group_gap": "6.0",
        "lines": [
            {"name": "Ascender", "pos": "9", "lw": "0.10", "style": "Solid"},
            {"name": "X-Height", "pos": "3", "lw": "0.10", "style": "Solid"},
            {"name": "Base", "pos": "0", "lw": "0.30", "style": "Solid"},
            {"name": "Descender", "pos": "-6", "lw": "0.10", "style": "Solid"}
        ],
        "slants": [{"angle": "38", "spacing": "6 mm", "lw": "0.10", "style": "Dotted"}],
        "oval_enabled": True, "oval_top": "X-Height", "oval_bot": "Base", "oval_ratio": "0.4"
    }
}

try: ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception: pass

# -----------------------------------------------------------------------------
# Core Architecture: Data & Engine
# -----------------------------------------------------------------------------

@dataclass
class GridConfig:
    pw: float; ph: float; mv: float; mh: float
    pen_w: float; gap: float
    lines: List[Dict]; slants: List[Dict]; oval_data: Dict
    line_color: str; dot_color: str
    show_center: bool; show_x_marker: bool; dot_gap: float; dot_size: float
    radial: bool; radius: float

@dataclass
class RenderData:
    page_width: float
    page_height: float
    margin_v: float
    margin_h: float
    line_color: str
    dot_color: str
    show_center: bool
    dot_gap: float
    dot_size: float
    radial: bool
    radius: float
    slants: List[Tuple] = field(default_factory=list)
    horizontals: List[Tuple] = field(default_factory=list)
    arcs: List[Tuple] = field(default_factory=list)
    ovals: List[Tuple] = field(default_factory=list) 
    markers: List[Tuple] = field(default_factory=list)

class GeometryEngine:
    @staticmethod
    def calculate(c: GridConfig) -> RenderData:
        rd = RenderData(c.pw, c.ph, c.mv, c.mh, c.line_color, c.dot_color, 
                        c.show_center, c.dot_gap, c.dot_size, c.radial, c.radius)
        if not c.lines: return rd

        pw_vals = [ld["pos"] for ld in c.lines]
        max_pw, min_pw = max(pw_vals), min(pw_vals)
        group_h_mm = (max_pw - min_pw) * c.pen_w
        
        pos_map = {ld["name"].lower(): ld["pos"] for ld in c.lines}

        current_top_y_mm = c.mv
        while current_top_y_mm + group_h_mm <= c.ph - c.mv:
            y_min = current_top_y_mm
            y_max = current_top_y_mm + group_h_mm
            
            x_min_clip = c.mh
            x_max_clip = c.pw - c.mh
            
            if c.radial:
                # ==========================================
                # 1. RADIUS ANCHORING
                # ==========================================
                # Anchor user's Radius input strictly to the 'Base' line.
                base_pos = pos_map.get("base", 0)
                y_base = current_top_y_mm + (max_pw - base_pos) * c.pen_w
        
                cx = c.pw / 2
                cy = y_base + c.radius
                
                # Bounding radii for this specific group (used later for precise slant intersections)
                r_top_bound = cy - current_top_y_mm
                r_bot_bound = cy - (current_top_y_mm + group_h_mm)
                
                for ld in c.lines:
                    y_line = current_top_y_mm + (max_pw - ld["pos"]) * c.pen_w
                    r = cy - y_line
                    rd.arcs.append((cx, cy, r, ld["lw"], ld["style"]))
                
                # ==========================================
                # 3. MISSING X-MARKS (RADIAL)
                # ==========================================
                if c.show_x_marker:
                    try:
                        t_pos = pos_map[c.oval_data["top"].lower()]
                        b_pos = pos_map[c.oval_data["bot"].lower()]
                        
                        r_top_mark = cy - (current_top_y_mm + (max_pw - max(t_pos, b_pos)) * c.pen_w)
                        r_bot_mark = cy - (current_top_y_mm + (max_pw - min(t_pos, b_pos)) * c.pen_w)
                        
                        r_mid = (r_top_mark + r_bot_mark) / 2.0
                        gap_height = abs(r_top_mark - r_bot_mark)
                        m_size = (gap_height / 2.0) * 0.8
                        
                        mx = x_min_clip + m_size + 0.5
                        dx = mx - cx
                        
                        # Calculate angular position on the arc so it sits perfectly on the margin
                        if abs(dx) <= r_mid:
                            angle = math.asin(dx / r_mid)
                            my = cy - r_mid * math.cos(angle)
                            # Passing rotation angle ensures the 'X' rotates perpendicular to the arc curvature
                            rd.markers.append((mx, my, m_size, angle))
                    except KeyError:
                        pass

                # ==========================================
                # 2. SLANT LINE COVERAGE
                # ==========================================
                for s in c.slants:
                    alpha = math.radians(s["angle"])
                    spacing_mm = s["spacing"]
                    
                    # Convert straight horizontal spacing into polar arc spacing along the Base Radius
                    d_phi = spacing_mm / c.radius
                    max_phi = c.pw / c.radius + 0.5
                    n_lines = int(max_phi / d_phi) + 2
                    
                    for i in range(-n_lines, n_lines + 1):
                        phi = i * d_phi
                        x_base = cx + c.radius * math.sin(phi)
                        y_base = cy - c.radius * math.cos(phi)
                        
                        # Optimization: cull slants totally off-page
                        if x_base < x_min_clip - 50 or x_base > x_max_clip + 50:
                            continue
                        
                        # Find intersection 't' distances along parametric slant line
                        # Project slants perfectly between uppermost and lowermost radii of this group.
                        disc_top = r_top_bound**2 - (c.radius * math.sin(alpha))**2
                        disc_bot = r_bot_bound**2 - (c.radius * math.sin(alpha))**2
                        
                        if disc_top < 0 or disc_bot < 0:
                            continue  # Slant angle too extreme to intersect bounds
                            
                        t_top = -c.radius * math.cos(alpha) + math.sqrt(disc_top)
                        t_bot = -c.radius * math.cos(alpha) + math.sqrt(disc_bot)
                        
                        p1_x = x_base + t_top * math.sin(phi + alpha)
                        p1_y = y_base - t_top * math.cos(phi + alpha)
                        
                        p2_x = x_base + t_bot * math.sin(phi + alpha)
                        p2_y = y_base - t_bot * math.cos(phi + alpha)
                        
                        if max(p1_x, p2_x) >= x_min_clip and min(p1_x, p2_x) <= x_max_clip:
                            rd.slants.append((p1_x, p1_y, p2_x, p2_y, s["lw"], s["style"]))
                            
                # ==========================================
                # 4. OVALS [Graceful Degradation]
                # ==========================================
                # True radial ovals require complex non-linear polar warping (bezier curves mapped to polar space).
                # Drawing standard Cartesian ellipses on a curved baseline produces visually broken/tilted results.
                # Therefore, we gracefully skip oval generation here to preserve the mathematical integrity of the sheet.

            else:
                # --- Standard Linear Mode ---
                try:
                    t_pos = pos_map[c.oval_data["top"].lower()]
                    b_pos = pos_map[c.oval_data["bot"].lower()]
                    
                    m_top_y = current_top_y_mm + (max_pw - max(t_pos, b_pos)) * c.pen_w
                    m_bot_y = current_top_y_mm + (max_pw - min(t_pos, b_pos)) * c.pen_w
                    
                    # Generate dynamic X-marker
                    if c.show_x_marker:
                        mid_y_mm = (m_top_y + m_bot_y) / 2.0
                        gap_height = m_bot_y - m_top_y
                        m_size_mm = (gap_height / 2.0) * 0.6
                        marker_x = x_min_clip + m_size_mm + 0.5
                        rd.markers.append((marker_x, mid_y_mm, m_size_mm, 0.0))
                    
                    # Generate Ovals
                    if c.oval_data["enabled"]:
                        o_top_y, o_bot_y = m_top_y, m_bot_y
                        o_h = o_bot_y - o_top_y
                        o_w = o_h * c.oval_data["ratio"]
                    else:
                        o_w, o_h = None, None
                        
                except KeyError:
                    o_w, o_h = None, None # Skip if lines are invalid

                for ld in c.lines:
                    y_line = current_top_y_mm + (max_pw - ld["pos"]) * c.pen_w
                    rd.horizontals.append((y_line, ld["lw"], ld["style"]))

                slant_anchor_x_mm = x_min_clip + 1.0
                base_y_mm = current_top_y_mm + (max_pw - pos_map.get("base", 0)) * c.pen_w

                for s in c.slants:
                    rad = math.radians(s["angle"])
                    spacing = s["spacing"]
                    n_min = math.floor((x_min_clip - slant_anchor_x_mm) / spacing) - 2
                    n_max = math.ceil((x_max_clip - slant_anchor_x_mm) / spacing) + 2
                    
                    for n in range(n_min, n_max):
                        x_cross = slant_anchor_x_mm + n * spacing
                        x_top = x_cross + (base_y_mm - y_min) * math.tan(rad)
                        x_bottom = x_cross + (base_y_mm - y_max) * math.tan(rad)
                        
                        p1_x, p1_y, p2_x, p2_y = x_top, y_min, x_bottom, y_max
                        if p1_x > p2_x: p1_x, p1_y, p2_x, p2_y = p2_x, p2_y, p1_x, p1_y
                            
                        if p2_x < x_min_clip or p1_x > x_max_clip: continue
                            
                        orig_p1_x, orig_p1_y, orig_p2_x, orig_p2_y = p1_x, p1_y, p2_x, p2_y
                        
                        if p1_x < x_min_clip:
                            if orig_p2_x != orig_p1_x: p1_y = orig_p1_y + (orig_p2_y - orig_p1_y) * (x_min_clip - orig_p1_x) / (orig_p2_x - orig_p1_x)
                            p1_x = x_min_clip
                        if p2_x > x_max_clip:
                            if orig_p2_x != orig_p1_x: p2_y = orig_p1_y + (orig_p2_y - orig_p1_y) * (x_max_clip - orig_p1_x) / (orig_p2_x - orig_p1_x)
                            p2_x = x_max_clip
                            
                        rd.slants.append((p1_x, p1_y, p2_x, p2_y, s["lw"], s["style"]))
                            
                        if o_w and o_h:
                            cx_o = x_cross + (base_y_mm - (o_top_y + o_h/2)) * math.tan(rad)
                            cy_o = o_top_y + o_h/2
                            shear_offset = (o_h/2) * math.tan(rad)
                            if (cx_o - o_w/2 - abs(shear_offset)) >= x_min_clip and (cx_o + o_w/2 + abs(shear_offset)) <= x_max_clip:
                                rd.ovals.append((cx_o, cy_o, o_w, o_h, rad, 0.1))

            current_top_y_mm += group_h_mm + c.gap
        return rd

class SvgExporter:
    @staticmethod
    def generate(rd: RenderData, json_str="{}") -> str:
        pw, ph = rd.page_width, rd.page_height
        svg = []
        
        svg.append(f'<?xml version="1.0" encoding="UTF-8" standalone="no"?>')
        svg.append(f'<svg width="{pw}mm" height="{ph}mm" viewBox="0 0 {pw} {ph}" xmlns="http://www.w3.org/2000/svg">')
        svg.append(f'  <desc id="calligraphy-metadata">{json_str}</desc>')
        svg.append(f'  <rect width="{pw}" height="{ph}" fill="none"/>')

        def get_dash(style):
            if style == "Dashed": return 'stroke-dasharray="2, 1"'
            if style == "Dotted": return 'stroke-dasharray="1, 1"'
            return ''

        for y, lw, style in rd.horizontals:
            svg.append(f'  <line x1="{rd.margin_h}" y1="{y}" x2="{pw - rd.margin_h}" y2="{y}" '
                       f'stroke="{rd.line_color}" stroke-width="{lw}" {get_dash(style)}/>')
        
        for x1, y1, x2, y2, lw, style in rd.slants:
            svg.append(f'  <line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" '
                       f'stroke="{rd.line_color}" stroke-width="{lw}" {get_dash(style)}/>')
            
        for cx, cy, r, lw, style in rd.arcs:
            dx_clip = (pw/2 - rd.margin_h)
            if dx_clip > r: dx_clip = r
            ang = math.asin(dx_clip / r)
            x1 = cx - r * math.sin(ang)
            x2 = cx + r * math.sin(ang)
            y_arc = cy - r * math.cos(ang)
            # Uses SVG Arc Path (A): rx ry x-axis-rotation large-arc-flag sweep-flag x y
            svg.append(f'  <path d="M {x1} {y_arc} A {r} {r} 0 0 1 {x2} {y_arc}" '
                       f'fill="none" stroke="{rd.line_color}" stroke-width="{lw}" {get_dash(style)}/>')

        for cx, cy, w, h, rad, lw in rd.ovals:
            deg = math.degrees(rad)
            transform = f'transform="translate({cx} {cy}) skewX({-deg}) translate({-cx} {-cy})"'
            svg.append(f'  <ellipse cx="{cx}" cy="{cy}" rx="{w/2}" ry="{h/2}" '
                       f'stroke="{rd.line_color}" stroke-width="{lw}" fill="none" {transform}/>')
            
        for mx, my, m_size, angle in rd.markers:
            deg = math.degrees(angle)
            svg.append(f'  <g transform="rotate({deg} {mx} {my})">')
            svg.append(f'    <line x1="{mx - m_size}" y1="{my - m_size}" '
                       f'x2="{mx + m_size}" y2="{my + m_size}" '
                       f'stroke="{rd.line_color}" stroke-width="0.5"/>')
            svg.append(f'    <line x1="{mx - m_size}" y1="{my + m_size}" '
                       f'x2="{mx + m_size}" y2="{my - m_size}" '
                       f'stroke="{rd.line_color}" stroke-width="0.5"/>')
            svg.append(f'  </g>')
            
        svg.append('</svg>')
        return "\n".join(svg)

# -----------------------------------------------------------------------------
# User Interface & Interaction
# -----------------------------------------------------------------------------

class CalligraphyApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Calligraphy Guide Sheet Generator")
        self.geometry("1450x950")
        ctk.set_appearance_mode("dark")
        
        self.backend_state: Optional[GridConfig] = None
        self._update_job = None
        self.line_rows = []
        self.slant_rows = []
        self.current_line_color = CONFIG["default_line_color"]
        self.current_dot_color = CONFIG["default_dot_color"]
        
        self.zoom_level = 1.0
        self.pan_x = 0; self.pan_y = 0
        self._drag_data = {"x": 0, "y": 0}
        
        self._setup_layout()
        self._build_sidebar()
        self._build_canvas()
        self._bind_numeric_inputs()
        
        self.load_preset("Italic (10°)")
        self.bind("<Configure>", self._on_resize)
        self.after(200, self.update_preview)

    def _setup_layout(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.sidebar_frame = ctk.CTkScrollableFrame(self, width=480, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=CONFIG["bg_color"])
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)

    def _build_sidebar(self):
        r = 0
        
        hdr_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        hdr_frame.grid(row=r, column=0, padx=20, pady=(20, 10), sticky="ew"); r += 1
        ctk.CTkLabel(hdr_frame, text="Parameters", font=ctk.CTkFont(size=20, weight="bold")).pack(side="left")
        ctk.CTkOptionMenu(hdr_frame, values=list(PRESETS.keys()), command=self.load_preset, width=150).pack(side="right")
        
        bg_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        bg_frame.grid(row=r, column=0, padx=20, pady=5, sticky="ew"); r += 1
        
        ctk.CTkLabel(bg_frame, text="Line Color:").grid(row=0, column=0, sticky="w", pady=2)
        self.btn_color = ctk.CTkButton(bg_frame, text=self.current_line_color, fg_color=self.current_line_color, text_color="black", width=80, command=lambda: self._pick_color("line"))
        self.btn_color.grid(row=0, column=1, padx=10, pady=2)
        
        ctk.CTkLabel(bg_frame, text="Dot Color:").grid(row=0, column=2, sticky="w", pady=2)
        self.btn_dot_color = ctk.CTkButton(bg_frame, text=self.current_dot_color, fg_color=self.current_dot_color, text_color="black", width=80, command=lambda: self._pick_color("dot"))
        self.btn_dot_color.grid(row=0, column=3, padx=10, pady=2)
        
        self.chk_center = ctk.CTkCheckBox(bg_frame, text="Center Line", command=self._debounce_update)
        self.chk_center.grid(row=1, column=0, columnspan=2, sticky="w", pady=10)
        
        ctk.CTkLabel(bg_frame, text="Dot Grid (mm):").grid(row=1, column=2, sticky="w", pady=10)
        self.ent_dot_gap = ctk.CTkEntry(bg_frame, width=50); self.ent_dot_gap.insert(0, "0")
        self.ent_dot_gap.grid(row=1, column=3, sticky="w", padx=10, pady=10)

        ctk.CTkLabel(bg_frame, text="Dot Size (mm):").grid(row=2, column=2, sticky="w", pady=(0, 10))
        self.ent_dot_size = ctk.CTkEntry(bg_frame, width=50); self.ent_dot_size.insert(0, "0.2")
        self.ent_dot_size.grid(row=2, column=3, sticky="w", padx=10, pady=(0, 10))

        ctk.CTkLabel(self.sidebar_frame, text="Page & Layout (mm or in)", font=ctk.CTkFont(weight="bold")).grid(row=r, column=0, padx=20, pady=(10, 0), sticky="w"); r += 1
        frame_page = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        frame_page.grid(row=r, column=0, padx=20, pady=5, sticky="ew"); r += 1
        
        ctk.CTkLabel(frame_page, text="Width:").grid(row=0, column=0, sticky="w", pady=5)
        self.ent_pw = ctk.CTkEntry(frame_page, width=70); self.ent_pw.insert(0, CONFIG["default_page_width"]); self.ent_pw.grid(row=0, column=1, sticky="e")
        ctk.CTkLabel(frame_page, text="Height:").grid(row=0, column=2, sticky="w", padx=(10,0)); 
        self.ent_ph = ctk.CTkEntry(frame_page, width=70); self.ent_ph.insert(0, CONFIG["default_page_height"]); self.ent_ph.grid(row=0, column=3, sticky="e")
        ctk.CTkLabel(frame_page, text="V Marg:").grid(row=1, column=0, sticky="w", pady=5)
        self.ent_mv = ctk.CTkEntry(frame_page, width=70); self.ent_mv.insert(0, CONFIG["default_margin_v"]); self.ent_mv.grid(row=1, column=1, sticky="e")
        ctk.CTkLabel(frame_page, text="H Marg:").grid(row=1, column=2, sticky="w", padx=(10,0)); 
        self.ent_mh = ctk.CTkEntry(frame_page, width=70); self.ent_mh.insert(0, CONFIG["default_margin_h"]); self.ent_mh.grid(row=1, column=3, sticky="e")
        
        ctk.CTkLabel(self.sidebar_frame, text="Metrics, Markers & Ovals", font=ctk.CTkFont(weight="bold")).grid(row=r, column=0, padx=20, pady=(15, 0), sticky="w"); r += 1
        frame_metrics = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        frame_metrics.grid(row=r, column=0, padx=20, pady=5, sticky="ew"); r += 1
        
        ctk.CTkLabel(frame_metrics, text="Pen Width:").grid(row=0, column=0, sticky="w")
        self.ent_pen = ctk.CTkEntry(frame_metrics, width=70); self.ent_pen.grid(row=0, column=1, sticky="e")
        ctk.CTkLabel(frame_metrics, text="Gap:").grid(row=0, column=2, sticky="w", padx=(10,0))
        self.ent_gap = ctk.CTkEntry(frame_metrics, width=70); self.ent_gap.grid(row=0, column=3, sticky="e")
        
        # Ovals and X-Marker Toggles
        self.chk_oval = ctk.CTkCheckBox(frame_metrics, text="Draw Ovals", command=self._debounce_update)
        self.chk_oval.grid(row=1, column=0, sticky="w", pady=(10,0))
        
        self.chk_x_marker = ctk.CTkCheckBox(frame_metrics, text="Draw X-Mark", command=self._debounce_update)
        self.chk_x_marker.grid(row=1, column=1, columnspan=2, sticky="w", padx=(5,0), pady=(10,0))
        self.chk_x_marker.select() # Default Enabled
        
        ctk.CTkLabel(frame_metrics, text="Ratio:").grid(row=2, column=0, sticky="w", pady=(10,0))
        self.ent_oval_ratio = ctk.CTkEntry(frame_metrics, width=70); self.ent_oval_ratio.grid(row=2, column=1, sticky="e", pady=(10,0))
        ctk.CTkLabel(frame_metrics, text="Bound Top:").grid(row=3, column=0, sticky="w", pady=(5,0))
        self.ent_oval_top = ctk.CTkEntry(frame_metrics, width=70); self.ent_oval_top.grid(row=3, column=1, sticky="e", pady=(5,0))
        ctk.CTkLabel(frame_metrics, text="Bound Bot:").grid(row=3, column=2, sticky="w", padx=(10,0), pady=(5,0))
        self.ent_oval_bot = ctk.CTkEntry(frame_metrics, width=70); self.ent_oval_bot.grid(row=3, column=3, sticky="e", pady=(5,0))
        
        self.chk_radial = ctk.CTkCheckBox(frame_metrics, text="Envelope Mode (Radial)", command=self._debounce_update)
        self.chk_radial.grid(row=4, column=0, columnspan=2, sticky="w", pady=(10,0))
        ctk.CTkLabel(frame_metrics, text="Radius(mm):").grid(row=4, column=2, sticky="w", padx=(10,0), pady=(10,0))
        self.ent_radius = ctk.CTkEntry(frame_metrics, width=70); self.ent_radius.insert(0, "200")
        self.ent_radius.grid(row=4, column=3, sticky="e", pady=(10,0))

        ctk.CTkLabel(self.sidebar_frame, text="Horizontal Lines", font=ctk.CTkFont(weight="bold")).grid(row=r, column=0, padx=20, pady=(15, 0), sticky="w"); r += 1
        self.frame_lines_grid = ctk.CTkFrame(self.sidebar_frame, fg_color="#2B2B2B")
        self.frame_lines_grid.grid(row=r, column=0, padx=20, pady=5, sticky="ew"); r += 1
        ctk.CTkButton(self.sidebar_frame, text="➕ Add Line", command=self._add_line_row, width=100, fg_color="#444444").grid(row=r, column=0, padx=20, pady=(0,10), sticky="w"); r += 1

        ctk.CTkLabel(self.sidebar_frame, text="Slant Overlays", font=ctk.CTkFont(weight="bold")).grid(row=r, column=0, padx=20, pady=(15, 0), sticky="w"); r += 1
        self.frame_slants_grid = ctk.CTkFrame(self.sidebar_frame, fg_color="#2B2B2B")
        self.frame_slants_grid.grid(row=r, column=0, padx=20, pady=5, sticky="ew"); r += 1
        ctk.CTkButton(self.sidebar_frame, text="➕ Add Slant", command=self._add_slant_row, width=100, fg_color="#444444").grid(row=r, column=0, padx=20, pady=(0,10), sticky="w"); r += 1
        
        frame_actions = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        frame_actions.grid(row=r, column=0, padx=20, pady=20, sticky="ew"); r += 1
        frame_actions.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(frame_actions, text="💾 Save SVG", command=self.save_svg, fg_color="#1E90FF").grid(row=0, column=0, pady=5, sticky="ew")
        ctk.CTkButton(frame_actions, text="📂 Load SVG", command=self.load_svg, fg_color="#555555").grid(row=1, column=0, pady=5, sticky="ew")
        ctk.CTkButton(frame_actions, text="🖨️ Print", command=self.print_svg, fg_color="#2E8B57").grid(row=2, column=0, pady=(15, 5), sticky="ew")

    def _bind_numeric_inputs(self):
        """DRY Principle: Bind all numeric input widgets to the update handler."""
        self.numeric_inputs = [
            self.ent_pw, self.ent_ph, self.ent_mv, self.ent_mh,
            self.ent_pen, self.ent_gap, self.ent_dot_gap, self.ent_dot_size,
            self.ent_oval_ratio, self.ent_oval_top, self.ent_oval_bot, self.ent_radius
        ]
        for ent in self.numeric_inputs:
            ent.bind("<KeyRelease>", self._debounce_update)

    def _build_canvas(self):
        self.toolbar = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.toolbar.grid(row=0, column=0, sticky="ew", padx=20, pady=(15, 0))
        ctk.CTkButton(self.toolbar, text="⟲ Reset View", command=self._reset_zoom, width=80, fg_color="#444444").pack(side="right", padx=5)
        ctk.CTkButton(self.toolbar, text="🔍 +", command=self._zoom_in, width=40, fg_color="#444444").pack(side="right", padx=5)
        ctk.CTkButton(self.toolbar, text="🔍 -", command=self._zoom_out, width=40, fg_color="#444444").pack(side="right", padx=5)

        self.canvas = tk.Canvas(self.main_frame, bg=CONFIG["bg_color"], highlightthickness=0)
        self.canvas.grid(row=1, column=0, sticky="nsew", padx=20, pady=15)
        
        self.canvas.bind("<ButtonPress-1>", self._on_drag_start)
        self.canvas.bind("<B1-Motion>", self._on_drag_motion)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)

    # --- Interaction & Presets ---
    def _pick_color(self, target):
        init_color = self.current_line_color if target == "line" else self.current_dot_color
        color = colorchooser.askcolor(title=f"Choose {target.title()} Color", initialcolor=init_color)[1]
        if color:
            if target == "line":
                self.current_line_color = color
                self.btn_color.configure(text=color, fg_color=color)
            else:
                self.current_dot_color = color
                self.btn_dot_color.configure(text=color, fg_color=color)
            self._debounce_update()

    def load_preset(self, preset_name):
        data = PRESETS.get(preset_name)
        if not data: return
        self._set_ui_state(data, partial=True)

    def _add_line_row(self, data=None):
        if data is None: data = {"name": "", "pos": "", "lw": "0.10", "style": "Solid"}
        row = ctk.CTkFrame(self.frame_lines_grid, fg_color="transparent")
        row.pack(fill="x", padx=5, pady=2)
        r_dict = {"frame": row}
        
        r_dict["name"] = ctk.CTkEntry(row, width=80); r_dict["name"].insert(0, data["name"]); r_dict["name"].grid(row=0, column=0, padx=2)
        r_dict["pos"] = ctk.CTkEntry(row, width=50); r_dict["pos"].insert(0, data["pos"]); r_dict["pos"].grid(row=0, column=1, padx=2)
        r_dict["lw"] = ctk.CTkEntry(row, width=50); r_dict["lw"].insert(0, data["lw"]); r_dict["lw"].grid(row=0, column=2, padx=2)
        r_dict["style"] = ctk.CTkOptionMenu(row, values=["Solid", "Dashed", "Dotted"], width=80, command=self._debounce_update); r_dict["style"].set(data.get("style", "Solid")); r_dict["style"].grid(row=0, column=3, padx=2)
        
        btn_del = ctk.CTkButton(row, text="🗑", width=30, fg_color="#8B0000", command=lambda: self._delete_row(r_dict, self.line_rows))
        btn_del.grid(row=0, column=4, padx=5)
        
        for k in ["name", "pos", "lw"]: r_dict[k].bind("<KeyRelease>", self._debounce_update)
        self.line_rows.append(r_dict); self._debounce_update()

    def _add_slant_row(self, data=None):
        if data is None: data = {"angle": "10", "spacing": "5", "lw": "0.10", "style": "Solid"}
        row = ctk.CTkFrame(self.frame_slants_grid, fg_color="transparent")
        row.pack(fill="x", padx=5, pady=2)
        r_dict = {"frame": row}
        
        r_dict["angle"] = ctk.CTkEntry(row, width=60); r_dict["angle"].insert(0, data["angle"]); r_dict["angle"].grid(row=0, column=0, padx=2)
        r_dict["spacing"] = ctk.CTkEntry(row, width=70); r_dict["spacing"].insert(0, data["spacing"]); r_dict["spacing"].grid(row=0, column=1, padx=2)
        r_dict["lw"] = ctk.CTkEntry(row, width=50); r_dict["lw"].insert(0, data["lw"]); r_dict["lw"].grid(row=0, column=2, padx=2)
        r_dict["style"] = ctk.CTkOptionMenu(row, values=["Solid", "Dashed", "Dotted"], width=80, command=self._debounce_update); r_dict["style"].set(data.get("style", "Solid")); r_dict["style"].grid(row=0, column=3, padx=2)
        
        btn_del = ctk.CTkButton(row, text="🗑", width=30, fg_color="#8B0000", command=lambda: self._delete_row(r_dict, self.slant_rows))
        btn_del.grid(row=0, column=4, padx=5)
        
        for k in ["angle", "spacing", "lw"]: r_dict[k].bind("<KeyRelease>", self._debounce_update)
        self.slant_rows.append(r_dict); self._debounce_update()

    def _delete_row(self, r_dict, arr):
        r_dict["frame"].destroy(); arr.remove(r_dict); self._debounce_update()

    # --- Viewport ---
    def _on_drag_start(self, event): self._drag_data["x"] = event.x; self._drag_data["y"] = event.y
    def _on_drag_motion(self, event):
        self.pan_x += event.x - self._drag_data["x"]; self.pan_y += event.y - self._drag_data["y"]
        self._drag_data["x"] = event.x; self._drag_data["y"] = event.y; self.update_preview()
    def _on_mousewheel(self, event): self._zoom_math(1.15 if event.delta > 0 else 1/1.15, event.x, event.y)
    def _zoom_in(self): self._zoom_math(1.15, self.canvas.winfo_width()/2, self.canvas.winfo_height()/2)
    def _zoom_out(self): self._zoom_math(1/1.15, self.canvas.winfo_width()/2, self.canvas.winfo_height()/2)
    def _reset_zoom(self): self.zoom_level = 1.0; self.pan_x = 0; self.pan_y = 0; self.update_preview()
    
    def _zoom_math(self, factor, mx, my):
        if not self.backend_state: return
        pw, ph = self.backend_state.pw, self.backend_state.ph
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        base_scale = min((ch - 40) / ph, (cw - 40) / pw)
        
        old_sc = base_scale * self.zoom_level
        old_ox = (cw - pw * old_sc) / 2 + self.pan_x
        old_oy = (ch - ph * old_sc) / 2 + self.pan_y

        self.zoom_level *= factor
        new_sc = base_scale * self.zoom_level

        self.pan_x = (mx - (mx - old_ox) * (new_sc / old_sc)) - (cw - pw * new_sc) / 2
        self.pan_y = (my - (my - old_oy) * (new_sc / old_sc)) - (ch - ph * new_sc) / 2
        self.update_preview()

    def _map_coords(self, x: float, y: float, scale: float, offset_x: float, offset_y: float) -> Tuple[float, float]:
        """Translates mm coordinates to canvas pixels respecting zoom and pan."""
        return offset_x + (x * scale), offset_y + (y * scale)

    # --- State Management ---
    def _parse_val(self, val_str):
        val_str = str(val_str).strip().lower()
        if val_str.endswith(('in', 'inch', '"')):
            try: return float(re.sub(r'[a-z"]', '', val_str).strip()) * CONFIG["in_to_mm"]
            except ValueError: return None
        elif val_str.endswith('mm'):
            try: return float(re.sub(r'mm', '', val_str).strip())
            except ValueError: return None
        else:
            try: return float(val_str)
            except ValueError: return None

    def _get_ui_state(self):
        return {
            "page_width": self.ent_pw.get().strip(), "page_height": self.ent_ph.get().strip(),
            "margin_v": self.ent_mv.get().strip(), "margin_h": self.ent_mh.get().strip(),
            "pen_width": self.ent_pen.get().strip(), "group_gap": self.ent_gap.get().strip(),
            "line_color": self.current_line_color, "dot_color": self.current_dot_color,
            "show_center": self.chk_center.get() == 1,
            "show_x_marker": self.chk_x_marker.get() == 1,
            "dot_gap": self.ent_dot_gap.get().strip(), "dot_size": self.ent_dot_size.get().strip(),
            "radial": self.chk_radial.get() == 1, "radius": self.ent_radius.get().strip(),
            "oval_enabled": self.chk_oval.get() == 1,
            "oval_ratio": self.ent_oval_ratio.get().strip(),
            "oval_top": self.ent_oval_top.get().strip(), "oval_bot": self.ent_oval_bot.get().strip(),
            "lines": [{"name": r["name"].get().strip(), "pos": r["pos"].get().strip(), "lw": r["lw"].get().strip(), "style": r["style"].get()} for r in self.line_rows],
            "slants": [{"angle": r["angle"].get().strip(), "spacing": r["spacing"].get().strip(), "lw": r["lw"].get().strip(), "style": r["style"].get()} for r in self.slant_rows]
        }

    def _set_ui_state(self, state, partial=False):
        def set_ent(ent, key): 
            if key in state: ent.delete(0, tk.END); ent.insert(0, str(state[key]))
        
        if not partial:
            set_ent(self.ent_pw, "page_width"); set_ent(self.ent_ph, "page_height")
            set_ent(self.ent_mv, "margin_v"); set_ent(self.ent_mh, "margin_h")
            if "line_color" in state:
                self.current_line_color = state["line_color"]
                self.btn_color.configure(text=self.current_line_color, fg_color=self.current_line_color)
            if "dot_color" in state:
                self.current_dot_color = state["dot_color"]
                self.btn_dot_color.configure(text=self.current_dot_color, fg_color=self.current_dot_color)
            if state.get("show_center", False): self.chk_center.select()
            else: self.chk_center.deselect()
            if state.get("show_x_marker", True): self.chk_x_marker.select()
            else: self.chk_x_marker.deselect()
            if state.get("radial", False): self.chk_radial.select()
            else: self.chk_radial.deselect()
            set_ent(self.ent_dot_gap, "dot_gap")
            set_ent(self.ent_dot_size, "dot_size")
            set_ent(self.ent_radius, "radius")

        set_ent(self.ent_pen, "pen_width"); set_ent(self.ent_gap, "group_gap")
        
        if "oval_enabled" in state:
            if state["oval_enabled"]: self.chk_oval.select()
            else: self.chk_oval.deselect()
        set_ent(self.ent_oval_ratio, "oval_ratio")
        set_ent(self.ent_oval_top, "oval_top"); set_ent(self.ent_oval_bot, "oval_bot")
        
        if "lines" in state:
            for r in list(self.line_rows): self._delete_row(r, self.line_rows)
            for ld in state["lines"]: self._add_line_row(ld)
        if "slants" in state:
            for r in list(self.slant_rows): self._delete_row(r, self.slant_rows)
            for sd in state["slants"]: self._add_slant_row(sd)
        
        self.update_preview()

    def _parse_inputs_to_mm(self) -> Optional[GridConfig]:
        pw, ph = self._parse_val(self.ent_pw.get()), self._parse_val(self.ent_ph.get())
        mv, mh = self._parse_val(self.ent_mv.get()), self._parse_val(self.ent_mh.get())
        pen, gap = self._parse_val(self.ent_pen.get()), self._parse_val(self.ent_gap.get())
        dot_gap = self._parse_val(self.ent_dot_gap.get()) or 0.0
        dot_size = self._parse_val(self.ent_dot_size.get()) or 0.2
        radius = self._parse_val(self.ent_radius.get()) or 200.0
        
        if None in (pw, ph, mv, mh, pen, gap) or any(v <= 0 for v in [pw, ph, pen]) or mv < 0 or mh < 0 or gap < 0: return None 

        oval_ratio_val = self._parse_val(self.ent_oval_ratio.get())
        if oval_ratio_val is None or oval_ratio_val <= 0: return None

        oval_data = {
            "enabled": self.chk_oval.get() == 1,
            "ratio": oval_ratio_val,
            "top": self.ent_oval_top.get().strip(),
            "bot": self.ent_oval_bot.get().strip()
        }

        l_data, s_data = [], []
        for r in self.line_rows:
            n = r["name"].get().strip()
            if not n: continue
            try:
                lw = self._parse_val(r["lw"].get())
                if lw is not None: l_data.append({"name": n, "pos": float(r["pos"].get()), "lw": lw, "style": r["style"].get()})
            except ValueError: continue
            
        for r in self.slant_rows:
            try:
                spc, lw = self._parse_val(r["spacing"].get()), self._parse_val(r["lw"].get())
                if spc is not None and spc > 0 and lw is not None:
                    s_data.append({"angle": float(r["angle"].get()), "spacing": spc, "lw": lw, "style": r["style"].get()})
            except ValueError: continue
                    
        return GridConfig(
            pw, ph, mv, mh, pen, gap, l_data, s_data, oval_data,
            self.current_line_color, self.current_dot_color,
            self.chk_center.get() == 1, 
            self.chk_x_marker.get() == 1,
            dot_gap, dot_size,
            self.chk_radial.get() == 1, radius
        )

    # --- Live Preview Engine ---
    def _debounce_update(self, event=None):
        if self._update_job: self.after_cancel(self._update_job)
        self._update_job = self.after(300, self.update_preview)

    def _on_resize(self, event):
        if event.widget == self: self._debounce_update()

    def update_preview(self):
        self.canvas.delete("all")
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        if cw <= 1 or ch <= 1: return
            
        self.backend_state = self._parse_inputs_to_mm()
        if not self.backend_state: return

        rd = GeometryEngine.calculate(self.backend_state)
        
        base_sc = min((ch - 40) / rd.page_height, (cw - 40) / rd.page_width)
        sc = base_sc * self.zoom_level
        ox = (cw - (rd.page_width * sc)) / 2 + self.pan_x
        oy = (ch - (rd.page_height * sc)) / 2 + self.pan_y
        
        p_x1, p_y1 = self._map_coords(0, 0, sc, ox, oy)
        p_x2, p_y2 = self._map_coords(rd.page_width, rd.page_height, sc, ox, oy)
        m_x1, m_y1 = self._map_coords(0, rd.margin_v, sc, ox, oy)
        m_x2, m_y2 = self._map_coords(0, rd.page_height - rd.margin_v, sc, ox, oy)
        
        self.canvas.create_rectangle(p_x1, p_y1, p_x2, p_y2, fill=CONFIG["page_color"], outline="")
        
        if rd.dot_gap > 0:
            nx = int((rd.page_width - 2*rd.margin_h) / rd.dot_gap)
            ny = int((rd.page_height - 2*rd.margin_v) / rd.dot_gap)
            r_px = max(0.5, rd.dot_size * sc)
            for i in range(nx + 1):
                for j in range(ny + 1):
                    x = rd.margin_h + i * rd.dot_gap
                    y = rd.margin_v + j * rd.dot_gap
                    mx, my = self._map_coords(x, y, sc, ox, oy)
                    self.canvas.create_oval(mx-r_px, my-r_px, mx+r_px, my+r_px, fill=rd.dot_color, outline="")

        if rd.show_center:
            mx, _ = self._map_coords(rd.page_width/2, 0, sc, ox, oy)
            self.canvas.create_line(mx, m_y1, mx, m_y2, fill=rd.line_color, dash=(4, 4))

        # Render Radial Arcs via robust polyline to avoid Tkinter create_arc bounding box quirks
        for (cx, cy, r, lw, style) in rd.arcs:
            cw_lw = max(1, lw * sc)
            tk_dash = tuple(max(1, int((d * sc) / cw_lw)) for d in CONFIG["style_map_ui"].get(style, ())) if style != "Solid" else ()
            
            points = []
            x_start = max(cx - r, rd.margin_h)
            x_end = min(cx + r, rd.page_width - rd.margin_h)
            
            if x_start >= x_end: continue
                
            steps = 60
            for i in range(steps + 1):
                px = x_start + (x_end - x_start) * (i / steps)
                dx = px - cx
                # floating point safe ceiling
                if abs(dx) > r: dx = r if dx > 0 else -r
                py = cy - math.sqrt(r**2 - dx**2)
                points.extend(self._map_coords(px, py, sc, ox, oy))
                
            if len(points) >= 4:
                self.canvas.create_line(points, fill=rd.line_color, width=cw_lw, dash=tk_dash)

        for (cx, cy, w, h, rad, lw) in rd.ovals:
            pts = []
            steps = 30
            for i in range(steps):
                t = (i / steps) * 2 * math.pi
                px = (w/2) * math.cos(t)
                py = (h/2) * math.sin(t)
                sx = cx + px - py * math.tan(rad) 
                sy = cy + py
                pts.extend(self._map_coords(sx, sy, sc, ox, oy))
            self.canvas.create_polygon(pts, fill="", outline=rd.line_color, width=max(1, lw*sc), smooth=True)

        for (x1, y1, x2, y2, lw, style) in rd.slants:
            cw_lw = max(1, lw * sc)
            tk_dash = tuple(max(1, int((d * sc) / cw_lw)) for d in CONFIG["style_map_ui"].get(style, ())) if style != "Solid" else ()
            p1 = self._map_coords(x1, y1, sc, ox, oy)
            p2 = self._map_coords(x2, y2, sc, ox, oy)
            self.canvas.create_line(p1[0], p1[1], p2[0], p2[1], fill=rd.line_color, width=cw_lw, dash=tk_dash)

        for (y, lw, style) in rd.horizontals:
            cw_lw = max(1, lw * sc)
            tk_dash = tuple(max(1, int((d * sc) / cw_lw)) for d in CONFIG["style_map_ui"].get(style, ())) if style != "Solid" else ()
            p1 = self._map_coords(rd.margin_h, y, sc, ox, oy)
            p2 = self._map_coords(rd.page_width - rd.margin_h, y, sc, ox, oy)
            self.canvas.create_line(p1[0], p1[1], p2[0], p2[1], fill=rd.line_color, dash=tk_dash, width=cw_lw)

        # Dynamic Rotatable X-Markers
        marker_lw = max(1, int(0.5 * sc)) 
        for mx_mm, my_mm, m_size_mm, angle in rd.markers:
            m_size_px = m_size_mm * sc 
            cx_px, cy_px = self._map_coords(mx_mm, my_mm, sc, ox, oy)
            
            s_ang = math.sin(angle)
            c_ang = math.cos(angle)
            
            def rot(px, py):
                return px * c_ang - py * s_ang, px * s_ang + py * c_ang
                
            p1x, p1y = rot(-m_size_px, -m_size_px)
            p2x, p2y = rot(m_size_px, m_size_px)
            p3x, p3y = rot(-m_size_px, m_size_px)
            p4x, p4y = rot(m_size_px, -m_size_px)
            
            self.canvas.create_line(cx_px + p1x, cy_px + p1y, 
                                    cx_px + p2x, cy_px + p2y, 
                                    fill=rd.line_color, width=marker_lw)
            self.canvas.create_line(cx_px + p3x, cy_px + p3y, 
                                    cx_px + p4x, cy_px + p4y, 
                                    fill=rd.line_color, width=marker_lw)

        self.canvas.create_rectangle(p_x1, p_y1, p_x2, p_y2, fill="", outline="#888888")
        self.canvas.create_rectangle(0, 0, cw, p_y1, fill=CONFIG["bg_color"], outline="") 
        self.canvas.create_rectangle(0, p_y2, cw, ch, fill=CONFIG["bg_color"], outline="") 
        self.canvas.create_rectangle(0, p_y1, p_x1, p_y2, fill=CONFIG["bg_color"], outline="") 
        self.canvas.create_rectangle(p_x2, p_y1, cw, p_y2, fill=CONFIG["bg_color"], outline="") 

    # --- IO & Printing ---
    def save_svg(self):
        config = self._parse_inputs_to_mm()
        if not config: return messagebox.showerror("Error", "Invalid parameters.")
        rd = GeometryEngine.calculate(config)
        filepath = filedialog.asksaveasfilename(defaultextension=".svg", filetypes=[("SVG Vector", "*.svg")])
        if filepath:
            with open(filepath, 'w', encoding="utf-8") as f: f.write(SvgExporter.generate(rd, json.dumps(self._get_ui_state())))
            messagebox.showinfo("Success", "SVG saved!")

    def load_svg(self):
        filepath = filedialog.askopenfilename(filetypes=[("SVG Vector", "*.svg")])
        if filepath:
            try:
                with open(filepath, 'r', encoding="utf-8", errors="ignore") as f: content = f.read()
                match = re.search(r'<desc id="calligraphy-metadata">(.*?)</desc>', content, re.DOTALL)
                
                if match:
                    json_str = match.group(1).strip()
                    self._set_ui_state(json.loads(json_str))
                    return
                        
                messagebox.showwarning("Warning", "No readable metadata found in this file.")
            except Exception as e:
                messagebox.showerror("Parse Error", f"Failed to load file state.\n{e}")

    def print_svg(self):
        config = self._parse_inputs_to_mm()
        if not config: return messagebox.showerror("Error", "Invalid parameters.")
        rd = GeometryEngine.calculate(config)
        
        temp_svg = os.path.join(os.environ.get("TEMP", "."), "calligraphy_temp.svg")
        temp_pdf = os.path.join(os.environ.get("TEMP", "."), "calligraphy_temp.pdf")
        
        with open(temp_svg, 'w', encoding="utf-8") as f: f.write(SvgExporter.generate(rd, json.dumps(self._get_ui_state())))
            
        inkscape_cmd = "inkscape"
        if os.name == 'nt':
            paths = [
                r"C:\Program Files\Inkscape\bin\inkscape.exe",
                r"C:\Program Files\Inkscape\inkscape.exe",
                r"C:\Program Files (x86)\Inkscape\bin\inkscape.exe"
            ]
            for p in paths:
                if os.path.exists(p):
                    inkscape_cmd = p
                    break

        try:
            subprocess.run([inkscape_cmd, "--export-filename=" + temp_pdf, temp_svg], check=True, capture_output=True)
            
            sumatra_path = r".\SumatraPDF-3.6.1-64.exe"
            try:
                if os.path.exists(sumatra_path):
                    subprocess.run([sumatra_path, "-print-to-default", "-silent", temp_pdf], check=True)
                else:
                    os.startfile(temp_pdf, "print")
            except OSError as e:
                if getattr(e, 'winerror', None) == 1155:
                    messagebox.showinfo("Manual Print", "No default PDF printer associated. Opening the file to print manually.")
                    os.startfile(temp_pdf)
                else:
                    print(f"Printing failed: {e}")

        except FileNotFoundError:
            messagebox.showerror("Dependency Error", "Inkscape not found. Please verify your Inkscape installation or use 'Save SVG' and print manually.")
        except subprocess.CalledProcessError as e: 
            messagebox.showerror("Print Cancelled", f"Conversion process failed.\n{e.stderr.decode('utf-8') if e.stderr else e}")


if __name__ == "__main__":
    app = CalligraphyApp()
    app.mainloop()