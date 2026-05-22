"""
Rasyon Hesaplama ve Optimizasyon Uygulaması v4.0 - Enhanced Edition
====================================================================
Gelişmiş özellikler:
  • Modern glassmorphism UI
  • PDF rapor çıktısı
  • Kısıt hazır ayarları (presets)
  • Optimizasyon geçmişi
  • Gelişmiş grafik animasyonları
  • Klavye kısayolları
  • Yardım sistemi

Gereksinimler:
    pip install PyQt5 pandas openpyxl pulp reportlab

Çalıştırma:
    python rasyon_optimizer.py
"""

import sys
import math
import json
import os
from datetime import datetime
from io import StringIO

import pandas as pd
from pulp import (
    LpProblem, LpMinimize, LpVariable, lpSum,
    LpStatus, value, PULP_CBC_CMD,
)

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QDoubleSpinBox, QGroupBox, QSplitter, QFileDialog, QMessageBox,
    QCheckBox, QFrame, QScrollArea, QStatusBar, QSizePolicy,
    QAbstractItemView, QTabWidget, QListWidget, QListWidgetItem,
    QDialog, QVBoxLayout as QVBoxLayout2, QTextEdit, QLineEdit,
    QComboBox, QProgressBar, QToolTip,
)
from PyQt5.QtCore import Qt, QRect, QSize, QPoint, QRectF, QTimer, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import (
    QFont, QColor, QPalette, QPainter, QLinearGradient,
    QPen, QBrush, QPainterPath, QFontMetrics, QIcon, QCursor,
    QLinearGradient, QConicalGradient,
)


# ══════════════════════════════════════════════════════════════════════
#  RENK & BOYUT SABİTLERİ
# ══════════════════════════════════════════════════════════════════════
C = {
    # Zemin katmanları - Modern glassmorphism palette
    "bg0":        "#0A0E17",   # en koyu — pencere zemini
    "bg1":        "#111827",   # panel zemini
    "bg2":        "#1A2235",   # kart / groupbox
    "bg3":        "#222D42",   # input / tablo satırı
    "bg_alt":     "#1E2A3E",   # alternatif tablo satırı
    "bg_hover":   "#2A3855",   # hover
    "bg_sel":     "#1E3A6E",   # seçili satır

    # Glassmorphism efekti
    "glass_bg":    "#1E293B",   # cam efekti arka plan
    "glass_border": "#334155",   # cam efekti kenarlık
    "glass_glow":   "#3B82F6",  # parlama efekti

    # Kenarlıklar
    "brd":        "#2E3F5C",   # normal border
    "brd_focus":  "#4A7FD4",   # focus border

    # Aksanlar - Gelişmiş renk paleti
    "green":      "#22C55E",   # başarı / kaba yem
    "green_dim":  "#15803D",
    "green_light": "#4ADE80",   # parlak yeşil
    "blue":       "#3B82F6",   # import / bilgi
    "blue_dim":   "#1D4ED8",
    "blue_light": "#60A5FA",    # parlak mavi
    "amber":      "#F59E0B",   # uyarı / fiyat
    "amber_dim":  "#B45309",
    "amber_light": "#FBBF24",   # parlak amber
    "red":        "#EF4444",   # hata
    "red_light":  "#F87171",    # parlak kırmızı
    "cyan":       "#06B6D4",   # enerji metriki
    "cyan_light": "#22D3EE",    # parlak cyan
    "purple":     "#A78BFA",   # protein metriki
    "purple_light": "#C4B5FD",  # parlak mor
    "pink":       "#EC4899",   # özel vurgular
    "pink_light": "#F472B6",

    # Metin
    "t1":         "#F1F5F9",   # birincil metin
    "t2":         "#94A3B8",   # ikincil / etiket
    "t3":         "#475569",   # devre dışı / placeholder
    "t_head":     "#CBD5E1",   # tablo başlık

    # Gradient renkleri
    "grad_blue_start":  "#1E40AF",
    "grad_blue_end":    "#3B82F6",
    "grad_green_start": "#059669",
    "grad_green_end":   "#10B981",
    "grad_purple_start":"#7C3AED",
    "grad_purple_end":  "#A78BFA",

    # Çizgi grafik çubuklarının renkleri (12 renk döngüsü)
    "bar_colors": [
        "#3B82F6","#22C55E","#F59E0B","#A78BFA",
        "#06B6D4","#F97316","#EC4899","#84CC16",
        "#14B8A6","#8B5CF6","#E879F9","#38BDF8",
    ],
}

# Font boyutları
F_TINY   = 10
F_SMALL  = 11
F_BASE   = 13
F_MED    = 14
F_LARGE  = 16
F_XLARGE = 22
F_HUGE   = 28


# ══════════════════════════════════════════════════════════════════════
#  GLOBAL STYLESHEET - Enhanced v4.0
# ══════════════════════════════════════════════════════════════════════
STYLE = f"""
/* ── Genel ── */
QMainWindow, QWidget {{
    background-color: {C['bg0']};
    color: {C['t1']};
    font-family: 'Segoe UI', 'SF Pro Text', 'Helvetica Neue', sans-serif;
    font-size: {F_BASE}px;
}}

/* ── Glass Panel Effect ── */
QFrame[glass="true"] {{
    background: rgba(30, 41, 59, 0.7);
    border: 1px solid {C['glass_border']};
    border-radius: 12px;
    backdrop-filter: blur(10px);
}}

/* ── GroupBox ── */
QGroupBox {{
    background-color: {C['bg2']};
    border: 1px solid {C['brd']};
    border-radius: 12px;
    margin-top: 20px;
    padding: 18px 16px 16px 16px;
    /* Glassmorphism effect */
    border: 1px solid rgba(59, 130, 246, 0.2);
    background: linear-gradient(135deg, rgba(30,41,59,0.9) 0%, rgba(26,34,53,0.95) 100%);
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 14px;
    padding: 2px 12px;
    background-color: {C['bg2']};
    color: {C['t2']};
    font-size: {F_SMALL}px;
    font-weight: 700;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    border-radius: 6px;
    border: 1px solid {C['brd']};
}}

/* ── Etiketler ── */
QLabel {{ background: transparent; color: {C['t1']}; }}
QLabel[role="caption"] {{
    color: {C['t2']};
    font-size: {F_SMALL}px;
}}
QLabel[role="section"] {{
    color: {C['t_head']};
    font-size: {F_BASE}px;
    font-weight: 700;
}}
QLabel[role="title"] {{
    font-size: {F_LARGE}px;
    font-weight: 700;
    color: {C['t1']};
}}

/* ── Butonlar (varsayılan) ── */
QPushButton {{
    background-color: {C['bg3']};
    color: {C['t1']};
    border: 1px solid {C['brd']};
    border-radius: 8px;
    padding: 8px 20px;
    font-size: {F_BASE}px;
    font-weight: 500;
    min-height: 36px;
    /* Subtle glow */
    border: 1px solid rgba(74, 127, 212, 0.3);
}}
QPushButton:hover {{
    background-color: {C['bg_hover']};
    border-color: {C['brd_focus']};
    /* Add glow effect */
    border: 1px solid {C['blue']};
    box-shadow: 0 0 10px rgba(59, 130, 246, 0.3);
}}
QPushButton:pressed {{ background-color: {C['bg0']}; }}

/* ── Optimize Et butonu ── */
QPushButton#btnOptimize {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                                stop:0 {C['green_dim']}, 
                                stop:0.5 {C['green']}, 
                                stop:1 {C['green_light']});
    color: #ffffff;
    border: none;
    border-radius: 10px;
    font-size: {F_MED}px;
    font-weight: 800;
    letter-spacing: 1px;
    min-height: 50px;
    padding: 0 32px;
    /* Glow effect */
    border: 2px solid rgba(34, 197, 94, 0.5);
    box-shadow: 0 4px 15px rgba(34, 197, 94, 0.4);
}}
QPushButton#btnOptimize:hover  {{ 
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                                stop:0 {C['green']}, 
                                stop:1 {C['green_light']});
    box-shadow: 0 6px 25px rgba(34, 197, 94, 0.6);
    border: 2px solid {C['green_light']};
}}
QPushButton#btnOptimize:pressed{{ background-color: {C['green_dim']}; }}
QPushButton#btnOptimize:disabled {{
    background: {C['bg3']};
    color: {C['t3']};
    border: 1px solid {C['brd']};
    box-shadow: none;
}}

/* ── Import butonu ── */
QPushButton#btnImport {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                                stop:0 {C['blue_dim']}, 
                                stop:1 {C['blue']});
    color: #fff;
    border: none;
    border-radius: 8px;
    font-size: {F_BASE}px;
    font-weight: 700;
    min-height: 38px;
    padding: 0 24px;
    border: 1px solid rgba(59, 130, 246, 0.4);
    box-shadow: 0 3px 12px rgba(59, 130, 246, 0.3);
}}
QPushButton#btnImport:hover  {{ 
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                                stop:0 {C['blue']}, 
                                stop:1 {C['blue_light']});
    box-shadow: 0 5px 20px rgba(59, 130, 246, 0.5);
}}
QPushButton#btnImport:pressed{{ background-color: {C['blue_dim']}; }}

/* ── Sıfırla butonu ── */
QPushButton#btnReset {{
    background-color: transparent;
    color: {C['red']};
    border: 1px solid {C['red']};
    border-radius: 8px;
    font-size: {F_BASE}px;
    font-weight: 600;
    min-height: 50px;
    padding: 0 20px;
}}
QPushButton#btnReset:hover {{ 
    background-color: rgba(239,68,68,0.12);
    box-shadow: 0 0 12px rgba(239, 68, 68, 0.3);
}}

/* ── Export butonu ── */
QPushButton#btnExport {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                                stop:0 {C['amber_dim']}, 
                                stop:1 {C['amber']});
    color: #000;
    border: none;
    border-radius: 8px;
    font-size: {F_BASE}px;
    font-weight: 700;
    min-height: 36px;
    padding: 0 22px;
    box-shadow: 0 3px 12px rgba(245, 158, 11, 0.3);
}}
QPushButton#btnExport:hover  {{ 
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                                stop:0 {C['amber']}, 
                                stop:1 {C['amber_light']});
    box-shadow: 0 5px 18px rgba(245, 158, 11, 0.5);
}}
QPushButton#btnExport:disabled {{
    background-color: {C['bg3']};
    color: {C['t3']};
}}

/* ── Preset butonları ── */
QPushButton#btnPreset {{
    background: {C['bg3']};
    color: {C['purple']};
    border: 1px solid {C['purple']};
    border-radius: 8px;
    padding: 6px 16px;
    font-size: {F_SMALL}px;
    font-weight: 600;
    min-height: 32px;
}}
QPushButton#btnPreset:hover {{
    background: rgba(167, 139, 250, 0.15);
    box-shadow: 0 0 12px rgba(167, 139, 250, 0.4);
}}

/* ── Tümünü Seç / Temizle ── */
QPushButton#btnSelAll {{
    background-color: transparent;
    color: {C['green']};
    border: 1px solid {C['green']};
    border-radius: 6px;
    padding: 4px 16px;
    font-size: {F_SMALL}px;
    font-weight: 600;
    min-height: 30px;
}}
QPushButton#btnSelAll:hover {{ 
    background-color: rgba(34,197,94,0.15);
    box-shadow: 0 0 10px rgba(34, 197, 94, 0.3);
}}
QPushButton#btnDesel {{
    background-color: transparent;
    color: {C['t2']};
    border: 1px solid {C['brd']};
    border-radius: 6px;
    padding: 4px 16px;
    font-size: {F_SMALL}px;
    font-weight: 600;
    min-height: 30px;
}}
QPushButton#btnDesel:hover {{ 
    background-color: {C['bg_hover']};
    border-color: {C['blue']};
}}

/* ── Tablolar ── */
QTableWidget {{
    background-color: {C['bg1']};
    border: 1px solid {C['brd']};
    border-radius: 10px;
    gridline-color: {C['brd']};
    color: {C['t1']};
    selection-background-color: {C['bg_sel']};
    selection-color: {C['t1']};
    outline: none;
    font-size: {F_BASE}px;
    /* Subtle inner glow */
    border: 1px solid rgba(74, 127, 212, 0.2);
}}
QTableWidget::item {{
    padding: 10px 14px;
    border: none;
}}
QTableWidget::item:selected {{ background-color: {C['bg_sel']}; }}
QTableWidget::item:alternate {{ background-color: {C['bg_alt']}; }}
QTableWidget::item:hover {{
    background-color: rgba(59, 130, 246, 0.1);
}}
QHeaderView::section {{
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                stop:0 {C['bg2']}, 
                                stop:1 {C['bg3']});
    color: {C['t2']};
    padding: 12px 14px;
    border: none;
    border-right: 1px solid {C['brd']};
    border-bottom: 2px solid {C['blue']};
    font-size: {F_SMALL}px;
    font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
}}
QHeaderView::section:last {{ border-right: none; }}
QHeaderView::section:hover {{
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                stop:0 {C['bg_hover']}, 
                                stop:1 {C['bg2']});
}}

/* ── SpinBox ── */
QDoubleSpinBox, QSpinBox {{
    background-color: {C['bg3']};
    color: {C['t1']};
    border: 1.5px solid {C['brd']};
    border-radius: 8px;
    padding: 8px 12px;
    font-size: {F_MED}px;
    min-width: 110px;
    min-height: 36px;
}}
QDoubleSpinBox:focus, QSpinBox:focus {{
    border-color: {C['blue']};
    background-color: {C['bg_hover']};
    box-shadow: 0 0 10px rgba(59, 130, 246, 0.3);
}}
QDoubleSpinBox::up-button, QDoubleSpinBox::down-button,
QSpinBox::up-button,       QSpinBox::down-button {{
    background-color: {C['bg_hover']};
    border: none;
    width: 22px;
    border-radius: 6px;
}}
QDoubleSpinBox::up-button:hover, QDoubleSpinBox::down-button:hover,
QSpinBox::up-button:hover,       QSpinBox::down-button:hover {{
    background-color: {C['blue']};
}}

/* ── ScrollBar ── */
QScrollBar:vertical {{
    background: {C['bg0']};
    width: 8px;
    border-radius: 4px;
    margin: 3px 2px;
}}
QScrollBar::handle:vertical {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                                stop:0 {C['brd']}, 
                                stop:1 {C['blue_dim']});
    border-radius: 4px;
    min-height: 32px;
}}
QScrollBar::handle:vertical:hover {{ background: {C['blue']}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar:horizontal {{
    background: {C['bg0']};
    height: 8px;
    border-radius: 4px;
}}
QScrollBar::handle:horizontal {{
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                stop:0 {C['brd']}, 
                                stop:1 {C['blue_dim']});
    border-radius: 4px;
}}

/* ── Tab ── */
QTabWidget::pane {{
    border: 1px solid {C['brd']};
    border-radius: 10px;
    background: {C['bg1']};
}}
QTabBar::tab {{
    background: {C['bg2']};
    color: {C['t2']};
    border: 1px solid {C['brd']};
    border-bottom: none;
    border-radius: 8px 8px 0 0;
    padding: 10px 24px;
    font-size: {F_BASE}px;
    font-weight: 600;
    margin-right: 4px;
}}
QTabBar::tab:selected {{
    background: {C['bg1']};
    color: {C['t1']};
    border-bottom: 3px solid {C['blue']};
    /* Glow effect */
    box-shadow: 0 0 15px rgba(59, 130, 246, 0.3);
}}
QTabBar::tab:hover:!selected {{ 
    background: {C['bg_hover']};
    color: {C['t1']};
}}

/* ── Status Bar ── */
QStatusBar {{
    background-color: {C['bg1']};
    color: {C['t2']};
    border-top: 1px solid {C['brd']};
    font-size: {F_SMALL}px;
    padding: 4px 12px;
    min-height: 28px;
}}

/* ── Splitter ── */
QSplitter::handle {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                                stop:0 {C['bg0']}, 
                                stop:0.5 {C['blue_dim']}, 
                                stop:1 {C['bg0']});
}}
QSplitter::handle:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                                stop:0 {C['bg0']}, 
                                stop:0.5 {C['blue']}, 
                                stop:1 {C['bg0']});
}}

/* ── CheckBox ── */
QCheckBox::indicator {{
    width: 18px; height: 18px;
    border-radius: 5px;
    border: 2px solid {C['brd']};
    background: {C['bg3']};
}}
QCheckBox::indicator:checked {{
    background-color: {C['green']};
    border-color: {C['green_light']};
    /* Glow effect */
    box-shadow: 0 0 8px rgba(34, 197, 94, 0.5);
}}
QCheckBox::indicator:hover {{ 
    border-color: {C['blue']};
    box-shadow: 0 0 6px rgba(59, 130, 246, 0.3);
}}

/* ── ComboBox ── */
QComboBox {{
    background-color: {C['bg3']};
    color: {C['t1']};
    border: 1px solid {C['brd']};
    border-radius: 8px;
    padding: 8px 14px;
    font-size: {F_BASE}px;
    min-height: 36px;
}}
QComboBox:hover {{
    border-color: {C['blue']};
    box-shadow: 0 0 8px rgba(59, 130, 246, 0.2);
}}
QComboBox::drop-down {{
    border: none;
    border-radius: 0 8px 8px 0;
    background: {C['bg_hover']};
    width: 30px;
}}
QComboBox::down-arrow {{
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid {C['t2']};
    margin-right: 8px;
}}
QComboBox QAbstractItemView {{
    background: {C['bg2']};
    border: 1px solid {C['brd']};
    border-radius: 8px;
    color: {C['t1']};
    selection-background-color: {C['bg_sel']};
    padding: 4px;
}}

/* ── List Widget ── */
QListWidget {{
    background-color: {C['bg1']};
    border: 1px solid {C['brd']};
    border-radius: 8px;
    color: {C['t1']};
    outline: none;
}}
QListWidget::item {{
    padding: 10px 14px;
    border-radius: 6px;
    margin: 2px;
}}
QListWidget::item:selected {{
    background: {C['bg_sel']};
    color: {C['t1']};
}}
QListWidget::item:hover {{
    background: rgba(59, 130, 246, 0.15);
}}

/* ── Progress Bar ── */
QProgressBar {{
    background-color: {C['bg3']};
    border: none;
    border-radius: 6px;
    height: 8px;
    text-align: center;
}}
QProgressBar::chunk {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                                stop:0 {C['blue']}, 
                                stop:1 {C['cyan']});
    border-radius: 6px;
}}

/* ── ToolTip ── */
QToolTip {{
    background: {C['bg2']};
    color: {C['t1']};
    border: 1px solid {C['blue']};
    border-radius: 6px;
    padding: 8px 12px;
    font-size: {F_SMALL}px;
}}
"""


# ══════════════════════════════════════════════════════════════════════
#  OPTİMİZASYON MOTORU
# ══════════════════════════════════════════════════════════════════════
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False


class RasyonOptimizer:
    KABA_KEYS = [
        "saman","kuru ot","silaj","mısır silaj","yonca","çayır",
        "tritikale","arpa saman","buğday saman","oat hay","hay",
        "silage","roughage","straw","grass","çim",
    ]

    @staticmethod
    def is_kaba(name: str) -> bool:
        n = name.lower()
        return any(k in n for k in RasyonOptimizer.KABA_KEYS)

    def optimize(self, feeds, total_kg,
                 min_prot, max_prot,
                 min_en,   max_en,
                 kaba_lo=0.40, kaba_hi=0.70):

        if not feeds:
            return {"status": "NoFeeds"}

        prob = LpProblem("LCR", LpMinimize)
        X = {f["name"]: LpVariable(f"x{i}", lowBound=0, upBound=total_kg)
             for i, f in enumerate(feeds)}

        prob += lpSum(X[f["name"]] * f["price"] for f in feeds), "Cost"

        prob += lpSum(X[f["name"]] for f in feeds) == total_kg,          "TotKg"
        prob += lpSum(X[f["name"]] * f["protein"]/100 for f in feeds) >= (min_prot/100)*total_kg, "PrMin"
        prob += lpSum(X[f["name"]] * f["protein"]/100 for f in feeds) <= (max_prot/100)*total_kg, "PrMax"
        prob += lpSum(X[f["name"]] * f["energy"]      for f in feeds) >= min_en*total_kg,          "EnMin"
        prob += lpSum(X[f["name"]] * f["energy"]      for f in feeds) <= max_en*total_kg,          "EnMax"

        kaba = [f for f in feeds if f["is_kaba"]]
        if kaba:
            ks = lpSum(X[f["name"]] for f in kaba)
            prob += ks >= kaba_lo * total_kg, "KabaLo"
            prob += ks <= kaba_hi * total_kg, "KabaHi"

        for f in feeds:
            nl = f["name"].lower()
            if "üre" in nl or ("ure" in nl and "pure" not in nl):
                prob += X[f["name"]] <= 0.15,              f"UreKg_{f['name']}"
                prob += X[f["name"]] <= 0.01 * total_kg,   f"UreKM_{f['name']}"

        prob.solve(PULP_CBC_CMD(msg=0))
        status = LpStatus[prob.status]
        if status != "Optimal":
            return {"status": status}

        results, tot_cost, tot_prot_kg, tot_en = [], 0.0, 0.0, 0.0
        for f in feeds:
            kg = max(0.0, round(value(X[f["name"]]) or 0, 4))
            if kg < 1e-4:
                continue
            cost = kg * f["price"]
            results.append({
                "name":    f["name"],
                "kg":      kg,
                "pct":     kg / total_kg * 100,
                "cost":    cost,
                "protein": f["protein"],
                "energy":  f["energy"],
                "is_kaba": f["is_kaba"],
            })
            tot_cost     += cost
            tot_prot_kg  += kg * f["protein"] / 100
            tot_en       += kg * f["energy"]

        return {
            "status":   "Optimal",
            "results":  results,
            "total_cost":        round(tot_cost, 3),
            "achieved_protein":  round(tot_prot_kg / total_kg * 100, 2),
            "achieved_energy":   round(tot_en / total_kg, 3),
            "total_protein_kg":  round(tot_prot_kg, 3),   # kg ham protein
            "total_energy_mj":   round(tot_en, 2),         # toplam MJ
            "total_kg":          total_kg,
            "min_prot":          min_prot,
            "max_prot":          max_prot,
            "min_en":            min_en,
            "max_en":            max_en,
        }


# ══════════════════════════════════════════════════════════════════════
#  YARDIMCI: Badge etiketi - Enhanced
# ══════════════════════════════════════════════════════════════════════
class Badge(QLabel):
    """Küçük renkli pill etiketi - Animated Glow Effect."""
    def __init__(self, text, bg="#22C55E", fg="#ffffff", parent=None):
        super().__init__(text, parent)
        self._bg = bg
        self._fg = fg
        self._glow_enabled = True
        self._animation_timer = None
        self._glow_opacity = 0.3
        self._update_stylesheet()

    def _update_stylesheet(self):
        glow_color = self._bg.replace("#", "") if len(self._bg) == 7 else "22C55E"
        self.setStyleSheet(f"""
            QLabel {{
                background-color: {self._bg};
                color: {self._fg};
                border-radius: 12px;
                padding: 3px 12px;
                font-size: {F_TINY}px;
                font-weight: 700;
                letter-spacing: 0.8px;
                border: 1px solid rgba(255,255,255,0.1);
            }}
        """)
        self.setFixedHeight(22)

    def start_glow_animation(self):
        """Başlat glow animasyonu"""
        if self._animation_timer:
            self._animation_timer.stop()
        self._animation_timer = QTimer(self)
        self._glow_opacity = 0.3
        def animate():
            self._glow_opacity = (self._glow_opacity + 0.05) % 0.8
            # Basit opacity animasyonu
        self._animation_timer.timeout.connect(animate)
        self._animation_timer.start(100)


# ══════════════════════════════════════════════════════════════════════
#  YARDIMCI: Metrik kartı - Enhanced with Animation
# ══════════════════════════════════════════════════════════════════════
class MetricCard(QFrame):
    """Metrik kartı - Animated with shimmer effect."""
    def __init__(self, icon, title, default="—", accent=C["blue"], parent=None):
        super().__init__(parent)
        self.accent = accent
        self._target_value = default
        self._animation_progress = 0
        self.setProperty("glass", "true")
        self.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                    stop:0 rgba(30,41,59,0.9),
                                    stop:1 rgba(26,34,53,0.95));
                border: 1px solid {C['brd']};
                border-radius: 14px;
                border-left: 5px solid {accent};
            }}
            QFrame:hover {{
                border: 1px solid {accent};
                border-left: 5px solid {accent};
            }}
        """)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(18, 14, 18, 14)
        lay.setSpacing(6)

        # Üst satır: ikon + başlık
        top = QHBoxLayout()
        ico = QLabel(icon)
        ico.setFont(QFont("Segoe UI Emoji", 17))
        ico.setFixedWidth(30)
        lbl = QLabel(title)
        lbl.setProperty("role", "caption")
        lbl.setStyleSheet(f"color:{C['t2']}; font-size:{F_SMALL}px; font-weight:600; letter-spacing:0.5px;")
        top.addWidget(ico)
        top.addWidget(lbl)
        top.addStretch()
        lay.addLayout(top)

        # Ana değer
        self.val_lbl = QLabel(default)
        self.val_lbl.setFont(QFont("Segoe UI", F_XLARGE + 2, QFont.Bold))
        self.val_lbl.setStyleSheet(f"""
            color: {accent};
            font-weight: 800;
            letter-spacing: -0.5px;
        """)
        lay.addWidget(self.val_lbl)

        # Alt açıklama
        self.sub_lbl = QLabel("")
        self.sub_lbl.setStyleSheet(f"color:{C['t3']}; font-size:{F_TINY}px;")
        lay.addWidget(self.sub_lbl)

        # Ekstra metin (absürt değer)
        self.abs_lbl = QLabel("")
        self.abs_lbl.setStyleSheet(
            f"color:{C['t2']}; font-size:{F_SMALL}px; font-weight:600;"
        )
        lay.addWidget(self.abs_lbl)

        # Animasyon timer
        self._anim_timer = QTimer(self)
        self._anim_timer.timeout.connect(self._update_animation)
        self._anim_value = 0
        self._target_numeric = 0
        self._is_numeric = False

    def _update_animation(self):
        if self._is_numeric:
            diff = self._target_numeric - self._anim_value
            step = max(0.01, abs(diff) * 0.1)
            if abs(diff) < step:
                self._anim_value = self._target_numeric
                self._anim_timer.stop()
            else:
                self._anim_value += step if diff > 0 else -step

    def set_value(self, val, sub="", absolute=""):
        self.val_lbl.setText(str(val))
        self.sub_lbl.setText(sub)
        self.abs_lbl.setText(absolute)

        # Sayısal animasyon kontrolü
        try:
            self._target_numeric = float(str(val).replace(",", "").replace("—", "0"))
            self._is_numeric = True
            if not self._anim_timer.isActive():
                self._anim_timer.start(30)
        except (ValueError, AttributeError):
            self._is_numeric = False

    def set_accent(self, color):
        self.accent = color
        self.val_lbl.setStyleSheet(f"color:{color}; font-weight:800;")
        self.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                    stop:0 rgba(30,41,59,0.9),
                                    stop:1 rgba(26,34,53,0.95));
                border: 1px solid {C['brd']};
                border-radius: 14px;
                border-left: 5px solid {color};
            }}
            QFrame:hover {{
                border: 1px solid {color};
            }}
        """)


# ══════════════════════════════════════════════════════════════════════
#  YARDIMCI: Yatay bar grafik - Enhanced with Animation
# ══════════════════════════════════════════════════════════════════════
class BarChartWidget(QWidget):
    """Rasyon bileşenlerini gösteren mini yatay çubuk grafik - Animated."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.data: list[dict] = []
        self._animation_progress = 0.0
        self._target_progress = 1.0
        self._anim_timer = QTimer(self)
        self._anim_timer.timeout.connect(self._animate)
        self._anim_speed = 0.03
        self.setMinimumHeight(200)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.setStyleSheet(f"""
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 {C['bg2']}, 
                                stop:1 {C['bg3']});
            border-radius: 12px;
            border: 1px solid {C['brd']};
        """)

    def set_data(self, items: list[dict]):
        self.data = items
        self._animation_progress = 0.0
        self._target_progress = 1.0
        if items and not self._anim_timer.isActive():
            self._anim_timer.start(16)  # ~60fps
        self.update()

    def _animate(self):
        if self._animation_progress < self._target_progress:
            self._animation_progress += self._anim_speed
            if self._animation_progress >= self._target_progress:
                self._animation_progress = self._target_progress
                self._anim_timer.stop()
        self.update()

    def paintEvent(self, event):
        if not self.data:
            return
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.TextAntialiasing)

        W, H   = self.width(), self.height()
        pad_l  = 190   # isim alanı
        pad_r  = 75    # % yazısı
        pad_t  = 18
        bar_h  = 26
        gap    = 12
        n      = len(self.data)
        total_h = n * (bar_h + gap) - gap + pad_t * 2
        y0     = max(pad_t, (H - total_h) // 2 + pad_t)

        avail_w = W - pad_l - pad_r - 28

        name_font = QFont("Segoe UI", F_SMALL, QFont.Medium)
        pct_font  = QFont("Segoe UI", F_SMALL + 1, QFont.Bold)
        
        for i, item in enumerate(self.data):
            y = y0 + i * (bar_h + gap)
            
            # Animasyonlu ilerleme
            anim_pct = min(item["pct"] * self._animation_progress, item["pct"])
            pct = max(anim_pct, 0.5)

            # Arka plan çubuğu - Glow efekti ile
            bg_rect = QRectF(pad_l, y, avail_w, bar_h)
            p.setPen(Qt.NoPen)
            p.setBrush(QColor(C["bg3"]))
            p.drawRoundedRect(bg_rect, 6, 6)

            # Değer çubuğu - Gradient + glow
            fill_w = avail_w * (pct / 100.0)
            if fill_w > 0:
                fill_rect = QRectF(pad_l, y, fill_w, bar_h)
                
                # Gradient
                grad = QLinearGradient(pad_l, y, pad_l + fill_w, y)
                color = QColor(item["color"])
                lighter = color.lighter(160)
                grad.setColorAt(0.0, lighter)
                grad.setColorAt(0.5, color)
                grad.setColorAt(1.0, color.lighter(120))
                p.setBrush(QBrush(grad))
                
                # Glow efekti
                p.save()
                p.setCompositionMode(QPainter.CompositionMode_Plus)
                glow_color = QColor(item["color"])
                glow_color.setAlpha(40)
                p.setPen(Qt.NoPen)
                p.setBrush(glow_color)
                glow_rect = QRectF(pad_l - 2, y - 2, fill_w + 4, bar_h + 4)
                p.drawRoundedRect(glow_rect, 8, 8)
                p.restore()
                
                p.setBrush(QBrush(grad))
                p.drawRoundedRect(fill_rect, 6, 6)

            # Yem adı (sol) - modern font
            p.setPen(QColor(C["t1"]))
            p.setFont(name_font)
            name = item["name"]
            fm = QFontMetrics(name_font)
            max_name_w = pad_l - 16
            elided = fm.elidedText(name, Qt.ElideRight, max_name_w)
            p.drawText(QRect(10, y, max_name_w, bar_h),
                       Qt.AlignRight | Qt.AlignVCenter, elided)

            # Yüzde (sağ) - vurgulu renk
            p.setFont(pct_font)
            pct_color = QColor(item["color"])
            p.setPen(pct_color)
            pct_text = f"%{anim_pct:.1f}"
            
            # Glow efekti yüzde için
            p.save()
            p.setCompositionMode(QPainter.CompositionMode_Plus)
            glow_pen = QPen(QColor(item["color"]).lighter(150))
            glow_pen.setWidthF(3)
            p.setPen(glow_pen)
            p.drawText(QRect(W - pad_r + 6, y, pad_r - 12, bar_h),
                       Qt.AlignLeft | Qt.AlignVCenter, pct_text)
            p.restore()
            
            p.setPen(pct_color)
            p.drawText(QRect(W - pad_r + 6, y, pad_r - 12, bar_h),
                       Qt.AlignLeft | Qt.AlignVCenter, pct_text)

        p.end()


# ══════════════════════════════════════════════════════════════════════
#  BAŞLIK WIDGET - Enhanced with Gradient & Animation
# ══════════════════════════════════════════════════════════════════════
class AnimatedBackground(QWidget):
    """Animasyonlu arka plan efekti"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self._angle = 0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._update_gradient)
        self._timer.start(50)
        
    def _update_gradient(self):
        self._angle = (self._angle + 2) % 360
        self.update()
        
    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        
        w, h = self.width(), self.height()
        
        # Gradient animasyonu
        grad = QLinearGradient(0, 0, w, h)
        grad.setColorAt(0, QColor(C["bg0"]))
        grad.setColorAt(0.3 + 0.1 * math.sin(math.radians(self._angle)), QColor("#0f172a"))
        grad.setColorAt(1, QColor(C["bg0"]))
        
        p.fillRect(0, 0, w, h, grad)
        
        # Parlak çizgiler
        p.setPen(QPen(QColor(C["blue"]).lighter(200).lighter(200), 1, Qt.DashLine))
        p.setOpacity(0.1)
        for i in range(3):
            offset = (self._angle + i * 120) % 360 / 360
            y = int(h * offset)
            p.drawLine(0, y, w, y)


class HeaderWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.setFixedHeight(88)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        # Animasyonlu gradient efekti
        self._gradient_angle = 0
        self._header_timer = QTimer(self)
        self._header_timer.timeout.connect(self._animate_header)
        self._header_timer.start(30)
        
        self._build_ui()

    def _animate_header(self):
        self._gradient_angle = (self._gradient_angle + 1) % 360
        self.update()

    def _build_ui(self):
        lay = QHBoxLayout(self)
        lay.setContentsMargins(28, 0, 28, 0)
        lay.setSpacing(18)

        # İkon - Modern stil
        ico_frame = QFrame()
        ico_frame.setFixedSize(56, 56)
        ico_frame.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                    stop:0 {C['green']}, stop:1 {C['green_dim']});
                border-radius: 14px;
                border: 2px solid {C['green_light']};
            }}
        """)
        ico_lay = QHBoxLayout(ico_frame)
        ico_lay.setContentsMargins(0, 0, 0, 0)
        ico_lay.setAlignment(Qt.AlignCenter)
        ico = QLabel("🌾")
        ico.setFont(QFont("Segoe UI Emoji", 26))
        ico_lay.addWidget(ico)
        lay.addWidget(ico_frame)

        # Metin grubu - Gradient başlık
        col = QVBoxLayout()
        col.setSpacing(4)
        
        t1 = QLabel("Rasyon Optimizasyon Sistemi")
        t1.setFont(QFont("Segoe UI", F_LARGE + 4, QFont.Bold))
        t1.setStyleSheet(f"""
            color: {C['t1']};
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                stop:0 {C['blue_light']}, stop:1 {C['cyan']});
            background-clip: text;
            -webkit-background-clip: text;
        """)
        
        t2 = QLabel("✨ Doğrusal Programlama (PuLP)  ·  En Düşük Maliyetli Rasyon Hesaplama")
        t2.setStyleSheet(f"color:{C['t2']}; font-size:{F_SMALL}px; font-weight:500;")
        
        col.addWidget(t1)
        col.addWidget(t2)
        lay.addLayout(col)
        lay.addStretch()

        # Sağ: badge'ler - Modern stil
        badges = QVBoxLayout()
        badges.setSpacing(8)
        badges.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        
        # Version badge with glow
        ver_frame = QFrame()
        ver_frame.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                    stop:0 {C['blue_dim']}, stop:1 {C['blue']});
                border-radius: 10px;
                border: 1px solid {C['blue_light']};
            }}
        """)
        ver_lay = QHBoxLayout(ver_frame)
        ver_lay.setContentsMargins(12, 4, 12, 4)
        ver_lay.setSpacing(6)
        ver_icon = QLabel("🚀")
        ver_icon.setFont(QFont("Segoe UI Emoji", 12))
        ver_text = QLabel("v4.0 Enhanced")
        ver_text.setFont(QFont("Segoe UI", F_TINY, QFont.Bold))
        ver_text.setStyleSheet(f"color:{C['t1']};")
        ver_lay.addWidget(ver_icon)
        ver_lay.addWidget(ver_text)
        
        tech_frame = QFrame()
        tech_frame.setStyleSheet(f"""
            QFrame {{
                background: {C['bg3']};
                border-radius: 10px;
                border: 1px solid {C['brd']};
            }}
        """)
        tech_lay = QHBoxLayout(tech_frame)
        tech_lay.setContentsMargins(12, 4, 12, 4)
        tech_lay.setSpacing(6)
        tech_icon = QLabel("⚙️")
        tech_icon.setFont(QFont("Segoe UI Emoji", 12))
        tech_text = QLabel("PuLP · PyQt5 · ReportLab")
        tech_text.setFont(QFont("Segoe UI", F_TINY, QFont.Medium))
        tech_text.setStyleSheet(f"color:{C['t2']};")
        tech_lay.addWidget(tech_icon)
        tech_lay.addWidget(tech_text)
        
        badges.addWidget(ver_frame)
        badges.addWidget(tech_frame)
        lay.addLayout(badges)

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        
        # Gradient arka plan
        w, h = self.width(), self.height()
        grad = QLinearGradient(0, 0, w, 0)
        grad.setColorAt(0, QColor(C["bg1"]))
        grad.setColorAt(0.5, QColor("#1a2337"))
        grad.setColorAt(1, QColor(C["bg0"]))
        p.fillRect(0, 0, w, h, grad)
        
        # Alt kenar çizgisi - gradient
        line_grad = QLinearGradient(0, 0, w, 0)
        line_grad.setColorAt(0, QColor(C["blue"]).lighter(200))
        line_grad.setColorAt(0.5, QColor(C["cyan"]).lighter(200))
        line_grad.setColorAt(1, QColor(C["blue"]).lighter(200))
        p.setPen(QPen(line_grad, 2))
        p.drawLine(0, h - 2, w, h - 2)
        
        # Üst köşe dekorasyonları
        p.setPen(QPen(QColor(C["blue"]).lighter(200).lighter(200), 1, Qt.DashLine))
        p.setOpacity(0.3)
        p.drawLine(0, 20, 40, 0)
        p.drawLine(w - 40, 0, w, 20)


# ══════════════════════════════════════════════════════════════════════
#  SOL PANEL — YEM TABLOSU
# ══════════════════════════════════════════════════════════════════════
class FeedTablePanel(QGroupBox):
    def __init__(self):
        super().__init__("  Yem Veritabanı")
        self._col_map: dict = {}
        self.feeds_df: pd.DataFrame | None = None
        self._build()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setSpacing(12)
        lay.setContentsMargins(14, 22, 14, 14)

        # ── Üst bar ──
        top = QHBoxLayout()
        top.setSpacing(8)

        self.btn_import = QPushButton("XLSX Yukle")
        self.btn_import.setObjectName("btnImport")
        self.btn_import.setFixedHeight(36)
        self.btn_import.clicked.connect(self._import)

        # Arama cubugu
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Yem ara...")
        self.search_box.setFixedHeight(36)
        self.search_box.setStyleSheet(f"""
            QLineEdit {{
                background: {C['bg3']};
                border: 1px solid {C['brd']};
                border-radius: 8px;
                padding: 6px 12px;
                color: {C['t1']};
                font-size: {F_BASE}px;
            }}
            QLineEdit:focus {{
                border: 1px solid {C['blue']};
            }}
        """)
        self.search_box.textChanged.connect(self._filter_table)
        
        self.btn_sel   = QPushButton("Tumunu Sec")
        self.btn_sel.setObjectName("btnSelAll")
        self.btn_sel.setFixedHeight(36)
        self.btn_sel.clicked.connect(self._sel_all)

        self.btn_desel = QPushButton("Secimi Kaldir")
        self.btn_desel.setObjectName("btnDesel")
        self.btn_desel.setFixedHeight(36)
        self.btn_desel.clicked.connect(self._desel_all)

        self.lbl_stat = QLabel("Henuz veri yuklenmedi")
        self.lbl_stat.setStyleSheet(f"color:{C['t2']}; font-size:{F_SMALL}px;")

        top.addWidget(self.btn_import)
        top.addWidget(self.search_box, stretch=1)
        top.addWidget(self.btn_sel)
        top.addWidget(self.btn_desel)
        top.addWidget(self.lbl_stat)
        lay.addLayout(top)

        # Sütun eşleşme bilgisi
        self.lbl_cols = QLabel("")
        self.lbl_cols.setStyleSheet(
            f"color:{C['t2']}; font-size:{F_TINY}px; "
            f"background:{C['bg3']}; border-radius:5px; padding:5px 10px;"
        )
        self.lbl_cols.setWordWrap(True)
        self.lbl_cols.hide()
        lay.addWidget(self.lbl_cols)

        # Tablo
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(
            QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked
        )
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setHighlightSections(False)
        self.table.setShowGrid(True)
        self.table.setSortingEnabled(True)
        lay.addWidget(self.table)

        # Alt not
        note = QLabel(
            "Fiyat hücrelerine cift tiklayarak duzenleyebilirsiniz. "
            "Kaba yem tespiti yem adindan otomatik yapilir."
        )
        note.setStyleSheet(f"color:{C['t3']}; font-size:{F_TINY}px;")
        note.setWordWrap(True)
        lay.addWidget(note)

    def _filter_table(self, text):
        """Tablodaki satirlari filtrele"""
        for row in range(self.table.rowCount()):
            name_item = self.table.item(row, 1)
            if name_item:
                name = name_item.text().lower()
                should_show = text.lower() in name if text else True
                self.table.setRowHidden(row, not should_show)

    # ── Sütun tespiti ──
    def _detect(self, df, candidates):
        low = {c.lower().strip(): c for c in df.columns}
        for cand in candidates:
            if cand.lower() in low:
                return low[cand.lower()]
        for cand in candidates:
            for k, orig in low.items():
                if cand.lower() in k:
                    return orig
        return None

    def _import(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Yem Veritabanı Seç", "",
            "Excel Dosyaları (*.xlsx *.xls);;Tüm Dosyalar (*)"
        )
        if not path:
            return
        try:
            df = pd.read_excel(path)
            self._load(df)
        except Exception as e:
            QMessageBox.critical(self, "Dosya Hatası", f"Dosya okunamadı:\n{e}")

    def _load(self, df: pd.DataFrame):
        df.columns = [str(c).strip() for c in df.columns]
        cn = self._detect(df, ["yem adı","yem_adi","yem","feed","name","ad"])
        cp = self._detect(df, ["ham protein","ham_protein","protein","hp","cp"])
        ce = self._detect(df, ["me","me (mj/kg)","metabolik enerji","enerji","energy","mj/kg"])
        cf = self._detect(df, ["fiyat","price","maliyet","cost","₺/kg","tl/kg"])

        missing = ([("Yem Adı", cn)] + [("Ham Protein", cp)] + [("ME", ce)])
        missing = [m[0] for m in missing if not m[1]]
        if missing:
            QMessageBox.warning(
                self, "Sütun Bulunamadı",
                f"Zorunlu sütunlar tespit edilemedi:\n• " + "\n• ".join(missing) +
                f"\n\nDosyadaki sütunlar:\n{', '.join(df.columns.tolist())}"
            )
            return

        self.feeds_df = df.copy()
        self._col_map = {"name": cn, "protein": cp, "energy": ce, "price": cf}

        info = (f"✔  Ad → '{cn}'   Protein → '{cp}'   Enerji → '{ce}'"
                + (f"   Fiyat → '{cf}'" if cf else "   Fiyat → bulunamadı (0)"))
        self.lbl_cols.setText(info)
        self.lbl_cols.show()
        self._fill_table()

    def _fill_table(self):
        df = self.feeds_df
        cm = self._col_map
        headers = ["", "Yem Adı", "Protein\n(%)", "ME\n(MJ/kg)", "Fiyat\n(₺/kg)", "Tür"]
        self.table.setSortingEnabled(False)
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(len(df))
        self.table.verticalHeader().setDefaultSectionSize(38)

        for row, (_, rec) in enumerate(df.iterrows()):
            # Checkbox
            cw = QWidget()
            cl = QHBoxLayout(cw)
            cl.setAlignment(Qt.AlignCenter)
            cl.setContentsMargins(0, 0, 0, 0)
            chk = QCheckBox()
            chk.setChecked(True)
            cl.addWidget(chk)
            self.table.setCellWidget(row, 0, cw)

            # Ad
            name = str(rec.get(cm["name"], ""))
            ni = QTableWidgetItem(name)
            ni.setFont(QFont("Segoe UI", F_BASE))
            ni.setFlags(ni.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 1, ni)

            # Protein
            try:   pv = float(rec[cm["protein"]])
            except: pv = 0.0
            pi = QTableWidgetItem(f"{pv:.2f}")
            pi.setTextAlignment(Qt.AlignCenter)
            pi.setFont(QFont("Segoe UI", F_BASE))
            pi.setFlags(pi.flags() & ~Qt.ItemIsEditable)
            pi.setForeground(QColor(C["purple"]))
            self.table.setItem(row, 2, pi)

            # Enerji
            try:   ev = float(rec[cm["energy"]])
            except: ev = 0.0
            ei = QTableWidgetItem(f"{ev:.3f}")
            ei.setTextAlignment(Qt.AlignCenter)
            ei.setFont(QFont("Segoe UI", F_BASE))
            ei.setFlags(ei.flags() & ~Qt.ItemIsEditable)
            ei.setForeground(QColor(C["cyan"]))
            self.table.setItem(row, 3, ei)

            # Fiyat (düzenlenebilir)
            try:   fv = float(rec[cm["price"]]) if cm["price"] else 0.0
            except: fv = 0.0
            fi = QTableWidgetItem(f"{fv:.4f}")
            fi.setTextAlignment(Qt.AlignCenter)
            fi.setFont(QFont("Segoe UI", F_BASE, QFont.Bold))
            fi.setForeground(QColor(C["amber"]))
            self.table.setItem(row, 4, fi)

            # Tür badge metni
            is_k = RasyonOptimizer.is_kaba(name)
            ki = QTableWidgetItem("Kaba" if is_k else "Kesif")
            ki.setTextAlignment(Qt.AlignCenter)
            ki.setFont(QFont("Segoe UI", F_SMALL, QFont.Bold))
            ki.setFlags(ki.flags() & ~Qt.ItemIsEditable)
            ki.setForeground(QColor(C["green"] if is_k else C["t2"]))
            if is_k:
                ki.setBackground(QColor(21, 128, 61, 40))
            self.table.setItem(row, 5, ki)

        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.Fixed)
        self.table.setColumnWidth(0, 44)
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)
        for c in [2, 3, 4, 5]:
            hdr.setSectionResizeMode(c, QHeaderView.ResizeToContents)

        self.table.setSortingEnabled(True)
        n = len(df)
        self.lbl_stat.setText(f"<b style='color:{C['t1']}'>{n}</b> yem yüklendi")

    def _sel_all(self):
        for r in range(self.table.rowCount()):
            w = self.table.cellWidget(r, 0)
            if w:
                w.findChild(QCheckBox).setChecked(True)

    def _desel_all(self):
        for r in range(self.table.rowCount()):
            w = self.table.cellWidget(r, 0)
            if w:
                w.findChild(QCheckBox).setChecked(False)

    def get_selected(self) -> list[dict]:
        out = []
        for r in range(self.table.rowCount()):
            w = self.table.cellWidget(r, 0)
            if not w:
                continue
            if not w.findChild(QCheckBox).isChecked():
                continue
            try:
                out.append({
                    "name":    self.table.item(r, 1).text(),
                    "protein": float(self.table.item(r, 2).text()),
                    "energy":  float(self.table.item(r, 3).text()),
                    "price":   float(self.table.item(r, 4).text()),
                    "is_kaba": "Kaba" in self.table.item(r, 5).text(),
                })
            except (ValueError, AttributeError):
                continue
        return out


# ══════════════════════════════════════════════════════════════════════
#  SAĞ ÜST PANEL — KISITLAR - Enhanced with Presets
# ══════════════════════════════════════════════════════════════════════

# Preset tanımları
PRESETS = {
    "Sığır (Süt)": {
        "total_kg": 22.0,
        "min_prot": 14.0, "max_prot": 17.0,
        "min_en": 10.5, "max_en": 12.5,
        "kaba_lo": 40, "kaba_hi": 70,
    },
    "Sığır (Et)": {
        "total_kg": 18.0,
        "min_prot": 11.0, "max_prot": 14.0,
        "min_en": 9.5, "max_en": 11.5,
        "kaba_lo": 50, "kaba_hi": 80,
    },
    "Koyun (Süt)": {
        "total_kg": 4.0,
        "min_prot": 15.0, "max_prot": 18.0,
        "min_en": 10.0, "max_en": 12.0,
        "kaba_lo": 45, "kaba_hi": 70,
    },
    "Keçi (Süt)": {
        "total_kg": 3.5,
        "min_prot": 14.0, "max_prot": 17.0,
        "min_en": 10.0, "max_en": 12.0,
        "kaba_lo": 35, "kaba_hi": 65,
    },
    "Besi (Dana)": {
        "total_kg": 10.0,
        "min_prot": 12.0, "max_prot": 15.0,
        "min_en": 10.0, "max_en": 11.5,
        "kaba_lo": 30, "kaba_hi": 60,
    },
}


class ConstraintPanel(QGroupBox):
    def __init__(self):
        super().__init__("  Optimizasyon Kısıtları")
        self._build()

    def _spin(self, lo, hi, val, dec=2, suf=""):
        s = QDoubleSpinBox()
        s.setRange(lo, hi)
        s.setValue(val)
        s.setDecimals(dec)
        s.setSuffix(suf)
        s.setSingleStep(0.5)
        return s

    def _section(self, title, icon=""):
        lbl = QLabel(f"{icon}  {title}" if icon else title)
        lbl.setFont(QFont("Segoe UI", F_BASE, QFont.Bold))
        lbl.setStyleSheet(f"color:{C['t_head']}; margin-top:6px;")
        return lbl

    def _row(self, label, widget, unit_hint=""):
        row = QHBoxLayout()
        lbl = QLabel(label)
        lbl.setStyleSheet(f"color:{C['t2']}; font-size:{F_BASE}px;")
        lbl.setMinimumWidth(175)
        row.addWidget(lbl)
        row.addWidget(widget)
        if unit_hint:
            u = QLabel(unit_hint)
            u.setStyleSheet(f"color:{C['t3']}; font-size:{F_TINY}px; margin-left:4px;")
            row.addWidget(u)
        row.addStretch()
        return row

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setSpacing(10)
        lay.setContentsMargins(16, 22, 16, 16)

        # ── Preset Seçici ──
        preset_frame = QFrame()
        preset_frame.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                    stop:0 {C['purple']}20, stop:1 {C['blue']}20);
                border: 1px solid {C['purple']}40;
                border-radius: 10px;
                padding: 8px;
            }}
        """)
        preset_lay = QVBoxLayout(preset_frame)
        preset_lay.setSpacing(6)
        
        preset_header = QHBoxLayout()
        preset_icon = QLabel("📋")
        preset_icon.setFont(QFont("Segoe UI Emoji", 14))
        preset_title = QLabel("Hızlı Ayarlar (Presets)")
        preset_title.setFont(QFont("Segoe UI", F_SMALL, QFont.Bold))
        preset_title.setStyleSheet(f"color:{C['purple_light']};")
        preset_header.addWidget(preset_icon)
        preset_header.addWidget(preset_title)
        preset_header.addStretch()
        preset_lay.addLayout(preset_header)
        
        # Preset butonları
        preset_btn_lay = QHBoxLayout()
        preset_btn_lay.setSpacing(6)
        
        self.preset_combo = QComboBox()
        self.preset_combo.addItems(list(PRESETS.keys()))
        self.preset_combo.setFixedHeight(32)
        self.preset_combo.setStyleSheet(f"""
            QComboBox {{
                background: {C['bg3']};
                color: {C['t1']};
                border: 1px solid {C['brd']};
                border-radius: 8px;
                padding: 6px 12px;
                font-size: {F_SMALL}px;
            }}
        """)
        
        btn_apply = QPushButton("✓ Uygula")
        btn_apply.setObjectName("btnPreset")
        btn_apply.setFixedHeight(32)
        btn_apply.setFixedWidth(80)
        btn_apply.clicked.connect(self._apply_preset)
        
        preset_btn_lay.addWidget(self.preset_combo)
        preset_btn_lay.addWidget(btn_apply)
        preset_btn_lay.addStretch()
        preset_lay.addLayout(preset_btn_lay)
        
        lay.addWidget(preset_frame)

        # ── Toplam Rasyon Card ──
        f1 = QFrame()
        f1.setStyleSheet(f"""
            QFrame {{
                background: {C['bg3']};
                border: 1px solid {C['brd']};
                border-radius: 10px;
            }}
        """)
        l1 = QVBoxLayout(f1)
        l1.setContentsMargins(14, 12, 14, 12)
        l1.setSpacing(8)
        l1.addWidget(self._section("Toplam Rasyon", "⚖️"))
        self.sp_total = self._spin(0.1, 9999, 20.0, 1, " kg/gün")
        l1.addLayout(self._row("Günlük Rasyon Miktarı:", self.sp_total))
        lay.addWidget(f1)

        # ── Protein Card ──
        f2 = QFrame()
        f2.setStyleSheet(f"""
            QFrame {{
                background: {C['bg3']};
                border: 1px solid {C['brd']};
                border-radius: 10px;
            }}
        """)
        l2 = QVBoxLayout(f2)
        l2.setContentsMargins(14, 12, 14, 12)
        l2.setSpacing(8)
        l2.addWidget(self._section("Ham Protein Kısıtları", "🧬"))
        self.sp_pmin = self._spin(0, 100, 14.0, 2, " %")
        self.sp_pmax = self._spin(0, 100, 18.0, 2, " %")
        l2.addLayout(self._row("Minimum Protein:", self.sp_pmin))
        l2.addLayout(self._row("Maksimum Protein:", self.sp_pmax))
        lay.addWidget(f2)

        # ── Enerji Card ──
        f3 = QFrame()
        f3.setStyleSheet(f"""
            QFrame {{
                background: {C['bg3']};
                border: 1px solid {C['brd']};
                border-radius: 10px;
            }}
        """)
        l3 = QVBoxLayout(f3)
        l3.setContentsMargins(14, 12, 14, 12)
        l3.setSpacing(8)
        l3.addWidget(self._section("Metabolik Enerji Kısıtları", "⚡"))
        self.sp_emin = self._spin(0, 30, 10.0, 3, " MJ/kg")
        self.sp_emax = self._spin(0, 30, 13.0, 3, " MJ/kg")
        l3.addLayout(self._row("Minimum ME:", self.sp_emin))
        l3.addLayout(self._row("Maksimum ME:", self.sp_emax))
        lay.addWidget(f3)

        # ── Kaba Yem Card ──
        f4 = QFrame()
        f4.setStyleSheet(f"""
            QFrame {{
                background: {C['bg3']};
                border: 1px solid {C['brd']};
                border-radius: 10px;
            }}
        """)
        l4 = QVBoxLayout(f4)
        l4.setContentsMargins(14, 12, 14, 12)
        l4.setSpacing(8)
        l4.addWidget(self._section("Kaba Yem Oranı (Ruminant)", "🌿"))
        self.sp_kmin = self._spin(0, 100, 40.0, 1, " %")
        self.sp_kmax = self._spin(0, 100, 70.0, 1, " %")
        l4.addLayout(self._row("Minimum Kaba Yem:", self.sp_kmin))
        l4.addLayout(self._row("Maksimum Kaba Yem:", self.sp_kmax))
        note_k = QLabel("💡 Kaba yem yoksa bu kısıt uygulanmaz.")
        note_k.setStyleSheet(f"color:{C['t3']}; font-size:{F_TINY}px;")
        l4.addWidget(note_k)
        lay.addWidget(f4)

        # ── Üre uyarı kutusu ──
        ure = QLabel("🔒  Üre Güvenlik Kısıtı Aktif\n"
                     "Rasyona dahil edilirse max %1 KM veya 0.15 kg ile sınırlanır.")
        ure.setWordWrap(True)
        ure.setStyleSheet(f"""
            color: {C['amber']};
            background: rgba(245,158,11,0.08);
            border: 1px solid rgba(245,158,11,0.35);
            border-radius: 10px;
            padding: 12px 16px;
            font-size: {F_SMALL}px;
            font-weight: 600;
            line-height: 1.5;
        """)
        lay.addWidget(ure)
        lay.addStretch()

    def _apply_preset(self):
        preset_name = self.preset_combo.currentText()
        if preset_name in PRESETS:
            p = PRESETS[preset_name]
            self.sp_total.setValue(p["total_kg"])
            self.sp_pmin.setValue(p["min_prot"])
            self.sp_pmax.setValue(p["max_prot"])
            self.sp_emin.setValue(p["min_en"])
            self.sp_emax.setValue(p["max_en"])
            self.sp_kmin.setValue(p["kaba_lo"])
            self.sp_kmax.setValue(p["kaba_hi"])

    def get(self) -> dict:
        return {
            "total_kg": self.sp_total.value(),
            "min_prot": self.sp_pmin.value(),
            "max_prot": self.sp_pmax.value(),
            "min_en":   self.sp_emin.value(),
            "max_en":   self.sp_emax.value(),
            "kaba_lo":  self.sp_kmin.value() / 100,
            "kaba_hi":  self.sp_kmax.value() / 100,
        }


# ══════════════════════════════════════════════════════════════════════
#  YARDIMCI: Besin özeti — progress bar'lı özet şerit
# ══════════════════════════════════════════════════════════════════════
class NutritionSummaryWidget(QWidget):
    """
    Protein ve enerji kullanımını görsel progress bar + sayısal değerler
    olarak gösteren kompakt şerit.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self._data: dict | None = None
        self.setStyleSheet(f"""
            background:{C['bg2']};
            border:1px solid {C['brd']};
            border-radius:10px;
        """)
        self.setFixedHeight(100)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

    def set_data(self, data: dict | None):
        self._data = data
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        W, H = self.width(), self.height()
        pad = 18
        mid = H // 2

        if not self._data:
            p.setPen(QColor(C["t3"]))
            p.setFont(QFont("Segoe UI", F_SMALL))
            p.drawText(QRect(0, 0, W, H), Qt.AlignCenter,
                       "Optimizasyon sonrası besin özeti burada görünecek")
            p.end()
            return

        d = self._data
        bar_h   = 14
        label_w = 160
        val_w   = 200
        bar_w   = W - pad * 2 - label_w - val_w

        # ── Protein satırı (üst) ──
        y_prot = mid - bar_h - 10
        self._draw_row(
            p,
            label    = "🧬  Ham Protein",
            val_pct  = d["achieved_protein"],
            val_abs  = d["total_protein_kg"],
            unit_pct = "%",
            unit_abs = "kg",
            lo       = d["min_prot"],
            hi       = d["max_prot"],
            x0       = pad,
            y        = y_prot,
            bar_w    = bar_w,
            bar_h    = bar_h,
            label_w  = label_w,
            val_w    = val_w,
            fill_color = C["purple"],
            range_color= "#6D28D9",
        )

        # ── Enerji satırı (alt) ──
        y_en = mid + 10
        self._draw_row(
            p,
            label    = "⚡  Metabolik Enerji",
            val_pct  = d["achieved_energy"],
            val_abs  = d["total_energy_mj"],
            unit_pct = "MJ/kg",
            unit_abs = "MJ toplam",
            lo       = d["min_en"],
            hi       = d["max_en"],
            x0       = pad,
            y        = y_en,
            bar_w    = bar_w,
            bar_h    = bar_h,
            label_w  = label_w,
            val_w    = val_w,
            fill_color = C["cyan"],
            range_color= "#0E7490",
        )
        p.end()

    def _draw_row(self, p, label, val_pct, val_abs, unit_pct, unit_abs,
                  lo, hi, x0, y, bar_w, bar_h, label_w, val_w,
                  fill_color, range_color):
        # Etiket
        p.setFont(QFont("Segoe UI", F_SMALL, QFont.Bold))
        p.setPen(QColor(C["t2"]))
        p.drawText(QRect(x0, y, label_w, bar_h),
                   Qt.AlignLeft | Qt.AlignVCenter, label)

        bx = x0 + label_w + 8

        # Arka plan çubuğu
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(C["bg3"]))
        p.drawRoundedRect(QRectF(bx, y, bar_w, bar_h), 5, 5)

        # Hedef aralık (lo–hi) — açık renkli bölge
        if hi > 0:
            lo_x = bx + bar_w * (lo / hi) if hi > 0 else bx
            range_fill = QColor(range_color)
            range_fill.setAlpha(55)
            p.setBrush(range_fill)
            p.drawRoundedRect(QRectF(lo_x, y, bar_w - (lo_x - bx), bar_h), 5, 5)

        # Gerçekleşen değer çubuğu
        # Protein için max=hi(%)*1.2; enerji için max=hi*1.2
        bar_max = hi * 1.25 if hi > 0 else 1
        ratio   = min(val_pct / bar_max, 1.0)
        fill_w  = bar_w * ratio

        grad = QLinearGradient(bx, y, bx + fill_w, y)
        fc = QColor(fill_color)
        grad.setColorAt(0.0, fc.lighter(150))
        grad.setColorAt(1.0, fc)
        p.setBrush(QBrush(grad))
        if fill_w > 2:
            p.drawRoundedRect(QRectF(bx, y, fill_w, bar_h), 5, 5)

        # Değer metni
        val_x = bx + bar_w + 10
        p.setFont(QFont("Segoe UI", F_SMALL, QFont.Bold))
        p.setPen(QColor(fill_color))
        p.drawText(
            QRect(int(val_x), y, int(val_w // 2 - 4), bar_h),
            Qt.AlignLeft | Qt.AlignVCenter,
            f"{val_pct:.2f} {unit_pct}"
        )
        p.setFont(QFont("Segoe UI", F_TINY))
        p.setPen(QColor(C["t2"]))
        p.drawText(
            QRect(int(val_x + val_w // 2), y, int(val_w // 2), bar_h),
            Qt.AlignLeft | Qt.AlignVCenter,
            f"{val_abs:.2f} {unit_abs}"
        )


# ══════════════════════════════════════════════════════════════════════
#  SAĞ ALT PANEL — SONUÇLAR - Enhanced with PDF Export
# ══════════════════════════════════════════════════════════════════════
class ResultPanel(QGroupBox):
    def __init__(self):
        super().__init__("  Optimizasyon Sonuçları")
        self._last: list[dict] = []
        self._history: list[dict] = []
        self._build()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setSpacing(12)
        lay.setContentsMargins(14, 22, 14, 14)

        # ── Metrik kartları ──
        card_row = QHBoxLayout()
        card_row.setSpacing(10)
        self.c_cost   = MetricCard("💰", "Toplam Maliyet",    "—",         C["green"])
        self.c_prot   = MetricCard("🧬", "Ham Protein",      "—",        C["purple"])
        self.c_energy = MetricCard("⚡", "Metabolik Enerji",  "—",         C["cyan"])
        self.c_status = MetricCard("📊", "Durum",             "Bekliyor",  C["t2"])
        for c in [self.c_cost, self.c_prot, self.c_energy, self.c_status]:
            card_row.addWidget(c)
        lay.addLayout(card_row)

        # ── Besin Özeti şeridi ──
        self.nutrition_bar = NutritionSummaryWidget()
        lay.addWidget(self.nutrition_bar)

        # ── Tab: Tablo | Grafik | Geçmiş ──
        self.tabs = QTabWidget()

        # Tablo sekmesi
        tbl_tab = QWidget()
        tbl_lay = QVBoxLayout(tbl_tab)
        tbl_lay.setContentsMargins(0, 8, 0, 0)
        tbl_lay.setSpacing(8)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setHighlightSections(False)
        self.table.setSortingEnabled(True)
        self.table.verticalHeader().setDefaultSectionSize(40)

        # Tablo başlıkları güncellendi - kütle bilgileri eklendi
        hdrs = ["Yem Adi", "Miktar\n(kg)", "Pay\n(%)",
                "Protein\n(%)", "Prot(kg)",  # Protein % ve kg
                "ME\n(MJ/kg)", "Enerji(MJ)",  # ME ve toplam MJ
                "Maliyet\n(₺)", "Tur"]  # Maliyet ve tür
        self.table.setColumnCount(len(hdrs))
        self.table.setHorizontalHeaderLabels(hdrs)
        h = self.table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.Stretch)
        for ci in range(1, len(hdrs)):
            h.setSectionResizeMode(ci, QHeaderView.ResizeToContents)

        tbl_lay.addWidget(self.table)

        bot = QHBoxLayout()
        self.btn_pdf = QPushButton("📄  PDF Raporu")
        self.btn_pdf.setObjectName("btnExport")
        self.btn_pdf.setEnabled(False)
        self.btn_pdf.clicked.connect(self._export_pdf)
        self.btn_pdf.setToolTip("PDF formatında rapor oluştur")
        
        self.btn_export = QPushButton("📊  Excel'e Aktar")
        self.btn_export.setObjectName("btnExport")
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self._export)
        bot.addStretch()
        bot.addWidget(self.btn_pdf)
        bot.addWidget(self.btn_export)
        tbl_lay.addLayout(bot)
        self.tabs.addTab(tbl_tab, "  📋  Tablo  ")

        # Grafik sekmesi
        chart_tab = QWidget()
        chart_lay = QVBoxLayout(chart_tab)
        chart_lay.setContentsMargins(8, 12, 8, 8)
        self.chart = BarChartWidget()
        chart_lay.addWidget(self.chart)
        self.tabs.addTab(chart_tab, "  📊  Grafik  ")

        # Geçmiş sekmesi
        history_tab = QWidget()
        history_lay = QVBoxLayout(history_tab)
        history_lay.setContentsMargins(4, 8, 4, 4)
        history_lay.setSpacing(6)
        
        hist_header = QLabel("📜 Optimizasyon Geçmişi")
        hist_header.setFont(QFont("Segoe UI", F_SMALL, QFont.Bold))
        hist_header.setStyleSheet(f"color:{C['t_head']};")
        history_lay.addWidget(hist_header)
        
        self.history_list = QListWidget()
        self.history_list.setAlternatingRowColors(True)
        self.history_list.itemClicked.connect(self._load_from_history)
        history_lay.addWidget(self.history_list)
        
        hist_btn_row = QHBoxLayout()
        btn_clear_hist = QPushButton("🗑️ Temizle")
        btn_clear_hist.setFixedHeight(28)
        btn_clear_hist.clicked.connect(self._clear_history)
        hist_btn_row.addStretch()
        hist_btn_row.addWidget(btn_clear_hist)
        history_lay.addLayout(hist_btn_row)
        
        self.tabs.addTab(history_tab, "  📜  Geçmiş  ")

        lay.addWidget(self.tabs)

    def _add_to_history(self, result: dict):
        timestamp = datetime.now().strftime("%H:%M:%S")
        cost = result.get("total_cost", 0)
        n_feeds = len(result.get("results", []))
        item_text = f"#{len(self._history)+1} - {timestamp} - ₺{cost:,.2f} - {n_feeds} yem"
        self._history.append(result.copy())
        self.history_list.addItem(item_text)

    def _load_from_history(self, item):
        idx = self.history_list.row(item)
        if 0 <= idx < len(self._history):
            result = self._history[idx]
            self.update(result, add_to_history=False)
            # Kaydır
            self.tabs.setCurrentIndex(0)

    def _clear_history(self):
        self._history.clear()
        self.history_list.clear()

    def update(self, result: dict, add_to_history=True):
        self._last = result.get("results", [])

        if result["status"] != "Optimal":
            self.c_status.set_value("❌ Çözüm Yok", "Infeasible")
            self.c_status.set_accent(C["red"])
            for c in [self.c_cost, self.c_prot, self.c_energy]:
                c.set_value("—")
            self.table.setRowCount(0)
            self.chart.set_data([])
            self.nutrition_bar.set_data(None)
            self.btn_export.setEnabled(False)
            self.btn_pdf.setEnabled(False)
            return

        if add_to_history:
            self._add_to_history(result)

        # Kartlar
        self.c_cost.set_value(
            f"TL {result['total_cost']:,.3f}",
            "Toplam rasyon maliyeti",
        )
        self.c_prot.set_value(
            f"%{result['achieved_protein']:.2f}",
            "Ham protein (KM)",
            absolute=f"{result['total_protein_kg']:.3f} kg",
        )
        self.c_energy.set_value(
            f"{result['achieved_energy']:.2f}",
            "MJ/kg metabolik enerji",
            absolute=f"{result['total_energy_mj']:.1f} MJ",
        )
        self.c_status.set_value("Optimal", f"{len(self._last)} yem")
        self.c_status.set_accent(C["green"])

        # Besin özeti progress bar
        self.nutrition_bar.set_data(result)

        # Tablo
        rows = self._last
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(rows))

        for r, item in enumerate(rows):
            def itm(txt, align=Qt.AlignCenter, bold=False, color=None):
                i = QTableWidgetItem(txt)
                i.setTextAlignment(align)
                f = QFont("Segoe UI", F_BASE)
                f.setBold(bold)
                i.setFont(f)
                if color:
                    i.setForeground(QColor(color))
                return i

            # Protein ve enerji kütle hesaplamalari
            prot_kg = item['kg'] * item['protein'] / 100
            enerji_mj = item['kg'] * item['energy']

            self.table.setItem(r, 0, itm(item["name"],
                                          Qt.AlignLeft | Qt.AlignVCenter,
                                          bold=True))
            self.table.setItem(r, 1, itm(f"{item['kg']:.3f}",   color=C["t1"]))
            self.table.setItem(r, 2, itm(f"{item['pct']:.2f}",  color=C["blue"], bold=True))
            self.table.setItem(r, 3, itm(f"{item['protein']:.2f}", color=C["purple"]))
            self.table.setItem(r, 4, itm(f"{prot_kg:.3f}",      color=C["purple"], bold=True))
            self.table.setItem(r, 5, itm(f"{item['energy']:.3f}",  color=C["cyan"]))
            self.table.setItem(r, 6, itm(f"{enerji_mj:.3f}",    color=C["cyan"], bold=True))
            self.table.setItem(r, 7, itm(f"{item['cost']:.3f}",    color=C["green"], bold=True))

            kind = "Kaba" if item["is_kaba"] else "Kesif"
            ki = itm(kind, color=C["green"] if item["is_kaba"] else C["t2"])
            if item["is_kaba"]:
                ki.setBackground(QColor(21, 128, 61, 35))
            self.table.setItem(r, 8, ki)

        self.table.setSortingEnabled(True)

        # Grafik verisi
        colors = C["bar_colors"]
        chart_data = [
            {
                "name":  item["name"],
                "pct":   item["pct"],
                "color": colors[i % len(colors)],
            }
            for i, item in enumerate(rows)
        ]
        self.chart.set_data(chart_data)
        self.btn_export.setEnabled(True)
        self.btn_pdf.setEnabled(True)

    def _export(self):
        if not self._last:
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Sonuclarini Kaydet", "rasyon_sonucu.xlsx",
            "Excel Dosyalari (*.xlsx)"
        )
        if not path:
            return
        try:
            # Excel icin yeni veri yapisi - ASCII karakterler
            export_data = []
            for item in self._last:
                prot_kg = item['kg'] * item['protein'] / 100
                enerji_mj = item['kg'] * item['energy']
                export_data.append({
                    "Yem Adi": item['name'],
                    "Miktar_kg": round(item['kg'], 3),
                    "Pay_yuzde": round(item['pct'], 2),
                    "Protein_yuzde": round(item['protein'], 2),
                    "Protein_kg": round(prot_kg, 3),
                    "ME_MJ_kg": round(item['energy'], 3),
                    "Enerji_MJ": round(enerji_mj, 3),
                    "Maliyet_tl": round(item['cost'], 3),
                    "Tur": "Kaba" if item["is_kaba"] else "Kesif"
                })
            df = pd.DataFrame(export_data)
            df.to_excel(path, index=False)
            QMessageBox.information(self, "Disa Aktarma Basarili",
                                    f"Dosya kaydedildi:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Disa Aktarma Hatasi", str(e))

    def _export_pdf(self):
        """PDF formatinda rapor olustur"""
        if not self._last:
            return
        if not HAS_REPORTLAB:
            QMessageBox.warning(self, "Eksik Modul",
                              "PDF raporu icin 'reportlab' paketi kurulu degil.\n"
                              "pip install reportlab")
            return
        
        path, _ = QFileDialog.getSaveFileName(
            self, "PDF Raporu Kaydet", "rasyon_raporu.pdf",
            "PDF Dosyalari (*.pdf)"
        )
        if not path:
            return
        
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.lib import colors
            from reportlab.lib.units import mm
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.enums import TA_CENTER, TA_RIGHT
            
            doc = SimpleDocTemplate(path, pagesize=A4, 
                                    rightMargin=20*mm, leftMargin=20*mm,
                                    topMargin=20*mm, bottomMargin=20*mm)
            
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'Title',
                parent=styles['Title'],
                fontSize=24,
                spaceAfter=20,
                alignment=TA_CENTER,
                textColor=colors.HexColor('#1E3A6E')
            )
            subtitle_style = ParagraphStyle(
                'Subtitle',
                parent=styles['Normal'],
                fontSize=12,
                spaceAfter=30,
                alignment=TA_CENTER,
                textColor=colors.grey
            )
            heading_style = ParagraphStyle(
                'Heading',
                parent=styles['Heading2'],
                fontSize=14,
                spaceBefore=15,
                spaceAfter=10,
                textColor=colors.HexColor('#3B82F6')
            )
            
            elements = []
            
            # Baslik
            elements.append(Paragraph("Rasyon Optimizasyon Raporu", title_style))
            elements.append(Paragraph(f"Olusturulma: {datetime.now().strftime('%d.%m.%Y %H:%M')}", subtitle_style))
            elements.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#3B82F6')))
            elements.append(Spacer(1, 15))
            
            # Ozet bilgiler
            if self._last:
                elements.append(Paragraph("Ozet Bilgiler", heading_style))
                
                total_prot_kg = sum(x['kg'] * x['protein'] / 100 for x in self._last)
                total_en_mj = sum(x['kg'] * x['energy'] for x in self._last)
                
                summary_data = [
                    ["Toplam Maliyet", f"TL {sum(x['cost'] for x in self._last):.3f}"],
                    ["Toplam Agirlik", f"{sum(x['kg'] for x in self._last):.3f} kg"],
                    ["Toplam Protein", f"{total_prot_kg:.3f} kg"],
                    ["Toplam Enerji", f"{total_en_mj:.3f} MJ"],
                    ["Kullanilan Yem Sayisi", str(len(self._last))],
                    ["Kaba Yem Sayisi", str(sum(1 for x in self._last if x['is_kaba']))],
                    ["Kesif Yem Sayisi", str(sum(1 for x in self._last if not x['is_kaba']))],
                ]
                summary_table = Table(summary_data, colWidths=[150, 150])
                summary_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8FAFC')),
                    ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#64748B')),
                    ('TEXTCOLOR', (1, 0), (1, -1), colors.HexColor('#1E293B')),
                    ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                    ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#E2E8F0')),
                    ('TOPPADDING', (0, 0), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                    ('LEFTPADDING', (0, 0), (-1, -1), 12),
                ]))
                elements.append(summary_table)
                elements.append(Spacer(1, 20))
                
                # Detay tablo - yeni sutunlar eklendi
                elements.append(Paragraph("Detayli Rasyon Bilesimi", heading_style))
                
                table_data = [["Yem Adi", "Miktar(kg)", "Pay(%)", "Prot(%)", "Prot(kg)", "ME(MJ/kg)", "Enerji(MJ)", "Maliyet(TL)", "Tur"]]
                for item in self._last:
                    prot_kg = item['kg'] * item['protein'] / 100
                    enerji_mj = item['kg'] * item['energy']
                    table_data.append([
                        item['name'][:25] + ('...' if len(item['name']) > 25 else ''),
                        f"{item['kg']:.3f}",
                        f"{item['pct']:.2f}",
                        f"{item['protein']:.2f}",
                        f"{prot_kg:.3f}",
                        f"{item['energy']:.3f}",
                        f"{enerji_mj:.3f}",
                        f"{item['cost']:.3f}",
                        "Kaba" if item['is_kaba'] else "Kesif"
                    ])
                
                col_widths = [55, 42, 35, 38, 38, 42, 42, 42, 28]
                detail_table = Table(table_data, colWidths=col_widths)
                detail_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3B82F6')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 7),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('ALIGN', (0, 1), (0, -1), 'LEFT'),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')]),
                    ('TOPPADDING', (0, 0), (-1, -1), 6),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                    ('LEFTPADDING', (0, 0), (-1, -1), 4),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ]))
                elements.append(detail_table)
            
            # Footer
            elements.append(Spacer(1, 30))
            elements.append(HRFlowable(width="100%", thickness=1, color=colors.grey))
            footer_text = Paragraph(
                f"<i>Rasyon Optimizasyon Sistemi v4.0 - {datetime.now().strftime('%d.%m.%Y')}</i>",
                ParagraphStyle('Footer', parent=styles['Normal'], fontSize=8, textColor=colors.grey, alignment=TA_CENTER)
            )
            elements.append(footer_text)
            
            doc.build(elements)
            QMessageBox.information(self, "PDF Basarili",
                                    f"PDF raporu olusturuldu:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "PDF Hatasi", str(e))


# ══════════════════════════════════════════════════════════════════════
#  ANA PENCERE - Enhanced with shortcuts & help
# ══════════════════════════════════════════════════════════════════════
class HelpDialog(QDialog):
    """Yardım dialog penceresi"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📖 Kullanım Kılavuzu")
        self.setMinimumSize(600, 500)
        self.setStyleSheet(f"""
            QDialog {{
                background: {C['bg1']};
                border: 1px solid {C['blue']};
                border-radius: 12px;
            }}
        """)
        
        lay = QVBoxLayout2(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(16)
        
        title = QLabel("🥬 Rasyon Optimizasyon Sistemi v4.0 - Kullanım Kılavuzu")
        title.setFont(QFont("Segoe UI", F_LARGE, QFont.Bold))
        title.setStyleSheet(f"color:{C['blue_light']};")
        lay.addWidget(title)
        
        text = QTextEdit()
        text.setReadOnly(True)
        text.setStyleSheet(f"""
            QTextEdit {{
                background: {C['bg2']};
                border: 1px solid {C['brd']};
                border-radius: 8px;
                color: {C['t1']};
                padding: 16px;
                font-size: {F_SMALL}px;
                line-height: 1.6;
            }}
        """)
        text.setHtml("""
        <h3 style="color: #60A5FA;">🎯 Genel Kullanım</h3>
        <ol style="color: #CBD5E1;">
        <li><b>Excel Dosyası Yükle:</b> Sol panelden "📂 XLSX Yükle" butonuna tıklayın.<br>
        Dosyada en az "Yem Adı", "Ham Protein" ve "ME" sütunları olmalıdır.</li>
        <li><b>Yem Seçimi:</b> Tablodaki onay kutularından kullanmak istediğiniz yemleri seçin.</li>
        <li><b>Kısıtları Ayarla:</b> Sağ panelden protein, enerji ve kaba yem kısıtlarını belirleyin.<br>
        Veya hazır ayarlardan (presets) birini seçip hızlıca uygulayın.</li>
        <li><b>Optimize Et:</b> "⚡ OPTİMİZE ET" butonuna tıklayarak en düşük maliyetli rasyonu hesaplayın.</li>
        <li><b>Sonuçları Görüntüle:</b> Sonuçlar tablo, grafik ve besin özeti olarak görüntülenir.</li>
        <li><b>Dışa Aktar:</b> Excel veya PDF formatında rapor oluşturabilirsiniz.</li>
        </ol>
        
        <h3 style="color: #60A5FA;">⌨️ Klavye Kısayolları</h3>
        <table style="color: #CBD5E1; border-collapse: collapse; width: 100%;">
        <tr><td style="padding: 6px;"><b>Ctrl+O</b></td><td>Excel dosyası aç</td></tr>
        <tr><td style="padding: 6px;"><b>Ctrl+Return</b></td><td>Optimizasyonu çalıştır</td></tr>
        <tr><td style="padding: 6px;"><b>Ctrl+E</b></td><td>Excel'e aktar</td></tr>
        <tr><td style="padding: 6px;"><b>Ctrl+P</b></td><td>PDF raporu oluştur</td></tr>
        <tr><td style="padding: 6px;"><b>F1</b></td><td>Bu yardım penceresini aç</td></tr>
        <tr><td style="padding: 6px;"><b>Escape</b></td><td>Sonuçları sıfırla</td></tr>
        </table>
        
        <h3 style="color: #60A5FA;">📋 Hazır Ayarlar (Presets)</h3>
        <ul style="color: #CBD5E1;">
        <li><b>Sığır (Süt):</b> 22 kg/gün, protein %14-17, ME 10.5-12.5 MJ/kg</li>
        <li><b>Sığır (Et):</b> 18 kg/gün, protein %11-14, ME 9.5-11.5 MJ/kg</li>
        <li><b>Koyun (Süt):</b> 4 kg/gün, protein %15-18, ME 10-12 MJ/kg</li>
        <li><b>Keçi (Süt):</b> 3.5 kg/gün, protein %14-17, ME 10-12 MJ/kg</li>
        <li><b>Besi (Dana):</b> 10 kg/gün, protein %12-15, ME 10-11.5 MJ/kg</li>
        </ul>
        
        <h3 style="color: #60A5FA;">📊 Excel Dosyası Formatı</h3>
        <p style="color: #94A3B8;">Sütun isimleri esnek algılanır. Örnek:<br>
        <code style="background: #1E293B; padding: 2px 6px; border-radius: 4px;">
        Yem Adı | Ham Protein | ME | Fiyat<br>
        Mısır | 8.5 | 13.1 | 4.50
        </code></p>
        
        <h3 style="color: #60A5FA;">⚠️ Önemli Notlar</h3>
        <ul style="color: #94A3B8;">
        <li>Kaba yem otomatik olarak yem adından tespit edilir (saman, ot, silaj vb.)</li>
        <li>Üre içeren yemler otomatik olarak %1 KM ile sınırlandırılır</li>
        <li>Kısıt aralıkları çok dar olursa "Infeasible" hatası alınabilir</li>
        </ul>
        """)
        lay.addWidget(text)
        
        close_btn = QPushButton("✕ Kapat")
        close_btn.setFixedHeight(36)
        close_btn.clicked.connect(self.close)
        lay.addWidget(close_btn)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Rasyon Optimizasyon Sistemi v4.0")
        self.setMinimumSize(1400, 880)
        self.resize(1550, 980)
        self.optimizer = RasyonOptimizer()
        self._build()
        self._setup_shortcuts()

    def _setup_shortcuts(self):
        """Klavye kısayollarını ayarla"""
        from PyQt5.QtWidgets import QShortcut
        from PyQt5.QtGui import QKeySequence
        
        # Ctrl+O: Dosya aç
        QShortcut(QKeySequence("Ctrl+O"), self, activated=lambda: self.feed_panel.btn_import.click())
        
        # Ctrl+Return: Optimize et
        QShortcut(QKeySequence("Ctrl+Return"), self, activated=self._run)
        
        # Ctrl+E: Excel'e aktar
        QShortcut(QKeySequence("Ctrl+E"), self, activated=lambda: self.result_panel.btn_export.click() if self.result_panel.btn_export.isEnabled() else None)
        
        # Ctrl+P: PDF raporu
        QShortcut(QKeySequence("Ctrl+P"), self, activated=lambda: self.result_panel.btn_pdf.click() if self.result_panel.btn_pdf.isEnabled() else None)
        
        # F1: Yardım
        QShortcut(QKeySequence("F1"), self, activated=self._show_help)
        
        # Escape: Sıfırla
        QShortcut(QKeySequence("Escape"), self, activated=self._reset)

    def _show_help(self):
        dlg = HelpDialog(self)
        dlg.exec_()

    def _build(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(HeaderWidget())

        body = QWidget()
        bl = QHBoxLayout(body)
        bl.setContentsMargins(18, 18, 18, 18)
        bl.setSpacing(18)

        # ── Sol: Yem tablosu ──
        self.feed_panel = FeedTablePanel()
        self.feed_panel.setMinimumWidth(560)

        # ── Sağ: Kısıtlar + butonlar + sonuçlar ──
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(14)
        right.setMinimumWidth(480)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.constraint_panel = ConstraintPanel()
        scroll.setWidget(self.constraint_panel)
        scroll.setMaximumHeight(480)

        # Buton satırı - Enhanced
        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)

        self.btn_help = QPushButton("📖 Yardım")
        self.btn_help.setFixedHeight(50)
        self.btn_help.setFixedWidth(90)
        self.btn_help.setStyleSheet(f"""
            QPushButton {{
                background: {C['bg3']};
                color: {C['cyan']};
                border: 1px solid {C['cyan']};
                border-radius: 10px;
                font-size: {F_SMALL}px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background: rgba(6, 182, 212, 0.15);
                box-shadow: 0 0 12px rgba(6, 182, 212, 0.3);
            }}
        """)
        self.btn_help.clicked.connect(self._show_help)
        
        self.btn_reset = QPushButton("↺  Sıfırla")
        self.btn_reset.setObjectName("btnReset")
        self.btn_reset.clicked.connect(self._reset)

        self.btn_opt = QPushButton("⚡  OPTİMİZE ET")
        self.btn_opt.setObjectName("btnOptimize")
        self.btn_opt.clicked.connect(self._run)

        btn_row.addWidget(self.btn_help)
        btn_row.addWidget(self.btn_reset)
        btn_row.addWidget(self.btn_opt, stretch=1)

        self.result_panel = ResultPanel()

        rl.addWidget(scroll)
        rl.addLayout(btn_row)
        rl.addWidget(self.result_panel, stretch=1)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.feed_panel)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([840, 580])

        bl.addWidget(splitter)
        root.addWidget(body, stretch=1)

        # Status bar - Enhanced
        self.sb = QStatusBar()
        self.setStatusBar(self.sb)
        
        # Status mesajı ile versiyon göstergesi
        self.sb.showMessage(
            "Hazır  ·  Ctrl+O: Dosya aç  ·  Ctrl+Return: Optimize  ·  F1: Yardım  ·  v4.0 Enhanced Edition"
        )

    def _run(self):
        feeds = self.feed_panel.get_selected()
        if not feeds:
            QMessageBox.warning(self, "Veri Yok",
                                "Optimizasyon için en az bir yem seçili olmalı.\n"
                                "Lütfen önce bir XLSX dosyası yükleyip yem seçin.")
            return

        c = self.constraint_panel.get()

        errs = []
        if c["min_prot"] > c["max_prot"]:
            errs.append("• Min Protein > Max Protein")
        if c["min_en"] > c["max_en"]:
            errs.append("• Min Enerji > Max Enerji")
        if c["kaba_lo"] > c["kaba_hi"]:
            errs.append("• Min Kaba Yem > Max Kaba Yem")
        if errs:
            QMessageBox.warning(self, "Geçersiz Kısıt",
                                "Aşağıdaki kısıtlar geçersiz:\n" + "\n".join(errs))
            return

        self.btn_opt.setEnabled(False)
        self.btn_opt.setText("⏳  Hesaplanıyor…")
        self.sb.showMessage("Optimizasyon çalışıyor, lütfen bekleyin…")
        QApplication.processEvents()

        try:
            res = self.optimizer.optimize(
                feeds=feeds,
                total_kg=c["total_kg"],
                min_prot=c["min_prot"], max_prot=c["max_prot"],
                min_en=c["min_en"],     max_en=c["max_en"],
                kaba_lo=c["kaba_lo"],   kaba_hi=c["kaba_hi"],
            )
        except Exception as e:
            QMessageBox.critical(self, "Optimizasyon Hatası", str(e))
            self.btn_opt.setEnabled(True)
            self.btn_opt.setText("⚡  OPTİMİZE ET")
            return

        self.btn_opt.setEnabled(True)
        self.btn_opt.setText("⚡  OPTİMİZE ET")

        if res["status"] != "Optimal":
            msgs = {
                "Infeasible": (
                    "Verilen kısıtlarla matematiksel çözüm bulunamadı (Infeasible).\n\n"
                    "Olası nedenler:\n"
                    "• Protein / enerji aralığı çok dar\n"
                    "• Kaba yem oranı sağlanamıyor (yeterli kaba yem seçili değil)\n"
                    "• Seçili yemlerin besin değerleri hedef aralıkta değil\n\n"
                    "→ Kısıt aralıklarını genişletin veya daha fazla yem seçin."
                ),
                "Unbounded": "Model sınırsız çözüm üretiyor; kısıtları kontrol edin.",
                "NoFeeds":   "Optimizasyona dahil edilecek yem bulunamadı.",
            }
            QMessageBox.warning(self, "Çözüm Bulunamadı",
                                msgs.get(res["status"], f"Durum: {res['status']}"))
            self.result_panel.update(res)
            self.sb.showMessage(f"Optimizasyon başarısız — {res['status']}")
            return

        self.result_panel.update(res)
        n = len(res["results"])
        self.sb.showMessage(
            f"✔ Optimal çözüm  ·  {n} yem  ·  "
            f"Maliyet: ₺{res['total_cost']:,.3f}  ·  "
            f"Protein: %{res['achieved_protein']:.2f}  ·  "
            f"ME: {res['achieved_energy']:.3f} MJ/kg"
        )

    def _reset(self):
        self.result_panel.update({"status": "Reset"})
        for card in [self.result_panel.c_cost,
                     self.result_panel.c_prot,
                     self.result_panel.c_energy]:
            card.set_value("—")
            card.sub_lbl.setText("")
        self.result_panel.c_status.set_value("Bekliyor", "")
        self.result_panel.c_status.set_accent(C["t2"])
        self.result_panel.table.setRowCount(0)
        self.result_panel.chart.set_data([])
        self.result_panel.btn_export.setEnabled(False)
        self.sb.showMessage("Sonuçlar sıfırlandı.")


# ══════════════════════════════════════════════════════════════════════
#  GİRİŞ NOKTASI
# ══════════════════════════════════════════════════════════════════════
def main():
    # High DPI ayarları QApplication'dan ÖNCE olmalı!
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setApplicationName("Rasyon Optimizasyon Sistemi")
    app.setStyleSheet(STYLE)

    pal = QPalette()
    pal.setColor(QPalette.Window,          QColor(C["bg0"]))
    pal.setColor(QPalette.WindowText,      QColor(C["t1"]))
    pal.setColor(QPalette.Base,            QColor(C["bg3"]))
    pal.setColor(QPalette.AlternateBase,   QColor(C["bg_alt"]))
    pal.setColor(QPalette.Text,            QColor(C["t1"]))
    pal.setColor(QPalette.Button,          QColor(C["bg_hover"]))
    pal.setColor(QPalette.ButtonText,      QColor(C["t1"]))
    pal.setColor(QPalette.Highlight,       QColor(C["bg_sel"]))
    pal.setColor(QPalette.HighlightedText, QColor(C["t1"]))
    app.setPalette(pal)

    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
